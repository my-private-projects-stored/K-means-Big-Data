"""thu thập và tiền xử lý dữ liệu nến 1 phút từ Binance Archive thành file Parquet."""

import os
import io
import sys
import time
import zipfile
import logging
import requests
import pandas as pd
from pathlib import Path

# danh sách 20 cặp coin có thanh khoản lớn
SYMBOLS = [
    "BTCUSDT", "ETHUSDT", "BNBUSDT", "SOLUSDT", "XRPUSDT",
    "DOGEUSDT", "ADAUSDT", "TRXUSDT", "LINKUSDT", "AVAXUSDT",
    "DOTUSDT", "LTCUSDT", "UNIUSDT", "ATOMUSDT", "NEARUSDT",
    "ETCUSDT", "XLMUSDT", "VETUSDT", "ICPUSDT", "ALGOUSDT",
]

# khoảng thời gian: 1 năm đầy đủ
YEAR = 2025
MONTHS = list(range(1, 13))  # 01 → 12

# đường dẫn thư mục dữ liệu
BASE_DIR = Path(__file__).resolve().parent.parent
RAW_DIR = BASE_DIR / "data" / "raw"
PROCESSED_DIR = BASE_DIR / "data" / "processed"
PARQUET_OUTPUT = PROCESSED_DIR / "crypto_20coins_1m.parquet"

# mẫu URL tải dữ liệu từ Binance Vision
URL_TEMPLATE = (
    "https://data.binance.vision/data/spot/monthly/klines/"
    "{symbol}/1m/{symbol}-1m-{year}-{month:02d}.zip"
)

# tên các cột nến OHLCV của Binance
COLUMN_NAMES = [
    "open_time", "open", "high", "low", "close", "volume",
    "close_time", "quote_asset_volume", "number_of_trades",
    "taker_buy_base_asset_volume", "taker_buy_quote_asset_volume", "ignore",
]

# các cột cần giữ lại cho phân cụm K-Means
KEEP_COLUMNS = ["symbol", "open_time", "open", "high", "low", "close", "volume"]

# cấu hình số lần thử lại khi tải thất bại
MAX_RETRIES = 3
RETRY_DELAY_SECONDS = 5

# cấu hình ghi log ra màn hình và file
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(BASE_DIR / "data" / "download_log.txt", encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)


def download_zip(url: str, retries: int = MAX_RETRIES) -> bytes | None:
    """tải file zip từ URL với cơ chế thử lại khi gặp lỗi mạng."""
    for attempt in range(1, retries + 1):
        try:
            resp = requests.get(url, timeout=60)
            if resp.status_code == 200:
                return resp.content
            elif resp.status_code == 404:
                log.warning(f"  404 Not Found: {url}")
                return None
            else:
                log.warning(f"  HTTP {resp.status_code} (lần {attempt}/{retries}): {url}")
        except requests.RequestException as e:
            log.warning(f"  Lỗi mạng (lần {attempt}/{retries}): {e}")
        if attempt < retries:
            time.sleep(RETRY_DELAY_SECONDS)
    log.error(f"  THẤT BẠI sau {retries} lần thử: {url}")
    return None


def extract_csv_from_zip(zip_bytes: bytes, symbol: str) -> pd.DataFrame | None:
    """giải nén file zip trong bộ nhớ và đọc CSV thành DataFrame."""
    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            csv_files = [f for f in zf.namelist() if f.endswith(".csv")]
            if not csv_files:
                log.warning(f"  Không tìm thấy file CSV trong zip của {symbol}")
                return None
            with zf.open(csv_files[0]) as csv_file:
                df = pd.read_csv(csv_file, header=None, names=COLUMN_NAMES)
                df["symbol"] = symbol
                return df[KEEP_COLUMNS]
    except (zipfile.BadZipFile, Exception) as e:
        log.error(f"  Lỗi giải nén zip {symbol}: {e}")
        return None


# bước 1: kiểm tra tính khả dụng của từng symbol

def test_symbols() -> list[str]:
    """tải thử tháng đầu tiên để kiểm tra tính khả dụng của từng symbol."""
    log.info("=" * 60)
    log.info("BƯỚC 1: KIỂM TRA 20 SYMBOL TRÊN BINANCE ARCHIVE")
    log.info("=" * 60)

    valid_symbols = []
    failed_symbols = []

    for symbol in SYMBOLS:
        url = URL_TEMPLATE.format(symbol=symbol, year=YEAR, month=1)
        log.info(f"  Testing {symbol}...")
        data = download_zip(url, retries=2)
        if data is not None:
            # thử giải nén để chắc chắn file không bị lỗi
            df = extract_csv_from_zip(data, symbol)
            if df is not None and len(df) > 0:
                log.info(f"  ✓ {symbol} — OK ({len(df):,} dòng tháng 01/{YEAR})")
                valid_symbols.append(symbol)
            else:
                log.warning(f"  ✗ {symbol} — file zip tải được nhưng CSV rỗng/lỗi")
                failed_symbols.append(symbol)
        else:
            log.warning(f"  ✗ {symbol} — không tải được")
            failed_symbols.append(symbol)

    log.info(f"\nKết quả kiểm tra: {len(valid_symbols)}/20 symbol hợp lệ")
    if failed_symbols:
        log.warning(f"Symbol thất bại: {', '.join(failed_symbols)}")
        log.warning("Các symbol này sẽ bị bỏ qua khi tải toàn bộ.")

    if len(valid_symbols) < 10:
        log.error("CHỈ CÓ DƯỚI 10 SYMBOL HỢP LỆ — dữ liệu sẽ quá ít. Dừng lại.")
        sys.exit(1)

    return valid_symbols


# bước 2: tải toàn bộ các file zip và gộp thành DataFrame

def download_all(valid_symbols: list[str]) -> pd.DataFrame:
    """tải toàn bộ dữ liệu 12 tháng của các symbol và gộp thành một DataFrame."""
    total_files = len(valid_symbols) * len(MONTHS)
    log.info("=" * 60)
    log.info(f"BƯỚC 2: TẢI TOÀN BỘ {total_files} FILE ZIP")
    log.info("=" * 60)

    all_dfs = []
    success_count = 0
    fail_count = 0
    download_log = []  # lưu log chi tiết cho từng file

    for i, symbol in enumerate(valid_symbols, 1):
        for month in MONTHS:
            url = URL_TEMPLATE.format(symbol=symbol, year=YEAR, month=month)
            label = f"[{success_count + fail_count + 1}/{total_files}] {symbol} {YEAR}-{month:02d}"

            zip_bytes = download_zip(url)
            if zip_bytes is None:
                log.warning(f"  {label} — THẤT BẠI")
                fail_count += 1
                download_log.append({
                    "symbol": symbol, "year": YEAR, "month": month,
                    "status": "FAILED", "rows": 0
                })
                continue

            df = extract_csv_from_zip(zip_bytes, symbol)
            if df is None or len(df) == 0:
                log.warning(f"  {label} — CSV rỗng/lỗi")
                fail_count += 1
                download_log.append({
                    "symbol": symbol, "year": YEAR, "month": month,
                    "status": "EMPTY", "rows": 0
                })
                continue

            all_dfs.append(df)
            success_count += 1
            rows = len(df)
            download_log.append({
                "symbol": symbol, "year": YEAR, "month": month,
                "status": "OK", "rows": rows
            })
            log.info(f"  {label} — OK ({rows:,} dòng)")

        log.info(f"  --- Hoàn thành {symbol} ({i}/{len(valid_symbols)} coin) ---")

    log.info(f"\nTổng kết download: {success_count} thành công, {fail_count} thất bại")

    # lưu thông tin chi tiết quá trình tải
    log_df = pd.DataFrame(download_log)
    log_path = BASE_DIR / "data" / "download_detail_log.csv"
    log_df.to_csv(log_path, index=False)
    log.info(f"Chi tiết download đã lưu: {log_path}")

    if not all_dfs:
        log.error("KHÔNG CÓ DỮ LIỆU NÀO ĐƯỢC TẢI THÀNH CÔNG. Dừng lại.")
        sys.exit(1)

    log.info("Đang gộp tất cả DataFrame...")
    combined_df = pd.concat(all_dfs, ignore_index=True)
    log.info(f"Tổng số dòng sau gộp: {len(combined_df):,}")

    return combined_df


# bước 3: lưu dữ liệu vào file Parquet

def save_parquet(df: pd.DataFrame) -> None:
    """chuyển đổi kiểu dữ liệu và lưu DataFrame sang định dạng Parquet."""
    log.info("=" * 60)
    log.info("BƯỚC 3: LƯU FILE PARQUET")
    log.info("=" * 60)

    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)

    # ép kiểu dữ liệu số phù hợp
    df["open_time"] = pd.to_numeric(df["open_time"], errors="coerce")
    for col in ["open", "high", "low", "close", "volume"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df.to_parquet(PARQUET_OUTPUT, index=False, engine="pyarrow")
    file_size_mb = PARQUET_OUTPUT.stat().st_size / (1024 * 1024)
    log.info(f"Đã lưu: {PARQUET_OUTPUT}")
    log.info(f"Dung lượng: {file_size_mb:.1f} MB")


# bước 4: kiểm tra tính hợp lệ và toàn vẹn của dữ liệu

def validate_data() -> None:
    """đọc lại file Parquet và kiểm tra các tiêu chí chất lượng dữ liệu."""
    log.info("=" * 60)
    log.info("BƯỚC 4: VALIDATION DỮ LIỆU")
    log.info("=" * 60)

    df = pd.read_parquet(PARQUET_OUTPUT, engine="pyarrow")
    total_rows = len(df)
    issues = []

    # kiểm tra số lượng coin
    unique_coins = sorted(df["symbol"].unique())
    num_coins = len(unique_coins)
    log.info(f"[1/6] Số coin: {num_coins} — {', '.join(unique_coins)}")
    if num_coins < 10:
        issues.append(f"CHỈ CÓ {num_coins} COIN — quá ít để benchmark có ý nghĩa")
    elif num_coins < 20:
        log.warning(f"  ⚠ Có {num_coins}/20 coin (thiếu {20 - num_coins} coin)")

    # kiểm tra phân bố tháng cho từng coin
    log.info("[2/6] Kiểm tra phân bố tháng cho mỗi coin:")
    # chuyển open_time (milliseconds) thành tháng
    df["_month"] = pd.to_datetime(df["open_time"], unit="ms").dt.month
    for coin in unique_coins:
        coin_months = sorted(df[df["symbol"] == coin]["_month"].unique())
        missing = [m for m in MONTHS if m not in coin_months]
        coin_rows = len(df[df["symbol"] == coin])
        if missing:
            log.warning(f"  ⚠ {coin}: {coin_rows:,} dòng, thiếu tháng {missing}")
        else:
            log.info(f"  ✓ {coin}: {coin_rows:,} dòng, đủ 12 tháng")
    df.drop(columns=["_month"], inplace=True)

    # kiểm tra các cột dữ liệu bắt buộc
    required_cols = {"symbol", "open_time", "open", "high", "low", "close", "volume"}
    actual_cols = set(df.columns)
    log.info(f"[3/6] Cột hiện có: {list(df.columns)}")
    if not required_cols.issubset(actual_cols):
        missing_cols = required_cols - actual_cols
        issues.append(f"THIẾU CỘT: {missing_cols}")

    # kiểm tra tổng số dòng dữ liệu
    log.info(f"[4/6] Tổng số dòng: {total_rows:,}")
    if total_rows < 10_000_000:
        log.warning(f"  ⚠ Chưa đạt 10M (thiếu {10_000_000 - total_rows:,} dòng)")
        if total_rows < 1_000_000:
            issues.append(f"CHỈ CÓ {total_rows:,} DÒNG — không đủ yêu cầu tối thiểu 1M")
    else:
        log.info(f"  ✓ Đạt mục tiêu ≥10M dòng")

    # kiểm tra tính hợp lệ của dữ liệu OHLCV
    log.info("[5/6] Kiểm tra tính hợp lệ dữ liệu:")
    null_count = df[["open", "high", "low", "close", "volume"]].isnull().sum().sum()
    open_zero = (df["open"] <= 0).sum()
    high_lt_low = (df["high"] < df["low"]).sum()
    vol_negative = (df["volume"] < 0).sum()

    log.info(f"  Null/NaN trong OHLCV: {null_count:,}")
    log.info(f"  open <= 0: {open_zero:,}")
    log.info(f"  high < low: {high_lt_low:,}")
    log.info(f"  volume < 0: {vol_negative:,}")

    if null_count > 0:
        log.warning(f"  ⚠ Có {null_count:,} giá trị null — sẽ bị loại ở bước tiền xử lý")
    if open_zero > 0:
        log.warning(f"  ⚠ Có {open_zero:,} dòng open<=0 — sẽ bị loại (tránh chia cho 0)")
    if high_lt_low > 0:
        log.warning(f"  ⚠ Có {high_lt_low:,} dòng high<low — sẽ bị loại")

    # tổng hợp kết quả kiểm tra
    log.info("=" * 60)
    log.info("[6/6] BÁO CÁO TỔNG HỢP")
    log.info("=" * 60)
    log.info(f"  Số coin:         {num_coins}")
    log.info(f"  Tổng số dòng:    {total_rows:,}")
    log.info(f"  File Parquet:     {PARQUET_OUTPUT}")
    log.info(f"  Dung lượng:      {PARQUET_OUTPUT.stat().st_size / (1024*1024):.1f} MB")

    if issues:
        log.error("=" * 60)
        log.error("CÓ LỖI NGHIÊM TRỌNG:")
        for issue in issues:
            log.error(f"  ✗ {issue}")
        log.error("Cần xử lý trước khi chạy main.py")
    else:
        log.info("")
        log.info("  ✅ VALIDATION THÀNH CÔNG — Sẵn sàng chạy main.py")
        log.info("")


# hàm thực thi chính

def main():
    """thực hiện tuần tự quy trình kiểm tra, tải, lưu và kiểm định dữ liệu."""
    start_time = time.time()

    log.info("╔══════════════════════════════════════════════════════════╗")
    log.info("║  DOWNLOAD DATA — K-Means Crypto Big Data               ║")
    log.info("║  20 coin × 12 tháng từ Binance Archive                  ║")
    log.info("╚══════════════════════════════════════════════════════════╝")
    log.info(f"Năm dữ liệu: {YEAR}")
    log.info(f"Số coin mục tiêu: {len(SYMBOLS)}")
    log.info(f"Tổng file zip dự kiến: {len(SYMBOLS) * len(MONTHS)}")
    log.info("")

    # khởi tạo thư mục lưu trữ
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)

    # bước 1: kiểm tra symbol
    valid_symbols = test_symbols()

    # bước 2: tải toàn bộ dữ liệu
    combined_df = download_all(valid_symbols)

    # bước 3: lưu thành Parquet
    save_parquet(combined_df)

    # giải phóng bộ nhớ RAM
    del combined_df

    # bước 4: kiểm định kết quả
    validate_data()

    elapsed = time.time() - start_time
    log.info(f"\nTổng thời gian: {elapsed:.1f} giây ({elapsed/60:.1f} phút)")


if __name__ == "__main__":
    main()
