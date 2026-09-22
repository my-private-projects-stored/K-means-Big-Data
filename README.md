# Tìm hiểu và triển khai thuật toán K-Means trên dữ liệu lớn

> **Ứng dụng:** Phân cụm các trạng thái biến động nến tiền điện tử (Crypto)  
> **Môn học:** Big Data  
> **Framework:** Python, Apache Spark (PySpark), Spark MLlib

---

## Mô tả đề tài

Đề tài sử dụng thuật toán K-Means để phân cụm các trạng thái biến động giá tiền điện tử dựa trên dữ liệu nến 1-phút (OHLCV). Mỗi bản ghi nến 1 phút của một đồng tiền là một đối tượng phân cụm, được biểu diễn bằng 3 đặc trưng: tỷ suất sinh lời (return), biên độ nến (range), và logarit khối lượng giao dịch (log volume).

**Quy mô dữ liệu:** 20 đồng coin × 12 tháng × 525.600 phút ≈ **10.512.000 bản ghi**

---

## Cấu trúc dự án

```
├── src/
│   ├── download_data.py    # Script tải dữ liệu từ Binance Archive
│   ├── main.py             # Pipeline chính: tiền xử lý → K-Means → benchmark
│   └── plot_results.py     # Vẽ 3 biểu đồ từ kết quả
├── data/
│   ├── raw/                # File zip/csv gốc
│   └── processed/          # Parquet đã tiền xử lý (10.5M dòng)
├── results/                # Kết quả: CSV + biểu đồ PNG
├── clustered_output/       # Output Parquet: nhãn cụm cho toàn bộ dữ liệu
├── requirements.txt
└── README.md
```

---

## Yêu cầu hệ thống

- **Python:** 3.11+
- **Java:** 8+ (cần cho PySpark)
- **RAM:** tối thiểu 6 GB khả dụng
- **Ổ cứng:** ~500 MB trống

---

## Hướng dẫn cài đặt

### 1. Cài đặt dependencies

```bash
pip install -r requirements.txt
```

Packages cần thiết: `pyspark`, `pandas`, `pyarrow`, `matplotlib`, `requests`

### 2. Cài đặt Hadoop winutils (Windows)

PySpark trên Windows cần `winutils.exe` để ghi file Parquet:

```bash
# Tải winutils.exe vào C:\hadoop\bin\
# Đặt biến môi trường: HADOOP_HOME=C:\hadoop
```

### 3. Kiểm tra Java

```bash
java -version
# Cần Java 8 trở lên
```

---

## Hướng dẫn chạy

### Bước 1: Tải dữ liệu

```bash
python src/download_data.py
```

- Tải 240 file zip từ Binance Archive (20 coin × 12 tháng)
- Giải nén → gộp thành 1 file Parquet (~200 MB)
- Chạy validation tự động: kiểm tra đủ coin, đủ tháng, ≥10M dòng
- Thời gian: ~5 phút

### Bước 2: Chạy pipeline K-Means

```bash
python src/main.py
```

- Đọc Parquet, tiền xử lý, tính 3 đặc trưng, chuẩn hóa
- Benchmark: đo thời gian K-Means trên 100K / 1M / toàn bộ 10.5M dòng
- Elbow Method + Silhouette Score (K=2..6)
- Chạy mô hình cuối trên toàn bộ dataset
- Thời gian: ~8-10 phút

### Bước 3: Vẽ biểu đồ

```bash
python src/plot_results.py
```

- Tạo 3 biểu đồ PNG trong thư mục `results/`
- Thời gian: ~10 giây

---

## Nguồn dữ liệu

- **Nguồn:** Binance Data Archive — `https://data.binance.vision/`
- **Loại:** OHLCV 1-minute klines
- **Thời gian:** Năm 2025 (01/2025 – 12/2025)
- **20 đồng coin:** BTCUSDT, ETHUSDT, BNBUSDT, SOLUSDT, XRPUSDT, DOGEUSDT, ADAUSDT, TRXUSDT, LINKUSDT, AVAXUSDT, DOTUSDT, LTCUSDT, UNIUSDT, ATOMUSDT, NEARUSDT, ETCUSDT, XLMUSDT, VETUSDT, ICPUSDT, ALGOUSDT

---

## Kết quả thực nghiệm

### Benchmark thời gian K-Means (K=3, maxIter=20, PySpark Local Mode)

| Mức dữ liệu | Số bản ghi | Thời gian (giây) |
|---|---|---|
| Nhỏ | 100.000 | 7.02 |
| Trung bình | 1.000.000 | 48.60 |
| Lớn (toàn bộ) | 10.512.000 | 83.67 |

### Chọn K tối ưu (Elbow + Silhouette)

| K | WCSS | Silhouette |
|---|---|---|
| 2 | 680,886.83 | 0.4313 |
| 3 | 517,013.78 | 0.5084 |
| **4** | **399,164.56** | **0.5649** |
| 5 | 348,897.24 | 0.5608 |
| 6 | 305,648.52 | 0.5624 |

**K=4 được chọn** dựa trên Silhouette Score cao nhất (0.5649) kết hợp với điểm khuỷu tay trên biểu đồ Elbow.
