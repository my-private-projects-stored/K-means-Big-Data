"""pipeline phân cụm dữ liệu tiền điện tử bằng thuật toán K-Means trên PySpark."""

import os
import sys
import time
import csv
from pathlib import Path

from pyspark.sql import SparkSession
from pyspark.sql.functions import col, log as spark_log
from pyspark.ml.feature import VectorAssembler, StandardScaler
from pyspark.ml.clustering import KMeans
from pyspark.ml.evaluation import ClusteringEvaluator

# thiết lập biến môi trường Hadoop để hỗ trợ ghi file Parquet trên Windows
os.environ["HADOOP_HOME"] = r"C:\hadoop"
os.environ["PATH"] = os.environ["PATH"] + r";C:\hadoop\bin"

# cấu hình đường dẫn dữ liệu và thư mục kết quả
BASE_DIR = Path(__file__).resolve().parent.parent
PARQUET_INPUT = BASE_DIR / "data" / "processed" / "crypto_20coins_1m.parquet"
RESULTS_DIR = BASE_DIR / "results"
OUTPUT_DIR = BASE_DIR / "clustered_output"

RESULTS_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# các mốc quy mô dữ liệu cho benchmark hiệu năng
BENCHMARK_SIZES = {
    "Nho_100K": 100_000,
    "TrungBinh_1M": 1_000_000,
    "Lon_ToanBo": None,  # None tương đương toàn bộ tập dữ liệu
}

# phạm vi số cụm K cần khảo sát
K_RANGE = range(2, 7)

# quy mô mẫu đại diện phục vụ phương pháp Elbow và hệ số Silhouette
ELBOW_SAMPLE_SIZE = 1_000_000
SILHOUETTE_SAMPLE_FRACTION = 0.1  # 10% của tập đại diện


def create_spark_session():
    """khởi tạo SparkSession với cấu hình tối ưu bộ nhớ cho máy cục bộ."""
    return SparkSession.builder \
        .appName("KMeans_Crypto_BigData") \
        .config("spark.driver.memory", "6g") \
        .config("spark.sql.shuffle.partitions", "8") \
        .getOrCreate()


def load_and_preprocess(spark):
    """đọc dữ liệu Parquet, làm sạch, trích xuất đặc trưng và chuẩn hóa StandardScaler."""
    print("=" * 60)
    print("DOC DU LIEU VA TIEN XU LY")
    print("=" * 60)

    # đọc dữ liệu từ file Parquet
    df_raw = spark.read.parquet(str(PARQUET_INPUT))
    raw_count = df_raw.count()
    print(f"So dong doc duoc: {raw_count:,}")

    # lọc các dòng nến OHLCV hợp lệ
    df_clean = df_raw.filter(
        (col("open") > 0) &
        (col("high") >= col("low")) &
        (col("volume") >= 0)
    ).dropna()
    clean_count = df_clean.count()
    print(f"So dong sau loc: {clean_count:,} (loai {raw_count - clean_count:,} dong)")

    # trích xuất 3 đặc trưng tài chính
    df_features = df_clean \
        .withColumn("return_pct", (col("close") - col("open")) / col("open")) \
        .withColumn("range_pct", (col("high") - col("low")) / col("open")) \
        .withColumn("log_volume", spark_log(col("volume") + 1.0))

    # gom các đặc trưng thành vector
    assembler = VectorAssembler(
        inputCols=["return_pct", "range_pct", "log_volume"],
        outputCol="raw_features"
    )
    assembled_df = assembler.transform(df_features)

    # chuẩn hóa đặc trưng về trung bình 0 và độ lệch chuẩn 1
    scaler = StandardScaler(
        inputCol="raw_features", outputCol="features",
        withStd=True, withMean=True
    )
    scaler_model = scaler.fit(assembled_df)
    final_df = scaler_model.transform(assembled_df).select(
        "symbol", "open_time",
        "return_pct", "range_pct", "log_volume",
        "features"
    ).cache()

    # kích hoạt lưu vào bộ nhớ RAM
    total_rows = final_df.count()
    print(f"So feature vectors: {total_rows:,}")
    print(f"Features: return_pct, range_pct, log_volume (da chuan hoa)")
    print()

    return final_df, total_rows


def run_benchmark(final_df, total_rows):
    """đo lường thời gian huấn luyện K-Means trên 3 quy mô dữ liệu."""
    print("=" * 60)
    print("BENCHMARK: DO THOI GIAN K-MEANS TREN 3 MUC DU LIEU")
    print("=" * 60)

    benchmark_results = []

    for level_name, num_rows in BENCHMARK_SIZES.items():
        if num_rows is None:
            num_rows = total_rows
            sample_df = final_df
        else:
            sample_df = final_df.limit(num_rows)

        # lưu tạm vào bộ nhớ và đếm số dòng
        sample_df = sample_df.cache()
        actual_count = sample_df.count()

        # đo thời gian huấn luyện mô hình K-Means
        start_time = time.time()
        kmeans = KMeans(k=4, seed=42, maxIter=20, featuresCol="features")
        model = kmeans.fit(sample_df)
        elapsed = time.time() - start_time

        wcss = model.summary.trainingCost

        benchmark_results.append({
            "level": level_name,
            "num_rows": actual_count,
            "kmeans_time_sec": round(elapsed, 2),
            "wcss": round(wcss, 2),
        })

        print(f"  [{level_name}] {actual_count:,} dong | "
              f"Thoi gian: {elapsed:.2f}s | WCSS: {wcss:.2f}")

        # giải phóng bộ nhớ nếu là tập con
        if num_rows != total_rows:
            sample_df.unpersist()

    # lưu kết quả benchmark ra file CSV
    csv_path = RESULTS_DIR / "benchmark.csv"
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["level", "num_rows", "kmeans_time_sec", "wcss"])
        writer.writeheader()
        writer.writerows(benchmark_results)

    print(f"\nDa luu: {csv_path}")
    print()

    return benchmark_results


def run_elbow_silhouette(final_df):
    """thực hiện phương pháp Elbow và hệ số Silhouette để chọn số cụm K tối ưu."""
    print("=" * 60)
    print("GIAI DOAN A: ELBOW + SILHOUETTE (K=2..6)")
    print("=" * 60)

    # lấy tập đại diện 1 triệu dòng cho Elbow
    elbow_sample = final_df.limit(ELBOW_SAMPLE_SIZE).cache()
    elbow_count = elbow_sample.count()
    print(f"Tap dai dien cho Elbow: {elbow_count:,} dong")

    # lấy mẫu 10% phục vụ tính điểm Silhouette
    sil_sample = elbow_sample.sample(fraction=SILHOUETTE_SAMPLE_FRACTION, seed=42).cache()
    sil_count = sil_sample.count()
    print(f"Mau cho Silhouette: {sil_count:,} dong (10% cua tap dai dien)")
    print()

    evaluator = ClusteringEvaluator(featuresCol="features")

    wcss_results = []
    sil_results = []

    for k in K_RANGE:
        # huấn luyện mô hình thử nghiệm với số cụm k
        kmeans = KMeans(k=k, seed=42, maxIter=20, featuresCol="features")
        model = kmeans.fit(elbow_sample)

        # tính WCSS
        wcss = model.summary.trainingCost
        wcss_results.append({"k": k, "wcss": round(wcss, 2)})

        # tính hệ số Silhouette
        preds = model.transform(sil_sample)
        sil_score = evaluator.evaluate(preds)
        sil_results.append({"k": k, "silhouette": round(sil_score, 4)})

        print(f"  K={k} | WCSS={wcss:,.2f} | Silhouette={sil_score:.4f}")

    # lưu kết quả WCSS
    wcss_path = RESULTS_DIR / "wcss.csv"
    with open(wcss_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["k", "wcss"])
        writer.writeheader()
        writer.writerows(wcss_results)

    # lưu kết quả Silhouette
    sil_path = RESULTS_DIR / "silhouette.csv"
    with open(sil_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["k", "silhouette"])
        writer.writeheader()
        writer.writerows(sil_results)

    print(f"\nDa luu: {wcss_path}")
    print(f"Da luu: {sil_path}")

    # in bảng tổng hợp các chỉ số theo K
    print()
    print("=" * 60)
    print("BANG TONG HOP - XEM DE CHON best_k")
    print("=" * 60)
    print(f"{'K':>3} | {'WCSS':>15} | {'Silhouette':>12}")
    print("-" * 40)
    for w, s in zip(wcss_results, sil_results):
        print(f"{w['k']:>3} | {w['wcss']:>15,.2f} | {s['silhouette']:>12.4f}")
    print()
    print("Huong dan chon K:")
    print("  - Elbow: chon K tai diem WCSS giam cham lai (khuyru tay)")
    print("  - Silhouette: K co diem cao nhat (nhung can doi voi y nghia)")
    print("  - Khong tu dong chon - can phan tich ca 2 chi so")
    print()

    elbow_sample.unpersist()
    sil_sample.unpersist()

    return wcss_results, sil_results


def run_final_model(final_df, best_k, total_rows):
    """huấn luyện mô hình K-Means cuối cùng trên toàn bộ dữ liệu và lưu kết quả."""
    print("=" * 60)
    print(f"GIAI DOAN B: MO HINH CUOI CUNG (K={best_k}) TREN TOAN BO DATASET")
    print("=" * 60)

    start_time = time.time()
    final_kmeans = KMeans(k=best_k, seed=42, maxIter=20, featuresCol="features")
    final_model = final_kmeans.fit(final_df)
    fit_time = time.time() - start_time

    predictions = final_model.transform(final_df)

    # tính toán và lưu tọa độ tâm cụm
    print(f"\nTAM CUM (CENTROIDS), K={best_k}:")
    centers = final_model.clusterCenters()
    centroid_data = []
    for i, center in enumerate(centers):
        print(f"  Cum {i}: Return={center[0]:.4f}, Range={center[1]:.4f}, "
              f"LogVolume={center[2]:.4f}")
        centroid_data.append({
            "cluster": i,
            "return_pct_scaled": round(center[0], 4),
            "range_pct_scaled": round(center[1], 4),
            "log_volume_scaled": round(center[2], 4),
        })

    centroid_path = RESULTS_DIR / "centroids.csv"
    with open(centroid_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=[
            "cluster", "return_pct_scaled", "range_pct_scaled", "log_volume_scaled"
        ])
        writer.writeheader()
        writer.writerows(centroid_data)

    # thống kê số lượng điểm dữ liệu theo từng cụm
    print(f"\nPHAN BO SO LUONG DIEM MOI CUM (TOAN BO {total_rows:,} DONG):")
    cluster_counts = predictions.groupBy("prediction").count().orderBy("prediction")
    cluster_counts.show()

    cluster_size_data = cluster_counts.collect()
    size_path = RESULTS_DIR / "cluster_sizes.csv"
    with open(size_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["cluster", "count"])
        writer.writeheader()
        for row in cluster_size_data:
            writer.writerow({"cluster": row["prediction"], "count": row["count"]})

    # thống kê trung bình các đặc trưng gốc theo cụm
    print("TRUNG BINH FEATURE (CHUA CHUAN HOA) THEO CUM:")
    predictions.groupBy("prediction").agg(
        {"return_pct": "avg", "range_pct": "avg", "log_volume": "avg"}
    ).orderBy("prediction").show()

    # thống kê phân bố các coin theo cụm
    print("PHAN BO COIN THEO TUNG CUM (TOP 50):")
    predictions.groupBy("prediction", "symbol").count() \
        .orderBy("prediction", "symbol").show(50)

    # lưu kết quả gán nhãn cụm ra định dạng Parquet
    output_path = str(OUTPUT_DIR)
    predictions.write.mode("overwrite").parquet(output_path)
    print(f"Da luu ket qua phan cum: {output_path}/")
    print(f"Cac cot: symbol, open_time, return_pct, range_pct, "
          f"log_volume, features, prediction")

    # lưu các thông số đánh giá của mô hình cuối cùng
    model_info_path = RESULTS_DIR / "final_model_info.csv"
    with open(model_info_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=[
            "best_k", "total_rows", "fit_time_sec", "wcss"
        ])
        writer.writeheader()
        writer.writerow({
            "best_k": best_k,
            "total_rows": total_rows,
            "fit_time_sec": round(fit_time, 2),
            "wcss": round(final_model.summary.trainingCost, 2),
        })

    print(f"\nDa luu: {centroid_path}")
    print(f"Da luu: {size_path}")
    print(f"Da luu: {model_info_path}")
    print(f"Thoi gian train mo hinh cuoi: {fit_time:.2f} giay")
    print()


def main():
    """hàm thực thi toàn bộ pipeline từ tiền xử lý, benchmark đến phân cụm cuối cùng."""
    overall_start = time.time()

    print("=" * 60)
    print("  K-MEANS CRYPTO BIG DATA - PIPELINE CHINH")
    print("  Du lieu: 20 coin x 12 thang ~ 10.5 trieu dong")
    print("=" * 60)
    print()

    # kiểm tra sự tồn tại của file dữ liệu đầu vào
    if not PARQUET_INPUT.exists():
        print(f"LOI: Khong tim thay file du lieu: {PARQUET_INPUT}")
        print("Hay chay download_data.py truoc.")
        sys.exit(1)

    # khởi tạo phiên làm việc Spark
    spark = create_spark_session()
    spark.sparkContext.setLogLevel("WARN")

    try:
        # bước 1: đọc và tiền xử lý dữ liệu
        final_df, total_rows = load_and_preprocess(spark)

        # bước 2: đo benchmark thời gian thực thi
        run_benchmark(final_df, total_rows)

        # bước 3: tìm số cụm K tối ưu bằng Elbow và Silhouette
        wcss_results, sil_results = run_elbow_silhouette(final_df)

        # bước 4: chọn K tối ưu và huấn luyện mô hình trên toàn bộ dữ liệu
        best_k = max(sil_results, key=lambda x: x["silhouette"])["k"]
        print(f"GIA TRI best_k GOI Y TU SILHOUETTE CAO NHAT: K={best_k}")
        print(f"(Can kiem tra lai voi bieu do Elbow truoc khi ket luan)")
        print()

        run_final_model(final_df, best_k, total_rows)

    finally:
        spark.stop()

    elapsed = time.time() - overall_start
    print("=" * 60)
    print(f"HOAN THANH TOAN BO PIPELINE")
    print(f"Tong thoi gian: {elapsed:.1f} giay ({elapsed/60:.1f} phut)")
    print("=" * 60)
    print()
    print("Buoc tiep theo: chay plot_results.py de ve bieu do")


if __name__ == "__main__":
    main()
