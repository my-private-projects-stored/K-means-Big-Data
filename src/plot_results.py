"""vẽ các biểu đồ kết quả đánh giá mô hình K-Means và benchmark hiệu năng."""

import csv
import sys
import io
from pathlib import Path

# cấu hình bảng mã UTF-8 cho console Windows
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

# cấu hình đường dẫn kết quả và font chữ
BASE_DIR = Path(__file__).resolve().parent.parent
RESULTS_DIR = BASE_DIR / "results"

# thiết lập font chữ hỗ trợ tiếng Việt
plt.rcParams['font.family'] = 'DejaVu Sans'


def read_csv(filename):
    """đọc nội dung file CSV thành danh sách dictionary."""
    path = RESULTS_DIR / filename
    with open(path, "r", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def plot_elbow():
    """vẽ biểu đồ phương pháp Elbow xác định số cụm K tối ưu."""
    data = read_csv("wcss.csv")
    ks = [int(row["k"]) for row in data]
    wcss = [float(row["wcss"]) for row in data]

    fig, ax = plt.subplots(figsize=(8, 5))
    ax.plot(ks, wcss, "bo-", linewidth=2, markersize=8)

    # đánh dấu điểm "khuỷu tay" (K có Silhouette cao nhất)
    sil_data = read_csv("silhouette.csv")
    best_k = max(sil_data, key=lambda x: float(x["silhouette"]))
    best_k_val = int(best_k["k"])
    best_k_idx = ks.index(best_k_val)
    ax.plot(best_k_val, wcss[best_k_idx], "r*", markersize=20,
            label=f"K tối ưu = {best_k_val}")

    ax.set_xlabel("Số cụm (K)", fontsize=12)
    ax.set_ylabel("WCSS (Tổng bình phương khoảng cách trong cụm)", fontsize=12)
    ax.set_title("Phương pháp Elbow — Chọn K tối ưu", fontsize=14, fontweight="bold")
    ax.set_xticks(ks)
    ax.legend(fontsize=11)
    ax.grid(True, alpha=0.3)

    ax.ticklabel_format(style="plain", axis="y")

    out_path = RESULTS_DIR / "elbow_chart.png"
    fig.tight_layout()
    fig.savefig(out_path, dpi=150)
    plt.close(fig)
    print(f"Da luu: {out_path}")


def plot_silhouette():
    """vẽ biểu đồ điểm Silhouette theo từng giá trị K."""
    data = read_csv("silhouette.csv")
    ks = [int(row["k"]) for row in data]
    scores = [float(row["silhouette"]) for row in data]

    fig, ax = plt.subplots(figsize=(8, 5))
    bars = ax.bar(ks, scores, color="#4ECDC4", edgecolor="#2C3E50", linewidth=1.2)

    # làm nổi bật cột có điểm Silhouette cao nhất
    best_idx = scores.index(max(scores))
    bars[best_idx].set_color("#E74C3C")
    bars[best_idx].set_edgecolor("#C0392B")

    for bar, score in zip(bars, scores):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.005,
                f"{score:.4f}", ha="center", va="bottom", fontsize=10, fontweight="bold")

    ax.set_xlabel("Số cụm (K)", fontsize=12)
    ax.set_ylabel("Điểm Silhouette", fontsize=12)
    ax.set_title("Điểm Silhouette theo K (mẫu 10% của 1 triệu dòng)", fontsize=14, fontweight="bold")
    ax.set_xticks(ks)
    ax.set_ylim(0, max(scores) * 1.15)
    ax.grid(True, alpha=0.3, axis="y")

    out_path = RESULTS_DIR / "silhouette_chart.png"
    fig.tight_layout()
    fig.savefig(out_path, dpi=150)
    plt.close(fig)
    print(f"Da luu: {out_path}")


def plot_benchmark():
    """vẽ biểu đồ so sánh thời gian huấn luyện theo quy mô dữ liệu."""
    data = read_csv("benchmark.csv")
    labels = [row["level"] for row in data]
    num_rows = [int(row["num_rows"]) for row in data]
    times = [float(row["kmeans_time_sec"]) for row in data]

    # nhãn tiếng Việt cho các mức dữ liệu
    vn_labels = {
        "Nho_100K": "Nhỏ",
        "TrungBinh_1M": "Trung bình",
        "Lon_ToanBo": "Lớn (toàn bộ)",
    }

    fig, ax = plt.subplots(figsize=(9, 5))

    colors = ["#3498DB", "#F39C12", "#E74C3C"]
    bars = ax.bar(range(len(labels)), times, color=colors, edgecolor="#2C3E50",
                  linewidth=1.2, width=0.6)

    for bar, t, n in zip(bars, times, num_rows):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 1,
                f"{t:.2f} giây", ha="center", va="bottom", fontsize=12, fontweight="bold")

    x_labels = [f"{vn_labels.get(label, label)}\n({n:,} dòng)" for label, n in zip(labels, num_rows)]
    ax.set_xticks(range(len(labels)))
    ax.set_xticklabels(x_labels, fontsize=10)

    ax.set_ylabel("Thời gian huấn luyện K-Means (giây)", fontsize=12)
    ax.set_title("Benchmark: Thời gian thực thi K-Means theo quy mô dữ liệu\n(PySpark Local Mode, K=4)",
                 fontsize=13, fontweight="bold")
    ax.set_ylim(0, max(times) * 1.25)
    ax.grid(True, alpha=0.3, axis="y")

    out_path = RESULTS_DIR / "benchmark_chart.png"
    fig.tight_layout()
    fig.savefig(out_path, dpi=150)
    plt.close(fig)
    print(f"Da luu: {out_path}")


def main():
    """vẽ và lưu toàn bộ các biểu đồ đánh giá mô hình."""
    print("=" * 60)
    print("  VE 3 BIEU DO TU KET QUA K-MEANS PIPELINE")
    print("=" * 60)
    print()

    plot_elbow()
    plot_silhouette()
    plot_benchmark()

    print()
    print("Hoan thanh! 3 bieu do da luu trong:", RESULTS_DIR)


if __name__ == "__main__":
    main()
