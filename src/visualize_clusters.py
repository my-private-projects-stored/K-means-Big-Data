"""trực quan hóa chi tiết kết quả phân cụm K-Means thành các biểu đồ chuyên sâu."""

import csv
import sys
import io
import numpy as np
from pathlib import Path

# cấu hình bảng mã UTF-8 cho console Windows
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
from matplotlib.gridspec import GridSpec

# cấu hình thư mục lưu trữ và đường dẫn kết quả
BASE_DIR = Path(__file__).resolve().parent.parent
RESULTS_DIR = BASE_DIR / "results"
OUTPUT_DIR = BASE_DIR / "clustered_output"
VIZ_DIR = BASE_DIR / "results" / "visualizations"
VIZ_DIR.mkdir(parents=True, exist_ok=True)

# thiết lập font chữ hỗ trợ tiếng Việt
plt.rcParams['font.family'] = 'DejaVu Sans'

# thiết lập màu sắc và tên định danh các cụm
CLUSTER_COLORS = ["#3498DB", "#E74C3C", "#2ECC71", "#F39C12", "#9B59B6", "#1ABC9C"]
CLUSTER_NAMES = {
    0: "Cụm 0\n(Biến động TB)",
    1: "Cụm 1\n(Thị trường yên tĩnh)",
    2: "Cụm 2\n(Ngoại lai cực đoan)",
    3: "Cụm 3\n(Biến động mạnh)",
}

SAMPLE_SIZE = 50_000


def read_csv_file(filename):
    """đọc file CSV kết quả thành danh sách dictionary."""
    path = RESULTS_DIR / filename
    with open(path, "r", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def load_sample_data():
    """đọc dữ liệu phân cụm từ Parquet và trích xuất mẫu phân tầng."""
    import pandas as pd

    print("Đang đọc dữ liệu phân cụm từ Parquet...")
    df = pd.read_parquet(OUTPUT_DIR, engine="pyarrow",
                         columns=["symbol", "return_pct", "range_pct",
                                  "log_volume", "prediction"])

    total = len(df)
    print(f"Tổng số dòng: {total:,}")

    if total > SAMPLE_SIZE:
        samples = []
        for cluster_id in df["prediction"].unique():
            cluster_data = df[df["prediction"] == cluster_id]
            n_sample = max(10, int(SAMPLE_SIZE * len(cluster_data) / total))
            n_sample = min(n_sample, len(cluster_data))
            samples.append(cluster_data.sample(n=n_sample, random_state=42))
        sample = pd.concat(samples, ignore_index=True)
        print(f"Lấy mẫu: {len(sample):,} dòng (giữ tỷ lệ cụm)")
    else:
        sample = df

    return df, sample


# biểu đồ 1 & 2: phân tán 2D theo các cặp đặc trưng

def plot_scatter_2d(sample):
    """vẽ hai biểu đồ phân tán giữa tỷ suất sinh lời với biên độ nến và khối lượng."""
    fig, axes = plt.subplots(1, 2, figsize=(16, 7))

    feature_pairs = [
        ("return_pct", "range_pct",
         "Tỷ suất sinh lời (%) vs Biên độ nến (%)", (-0.05, 0.05), (-0.001, 0.02)),
        ("return_pct", "log_volume",
         "Tỷ suất sinh lời (%) vs Log khối lượng", (-0.05, 0.05), None),
    ]

    for ax, (fx, fy, title, xlim, ylim) in zip(axes, feature_pairs):
        for cluster_id in sorted(sample["prediction"].unique()):
            cluster_data = sample[sample["prediction"] == cluster_id]
            label = CLUSTER_NAMES.get(cluster_id, f"Cụm {cluster_id}")
            label_short = label.replace("\n", " ")
            ax.scatter(
                cluster_data[fx], cluster_data[fy],
                c=CLUSTER_COLORS[cluster_id % len(CLUSTER_COLORS)],
                alpha=0.3, s=5, label=f"{label_short} ({len(cluster_data):,})"
            )

        xlabel_map = {"return_pct": "Tỷ suất sinh lời", "range_pct": "Biên độ nến",
                      "log_volume": "Log khối lượng"}
        ax.set_xlabel(xlabel_map.get(fx, fx), fontsize=11)
        ax.set_ylabel(xlabel_map.get(fy, fy), fontsize=11)
        ax.set_title(title, fontsize=12, fontweight="bold")
        ax.legend(fontsize=8, markerscale=3, loc="upper right")
        ax.grid(True, alpha=0.2)

        if xlim:
            ax.set_xlim(xlim)
        if ylim:
            ax.set_ylim(ylim)

    fig.suptitle("Phân cụm K-Means: Biểu đồ phân tán 2D (dữ liệu mẫu)",
                 fontsize=15, fontweight="bold", y=1.02)
    fig.tight_layout()

    out_path = VIZ_DIR / "scatter_2d_clusters.png"
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"Đã lưu: {out_path}")


# biểu đồ 3: bản đồ nhiệt coin theo cụm

def plot_coin_cluster_heatmap(df_full):
    """vẽ bản đồ nhiệt thể hiện tỷ lệ phần trăm của mỗi coin trong từng cụm."""
    ct = df_full.groupby(["symbol", "prediction"]).size().unstack(fill_value=0)
    ct_pct = ct.div(ct.sum(axis=1), axis=0) * 100

    fig, ax = plt.subplots(figsize=(10, 12))

    data = ct_pct.values
    coins = ct_pct.index.tolist()
    clusters = [f"Cụm {c}" for c in ct_pct.columns.tolist()]

    cmap = plt.cm.YlOrRd
    im = ax.imshow(data, cmap=cmap, aspect="auto", vmin=0, vmax=100)

    for i in range(len(coins)):
        for j in range(len(clusters)):
            val = data[i, j]
            color = "white" if val > 60 else "black"
            ax.text(j, i, f"{val:.1f}%", ha="center", va="center",
                    fontsize=9, fontweight="bold", color=color)

    ax.set_xticks(range(len(clusters)))
    ax.set_xticklabels(clusters, fontsize=11, fontweight="bold")
    ax.set_yticks(range(len(coins)))
    ax.set_yticklabels(coins, fontsize=10)
    ax.set_title("Phân bố Coin theo từng Cụm (%)",
                 fontsize=14, fontweight="bold", pad=15)

    cbar = fig.colorbar(im, ax=ax, shrink=0.6, pad=0.02)
    cbar.set_label("Tỷ lệ (%)", fontsize=11)

    fig.tight_layout()
    out_path = VIZ_DIR / "coin_cluster_heatmap.png"
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"Đã lưu: {out_path}")


# biểu đồ 4: phân bố đặc trưng theo cụm dạng hộp

def plot_feature_boxplots(sample):
    """vẽ biểu đồ hộp phân bố 3 đặc trưng tài chính cho từng cụm."""
    features = ["return_pct", "range_pct", "log_volume"]
    titles = ["Tỷ suất sinh lời (%)", "Biên độ nến (%)", "Log khối lượng"]

    fig, axes = plt.subplots(1, 3, figsize=(18, 6))

    for ax, feat, title in zip(axes, features, titles):
        clusters = sorted(sample["prediction"].unique())
        data_by_cluster = []
        labels = []

        for c in clusters:
            cluster_data = sample[sample["prediction"] == c][feat].dropna()
            q01 = cluster_data.quantile(0.01)
            q99 = cluster_data.quantile(0.99)
            clipped = cluster_data[(cluster_data >= q01) & (cluster_data <= q99)]
            data_by_cluster.append(clipped.values)
            labels.append(f"C{c}")

        bp = ax.boxplot(data_by_cluster, tick_labels=labels, patch_artist=True,
                        showfliers=False, widths=0.6)

        for patch, color in zip(bp["boxes"], CLUSTER_COLORS[:len(clusters)]):
            patch.set_facecolor(color)
            patch.set_alpha(0.7)

        ax.set_title(title, fontsize=13, fontweight="bold")
        ax.set_xlabel("Cụm", fontsize=11)
        ax.grid(True, alpha=0.2, axis="y")

    fig.suptitle("Phân bố đặc trưng theo Cụm (phần vị 1% – 99%)",
                 fontsize=15, fontweight="bold", y=1.02)
    fig.tight_layout()

    out_path = VIZ_DIR / "feature_boxplots.png"
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"Đã lưu: {out_path}")


# biểu đồ 5: tỷ lệ kích thước các cụm

def plot_cluster_donut():
    """vẽ biểu đồ vành khăn thể hiện tỷ lệ phần trăm kích thước các cụm."""
    data = read_csv_file("cluster_sizes.csv")
    clusters = [int(row["cluster"]) for row in data]
    counts = [int(row["count"]) for row in data]
    total = sum(counts)

    fig, ax = plt.subplots(figsize=(8, 8))

    colors = [CLUSTER_COLORS[c % len(CLUSTER_COLORS)] for c in clusters]

    labels = []
    for c, count in zip(clusters, counts):
        pct = count / total * 100
        name = CLUSTER_NAMES.get(c, f"Cụm {c}").replace("\n", " ")
        labels.append(f"{name}\n{count:,} ({pct:.1f}%)")

    wedges, texts, autotexts = ax.pie(
        counts, labels=labels, colors=colors, autopct="",
        startangle=90, pctdistance=0.85,
        wedgeprops=dict(width=0.4, edgecolor="white", linewidth=2)
    )

    for text in texts:
        text.set_fontsize(10)
        text.set_fontweight("bold")

    centre_circle = plt.Circle((0, 0), 0.55, fc="white")
    ax.add_artist(centre_circle)

    ax.text(0, 0, f"Tổng cộng\n{total:,}", ha="center", va="center",
            fontsize=14, fontweight="bold")

    ax.set_title("Phân bố kích thước các Cụm (K=4)",
                 fontsize=14, fontweight="bold", pad=20)

    out_path = VIZ_DIR / "cluster_donut.png"
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"Đã lưu: {out_path}")


# biểu đồ 6: quá trình chọn K kết hợp Elbow và Silhouette

def plot_k_selection_combined():
    """vẽ kết hợp đường cong Elbow và điểm Silhouette trên biểu đồ hai trục."""
    wcss_data = read_csv_file("wcss.csv")
    sil_data = read_csv_file("silhouette.csv")

    ks = [int(row["k"]) for row in wcss_data]
    wcss = [float(row["wcss"]) for row in wcss_data]
    sil = [float(row["silhouette"]) for row in sil_data]

    fig, ax1 = plt.subplots(figsize=(10, 6))
    ax2 = ax1.twinx()

    line1 = ax1.plot(ks, wcss, "b-o", linewidth=2, markersize=8, label="WCSS (Elbow)")
    ax1.set_xlabel("Số cụm (K)", fontsize=12)
    ax1.set_ylabel("WCSS", fontsize=12, color="blue")
    ax1.tick_params(axis="y", labelcolor="blue")
    ax1.ticklabel_format(style="plain", axis="y")

    line2 = ax2.plot(ks, sil, "r-s", linewidth=2, markersize=8, label="Điểm Silhouette")
    ax2.set_ylabel("Điểm Silhouette", fontsize=12, color="red")
    ax2.tick_params(axis="y", labelcolor="red")

    best_k = 4
    ax1.axvline(x=best_k, color="green", linestyle="--", alpha=0.7, linewidth=2)

    lines = line1 + line2
    labels = [l.get_label() for l in lines]
    labels.append(f"K tối ưu = {best_k}")
    ax1.legend(lines + [ax1.axvline(x=best_k, color="green", linestyle="--", alpha=0)],
               labels, loc="center right", fontsize=11)

    ax1.set_xticks(ks)
    ax1.set_title("Quá trình chọn K: Phương pháp Elbow + Điểm Silhouette",
                  fontsize=14, fontweight="bold")
    ax1.grid(True, alpha=0.2)

    fig.tight_layout()
    out_path = VIZ_DIR / "k_selection_combined.png"
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"Đã lưu: {out_path}")


# biểu đồ 7: radar so sánh tâm các cụm

def plot_centroid_radar():
    """vẽ biểu đồ radar so sánh các đặc trưng đã chuẩn hóa tại tâm cụm."""
    data = read_csv_file("centroids.csv")

    categories = ["Tỷ suất sinh lời", "Biên độ nến", "Log khối lượng"]
    N = len(categories)

    fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(polar=True))

    angles = np.linspace(0, 2 * np.pi, N, endpoint=False).tolist()
    angles += angles[:1]

    for row in data:
        cluster_id = int(row["cluster"])
        values = [
            float(row["return_pct_scaled"]),
            float(row["range_pct_scaled"]),
            float(row["log_volume_scaled"]),
        ]
        values += values[:1]

        color = CLUSTER_COLORS[cluster_id % len(CLUSTER_COLORS)]
        name = CLUSTER_NAMES.get(cluster_id, f"Cụm {cluster_id}").replace("\n", " ")

        ax.plot(angles, values, "o-", linewidth=2, color=color, label=name)
        ax.fill(angles, values, alpha=0.1, color=color)

    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(categories, fontsize=12, fontweight="bold")
    ax.set_title("Tâm cụm (Đặc trưng đã chuẩn hóa)",
                 fontsize=14, fontweight="bold", pad=30)
    ax.legend(loc="upper right", bbox_to_anchor=(1.3, 1.1), fontsize=10)
    ax.grid(True, alpha=0.3)

    out_path = VIZ_DIR / "centroid_radar.png"
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"Đã lưu: {out_path}")


# hàm thực thi chính

def main():
    """đọc dữ liệu mẫu và vẽ toàn bộ 7 biểu đồ trực quan hóa."""
    print("=" * 60)
    print("  TRỰC QUAN HÓA KẾT QUẢ PHÂN CỤM K-MEANS")
    print("=" * 60)
    print()

    df_full, sample = load_sample_data()

    print("\n--- Biểu đồ 1-2: Phân tán 2D ---")
    plot_scatter_2d(sample)

    print("\n--- Biểu đồ 3: Bản đồ nhiệt Coin × Cụm ---")
    plot_coin_cluster_heatmap(df_full)

    print("\n--- Biểu đồ 4: Biểu đồ hộp đặc trưng ---")
    plot_feature_boxplots(sample)

    print("\n--- Biểu đồ 5: Biểu đồ tròn kích thước cụm ---")
    plot_cluster_donut()

    print("\n--- Biểu đồ 6: Quá trình chọn K ---")
    plot_k_selection_combined()

    print("\n--- Biểu đồ 7: Radar tâm cụm ---")
    plot_centroid_radar()

    print()
    print("=" * 60)
    print(f"HOÀN THÀNH! 7 biểu đồ đã lưu trong: {VIZ_DIR}")
    print("=" * 60)

    del df_full, sample


if __name__ == "__main__":
    main()
