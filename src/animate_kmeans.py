"""tạo ảnh động mô phỏng từng bước lặp và quá trình hội tụ của thuật toán K-Means."""

import numpy as np
import pandas as pd
import sys
import io
from pathlib import Path

# cấu hình bảng mã UTF-8 cho console Windows
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
plt.rcParams['font.family'] = 'DejaVu Sans'
from matplotlib.gridspec import GridSpec
from PIL import Image

# cấu hình số cụm và tham số thuật toán
BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "clustered_output"
VIZ_DIR = BASE_DIR / "results" / "visualizations"
VIZ_DIR.mkdir(parents=True, exist_ok=True)

K = 4
MAX_ITER = 15
SAMPLE_SIZE = 5000
SEED = 42

CLUSTER_COLORS = np.array([
    [0.204, 0.596, 0.859],  # xanh duong
    [0.906, 0.298, 0.235],  # do
    [0.180, 0.800, 0.443],  # xanh la
    [0.953, 0.612, 0.071],  # cam
])

FEATURE_NAMES = ["return_pct", "range_pct", "log_volume"]


def load_data():
    """đọc dữ liệu, chuẩn hóa đặc trưng và lấy mẫu phục vụ vẽ biểu đồ."""
    print("Dang doc du lieu...")
    df = pd.read_parquet(OUTPUT_DIR, engine="pyarrow",
                         columns=["return_pct", "range_pct", "log_volume"])

    X_raw = df[FEATURE_NAMES].values.astype(np.float64)
    n_total = len(X_raw)
    print(f"Tong so dong: {n_total:,}")

    # chuẩn hóa đặc trưng tương đương StandardScaler
    mean = X_raw.mean(axis=0)
    std = X_raw.std(axis=0)
    X = (X_raw - mean) / std
    print(f"Da chuan hoa: mean={mean.round(4)}, std={std.round(4)}")

    # lấy mẫu ngẫu nhiên cho biểu đồ phân tán
    np.random.seed(SEED)
    sample_idx = np.random.choice(n_total, size=SAMPLE_SIZE, replace=False)
    X_sample = X[sample_idx]
    print(f"Mau cho scatter plot: {SAMPLE_SIZE:,} diem")

    return X, X_sample, n_total


def manual_kmeans_step(X, centroids):
    """thực hiện một vòng lặp K-Means: gán cụm, cập nhật trọng tâm và tính WCSS."""
    # tính khoảng cách từ mỗi điểm đến từng tâm cụm
    dists = np.zeros((len(X), K), dtype=np.float64)
    for k in range(K):
        diff = X - centroids[k]
        dists[:, k] = np.sum(diff ** 2, axis=1)

    labels = np.argmin(dists, axis=1)

    # cập nhật tọa độ tâm cụm mới
    centroids_new = np.zeros_like(centroids)
    for k in range(K):
        mask = labels == k
        if mask.sum() > 0:
            centroids_new[k] = X[mask].mean(axis=0)
        else:
            centroids_new[k] = centroids[k]

    # tính tổng bình phương khoảng cách trong cụm (WCSS)
    wcss = sum(dists[np.arange(len(X)), labels])

    return labels, centroids_new, wcss


def run_kmeans_with_history(X, X_sample):
    """chạy lặp K-Means và lưu lại trạng thái từng vòng lặp."""
    print(f"\nChay K-Means thu cong (K={K}, {MAX_ITER} iterations)...")
    print(f"Du lieu: {len(X):,} diem (full), scatter: {len(X_sample):,} diem")

    # khởi tạo ngẫu nhiên K tâm cụm ban đầu
    np.random.seed(SEED)
    init_idx = np.random.choice(len(X), size=K, replace=False)
    centroids = X[init_idx].copy()

    history = []

    # lưu trạng thái khởi tạo ban đầu
    sample_dists = np.zeros((len(X_sample), K))
    for k in range(K):
        sample_dists[:, k] = np.sum((X_sample - centroids[k]) ** 2, axis=1)
    sample_labels = np.argmin(sample_dists, axis=1)

    history.append({
        "iter": 0,
        "centroids": centroids.copy(),
        "sample_labels": sample_labels.copy(),
        "wcss": None,
        "cluster_sizes": [int((sample_labels == k).sum()) for k in range(K)],
    })
    print(f"  Iter 0 (init): centroids initialized")

    # lặp qua từng vòng chạy
    for i in range(1, MAX_ITER + 1):
        labels_full, centroids, wcss = manual_kmeans_step(X, centroids)

        # gán nhãn cho tập mẫu theo tâm cụm mới
        sample_dists = np.zeros((len(X_sample), K))
        for k in range(K):
            sample_dists[:, k] = np.sum((X_sample - centroids[k]) ** 2, axis=1)
        sample_labels = np.argmin(sample_dists, axis=1)

        # thống kê kích thước cụm trên toàn bộ dữ liệu
        cluster_sizes = [int((labels_full == k).sum()) for k in range(K)]

        history.append({
            "iter": i,
            "centroids": centroids.copy(),
            "sample_labels": sample_labels.copy(),
            "wcss": wcss,
            "cluster_sizes": cluster_sizes,
        })

        # kiểm tra điều kiện hội tụ sớm
        if i > 1:
            prev_wcss = history[-2]["wcss"]
            if prev_wcss and abs(prev_wcss - wcss) / prev_wcss < 1e-6:
                print(f"  Hoi tu tai iteration {i}!")
                break

    return history


def create_animation_frames(X_sample, history):
    """vẽ các biểu đồ phân tán và bảng thông số cho từng vòng lặp."""
    print(f"\nDang tao {len(history)} frames...")

    frames_dir = VIZ_DIR / "frames"
    frames_dir.mkdir(exist_ok=True)

    plot_configs = [
        (0, 2, "Tỷ suất sinh lời (chuẩn hóa)", "Log khối lượng (chuẩn hóa)"),
        (0, 1, "Tỷ suất sinh lời (chuẩn hóa)", "Biên độ nến (chuẩn hóa)"),
    ]

    frame_paths = []

    for step in history:
        it = step["iter"]
        centroids = step["centroids"]
        sample_labels = step["sample_labels"]
        wcss = step["wcss"]
        cluster_sizes = step["cluster_sizes"]

        fig = plt.figure(figsize=(16, 7))
        gs = GridSpec(1, 3, width_ratios=[4, 4, 2.5], figure=fig)

        # vẽ 2 biểu đồ phân tán 2D
        for idx, (fx, fy, xlabel, ylabel) in enumerate(plot_configs):
            ax = fig.add_subplot(gs[0, idx])

            colors = CLUSTER_COLORS[sample_labels]
            ax.scatter(X_sample[:, fx], X_sample[:, fy],
                       c=colors, alpha=0.4, s=8, edgecolors="none")

            for k in range(K):
                ax.scatter(centroids[k, fx], centroids[k, fy],
                           c=[CLUSTER_COLORS[k]], marker="X", s=200,
                           edgecolors="black", linewidths=2, zorder=10)

            ax.set_xlabel(xlabel, fontsize=10)
            ax.set_ylabel(ylabel, fontsize=10)
            ax.set_xlim(-4, 6)
            if fy == 2:
                ax.set_ylim(-3, 4)
            else:
                ax.set_ylim(-1, 8)
            ax.grid(True, alpha=0.2)

            if idx == 0:
                ax.set_title("Sinh lời vs Log khối lượng", fontsize=11, fontweight="bold")
            else:
                ax.set_title("Sinh lời vs Biên độ nến", fontsize=11, fontweight="bold")

        # bảng hiển thị thông số chi tiết của vòng lặp
        ax_info = fig.add_subplot(gs[0, 2])
        ax_info.axis("off")

        info_text = f"Vòng lặp: {it}\n"
        info_text += f"{'=' * 25}\n\n"

        if wcss is not None:
            info_text += f"WCSS: {wcss:,.0f}\n\n"
        else:
            info_text += "WCSS: (đang khởi tạo)\n\n"

        info_text += "Kích thước cụm\n"
        info_text += f"(toàn bộ {10_512_000:,} dòng):\n"
        info_text += f"{'-' * 25}\n"

        cluster_labels = ["C0 Biến động TB", "C1 Yên tĩnh", "C2 Ngoại lai", "C3 Biến động mạnh"]
        for k in range(K):
            size = cluster_sizes[k]
            pct = size / sum(cluster_sizes) * 100 if sum(cluster_sizes) > 0 else 0
            info_text += f"{cluster_labels[k]}:\n  {size:>10,} ({pct:.1f}%)\n"

        ax_info.text(0.05, 0.95, info_text, transform=ax_info.transAxes,
                     fontsize=10, verticalalignment="top", fontfamily="monospace",
                     bbox=dict(boxstyle="round,pad=0.5", facecolor="lightyellow",
                                edgecolor="gray", alpha=0.9))

        # biểu đồ nhỏ thể hiện sự hội tụ của WCSS
        if it > 0:
            wcss_history = [h["wcss"] for h in history[:it+1] if h["wcss"] is not None]
            if len(wcss_history) > 1:
                ax_wcss = fig.add_axes([0.73, 0.08, 0.22, 0.2])
                ax_wcss.plot(range(1, len(wcss_history) + 1), wcss_history,
                             "b-o", markersize=4, linewidth=1.5)
                ax_wcss.set_xlabel("Vòng lặp", fontsize=8)
                ax_wcss.set_ylabel("WCSS", fontsize=8)
                ax_wcss.set_title("Quá trình hội tụ", fontsize=9, fontweight="bold")
                ax_wcss.tick_params(labelsize=7)
                ax_wcss.ticklabel_format(style="scientific", axis="y", scilimits=(0, 0))
                ax_wcss.grid(True, alpha=0.3)

        # tiêu đề chính của khung hình
        if it == 0:
            fig.suptitle("Thuật toán K-Means: Khởi tạo",
                         fontsize=14, fontweight="bold", y=0.98)
        else:
            fig.suptitle(f"Thuật toán K-Means: Vòng lặp {it}",
                         fontsize=14, fontweight="bold", y=0.98)

        fig.tight_layout(rect=[0, 0, 1, 0.95])

        frame_path = frames_dir / f"frame_{it:03d}.png"
        fig.savefig(frame_path, dpi=120, bbox_inches="tight",
                    facecolor="white", edgecolor="none")
        plt.close(fig)
        frame_paths.append(frame_path)

        print(f"  Frame {it} saved")

    return frame_paths


def compile_gif(frame_paths):
    """ghép các khung hình riêng lẻ thành ảnh động GIF hoàn chỉnh."""
    print("\nDang tao GIF animation...")

    images = [Image.open(p) for p in frame_paths]

    # thiết lập thời gian hiển thị cho từng khung hình
    durations = []
    for i in range(len(images)):
        if i == 0:
            durations.append(2000)  # 2 giay cho init
        elif i == len(images) - 1:
            durations.append(4000)  # 4 giay cho frame cuoi
        else:
            durations.append(800)   # 0.8 giay moi iteration

    gif_path = VIZ_DIR / "kmeans_animation.gif"
    images[0].save(
        gif_path,
        save_all=True,
        append_images=images[1:],
        duration=durations,
        loop=0,  # loop vo han
    )

    file_size = gif_path.stat().st_size / (1024 * 1024)
    print(f"Da luu: {gif_path} ({file_size:.1f} MB)")
    return gif_path


def main():
    """điều phối quy trình mô phỏng và tạo ảnh động K-Means."""
    print("=" * 60)
    print("  ANIMATION QUÁ TRÌNH K-MEANS")
    print("  Toàn bộ dữ liệu: 10.5M dòng | Mẫu scatter: 5000 điểm")
    print("=" * 60)
    print()

    X, X_sample, n_total = load_data()
    history = run_kmeans_with_history(X, X_sample)
    frame_paths = create_animation_frames(X_sample, history)
    gif_path = compile_gif(frame_paths)

    print()
    print("=" * 60)
    print("HOAN THANH!")
    print(f"  GIF: {gif_path}")
    print(f"  Frames: {len(frame_paths)} frames")
    print("=" * 60)


if __name__ == "__main__":
    main()
