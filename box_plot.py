import numpy as np
import pandas as pd
from matplotlib.axes import Axes
from matplotlib.lines import Line2D

def render_box_and_scatter_chart(
    ax: Axes, df: pd.DataFrame, sheet_name: str = "Data"
) -> None:
    # 1. 准备x轴的标签
    x_labels = []
    for col in df.columns:
        if isinstance(col, int):
            x_labels.append(f"Col {col}")
        else:
            x_labels.append(str(col))

    y_data_values = []
    all_data = []

    # 2. 从DataFrame中提取每列数据，并计算全局范围
    for col in df.columns:
        # 使用 pd.to_numeric 过滤非数字内容，再去除NaN值
        numeric_series = pd.to_numeric(df[col], errors="coerce")
        column_data = numeric_series.dropna().tolist()

        y_data_values.append(column_data)
        all_data.extend(column_data)

    if all_data:
        global_min = min(all_data)
        global_max = max(all_data)
    else:
        global_min, global_max = 0.0, 1.0

    data_range = global_max - global_min
    data_padding = data_range if data_range > 0 else 1.0

    # 将数据转换为适合绘制的字典格式
    combined_data = {label: values for label, values in zip(x_labels, y_data_values)}

    # 准备x轴位置 - 箱线图居中，散点图在左，直方图在右
    x_positions = np.arange(len(x_labels)) + 1
    x_positions_scatter = x_positions - 0.3
    x_positions_hist = x_positions + 0.3

    # 3. 绘制散点图
    scatter_color = "gray"
    for i, (label, values) in enumerate(combined_data.items()):
        if not values:
            continue
        # 增加水平抖动
        x_jittered = np.random.normal(x_positions_scatter[i], 0.02, size=len(values))
        ax.scatter(
            x_jittered,
            values,
            color=scatter_color,
            alpha=0.6,
            s=40,
            edgecolors="white",
            linewidth=0.5,
            label="Data Points" if i == 0 else "",
        )

    # 4. 绘制箱线图
    valid_data_for_box = list(combined_data.values())
    if any(valid_data_for_box):
        # 仅传入非空数据，对应其正确的位置
        valid_positions = [x_positions[i] for i, v in enumerate(valid_data_for_box) if v]
        valid_values = [v for v in valid_data_for_box if v]
        
        box = ax.boxplot(
            valid_values,
            positions=valid_positions,
            patch_artist=True,
            widths=0.3,
            showmeans=True,
            meanline=True,
            showfliers=False,
        )

        box_color = "royalblue"
        for patch in box["boxes"]:
            patch.set_facecolor(box_color)
            patch.set_alpha(0.5)
        for whisker in box["whiskers"]:
            whisker.set(color=box_color, linewidth=1.5)
        for cap in box["caps"]:
            cap.set(color=box_color, linewidth=1.5)
        for median in box["medians"]:
            median.set(color="red", linewidth=2)

    # 5. 绘制压缩后的横置直方图
    hist_color = "lightgrey"
    max_hist_height = 0.15

    # 预计算全局统一的bins，以保证每组数据的直方图是在同样的Y轴刻度下进行统计的
    if all_data and global_max > global_min:
        # 在最小值和最大值之间均匀划分50个区间
        global_bins = np.linspace(global_min, global_max, 50)
        global_bin_width = global_bins[1] - global_bins[0]
    else:
        global_bins = 50
        global_bin_width = 0.02

    for i, (label, values) in enumerate(combined_data.items()):
        if not values:
            continue
        
        # 使用全局统一的边缘计算频率
        hist, bin_edges = np.histogram(values, bins=global_bins)

        bin_width = bin_edges[1] - bin_edges[0] if len(bin_edges) > 1 else global_bin_width

        # 归一化限制最高高度，保证在横向不会遮挡或重叠其他图表
        if len(hist) > 0 and max(hist) > 0:
            hist = hist / max(hist) * max_hist_height

        ax.barh(
            bin_edges[:-1],
            hist,
            height=bin_width,
            left=x_positions_hist[i],
            color=hist_color,
            alpha=0.6,
            edgecolor="white",
            linewidth=0.5,
            label="Histogram" if i == 0 else "",
        )

    # 6. 计算和添加统计标签
    for i, (label, values) in enumerate(combined_data.items()):
        if not values:
            continue
        min_val = np.min(values)
        max_val = np.max(values)
        mean_val = np.mean(values)
        std_val = np.std(values)

        stats_text = (
            f"Min: {min_val:.3f}\n"
            f"Max: {max_val:.3f}\n"
            f"Avg: {mean_val:.3f}\n"
            f"Stdev: {std_val:.3f}"
        )

        # 动态放置文字标识于全局最大值上方，防止与数据重叠或互相挤压
        text_y_pos = global_max + 0.10 * data_padding
        
        ax.text(
            x_positions[i],
            text_y_pos,
            stats_text,
            fontsize=9,
            bbox=dict(facecolor="white", alpha=0.8, edgecolor="gray"),
            va="bottom",
            ha="center",
        )

    # 7. 设置Axes细节
    ax.set_xticks(x_positions)
    ax.set_xticklabels(x_labels, rotation=-45, ha="left", fontsize=10)
    ax.xaxis.grid(True, linestyle="-", which="major", color="lightgrey", alpha=0.5)
    ax.xaxis.set_ticks_position('bottom')

    ax.set_ylabel("Data Value", fontsize=12)
    
    # 根据全局范围动态计算 y极限
    # 顶部预留0.4至0.5额外的padding以容纳统计文本框
    ax.set_ylim(global_min - 0.1 * data_padding, global_max + 0.45 * data_padding)
    
    ax.yaxis.grid(True, linestyle="-", which="major", color="lightgrey", alpha=0.5)
    ax.set_title(f"Data Analysis_{sheet_name}", fontsize=14)

    # 8. 图例配置
    legend_elements = [
        Line2D(
            [0],
            [0],
            marker="o",
            color="w",
            label="Data Points",
            markerfacecolor="gray",
            markersize=10,
        ),
        Line2D([0], [0], color="royalblue", lw=4, label="Boxplot"),
        Line2D([0], [0], color="lightgrey", lw=4, label="Histogram"),
    ]
    ax.legend(handles=legend_elements, loc='lower right')

    # 设置图表边框
    for spine in ax.spines.values():
        spine.set_color("black")
        spine.set_linewidth(1)
