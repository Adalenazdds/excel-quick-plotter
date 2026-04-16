from __future__ import annotations

import numpy as np
import pandas as pd
from matplotlib.figure import Figure
from matplotlib.lines import Line2D
import matplotlib.cm as cm


def _is_numeric(val) -> bool:
    """辅助函数：判断单元格是否为纯数字"""
    if pd.isna(val) or val is None or str(val).strip() == "":
        return False
    try:
        float(val)
        return True
    except (ValueError, TypeError):
        return False


def render_line_chart(fig: Figure, df: pd.DataFrame, sheet_name: str = "Data") -> None:
    """绘制折线图：支持自动表头检测、Gap连接与断点强化显示"""

    fig.clear()
    ax = fig.add_subplot(111)

    if df is None or df.shape[0] == 0 or df.shape[1] == 0:
        ax.text(0.5, 0.5, "Empty data", ha="center", va="center", fontsize=12)
        return

    # --- 修改点 1: 自动表头检测逻辑 ---
    first_row = df.iloc[0].tolist()
    # 规则：第一行中只要有一个非空且非数字的单元格，即视为表头
    is_header = any(not _is_numeric(c) and not (pd.isna(c) or str(c).strip() == "") for c in first_row)

    if is_header:
        # 提取第一行为标签，处理空值为 "N/A"
        x_labels = [str(c).strip() if (not pd.isna(c) and str(c).strip() != "") else "N/A" for c in first_row]
        # 数据从第二行开始
        data_df = df.iloc[1:]
    else:
        # 使用默认列名或 Col N
        n_cols = df.shape[1]
        x_labels = [f"Col {i + 1}" for i in range(n_cols)]
        # 数据即为全部
        data_df = df

    if data_df.empty:
        ax.text(0.5, 0.5, "No data to plot", ha="center", va="center", fontsize=12)
        return

    n_rows = data_df.shape[0]
    n_cols = data_df.shape[1]
    x = np.arange(n_cols, dtype=float)

    # --- 动态颜色生成 ---
    if n_rows <= 20:
        cmap = cm.get_cmap("tab20")
        line_colors = [cmap(i) for i in range(n_rows)]
    else:
        cmap = cm.get_cmap("rainbow")
        line_colors = [cmap(i) for i in np.linspace(0, 1, n_rows)]

    for row_idx in range(n_rows):
        color = line_colors[row_idx]
        
        # 强制转换为数值
        y_series = pd.to_numeric(data_df.iloc[row_idx, :], errors="coerce")
        y = y_series.to_numpy(dtype=float)
        
        # 行标签（如果数据包含表头，行索引需要偏移）
        line_label = f"Row {row_idx + 1}"
        valid = ~np.isnan(y)

        # 1. 绘制底层连接桥（低透明虚线）
        if valid.any():
            ax.plot(
                x[valid],
                y[valid],
                color=color,
                lw=2,
                linestyle="--",
                alpha=0.4,
                zorder=1,
            )

        # 2. 绘制表层主折线（自动断开）
        ax.plot(
            x,
            y,
            color=color,
            lw=2,
            marker="o",
            markersize=5,
            label=line_label,
            zorder=2,
        )

        # 3. 强化断点绘制 (白底红十字)
        if valid.any():
            prev_invalid = np.r_[False, ~valid[:-1]]
            next_invalid = np.r_[~valid[1:], False]
            boundary = valid & (prev_invalid | next_invalid)
            boundary_idx = np.where(boundary)[0]

            if boundary_idx.size > 0:
                ax.scatter(
                    x[boundary_idx],
                    y[boundary_idx],
                    color="white",
                    s=160,
                    zorder=9,
                )
                ax.scatter(
                    x[boundary_idx],
                    y[boundary_idx],
                    color="#EE884C",
                    marker="X",
                    s=100,
                    edgecolors="white",
                    linewidth=1.5,
                    zorder=10,
                )

    # --- 样式与坐标轴设置 ---
    ax.set_title(f"Line Plot_{sheet_name}", fontsize=14)
    ax.set_ylabel("Data Value", fontsize=12)
    ax.set_xticks(x)
    ax.set_xticklabels(x_labels, rotation=-45, ha="left", fontsize=10)

    ax.xaxis.grid(True, linestyle="-", which="major", color="lightgrey", alpha=0.5)
    ax.yaxis.grid(True, linestyle="-", which="major", color="lightgrey", alpha=0.5)
    ax.xaxis.set_ticks_position('bottom')

    for spine in ax.spines.values():
        spine.set_color("black")
        spine.set_linewidth(1)
        spine.set_visible(True)

    # --- 图例集成 ---
    handles, labels = ax.get_legend_handles_labels()
    
    bridge_handle = Line2D([0], [0], color="gray", lw=2, linestyle="--", alpha=0.5, label="Gap Connection")
    handles.append(bridge_handle)
    labels.append("Gap Connection")

    break_handle = Line2D([0], [0], marker="X", linestyle="None", markerfacecolor="#EE884C", 
                          markeredgecolor="white", markeredgewidth=1.5, markersize=10, label="Data Break (N/A)")
    handles.append(break_handle)
    labels.append("Data Break (N/A)")
    
    ax.legend(handles=handles, labels=labels, loc="center left", bbox_to_anchor=(1.02, 0.5))

    fig.tight_layout()