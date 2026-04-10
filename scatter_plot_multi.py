from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd
import seaborn as sns
from matplotlib.colors import Normalize
from matplotlib.figure import Figure
import matplotlib.cm as cm


@dataclass(frozen=True)
class _XYGroup:
    x: np.ndarray
    y: np.ndarray
    cmap: str
    hist_color: str
    hist_edgecolor: str
    name: str


# 预设的“多巴胺”同系列色彩主题，支持无限循环拓展
THEMES = [
    {"cmap": "Blues",   "hist_color": "#A8D1E7", "hist_edgecolor": "#6D9DBB"},
    {"cmap": "Reds",    "hist_color": "#F19C9C", "hist_edgecolor": "#C26A6A"},
    {"cmap": "Greens",  "hist_color": "#A1D9B5", "hist_edgecolor": "#4E9A68"},
    {"cmap": "Purples", "hist_color": "#C8B8E0", "hist_edgecolor": "#735A9C"},
    {"cmap": "Oranges", "hist_color": "#F8C488", "hist_edgecolor": "#D07B28"},
    {"cmap": "Greys",   "hist_color": "#D5D8DC", "hist_edgecolor": "#5D6D7E"},
]


def _coerce_xy(df: pd.DataFrame, x_col: int, y_col: int) -> tuple[np.ndarray, np.ndarray]:
    if df.shape[1] <= max(x_col, y_col):
        raise ValueError(f"选区列数不足以提取第 {x_col+1} 和 {y_col+1} 列。")

    x_raw = pd.to_numeric(df.iloc[:, x_col], errors="coerce")
    y_raw = pd.to_numeric(df.iloc[:, y_col], errors="coerce")

    xy = pd.DataFrame({"x": x_raw, "y": y_raw}).dropna(how="any")
    if xy.shape[0] < 2:
        raise ValueError(f"第 {x_col+1} 和 {y_col+1} 列有效数值对不足 2 行。")

    return xy["x"].to_numpy(), xy["y"].to_numpy()


def render_multi_scatter_kde_chart(fig: Figure, df: pd.DataFrame, sheet_name: str = "Data") -> None:
    """Render a dynamic multi-group KDE scatter density chart."""
    
    fig.clear()
    groups: list[_XYGroup] = []
    
    # 自动解析多组数据：每相邻两列为一组
    num_cols = df.shape[1]
    for i in range(0, num_cols, 2):
        if i + 1 >= num_cols:
            break  # 最后一列落单，自动跳过
            
        try:
            x_arr, y_arr = _coerce_xy(df, i, i + 1)
            # 通过取模运算实现颜色的无限循环
            theme_idx = (i // 2) % len(THEMES)
            theme = THEMES[theme_idx]
            
            groups.append(
                _XYGroup(
                    x=x_arr,
                    y=y_arr,
                    cmap=theme["cmap"],
                    hist_color=theme["hist_color"],
                    hist_edgecolor=theme["hist_edgecolor"],
                    name=f"Group {(i // 2) + 1}"
                )
            )
        except Exception as e:
            print(f"[Warn] 跳过组 {i//2 + 1}: {e}")
            pass

    n_groups = len(groups)
    if n_groups == 0:
        ax = fig.add_subplot(111)
        ax.text(0.5, 0.5, "未找到有效的多组数据(请确保每两列成对出现数值)", ha='center', va='center', fontsize=12)
        return

    all_x = np.concatenate([g.x for g in groups])
    all_y = np.concatenate([g.y for g in groups])

    global_min = float(min(all_x.min(), all_y.min()))
    global_max = float(max(all_x.max(), all_y.max()))
    data_range = global_max - global_min
    data_padding = data_range if data_range > 0 else 1.0
    pad = 0.10 * data_padding

    if fig.get_size_inches().min() <= 0:
        fig.set_size_inches(10, 8)

    # 布局参数：右侧留白 15% 给 Colorbar
    fig.subplots_adjust(left=0.08, right=0.85, top=0.90, bottom=0.10)
    fig.suptitle(f"Scatter Density (Multi-Group) - {sheet_name}", fontsize=14, fontweight='bold', y=0.96)

    gs = fig.add_gridspec(
        4, 4,
        width_ratios=[1, 1, 1, 0.4],
        height_ratios=[0.4, 1, 1, 1],
        wspace=0.08, 
        hspace=0.08,
    )

    ax_top = fig.add_subplot(gs[0, 0:3])
    ax_main = fig.add_subplot(gs[1:4, 0:3])
    ax_right = fig.add_subplot(gs[1:4, 3], sharey=ax_main)

    # -----------------------------------------------------------
    # 动态计算云图透明度：组数越多，整体透明度越高（Alpha 越小），防止互相遮盖
    # -----------------------------------------------------------
    dynamic_alpha = max(0.85 - n_groups * 0.15, 0.25)

    # 1. 绘制中间主图
    for g in groups:
        sns.kdeplot(
            x=g.x, y=g.y,
            ax=ax_main,
            cmap=g.cmap,
            fill=True,
            thresh=0.05,
            alpha=dynamic_alpha, 
            levels=12,
        )

    ax_main.set_xlim(global_min - pad, global_max + pad)
    ax_main.set_ylim(global_min - pad, global_max + pad)
    ax_main.set_xlabel("X")
    ax_main.set_ylabel("Y")

    binwidth = max(data_range / 30.0, 1e-6)

    # 2. 绘制顶部边缘直方图
    for g in groups:
        sns.histplot(
            x=g.x,
            ax=ax_top,
            color=g.hist_color,
            edgecolor=g.hist_edgecolor,
            binwidth=binwidth,
            stat="probability",
            alpha=0.6,  # 直方图保持适中的透明度即可
            label=g.name,
        )
        
    ax_top.set_xlim(ax_main.get_xlim())
    ax_top.tick_params(labelbottom=False)
    ax_top.set_xlabel("")
    ax_top.set_ylabel("Freq", fontsize=10)
    
    if len(groups) > 1:
        ax_top.legend(loc="upper right", frameon=False)
        
    ax_top.spines["top"].set_visible(False)
    ax_top.spines["right"].set_visible(False)
    ax_top.spines["bottom"].set_visible(False) 

    # 3. 绘制右侧边缘直方图
    for g in groups:
        sns.histplot(
            y=g.y,
            ax=ax_right,
            color=g.hist_color,
            edgecolor=g.hist_edgecolor,
            binwidth=binwidth,
            stat="probability",
            alpha=0.6,
        )
        
    ax_right.tick_params(labelleft=False)
    ax_right.set_ylabel("")
    ax_right.set_xlabel("Freq", fontsize=10)
    ax_right.spines["top"].set_visible(False)
    ax_right.spines["right"].set_visible(False)
    ax_right.spines["left"].set_visible(False)

    # ==========================
    # 4. Colorbar 动态堆叠布局
    # ==========================
    fig.canvas.draw()
    pos_main = ax_main.get_position()
    pos_right = ax_right.get_position()

    cb_width = 0.012
    cb_x = pos_right.x1 + 0.02
    
    n = n_groups
    gap = 0.02
    
    # 如果只有1组，高度0.6；多组则平分高度，最高不超过0.4
    if n == 1:
        cb_h = pos_main.height * 0.60
    else:
        cb_h = min((pos_main.height - (n - 1) * gap) / n, 0.4)
    
    current_y = pos_main.y0 + pos_main.height - cb_h
    if n == 1:
        current_y = pos_main.y0 + pos_main.height * 0.20

    for g in groups:
        cax = fig.add_axes([cb_x, current_y, cb_width, cb_h])
        sm = cm.ScalarMappable(cmap=g.cmap, norm=Normalize(vmin=0, vmax=1.0))
        sm.set_array([])
        fig.colorbar(sm, cax=cax)
        current_y -= (cb_h + gap)