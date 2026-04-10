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
    alpha: float


def _coerce_xy(df: pd.DataFrame, x_col: int, y_col: int) -> tuple[np.ndarray, np.ndarray]:
    if df.shape[1] <= max(x_col, y_col):
        raise ValueError("选区列数不足：scatter_plot 至少需要 2 列 (X,Y)。")

    x_raw = pd.to_numeric(df.iloc[:, x_col], errors="coerce")
    y_raw = pd.to_numeric(df.iloc[:, y_col], errors="coerce")

    xy = pd.DataFrame({"x": x_raw, "y": y_raw}).dropna(how="any")
    if xy.shape[0] < 2:
        raise ValueError("有效数据不足：请确保 X/Y 两列至少有 2 行数值，并且不要只框选表头。")

    return xy["x"].to_numpy(), xy["y"].to_numpy()


def render_scatter_kde_chart(fig: Figure, df: pd.DataFrame, sheet_name: str = "Data") -> None:
    """Render a KDE-based scatter density chart with marginal histograms."""

    fig.clear()

    x1, y1 = _coerce_xy(df, 0, 1)
    groups: list[_XYGroup] = [
        _XYGroup(
            x=x1,
            y=y1,
            cmap="Blues",
            hist_color="#A8D1E7",
            hist_edgecolor="#6D9DBB",
            alpha=0.7,
        )
    ]

    has_second_group = df.shape[1] >= 4
    if has_second_group:
        try:
            x2, y2 = _coerce_xy(df, 2, 3)
            groups.append(
                _XYGroup(
                    x=x2,
                    y=y2,
                    cmap="Reds",
                    hist_color="#F19C9C",
                    hist_edgecolor="#C26A6A",
                    alpha=0.6,
                )
            )
        except Exception:
            # 第二组不满足有效数据条件时，退化为单组
            pass

    all_x = np.concatenate([g.x for g in groups])
    all_y = np.concatenate([g.y for g in groups])

    global_min = float(min(all_x.min(), all_y.min()))
    global_max = float(max(all_x.max(), all_y.max()))
    data_range = global_max - global_min
    data_padding = data_range if data_range > 0 else 1.0
    pad = 0.10 * data_padding

    if fig.get_size_inches().min() <= 0:
        fig.set_size_inches(10, 8)

    # 1. 设置画布预留空间: 给顶部(top)留空间放主标题, 右侧(right)留出15%的空间放色标(colorbar)
    fig.subplots_adjust(left=0.08, right=0.85, top=0.90, bottom=0.10)

    # 2. 将标题放在全局最顶端，而不是主图坐标轴内，避免和 top 直方图重叠
    fig.suptitle(f"Scatter Density - {sheet_name}", fontsize=14, fontweight='bold', y=0.96)

    # 3. 调整网格间距 (wspace/hspace)，让直方图与主图靠得更紧密（原本的 0.4 间隙太宽了）
    gs = fig.add_gridspec(
        4,
        4,
        width_ratios=[1, 1, 1, 0.4],
        height_ratios=[0.4, 1, 1, 1],
        wspace=0.08, 
        hspace=0.08,
    )

    ax_top = fig.add_subplot(gs[0, 0:3])
    ax_main = fig.add_subplot(gs[1:4, 0:3])
    ax_right = fig.add_subplot(gs[1:4, 3], sharey=ax_main)

    # KDE 主图
    for g in groups:
        sns.kdeplot(
            x=g.x,
            y=g.y,
            ax=ax_main,
            cmap=g.cmap,
            fill=True,
            thresh=0.05,
            alpha=g.alpha,
            levels=12,
        )

    ax_main.set_xlim(global_min - pad, global_max + pad)
    ax_main.set_ylim(global_min - pad, global_max + pad)
    ax_main.set_xlabel("X")
    ax_main.set_ylabel("Y")

    binwidth = max(data_range / 30.0, 1e-6)

    # 顶部边缘直方图 (X)
    for idx, g in enumerate(groups):
        sns.histplot(
            x=g.x,
            ax=ax_top,
            color=g.hist_color,
            edgecolor=g.hist_edgecolor,
            binwidth=binwidth,
            stat="probability",
            alpha=0.6,
            label=("Group 1" if idx == 0 else "Group 2"),
        )
    ax_top.set_xlim(ax_main.get_xlim())
    ax_top.tick_params(labelbottom=False)
    ax_top.set_xlabel("")
    ax_top.set_ylabel("Freq", fontsize=10)
    if len(groups) > 1:
        ax_top.legend(loc="upper right", frameon=False)
    ax_top.spines["top"].set_visible(False)
    ax_top.spines["right"].set_visible(False)
    # 底部边框也隐藏一下，让图表衔接更现代
    ax_top.spines["bottom"].set_visible(False) 

    # 右侧边缘直方图 (Y)
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
    # 4. Colorbar 自动布局修复
    # ==========================
    fig.canvas.draw()
    pos_main = ax_main.get_position()
    pos_right = ax_right.get_position()

    cb_width = 0.012
    # 关键点：Colorbar 应该放在右侧直方图的右边，而不是主图的右边！
    cb_x = pos_right.x1 + 0.02 

    if len(groups) == 1:
        cax = fig.add_axes([cb_x, pos_main.y0 + pos_main.height * 0.20, cb_width, pos_main.height * 0.60])
        sm = cm.ScalarMappable(cmap=groups[0].cmap, norm=Normalize(vmin=0, vmax=1.0))
        sm.set_array([])
        fig.colorbar(sm, cax=cax)
    else:
        cb_height = pos_main.height * 0.40
        cax_blue = fig.add_axes([cb_x, pos_main.y0 + pos_main.height * 0.55, cb_width, cb_height])
        sm_blue = cm.ScalarMappable(cmap="Blues", norm=Normalize(vmin=0, vmax=1.0))
        sm_blue.set_array([])
        fig.colorbar(sm_blue, cax=cax_blue)

        cax_red = fig.add_axes([cb_x, pos_main.y0 + pos_main.height * 0.05, cb_width, cb_height])
        sm_red = cm.ScalarMappable(cmap="Reds", norm=Normalize(vmin=0, vmax=1.0))
        sm_red.set_array([])
        fig.colorbar(sm_red, cax=cax_red)