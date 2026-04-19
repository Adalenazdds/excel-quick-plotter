import numpy as np
import pandas as pd
from matplotlib.axes import Axes
from matplotlib.lines import Line2D
import matplotlib.ticker as mtick
from typing import Optional

def _is_blank_cell(value) -> bool:
    if value is None: return True
    if pd.isna(value): return True
    if isinstance(value, str) and value.strip() == "": return True
    return False

def _is_numeric_type_cell(value) -> bool:
    if isinstance(value, bool): return False
    return isinstance(value, (int, float, np.number))

def _detect_header_row(df: pd.DataFrame) -> bool:
    if df is None or df.shape[0] == 0 or df.shape[1] == 0: return False
    first_row = df.iloc[0, :].tolist()
    for cell in first_row:
        if _is_blank_cell(cell): continue
        if not _is_numeric_type_cell(cell): return True
    return False

def render_pareto_chart(
    ax: Axes,
    df: pd.DataFrame,
    sheet_name: str = "Data",
    excel_start_row: Optional[int] = None,
) -> None:
    # 1. 识别表头
    has_header = _detect_header_row(df)
    if has_header:
        headers =["N/A" if _is_blank_cell(v) else str(v).strip() for v in df.iloc[0, :]]
        data_df = df.iloc[1:, :]
    else:
        headers =[f"Col {col}" if isinstance(col, int) else str(col) for col in df.columns]
        data_df = df

    # 2. 智能数据解析策略
    items = []
    if data_df.shape[1] == 1:
        # 策略 A：仅 1 列，视作分类文本，进行频次统计
        counts = data_df.iloc[:, 0].dropna().value_counts()
        for k, v in counts.items():
            if not _is_blank_cell(k):
                items.append((str(k).strip(), float(v)))
    elif data_df.shape[1] == 2:
        # 策略 B：2 列，尝试以第 1 列为标签，第 2 列为数值进行汇总
        col1 = data_df.iloc[:, 0]
        col2 = pd.to_numeric(data_df.iloc[:, 1], errors='coerce')
        if col2.notna().sum() > 0:
            temp_df = pd.DataFrame({'label': col1, 'val': col2}).dropna(subset=['val'])
            grouped = temp_df.groupby('label')['val'].sum()
            for k, v in grouped.items():
                if not _is_blank_cell(k) and v > 0:
                    items.append((str(k).strip(), float(v)))
                    
    if not items:
        # 策略 C：多列或上述策略失败，视每列为一个分类，汇总其所有有效数值
        for i, col in enumerate(data_df.columns):
            numeric_vals = pd.to_numeric(data_df[col], errors='coerce').dropna()
            val_sum = numeric_vals.sum()
            if val_sum > 0:
                items.append((headers[i], float(val_sum)))

    # 降序排列并计算累计百分比
    items.sort(key=lambda x: x[1], reverse=True)
    if not items:
        raise ValueError("未找到可用于绘制帕累托图的有效正数数据")

    labels = [x[0] for x in items]
    values = [x[1] for x in items]
    total = sum(values)
    cum_pct = np.cumsum(values) / total * 100
    x_positions = np.arange(len(labels))

    # 3. 绘制帕累托柱状图 (继承 box_plot 风格)
    bar_color = "royalblue"
    bars = ax.bar(
        x_positions, 
        values, 
        color=bar_color, 
        alpha=0.6, 
        edgecolor="white", 
        linewidth=1,
        width=0.6
    )

    # 4. 绘制累计百分比折线图 (双 Y 轴)
    ax2 = ax.twinx()
    line_color = "red"
    ax2.plot(
        x_positions, cum_pct, color=line_color, marker="o", 
        linewidth=2, markersize=6, markerfacecolor="white", 
        markeredgecolor=line_color, markeredgewidth=1.5
    )

    # 80% 黄金阈值虚线
    ax2.axhline(80, color="gray", linestyle="--", linewidth=1.5, alpha=0.7)
    
    # 5. 坐标轴与样式配置
    ax.set_xticks(x_positions)
    ax.set_xticklabels(labels, rotation=-45, ha="left", fontsize=10)
    ax.xaxis.set_ticks_position('bottom')
    
    ax.set_ylabel("Frequency / Sum", fontsize=12)
    ax2.set_ylabel("Cumulative Percentage (%)", fontsize=12)
    ax2.set_ylim(0, 105)
    ax2.yaxis.set_major_formatter(mtick.PercentFormatter())
    
    ax.yaxis.grid(True, linestyle="-", which="major", color="lightgrey", alpha=0.5)
    ax.set_title(f"Pareto Analysis_{sheet_name}", fontsize=14)

    for spine in ax.spines.values():
        spine.set_color("black")
        spine.set_linewidth(1)

    # 6. 图例外置配置
    legend_elements =[
        Line2D([0], [0], color=bar_color, lw=6, alpha=0.6, label="Value"),
        Line2D([0], [0], color=line_color, marker="o", lw=2, markerfacecolor="white", label="Cumulative %"),
        Line2D([0], [0], color="gray", linestyle="--", lw=1.5, label="80% Threshold")
    ]
    ax.legend(
        handles=legend_elements, 
        loc='center left', 
        bbox_to_anchor=(1.20, 0.5), # 距离拉宽以避开双 Y 轴标签
        borderaxespad=0.
    )
    ax.figure.tight_layout()

    # 7. 交互悬浮气泡
    try:
        import mplcursors
        prev_cursor = getattr(ax.figure, "_eqp_mplcursors_cursor", None)
        if prev_cursor is not None:
            try: prev_cursor.remove()
            except Exception: pass

        for i, bar in enumerate(bars):
            setattr(bar, "_eqp_meta", {"label": labels[i], "value": values[i], "cum_pct": cum_pct[i]})

        cursor = mplcursors.cursor(bars, hover=True)
        @cursor.connect("add")
        def _on_add(sel):
            meta = getattr(sel.artist, "_eqp_meta", None)
            if not meta: return
            sel.annotation.set_text(
                f"Category: {meta['label']}\n"
                f"Value: {meta['value']:g}\n"
                f"Cum. Pct: {meta['cum_pct']:.1f}%"
            )
        setattr(ax.figure, "_eqp_mplcursors_cursor", cursor)
    except Exception:
        pass
