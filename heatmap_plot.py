from __future__ import annotations

import copy

import numpy as np
import pandas as pd
import matplotlib.cm as cm
from matplotlib.colors import Normalize
from matplotlib.figure import Figure
from mpl_toolkits.axes_grid1 import make_axes_locatable

from numeric_coercion import normalize_numeric_like


def coerce_numeric_matrix(df: pd.DataFrame, *, missing: str = "keep") -> pd.DataFrame:
	"""Coerce a DataFrame to a numeric matrix.

	- Strips whitespace
	- Treats empty strings as missing
	- Supports common formats like "1,000" and "12%" (percent sign is stripped)
	- Converts non-numeric cells to NaN

	Parameters
	----------
	missing:
		"keep" (default): keep NaN; heatmap will mask them.
		"fill0": fill NaN with 0.
		"drop": drop rows/cols that contain any NaN.
	"""

	cleaned = normalize_numeric_like(df.copy())

	# Convert to numeric. Non-numeric becomes NaN.
	numeric_df = cleaned.apply(lambda col: pd.to_numeric(col, errors="coerce"))

	# Drop fully empty rows/cols after numeric coercion.
	numeric_df = numeric_df.dropna(axis=0, how="all").dropna(axis=1, how="all")

	if missing == "fill0":
		numeric_df = numeric_df.fillna(0)
	elif missing == "drop":
		numeric_df = numeric_df.dropna(axis=0, how="any").dropna(axis=1, how="any")
	elif missing == "keep":
		pass
	else:
		raise ValueError(f"unknown missing strategy: {missing!r}")

	if numeric_df.shape[0] == 0 or numeric_df.shape[1] == 0:
		raise ValueError("未找到有效数值矩阵")
	if not np.isfinite(numeric_df.to_numpy(dtype=float)).any():
		raise ValueError("未找到有效数值矩阵")

	return numeric_df


def render_heatmap_chart(fig: Figure, df: pd.DataFrame, sheet_name: str = "Data") -> None:
	"""Render a Matplotlib heatmap from a rectangular Excel selection.

	- No normalization.
	- Non-numeric cells are coerced to NaN.
	- Tick labels are hidden by default (clean look).
	- Values are annotated only for small matrices.
	"""

	fig.clear()

	numeric_df = coerce_numeric_matrix(df)
	data = numeric_df.to_numpy(dtype=float)

	rows, cols = int(data.shape[0]), int(data.shape[1])
	finite_mask = np.isfinite(data)

	vmin = float(np.nanmin(data))
	vmax = float(np.nanmax(data))

	# Guard against degenerate scale.
	if not np.isfinite(vmin) or not np.isfinite(vmax):
		raise ValueError("未找到有效数值矩阵")

	ax = fig.add_subplot(111)

	# 【修正 1】移除硬编码的 subplots_adjust，避免非对称边距导致整体偏右
	fig.suptitle(f"Heatmap - {sheet_name}", fontsize=14, fontweight="bold", y=0.95)

	cmap_base = cm.get_cmap("Blues")
	cmap = copy.copy(cmap_base)
	try:
		cmap.set_bad(color="#E5E7EB")
	except Exception:
		pass

	masked = np.ma.masked_invalid(data)
	im = ax.imshow(
		masked,
		cmap=cmap,
		interpolation="none",
		aspect="equal",
		vmin=vmin,
		vmax=vmax,
		resample=False,
	)

	# Pixel-aligned edges (half-pixel boundaries) for cleaner rendering.
	ax.set_xlim(-0.5, cols - 0.5)
	ax.set_ylim(rows - 0.5, -0.5)

	# 【修正 2】使用 make_axes_locatable 完美绑定 Heatmap 与 Colorbar，确保作为一个整体居中
	divider = make_axes_locatable(ax)
	# 在 heatmap 的右侧划出 5% 的宽度，间距 0.15 放置 colorbar
	cax = divider.append_axes("right", size="5%", pad=0.15)

	cbar = fig.colorbar(im, cax=cax)
	cbar.ax.tick_params(labelsize=9)

	# Clean look: hide ticks and labels by default.
	ax.set_xticks([])
	ax.set_yticks([])

	# Small-matrix annotations.
	annotate_threshold = 400
	if rows * cols <= annotate_threshold:
		norm = Normalize(vmin=vmin, vmax=vmax)
		for r in range(rows):
			for c in range(cols):
				if not finite_mask[r, c]:
					continue
				val = float(data[r, c])
				intensity = float(norm(val)) if vmax != vmin else 0.0
				text_color = "white" if intensity >= 0.60 else "black"
				ax.text(
					c,
					r,
					f"{val:.3g}",
					ha="center",
					va="center",
					fontsize=8,
					color=text_color,
				)

	for spine in ax.spines.values():
		spine.set_color("black")
		spine.set_linewidth(1)

