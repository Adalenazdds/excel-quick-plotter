import sys
import traceback
from typing import Optional

import pandas as pd
import xlwings as xw
from PyQt5.QtCore import QObject, QPoint, Qt, QThread, pyqtSignal
from PyQt5.QtGui import QMouseEvent
from PyQt5.QtWidgets import (
    QApplication,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSizePolicy,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

import matplotlib
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
from matplotlib.figure import Figure

from plotbox import render_box_and_scatter_chart

try:
	import pythoncom as _pythoncom
except Exception:
	_pythoncom = None


class ExcelFetchWorker(QObject):

	finished = pyqtSignal(object, dict)  # pandas.DataFrame, metadata dict
	failed = pyqtSignal(str)

	def __init__(self):
		super().__init__()

	def run(self):
		try:
			if _pythoncom is not None:
				try:
					_pythoncom.CoInitialize()
				except Exception:
					pass

			app = xw.apps.active
			if app is None:
				raise RuntimeError(
					"xlwings 未检测到活动 Excel 实例（常见原因：Excel 以管理员运行/不是微软 Excel/WPS/不同用户会话）"
				)

			book = app.books.active
			selection = book.app.selection
			values = selection.options(ndim=2).value
			if values is None:
				raise ValueError("empty selection")

			# 1) 打印活动 Excel/工作簿/选区信息（尽量使用基础属性，避免版本差异）
			try:
				book_name = getattr(book, "name", None)
				addr = None
				try:
					addr = selection.address
				except Exception:
					addr = None
				print("[Excel] Active workbook:", book_name)
				if addr:
					print("[Excel] Selection address:", addr)
			except Exception:
				pass

			# 2) 打印选区中的数值（数字类型）
			try:
				numbers = []
				for row in values:
					for cell in row:
						if isinstance(cell, bool):
							continue
						if isinstance(cell, (int, float)):
							# 过滤 NaN
							if isinstance(cell, float) and cell != cell:
								continue
							numbers.append(cell)
				print("[Excel] Selected numeric values:")
				print(numbers)
			except Exception:
				pass

			# 选区可能是空白区域：例如选到一大片空单元格时，values 会是 [[None], ...]
			has_any_value = False
			try:
				for row in values:
					for cell in row:
						if cell is None:
							continue
						if isinstance(cell, str) and cell.strip() == "":
							continue
						has_any_value = True
						break
					if has_any_value:
						break
			except Exception:
				has_any_value = True

			if not has_any_value:
				raise ValueError("empty selection")

			df = pd.DataFrame(values)
			try:
				# 兼容性输出：保留 DataFrame 供后续作图使用
				print("[Excel] DataFrame preview:")
				print(df)
			except Exception:
				pass

			meta = {
				"book_name": getattr(book, "name", "未知表"),
				"sheet_name": (
					getattr(selection.sheet, "name", "未知页")
					if hasattr(selection, "sheet")
					else "未知页"
				),
				"address": getattr(selection, "address", "未知选区"),
				"filepath": getattr(book, "fullname", ""),
			}
			self.finished.emit(df, meta)
		except Exception as exc:
			try:
				print("[ExcelFetchWorker] failed:")
				print(traceback.format_exc())
			except Exception:
				pass
			self.failed.emit(str(exc))
		finally:
			if _pythoncom is not None:
				try:
					_pythoncom.CoUninitialize()
				except Exception:
					pass


class FloatingToolWindow(QWidget):

	def __init__(self):
		super().__init__()

		self._drag_active = False
		self._drag_offset = None  # type: Optional[QPoint]
		self._excel_thread = None  # type: Optional[QThread]
		self._excel_worker = None  # type: Optional[ExcelFetchWorker]
		self._last_df = None
		self._chart_windows = []

		self._init_window()
		self._init_ui()

	def _init_window(self) -> None:
		self.setWindowTitle("EXCEL快速分析")
		self.setWindowFlags(Qt.Window | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
		self.setMinimumSize(250, 100)
		self.resize(250, 120)

	def _init_ui(self) -> None:
		root = QVBoxLayout(self)
		root.setContentsMargins(10, 10, 10, 10)
		root.setSpacing(10)

		self.top_bar = QWidget(self)
		top_layout = QHBoxLayout(self.top_bar)
		top_layout.setContentsMargins(0, 0, 0, 0)
		top_layout.setSpacing(8)

		self.status_label = QLabel("就绪", self.top_bar)
		self.status_label.setSizePolicy(
			QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
		)

		self.pin_button = QToolButton(self.top_bar)
		self.pin_button.setCheckable(True)
		self.pin_button.setChecked(True)
		self.pin_button.setToolTip("切换是否置顶")
		self.pin_button.setAutoRaise(True)
		self.pin_button.setFixedSize(28, 28)
		self._apply_pin_visual(True)
		self.pin_button.toggled.connect(self._set_always_on_top)

		top_layout.addWidget(self.status_label)
		top_layout.addWidget(self.pin_button)
		root.addWidget(self.top_bar)

		self.info_label = QLabel("当前未提取数据\n请选中Excel数据区后点击底部按钮", self)
		self.info_label.setWordWrap(True)
		self.info_label.setStyleSheet("color: #555; font-size: 13px;")
		self.info_label.setAlignment(Qt.AlignCenter)
		root.addWidget(self.info_label, 1)

		self.action_button = QPushButton("提取并作图", self)
		self.action_button.clicked.connect(self._on_extract_plot_clicked)
		root.addWidget(self.action_button)

	def set_status(self, text: str) -> None:
		self.status_label.setText(text)

	def _apply_pin_visual(self, pinned: bool) -> None:
		self.pin_button.setText("📌" if pinned else "📍")

	def _set_always_on_top(self, on: bool) -> None:
		self._apply_pin_visual(on)
		self.setWindowFlag(Qt.WindowStaysOnTopHint, on)
		self.show()

	def _on_extract_plot_clicked(self) -> None:
		if self._excel_thread is not None and self._excel_thread.isRunning():
			return

		self.set_status("读取中...")
		self.action_button.setEnabled(False)

		thread = QThread(self)
		worker = ExcelFetchWorker()
		worker.moveToThread(thread)

		thread.started.connect(worker.run)
		worker.finished.connect(self._on_excel_fetch_success)
		worker.failed.connect(self._on_excel_fetch_failed)
		worker.finished.connect(thread.quit)
		worker.failed.connect(thread.quit)
		thread.finished.connect(worker.deleteLater)
		thread.finished.connect(thread.deleteLater)

		self._excel_thread = thread
		self._excel_worker = worker
		thread.start()

	def _on_excel_fetch_success(self, df, meta) -> None:
		self._last_df = df

		# 更新悬浮窗的信息
		book_name = meta.get("book_name", "未知")
		address = meta.get("address", "未知")
		filepath = meta.get("filepath", "")

		info_text = f"文件: {book_name}\n范围: {address}"
		if filepath:
			info_text += f"\n路径: {filepath}"
		self.info_label.setText(info_text)

		try:
			print(df)
		except Exception:
			pass
		render_ok = True
		try:
			self._show_chart_window(df, meta)
		except Exception as exc:
			render_ok = False
			try:
				print("[UI] render chart failed:", repr(exc))
			except Exception:
				pass
			self.set_status("作图失败")
		if render_ok:
			self.set_status("就绪")
		self.action_button.setEnabled(True)
		self._excel_thread = None
		self._excel_worker = None

	def _on_excel_fetch_failed(self, _message: str) -> None:
		try:
			print("[UI] Excel fetch failed:", _message)
		except Exception:
			pass
		self.set_status("未检测到有效数据")
		self.action_button.setEnabled(True)
		self._excel_thread = None
		self._excel_worker = None

	def _hit_interactive_widget(self, local_pos: QPoint) -> bool:
		widget = self.childAt(local_pos)
		while widget is not None:
			if widget in (self.pin_button, self.action_button):
				return True
			widget = widget.parentWidget()
		return False

	def mousePressEvent(self, event: QMouseEvent) -> None:
		if event.button() == Qt.LeftButton:
			if not self._hit_interactive_widget(event.pos()):
				self._drag_active = True
				self._drag_offset = event.globalPos() - self.frameGeometry().topLeft()
				event.accept()
				return
		super().mousePressEvent(event)

	def mouseMoveEvent(self, event: QMouseEvent) -> None:
		if (
			self._drag_active
			and (event.buttons() & Qt.LeftButton)
			and self._drag_offset is not None
		):
			self.move(event.globalPos() - self._drag_offset)
			event.accept()
			return
		super().mouseMoveEvent(event)

	def mouseReleaseEvent(self, event: QMouseEvent) -> None:
		if event.button() == Qt.LeftButton:
			self._drag_active = False
			self._drag_offset = None
			event.accept()
			return
		super().mouseReleaseEvent(event)

	def closeEvent(self, event) -> None:
		thread = self._excel_thread
		if thread is not None and thread.isRunning():
			try:
				thread.requestInterruption()
			except Exception:
				pass
			thread.quit()
			thread.wait(1500)
		super().closeEvent(event)

	def _show_chart_window(self, df, meta) -> None:
		# 设置中文字体支持
		matplotlib.rcParams["font.sans-serif"] = [
			"Microsoft YaHei",
			"SimHei",
			"SimSun",
			"Arial Unicode MS",
		]
		matplotlib.rcParams["axes.unicode_minus"] = False

		# 当图表数据量较大时，自动计算动态图表大小
		num_cols = df.shape[1]
		fig_width = max(8, num_cols * 1.5)

		chart_win = QWidget()
		chart_win.setAttribute(Qt.WA_DeleteOnClose)
		chart_win.setWindowTitle(
			f"分析图表 - {meta.get('book_name', '')} | 范围: {meta.get('address', '')}"
		)

		# 设定窗口大小并显示
		chart_win.resize(int(fig_width * 100), 600)
		layout = QVBoxLayout(chart_win)

		fig = Figure(figsize=(fig_width, 6), dpi=100)
		canvas = FigureCanvas(fig)

		# 添加 Matplotlib 工具栏（自带缩放、保存等功能）
		toolbar = NavigationToolbar(canvas, chart_win)
		layout.addWidget(toolbar)
		layout.addWidget(canvas, 1)

		ax = fig.add_subplot(111)

		try:
			render_box_and_scatter_chart(
				ax, df, sheet_name=meta.get("sheet_name", "Data")
			)
		except Exception as exc:
			ax.text(0.5, 0.5, f"作图失败: {exc}", ha="center", va="center")
			print(f"[UI] plotbox render failed: {exc}")

		try:
			fig.tight_layout()
		except Exception:
			pass

		chart_win.show()

		# 注册销毁事件清理引用，防止内存泄露
		def on_destroyed():
			if chart_win in self._chart_windows:
				self._chart_windows.remove(chart_win)

		chart_win.destroyed.connect(on_destroyed)

		self._chart_windows.append(chart_win)


def main():
    app = QApplication(sys.argv)
    window = FloatingToolWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
