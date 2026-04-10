import sys
import traceback
from typing import Optional

import pandas as pd
import xlwings as xw
from PyQt5.QtCore import QObject, QPoint, Qt, QThread, pyqtSignal
from PyQt5.QtGui import QMouseEvent, QIcon
from PyQt5.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMenu,
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

# 确保你的同级目录下有这三个文件
from box_plot import render_box_and_scatter_chart
from scatter_plot import render_scatter_kde_chart
from scatter_plot_multi import render_multi_scatter_kde_chart

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

            meta = {
                "book_name": getattr(book, "name", "未知表"),
                "sheet_name": (
                    getattr(selection.sheet, "name", "未知页")
                    if hasattr(selection, "sheet")
                    else "未知页"
                ),
                "address": getattr(selection, "address", "未知选区"),
                "filepath": getattr(book, "fullname", ""),
                "nrows": int(getattr(df, "shape", (0, 0))[0]),
                "ncols": int(getattr(df, "shape", (0, 0))[1]),
            }
            self.finished.emit(df, meta)
        except Exception as exc:
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
        self._chart_type = "box"  # box | scatter | multi

        self._init_window()
        self._init_ui()

    def _init_window(self) -> None:
        self.setWindowTitle("EXCEL快速分析")
        self.setWindowIcon(QIcon('EXCEL-Quick-Plotter.ico'))
        self.setWindowFlags(Qt.Window | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setMinimumSize(280, 150)
        self.resize(320, 240)

    def _init_ui(self) -> None:
        base_layout = QVBoxLayout(self)
        base_layout.setContentsMargins(10, 10, 10, 10)

        self.main_frame = QFrame(self)
        self.main_frame.setObjectName("MainFrame")
        self.main_frame.setStyleSheet("""
            QFrame#MainFrame {
                background-color: #F8F9FA;
                border-radius: 16px;
                border: 1px solid #E9ECEF;
            }
        """)
        base_layout.addWidget(self.main_frame)

        root = QVBoxLayout(self.main_frame)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(12)

        # -- 顶部状态栏 --
        self.top_bar = QWidget(self.main_frame)
        top_layout = QHBoxLayout(self.top_bar)
        top_layout.setContentsMargins(0, 0, 0, 0)
        
        self.status_label = QLabel("就绪 🎈", self.top_bar)
        self.status_label.setStyleSheet("font-size:16px; font-weight:900; color:#2C3E50;")
        self.status_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        self.pin_button = QToolButton(self.top_bar)
        self.pin_button.setCheckable(True)
        self.pin_button.setChecked(True)
        self.pin_button.setToolTip("切换是否置顶")
        self.pin_button.setFixedSize(32, 32)
        self.pin_button.setStyleSheet("""
            QToolButton {
                background-color: transparent; border-radius: 16px; font-size: 16px;
            }
            QToolButton:hover { background-color: #E2E8F0; }
        """)
        self._apply_pin_visual(True)
        self.pin_button.toggled.connect(self._set_always_on_top)

        # 修复的图表选择按钮
        self.chart_button = QToolButton(self.top_bar)
        self.chart_button.setToolTip("选择图表类型")
        # 宽度调大一点以便显示文字和下拉箭头
        self.chart_button.setFixedSize(80, 32) 
        self.chart_button.setStyleSheet("""
            QToolButton {
                background-color: #E2E8F0; border-radius: 16px; font-size: 12px; font-weight: bold; color: #2C3E50;
            }
            QToolButton:hover { background-color: #CBD5E1; }
            QToolButton::menu-indicator { image: none; } /* 隐藏默认的难看的箭头 */
        """)
        
        # 使用 MenuButtonPopup，点击按钮触发主动作，点击边缘下拉
        # 但为了用户体验一致，我们直接将整个按钮变成触发下拉菜单的入口
        self.chart_button.setPopupMode(QToolButton.InstantPopup)
        
        chart_menu = QMenu(self.chart_button)
        action_box = chart_menu.addAction("Box Plot")
        action_scatter = chart_menu.addAction("Scatter (双组)")
        action_multi = chart_menu.addAction("Scatter (多组)")
        
        action_box.triggered.connect(lambda: self._set_chart_type("box"))
        action_scatter.triggered.connect(lambda: self._set_chart_type("scatter"))
        action_multi.triggered.connect(lambda: self._set_chart_type("multi"))
        
        self.chart_button.setMenu(chart_menu)
        self._apply_chart_visual()

        top_layout.addWidget(self.status_label)
        top_layout.addWidget(self.chart_button)
        top_layout.addWidget(self.pin_button)
        root.addWidget(self.top_bar)

        # -- 信息展示卡片 --
        self.info_card = QFrame(self.main_frame)
        self.info_card.setObjectName("InfoCard")
        self.info_card.setStyleSheet("""
            QFrame#InfoCard {
                background-color: #FFFFFF;
                border-radius: 12px;
            }
        """)
        info_layout = QVBoxLayout(self.info_card)
        info_layout.setContentsMargins(12, 12, 12, 12)
        info_layout.setSpacing(10)

        self.info_title = QLabel("📊 当前活动选区", self.info_card)
        self.info_title.setStyleSheet("font-size:14px; font-weight:800; color:#2C3E50;")
        info_layout.addWidget(self.info_title)

        self.info_hint = QLabel(self.info_card)
        self.info_hint.setWordWrap(True)
        self.info_hint.setStyleSheet("font-size:13px; color:#95A5A6;")
        info_layout.addWidget(self.info_hint)

        self.sheet_prefix, self.sheet_pill, _ = self._create_pill_row(
            info_layout, "工作表：", "SheetPill", "#C3BEF0", "#312C57" 
        )
        
        self.range_prefix, self.range_pill1, self.range_sep, self.range_pill2 = self._create_double_pill_row(
            info_layout, "范    围：", "Range", "#A8E6CF", "#1A4D3A" 
        )
        
        self.cells_prefix, self.cells_pill1, self.cells_sep, self.cells_pill2 = self._create_double_pill_row(
            info_layout, "单元格：", "Cells", "#FFD3B6", "#8A3C12" 
        )

        self.path_label = QLabel(self.info_card)
        self.path_label.setWordWrap(True)
        self.path_label.setStyleSheet("font-size:11px; color:#BDC3C7; font-family:Consolas, \"Courier New\";")
        info_layout.addWidget(self.path_label)

        self._set_info_placeholder()
        root.addWidget(self.info_card, 1)

        # -- 底部高亮主按钮 --
        self.action_button = QPushButton("✨ 提取并作图", self.main_frame)
        self.action_button.setCursor(Qt.PointingHandCursor)
        self.action_button.setStyleSheet("""
            QPushButton {
                background-color: #3DC2EC;
                color: #FFFFFF;
                font-size: 14px;
                font-weight: bold;
                border: none;
                border-radius: 18px;
                padding: 10px;
            }
            QPushButton:hover { background-color: #5ED1F4; }
            QPushButton:pressed { background-color: #2BAAD4; }
            QPushButton:disabled { background-color: #D1D8E0; color: #A5B1C2; }
        """)
        self.action_button.clicked.connect(self._on_extract_plot_clicked)
        root.addWidget(self.action_button)

    def _create_pill_row(self, parent_layout, label_text, object_name, bg_color, text_color):
        row_layout = QHBoxLayout()
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.setSpacing(6)

        prefix = QLabel(label_text)
        prefix.setFixedWidth(65)
        prefix.setStyleSheet("font-size:13px; font-weight:700; color:#34495E;")
        row_layout.addWidget(prefix)

        pill = QLabel()
        pill.setAlignment(Qt.AlignCenter)
        pill.setObjectName(object_name)
        pill.setStyleSheet(f"""
            QLabel#{object_name} {{
                background-color: {bg_color};
                color: {text_color};
                font-size: 13px;
                padding: 4px 18px;
                border: 1px solid transparent;
                border-radius: 12px;
                min-height: 16px;
            }}
        """)
        row_layout.addWidget(pill, 0, Qt.AlignVCenter)

        suffix = QLabel()
        suffix.setStyleSheet("font-size:12px; font-weight:600; color:#7F8C8D;")
        row_layout.addWidget(suffix, 0, Qt.AlignVCenter)
        
        row_layout.addStretch(1)
        parent_layout.addLayout(row_layout)
        return prefix, pill, suffix

    def _create_double_pill_row(self, parent_layout, label_text, obj_name_prefix, bg_color, text_color):
        row_layout = QHBoxLayout()
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.setSpacing(6)

        prefix = QLabel(label_text)
        prefix.setFixedWidth(65)
        prefix.setStyleSheet("font-size:13px; font-weight:700; color:#34495E;")
        row_layout.addWidget(prefix)

        pill_style = f"""
            QLabel {{
                background-color: {bg_color};
                color: {text_color};
                font-size: 13px;
                padding: 4px 14px;
                border: 1px solid transparent;
                border-radius: 12px;
                min-height: 16px;
            }}
        """

        pill1 = QLabel()
        pill1.setAlignment(Qt.AlignCenter)
        pill1.setStyleSheet(pill_style)
        row_layout.addWidget(pill1, 0, Qt.AlignVCenter)

        separator = QLabel(":")
        separator.setStyleSheet("font-size:14px; font-weight:700; color:#34495E;")
        row_layout.addWidget(separator, 0, Qt.AlignVCenter)

        pill2 = QLabel()
        pill2.setAlignment(Qt.AlignCenter)
        pill2.setStyleSheet(pill_style)
        row_layout.addWidget(pill2, 0, Qt.AlignVCenter)

        row_layout.addStretch(1)
        parent_layout.addLayout(row_layout)
        return prefix, pill1, separator, pill2

    def set_status(self, text: str) -> None:
        self.status_label.setText(text)

    def _apply_pin_visual(self, pinned: bool) -> None:
        self.pin_button.setText("📌" if pinned else "📍")

    def _apply_chart_visual(self) -> None:
        text_map = {
            "box": "Box ▾",
            "scatter": "Scatter ▾",
            "multi": "Multi ▾"
        }
        self.chart_button.setText(text_map.get(self._chart_type, "图表 ▾"))

    def _set_chart_type(self, chart_type: str) -> None:
        if chart_type not in ("box", "scatter", "multi"):
            return
        self._chart_type = chart_type
        self._apply_chart_visual()

    def _set_always_on_top(self, on: bool) -> None:
        self._apply_pin_visual(on)
        self.setWindowFlag(Qt.WindowStaysOnTopHint, on)
        self.show()

    def _on_extract_plot_clicked(self) -> None:
        if self._excel_thread is not None and self._excel_thread.isRunning():
            return

        self.set_status("🚀 读取中...")
        self.action_button.setEnabled(False)
        self.action_button.setText("读取中...")

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
        self._set_info_from_meta(meta)

        render_ok = True
        try:
            self._show_chart_window(df, meta)
        except Exception as exc:
            render_ok = False
            self.set_status("作图失败 ❌")
            print(traceback.format_exc())
            
        if render_ok:
            self.set_status("就绪 🎈")
        self.action_button.setEnabled(True)
        self.action_button.setText("✨ 提取并作图")
        self._excel_thread = None
        self._excel_worker = None

    def _on_excel_fetch_failed(self, _message: str) -> None:
        try:
            print("[UI] Excel fetch failed:", _message)
        except Exception:
            pass
        self.set_status("未检测到数据 ❌")
        self.action_button.setEnabled(True)
        self.action_button.setText("✨ 提取并作图")
        self._excel_thread = None
        self._excel_worker = None

    def _set_info_placeholder(self) -> None:
        self.info_title.setText("等待框选数据...")
        self.info_hint.setText("请在 Excel 中选中数据区域，然后点击下方按钮。")

        self.sheet_prefix.setVisible(False)
        self.sheet_pill.setVisible(False)
        
        self.range_prefix.setVisible(False)
        self.range_pill1.setVisible(False)
        self.range_sep.setVisible(False)
        self.range_pill2.setVisible(False)
        
        self.cells_prefix.setVisible(False)
        self.cells_pill1.setVisible(False)
        self.cells_sep.setVisible(False)
        self.cells_pill2.setVisible(False)
        
        self.path_label.setVisible(False)

    def _set_info_from_meta(self, meta: dict) -> None:
        sheet_name = str(meta.get("sheet_name", "未知"))
        address = str(meta.get("address", "未知"))
        filepath = str(meta.get("filepath", ""))
        nrows = meta.get("nrows")
        ncols = meta.get("ncols")
        try:
            nrows_int = int(nrows) if nrows is not None else 0
            ncols_int = int(ncols) if ncols is not None else 0
        except Exception:
            nrows_int, ncols_int = 0, 0

        self.info_title.setText("📊 当前活动选区")
        self.info_hint.setText("")
        self.info_hint.setVisible(False)

        for widget in [self.sheet_prefix, self.sheet_pill, 
                       self.range_prefix, self.range_pill1, self.range_sep, self.range_pill2,
                       self.cells_prefix, self.cells_pill1, self.cells_sep, self.cells_pill2,
                       self.path_label]:
            widget.setVisible(True)

        self.sheet_pill.setText(sheet_name)

        addr_clean = address.replace("$", "")
        if ":" in addr_clean:
            start_cell, end_cell = addr_clean.split(":", 1)
            self.range_pill1.setText(start_cell)
            self.range_pill2.setText(end_cell)
        else:
            self.range_pill1.setText(addr_clean)
            self.range_pill2.setVisible(False)
            self.range_sep.setVisible(False)

        if nrows_int > 0 and ncols_int > 0:
            self.cells_pill1.setText(f"{nrows_int}行")
            self.cells_pill2.setText(f"{ncols_int}列")
        else:
            self.cells_pill1.setText("未知")
            self.cells_pill2.setVisible(False)
            self.cells_sep.setVisible(False)

        self.path_label.setText(f"{filepath if filepath else '未获取路径'}")

    def _hit_interactive_widget(self, local_pos: QPoint) -> bool:
        widget = self.childAt(local_pos)
        while widget is not None:
            if widget in (self.pin_button, self.chart_button, self.action_button):
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
        if (self._drag_active and (event.buttons() & Qt.LeftButton) and self._drag_offset is not None):
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
        matplotlib.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "SimSun", "Arial Unicode MS"]
        matplotlib.rcParams["axes.unicode_minus"] = False

        num_cols = df.shape[1]
        fig_width = max(8, num_cols * 1.5)

        chart_win = QWidget()
        chart_win.setAttribute(Qt.WA_DeleteOnClose)
        chart_win.setWindowTitle(f"分析图表 - {meta.get('book_name', '')} | 范围: {meta.get('address', '')}")
        chart_win.resize(int(fig_width * 100), 600)
        
        layout = QVBoxLayout(chart_win)
        fig = Figure(figsize=(fig_width, 6), dpi=100)
        canvas = FigureCanvas(fig)

        toolbar = NavigationToolbar(canvas, chart_win)
        layout.addWidget(toolbar)
        layout.addWidget(canvas, 1)

        try:
            if self._chart_type == "box":
                ax = fig.add_subplot(111)
                render_box_and_scatter_chart(ax, df, sheet_name=meta.get("sheet_name", "Data"))
            elif self._chart_type == "scatter":
                render_scatter_kde_chart(fig, df, sheet_name=meta.get("sheet_name", "Data"))
            elif self._chart_type == "multi":
                render_multi_scatter_kde_chart(fig, df, sheet_name=meta.get("sheet_name", "Data"))
        except Exception as exc:
            fig.clear()
            ax = fig.add_subplot(111)
            ax.text(0.5, 0.5, f"作图失败: {exc}", ha="center", va="center")
            print(f"[UI] render failed: {exc}")

        if self._chart_type == "box":
            try:
                fig.tight_layout()
            except Exception:
                pass

        chart_win.show()

        def on_destroyed():
            if chart_win in self._chart_windows:
                self._chart_windows.remove(chart_win)

        chart_win.destroyed.connect(on_destroyed)
        self._chart_windows.append(chart_win)


def main():
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    
    app = QApplication(sys.argv)
    window = FloatingToolWindow()
    window.show()
    return app.exec()

if __name__ == "__main__":
    raise SystemExit(main())