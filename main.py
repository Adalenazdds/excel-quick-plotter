import os
import sys
import traceback
import io
from io import StringIO
from typing import Optional

import pandas as pd
import xlwings as xw
from PyQt5.QtCore import QObject, QPoint, Qt, QThread, pyqtSignal
from PyQt5.QtGui import QCursor, QMouseEvent, QIcon, QImage
from PyQt5.QtWidgets import (
    QApplication,
    QCheckBox,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QMenu,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QToolTip,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

import matplotlib
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
from matplotlib.figure import Figure

try:
    from pynput import keyboard as _pynput_keyboard
except Exception:
    _pynput_keyboard = None

try:
    import keyboard as _keyboard
except Exception:
    _keyboard = None

# 确保同级目录下有这些模块文件
from box_plot import render_box_and_scatter_chart
from scatter_plot import render_scatter_kde_chart
from scatter_plot_multi import render_multi_scatter_kde_chart
from line_plot import render_line_chart
from heatmap_plot import render_heatmap_chart, coerce_numeric_matrix

try:
    import pythoncom as _pythoncom
except Exception:
    _pythoncom = None


def resource_path(relative_path: str) -> str:
    if getattr(sys, "frozen", False):
        base_path = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


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


def _parse_tabular_text_to_df(text: str) -> pd.DataFrame:
    raw = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    raw = raw.strip("\n")
    if not raw.strip():
        raise ValueError("empty clipboard")

    # Google Sheets / Excel / many web tables: tab-separated rows
    if "\t" in raw:
        sep = "\t"
    else:
        # Fallback: simple CSV
        lines = [ln for ln in raw.split("\n") if ln.strip() != ""]
        if len(lines) >= 2 and any("," in ln for ln in lines):
            sep = ","
        else:
            raise ValueError("clipboard text is not a table")

    df = pd.read_csv(
        StringIO(raw),
        sep=sep,
        header=None,
        dtype=object,
        engine="python",
        keep_default_na=True,
    )

    # Strip strings, treat empty/whitespace-only as missing.
    df = df.apply(lambda col: col.map(lambda v: v.strip() if isinstance(v, str) else v))
    df = df.replace(r"^\s*$", pd.NA, regex=True)

    # Drop fully empty rows/cols.
    df = df.dropna(axis=0, how="all").dropna(axis=1, how="all")

    if df.shape[0] == 0 or df.shape[1] == 0:
        raise ValueError("empty table")
    return df


class ClipboardFetchWorker(QObject):
    finished = pyqtSignal(object, dict)  # pandas.DataFrame, metadata dict
    failed = pyqtSignal(str)

    def __init__(self, clipboard_text: str):
        super().__init__()
        self._clipboard_text = clipboard_text

    def run(self):
        try:
            df = _parse_tabular_text_to_df(self._clipboard_text)
            nrows, ncols = int(df.shape[0]), int(df.shape[1])
            meta = {
                "book_name": "Clipboard",
                "sheet_name": "Clipboard",
                "address": f"R1C1:R{nrows}C{ncols}",
                "filepath": "clipboard://",
                "nrows": nrows,
                "ncols": ncols,
            }
            self.finished.emit(df, meta)
        except Exception as exc:
            self.failed.emit(str(exc))


class ChartDashboardWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("数据分析画板 - 多图集")
        self.resize(1150, 700)
        self.setAttribute(Qt.WA_DeleteOnClose, False)

        # 1. 创建主垂直布局，消除边距
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # 2. 引入 QScrollArea（核心修改，当图表多时提供滚动条）
        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)  # 允许内部网格自动填充宽度
        self.scroll_area.setFrameShape(QFrame.NoFrame)
        self.scroll_area.setStyleSheet("QScrollArea { background-color: #F4F6F8; }")

        # 3. 创建真正装载网格图表的容器 widget
        self.grid_container = QWidget()
        self.grid_container.setStyleSheet("QWidget { background-color: #F4F6F8; }")
        self.grid_layout = QGridLayout(self.grid_container)
        self.grid_layout.setSpacing(15)
        self.grid_layout.setContentsMargins(15, 15, 15, 15)

        # 4. 将容器放入滚动区域
        self.scroll_area.setWidget(self.grid_container)
        main_layout.addWidget(self.scroll_area)

        self.chart_count = 0
        self.max_columns = 3  # 一行最多显示 3 张图

    def add_chart(self, canvas, toolbar, meta, chart_type):
        """将生成的图表添加到网格中"""
        container = QFrame()
        # 加个卡片背景，让多图看起来更清爽
        container.setStyleSheet(
            """
            QFrame {
                background-color: white;
                border-radius: 8px;
                border: 1px solid #D1D8E0;
            }
        """
        )
        # 【关键修改】为单张图表卡片设置最小尺寸，防止被无限压缩导致图表内部文字重叠！
        container.setMinimumSize(420, 360)

        vbox = QVBoxLayout(container)
        vbox.setContentsMargins(10, 10, 10, 10)

        title_text = f"[{chart_type.upper()}] {meta.get('sheet_name', '')} | {meta.get('address', '')}"
        title_label = QLabel(title_text)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet(
            "font-weight: bold; font-size: 13px; color: #2C3E50; border: none;"
        )

        vbox.addWidget(title_label)
        vbox.addWidget(toolbar)
        vbox.addWidget(canvas, 1)

        row = self.chart_count // self.max_columns
        col = self.chart_count % self.max_columns

        self.grid_layout.addWidget(container, row, col)
        self.chart_count += 1

        self.show()
        self.raise_()
        self.activateWindow()

    def closeEvent(self, event):
        """关闭时清空画板并隐藏，以便下次提取时是全新干净的画板"""
        for i in reversed(range(self.grid_layout.count())):
            widget_to_remove = self.grid_layout.itemAt(i).widget()
            self.grid_layout.removeWidget(widget_to_remove)
            widget_to_remove.setParent(None)
        self.chart_count = 0
        self.hide()
        event.ignore()


class _HotkeyBridge(QObject):
    triggered = pyqtSignal()


class _GlobalHotkeyManager:
    def __init__(self, bridge: _HotkeyBridge, shortcut: str = "<ctrl>+q"):
        self._bridge = bridge
        self._shortcut = shortcut
        self._listener = None
        self._keyboard_hotkey = None

    @property
    def available(self) -> bool:
        return _keyboard is not None or _pynput_keyboard is not None

    def start(self) -> bool:
        if not self.available:
            return False

        if self._listener is not None:
            return True

        # Prefer `keyboard` on Windows because it can suppress the keystrokes
        # so foreground apps (e.g. Excel) won't beep / consume Alt+Key.
        if _keyboard is not None:
            try:
                def _on_activate() -> None:
                    try:
                        self._bridge.triggered.emit()
                    except Exception:
                        pass

                # `keyboard` hotkey syntax differs from pynput.
                # We use left-alt specifically to avoid conflicting with right-alt (AltGr).
                hotkey = "left alt+k" if self._shortcut == "<alt_l>+k" else "alt+k"
                self._keyboard_hotkey = _keyboard.add_hotkey(
                    hotkey,
                    _on_activate,
                    suppress=True,
                    trigger_on_release=False,
                )
                return True
            except Exception:
                self._keyboard_hotkey = None
                # Fall back to pynput.

        try:
            def _on_activate() -> None:
                try:
                    self._bridge.triggered.emit()
                except Exception:
                    pass

            self._listener = _pynput_keyboard.GlobalHotKeys({self._shortcut: _on_activate})
            self._listener.start()
            return True
        except Exception:
            self._listener = None
            return False

    def stop(self) -> None:
        if _keyboard is not None and self._keyboard_hotkey is not None:
            try:
                _keyboard.remove_hotkey(self._keyboard_hotkey)
            except Exception:
                pass
            try:
                _keyboard.unhook_all()
            except Exception:
                pass
            self._keyboard_hotkey = None

        listener = self._listener
        self._listener = None
        if listener is None:
            return
        try:
            listener.stop()
        except Exception:
            pass


class FloatingToolWindow(QWidget):

    def __init__(self):
        super().__init__()

        self._drag_active = False
        self._drag_offset = None  # type: Optional[QPoint]
        self._excel_thread = None  # type: Optional[QThread]
        self._excel_worker = None  # type: Optional[QObject]
        self._fallback_clipboard_attempted = False
        self._last_df = None
        self._chart_windows = []
        self._chart_type = "box"  # box | scatter | multi | heatmap
        self._pending_hotkey_trigger = False

        # 【新增】初始化全局单一的画板窗口
        self.dashboard_window = ChartDashboardWindow()

        self._init_window()
        self._init_ui()

    def _init_window(self) -> None:
        self.setWindowTitle("EXCEL快速分析")
        self.setWindowIcon(QIcon(resource_path("icon.ico")))
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
        top_layout.setSpacing(6)
        
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
        self.chart_button.setFixedSize(70, 32) 
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
        action_heatmap = chart_menu.addAction("Heatmap")
        action_line = chart_menu.addAction("Line Plot (多行)")
        
        action_box.triggered.connect(lambda: self._set_chart_type("box"))
        action_scatter.triggered.connect(lambda: self._set_chart_type("scatter"))
        action_multi.triggered.connect(lambda: self._set_chart_type("multi"))
        action_heatmap.triggered.connect(lambda: self._set_chart_type("heatmap"))
        action_line.triggered.connect(lambda: self._set_chart_type("line"))
        
        self.chart_button.setMenu(chart_menu)
        self._apply_chart_visual()

        # [新增] 高亮离群点开关（仅 Box Plot 显示）
        self.highlight_outliers_toggle = QCheckBox("离群点", self.top_bar)
        self.highlight_outliers_toggle.setChecked(True)
        self.highlight_outliers_toggle.setCursor(Qt.PointingHandCursor)
        # 让指示器在右侧，更像现代 Toggle Switch
        self.highlight_outliers_toggle.setLayoutDirection(Qt.RightToLeft)
        self.highlight_outliers_toggle.setStyleSheet("""
            QCheckBox {
                font-size: 12px;
                spacing: 4px;
            }
            QCheckBox::indicator {
                width: 32px;
                height: 16px;
                border-radius: 8px;
                background-color: #D1D8E0;
                image: none;
            }
            QCheckBox::indicator:checked {
                background-color: #EE884C;
                image: none;
            }
            QCheckBox::indicator:unchecked {
                background-color: #D1D8E0;
                image: none;
            }
        """)
        self.highlight_outliers_toggle.setVisible(self._chart_type == "box")

        top_layout.addWidget(self.status_label)
        top_layout.addWidget(self.chart_button)
        top_layout.addWidget(self.highlight_outliers_toggle)
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
        self.action_button.setToolTip("点击按钮或按全局快捷键 左Alt+K")
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

        self.hotkey_hint_label = QLabel("全局快捷键：左Alt+K", self.main_frame)
        self.hotkey_hint_label.setStyleSheet("font-size:11px; color:#95A5A6;")
        self.hotkey_hint_label.setAlignment(Qt.AlignCenter)
        root.addWidget(self.hotkey_hint_label)

    def _on_hotkey_triggered(self) -> None:
        # Mark this extraction as hotkey-originated so the chart window can be
        # brought to front more aggressively on Windows.
        self._pending_hotkey_trigger = True
        self._on_extract_plot_clicked()

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
            "multi": "Multi ▾",
            "heatmap": "Heatmap ▾",
            "line": "Line ▾",
        }
        self.chart_button.setText(text_map.get(self._chart_type, "图表 ▾"))

    def _set_chart_type(self, chart_type: str) -> None:
        if chart_type not in ("box", "scatter", "multi", "heatmap", "line"):
            return
        self._chart_type = chart_type
        self._apply_chart_visual()

        # [新增] 仅当选择 Box Plot 时显示开关
        try:
            self.highlight_outliers_toggle.setVisible(self._chart_type == "box")
        except Exception:
            pass

    def _set_always_on_top(self, on: bool) -> None:
        self._apply_pin_visual(on)
        self.setWindowFlag(Qt.WindowStaysOnTopHint, on)
        self.show()

    def _on_extract_plot_clicked(self) -> None:
        if self._excel_thread is not None and self._excel_thread.isRunning():
            return

        self._fallback_clipboard_attempted = False
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
        self._pending_hotkey_trigger = False

    def _on_excel_fetch_failed(self, _message: str) -> None:
        try:
            print("[UI] Excel fetch failed:", _message)
        except Exception:
            pass

        # Auto fallback: try parsing clipboard tabular data (e.g. Google Sheets Ctrl+C selection)
        if not self._fallback_clipboard_attempted:
            self._fallback_clipboard_attempted = True
            self._excel_thread = None
            self._excel_worker = None
            self._start_clipboard_fetch()
            return

        self.set_status("未检测到数据 ❌")
        self.action_button.setEnabled(True)
        self.action_button.setText("✨ 提取并作图")
        self._excel_thread = None
        self._excel_worker = None
        self._pending_hotkey_trigger = False

    def _start_clipboard_fetch(self) -> None:
        try:
            clipboard = QApplication.clipboard()
            clipboard_text = clipboard.text() if clipboard is not None else ""
        except Exception:
            clipboard_text = ""

        self.set_status("Excel 未检测到数据，尝试读取剪贴板...")

        thread = QThread(self)
        worker = ClipboardFetchWorker(clipboard_text)
        worker.moveToThread(thread)

        thread.started.connect(worker.run)
        worker.finished.connect(self._on_clipboard_fetch_success)
        worker.failed.connect(self._on_clipboard_fetch_failed)
        worker.finished.connect(thread.quit)
        worker.failed.connect(thread.quit)
        thread.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)

        self._excel_thread = thread
        self._excel_worker = worker
        thread.start()

    def _on_clipboard_fetch_success(self, df, meta) -> None:
        self._last_df = df
        self._set_info_from_meta(meta)

        render_ok = True
        try:
            self._show_chart_window(df, meta)
        except Exception:
            render_ok = False
            self.set_status("作图失败 ❌")
            print(traceback.format_exc())

        if render_ok:
            self.set_status("就绪 🎈")
        self.action_button.setEnabled(True)
        self.action_button.setText("✨ 提取并作图")
        self._excel_thread = None
        self._excel_worker = None
        self._pending_hotkey_trigger = False

    def _on_clipboard_fetch_failed(self, _message: str) -> None:
        try:
            print("[UI] Clipboard fetch failed:", _message)
        except Exception:
            pass
        self.set_status("未检测到数据 ❌")
        self.action_button.setEnabled(True)
        self.action_button.setText("✨ 提取并作图")
        self._excel_thread = None
        self._excel_worker = None
        self._pending_hotkey_trigger = False

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
            if widget in (self.pin_button, self.chart_button, self.action_button, getattr(self, "highlight_outliers_toggle", None)):
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
        matplotlib.rcParams["font.sans-serif"] = [
            "Microsoft YaHei",
            "SimHei",
            "SimSun",
            "Arial Unicode MS",
        ]
        matplotlib.rcParams["axes.unicode_minus"] = False

        # 创建基础的 Figure 和 Canvas
        fig = Figure(figsize=(5, 4), dpi=100)  # 尺寸调小，以适应网格化展示
        canvas = FigureCanvas(fig)

        # 实例化 Toolbar，parent 设为 None，稍后由画板容器接管
        toolbar = NavigationToolbar(canvas, None)

        # 追加“复制图片”按钮（复制当前图表到剪贴板）
        toolbar.addSeparator()

        def _copy_plot_to_clipboard() -> None:
            try:
                # 【核心逻辑变更】抛弃由于截取 UI 导致的分辨率过低和变形
                # 直接调用 Matplotlib 渲染出高清无损、排版原生的图像到内存
                buf = io.BytesIO()
                # bbox_inches="tight" 能自动裁剪白边，dpi=250 保证 PPT 里看极致清晰
                fig.savefig(buf, format="png", dpi=250, bbox_inches="tight", facecolor="white")
                buf.seek(0)

                # 将内存中的 PNG 二进制流转换为 QImage 塞进剪贴板
                image = QImage.fromData(buf.getvalue())
                clipboard = QApplication.clipboard()
                if clipboard is not None:
                    clipboard.setImage(image)

                try:
                    QToolTip.showText(
                        QCursor.pos(),
                        "已复制高清原图 (推荐粘贴至PPT)!",
                        toolbar,
                    )
                except Exception:
                    pass
            except Exception as e:
                try:
                    QToolTip.showText(QCursor.pos(), f"复制失败: {e}", toolbar)
                except Exception:
                    pass

        copy_action = toolbar.addAction("复制图片")
        copy_action.setToolTip("复制当前图表到剪贴板")
        copy_action.triggered.connect(_copy_plot_to_clipboard)

        try:
            if self._chart_type == "box":
                ax = fig.add_subplot(111)
                render_box_and_scatter_chart(
                    ax,
                    df,
                    sheet_name=meta.get("sheet_name", "Data"),
                    highlight_outliers=bool(self.highlight_outliers_toggle.isChecked()),
                )
                try:
                    fig.tight_layout()
                except Exception:
                    pass
            elif self._chart_type == "scatter":
                render_scatter_kde_chart(fig, df, sheet_name=meta.get("sheet_name", "Data"))
            elif self._chart_type == "multi":
                render_multi_scatter_kde_chart(fig, df, sheet_name=meta.get("sheet_name", "Data"))
            elif self._chart_type == "line":
                render_line_chart(fig, df, sheet_name=meta.get("sheet_name", "Data"))
            elif self._chart_type == "heatmap":
                render_heatmap_chart(fig, df, sheet_name=meta.get("sheet_name", "Data"))
        except Exception as exc:
            fig.clear()
            ax = fig.add_subplot(111)
            ax.text(0.5, 0.5, f"作图失败: {exc}", ha="center", va="center")
            print(f"[UI] render failed: {exc}")

        # 【核心逻辑】将生成好的 canvas 统一扔进大画板窗口里
        self.dashboard_window.add_chart(canvas, toolbar, meta, self._chart_type)


def main():
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    
    app = QApplication(sys.argv)
    window = FloatingToolWindow()

    # Global hotkey (Left Alt+K) triggers the same action as clicking the button.
    hotkey_bridge = _HotkeyBridge()
    hotkey_bridge.triggered.connect(window._on_hotkey_triggered)
    hotkey_manager = _GlobalHotkeyManager(hotkey_bridge, shortcut="<alt_l>+k")
    started = hotkey_manager.start()
    if not started:
        try:
            window.hotkey_hint_label.setText("全局快捷键不可用")
        except Exception:
            pass

    def _cleanup_hotkey():
        hotkey_manager.stop()

    app.aboutToQuit.connect(_cleanup_hotkey)
    window.show()
    return app.exec()

if __name__ == "__main__":
    raise SystemExit(main())