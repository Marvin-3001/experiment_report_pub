import re
import sys
import json
import shutil
import tempfile
import numpy as np
from pathlib import Path
from PyQt6.QtCore import Qt
from types import MethodType
from PyQt6.QtCore import QDate
from reportlab.lib import colors
from reportlab.lib.units import mm
from xml.sax.saxutils import escape
from scipy.optimize import curve_fit
from matplotlib import font_manager
from matplotlib.figure import Figure
from reportlab.lib.pagesizes import A4
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.pdfbase import pdfmetrics
from PyQt6.QtGui import QShortcut, QKeySequence, QFont
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, Image, LongTable, KeepTogether)
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget,
    QLabel, QComboBox, QPushButton, QTableWidget, QTableWidgetItem,
    QMessageBox, QHeaderView, QSizePolicy, QApplication, QSplitter,
    QFormLayout, QLineEdit, QTextEdit, QFileDialog, QScrollArea,
    QDateEdit, QGroupBox, QTabWidget, QVBoxLayout, QHBoxLayout)



IMPORT_PICTURE_PREFIX = "import_picture"

APP_STATE = {
    "picture_counter": 0,
    "generated_pictures": {},
    "picture_sources": {},

    "import_picture_counter": 0,
    "import_pictures": {},
    "import_picture_sources": {}
}

CHINESE_FONT_NAME = "STSong-Light"
WESTERN_FONT_NAME = "TimesNewRoman"

def find_times_new_roman_path():
    """
    自动查找 Times New Roman 常规字体文件。
    优先使用 Windows Fonts 文件夹；
    如果找不到，再使用 matplotlib 的字体查找机制。
    """

    candidate_paths = [
        Path(r"C:\Windows\Fonts\times.ttf"),
        Path(r"C:\Windows\Fonts\Times.ttf"),
        Path(r"C:\Windows\Fonts\TIMES.TTF"),
    ]

    windows_dir = Path(str(Path.home().anchor)) / "Windows" / "Fonts"

    candidate_paths.extend([
        windows_dir / "times.ttf",
        windows_dir / "Times.ttf",
        windows_dir / "TIMES.TTF",
    ])

    for font_path in candidate_paths:
        if font_path.exists():
            return font_path

    try:
        found_path = font_manager.findfont(
            "Times New Roman",
            fallback_to_default=False
        )

        found_path = Path(found_path)

        if found_path.exists():
            return found_path

    except (ValueError, RuntimeError, OSError):
        pass

    raise FileNotFoundError(
        "没有找到 Times New Roman 字体文件。"
        "请确认系统已安装 Times New Roman，或手动指定 TIMES_NEW_ROMAN_PATH。"
    )


TIMES_NEW_ROMAN_PATH = find_times_new_roman_path()


# =========================================================
# 读取表格中的前 n 列数字
# =========================================================
def read_numeric_columns(table, required_cols):
    """
    从表格中读取前 required_cols 列数据，并转成 numpy 数组，自动跳过整行为空的行
    """
    data = []

    for row in range(table.rowCount()):
        row_data = []
        is_empty_row = True

        for col in range(required_cols):
            item = table.item(row, col)
            text = item.text().strip() if item else ""

            if text != "":
                is_empty_row = False

            row_data.append(text)

        # 如果整行都是空的，就跳过
        if is_empty_row:
            continue

        # 尝试把这一行转成 float
        try:
            numeric_row = [float(x) for x in row_data]
        except ValueError:
            raise ValueError(f"第 {row + 1} 行存在问题，请检查。")

        data.append(numeric_row)

    if not data:
        raise ValueError("表格中没有有效数据。")

    return np.array(data, dtype=float)


# =========================================================
# 从剪贴板粘贴数据到表格
# =========================================================
def paste_from_clipboard(table):
    """
    支持从 Excel 复制多行多列后，直接粘贴到 QTableWidget。
    可通过表格属性控制：
    table.report_allow_expand_rows = True / False
    table.report_allow_expand_cols = True / False
    默认：
    允许扩展行；
    不允许扩展列。
    """

    clipboard = QApplication.clipboard()
    text = clipboard.text()

    if not text.strip():
        return

    start_row = table.currentRow()
    start_col = table.currentColumn()

    if start_row < 0:
        start_row = 0
    if start_col < 0:
        start_col = 0

    readonly_cells = getattr(table, "report_readonly_cells", set())

    allow_expand_rows = getattr(table, "report_allow_expand_rows", True)
    allow_expand_cols = getattr(table, "report_allow_expand_cols", False)

    rows = text.strip().split("\n")

    for r, row_text in enumerate(rows):
        cols = row_text.split("\t")

        for c, value in enumerate(cols):
            row = start_row + r
            col = start_col + c

            # 行超出范围
            if row >= table.rowCount():
                if allow_expand_rows:
                    old_row_count = table.rowCount()
                    table.setRowCount(row + 1)

                    for new_row in range(old_row_count, table.rowCount()):
                        for new_col in range(table.columnCount()):
                            if table.item(new_row, new_col) is None:
                                table.setItem(new_row, new_col, QTableWidgetItem(""))
                else:
                    continue

            # 列超出范围
            if col >= table.columnCount():
                if allow_expand_cols:
                    old_col_count = table.columnCount()
                    table.setColumnCount(col + 1)

                    for existing_row in range(table.rowCount()):
                        for new_col in range(old_col_count, table.columnCount()):
                            if table.item(existing_row, new_col) is None:
                                table.setItem(existing_row, new_col, QTableWidgetItem(""))
                else:
                    continue

            # 跳过只读单元格
            if (row, col) in readonly_cells:
                continue

            item = table.item(row, col)

            if item is None:
                table.setItem(row, col, QTableWidgetItem(value.strip()))
            else:
                item.setText(value.strip())


# =========================================================
# 清空选中的单元格
# =========================================================
def clear_selected_cells(table):

    readonly_cells = getattr(table, "report_readonly_cells", set())
    for item in table.selectedItems():
        row = item.row()
        col = item.column()

        if (row, col) in readonly_cells:
            continue

        item.setText("")


# =========================================================
# 清空整张表格内容
# =========================================================
def clear_all_cells(table):
    for row in range(table.rowCount()):
        for col in range(table.columnCount()):
            item = table.item(row, col)
            if item:
                item.setText("")
            else:
                table.setItem(row, col, QTableWidgetItem(""))


# =========================================================
# 配置表格：根据不同实验，设置表头、行列数、样式
# =========================================================
def setup_table(table, config):
    table.clear()
    table.setRowCount(config["rows"])
    table.setColumnCount(config["cols"])
    table.setHorizontalHeaderLabels(config["headers"])

    # 表头自动拉伸
    table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

    # 行高
    table.verticalHeader().setDefaultSectionSize(24)

    # 交替行颜色
    table.setAlternatingRowColors(True)

    # 样式表
    table.setStyleSheet(config["style"])

    # 先把每个单元格初始化为空字符串，方便后续手动输入和读取
    for row in range(table.rowCount()):
        for col in range(table.columnCount()):
            table.setItem(row, col, QTableWidgetItem(""))


# =========================================================
# 清空右侧图像
# =========================================================
def clear_plot(state):
    figure = state["figure"]
    canvas = state["canvas"]

    figure.clear()
    figure.add_subplot(111)
    canvas.draw()

# =========================================================
# 切换实验时执行
# =========================================================
def change_experiment(state, experiment_name):
    config = state["experiments"][experiment_name]

    state["current_experiment"] = experiment_name
    setup_table(state["table"], config)
    clear_plot(state)


# =========================================================
# 点击“绘图”按钮时执行
# =========================================================
def save_generated_plot_to_picture_store(state):
    """
    每次绘图后，将当前 Matplotlib 图保存为一个可在报告正文中调用的图片变量。
    规则：
    1. {picture1}、{picture2}、{picture3} ... 分别对应每次点击“绘图”后保存的图片；
    """

    APP_STATE["picture_counter"] += 1
    picture_key = f"picture{APP_STATE['picture_counter']}"

    picture_dir = Path(tempfile.gettempdir()) / "combined_experiment_tool_pictures"
    picture_dir.mkdir(parents=True, exist_ok=True)

    picture_path = picture_dir / f"{picture_key}.png"
    state["figure"].savefig(str(picture_path), dpi=300, bbox_inches="tight")

    APP_STATE["generated_pictures"][picture_key] = str(picture_path)
    APP_STATE["picture_sources"][picture_key] = str(
        state.get("current_experiment", "")
    ).strip()


def is_import_picture_placeholder_key(key):
    """
    判断 key 是否是手动导入图片变量。
    支持：
    import_picture1、import_picture2、import_picture3 ...
    """

    key = str(key).strip()
    pattern = rf"{re.escape(IMPORT_PICTURE_PREFIX)}[0-9]+"

    return re.fullmatch(pattern, key) is not None


def save_import_picture_to_picture_store(image_path):
    """
    将手动导入的单张图片复制到临时目录，并生成图片变量。
    每导入一张图片，生成一个新变量：
    {import_picture1}、{import_picture2}、{import_picture3} ...
    """

    image_path = str(image_path).strip()

    if not image_path:
        raise ValueError("导入图片路径为空。")

    image_file = Path(image_path)

    if not image_file.exists():
        raise FileNotFoundError(f"导入图片不存在：{image_file}")

    image_suffix = image_file.suffix.lower()

    if image_suffix not in [".png", ".jpg", ".jpeg", ".bmp"]:
        raise ValueError(f"不支持的图片格式：{image_suffix}")

    APP_STATE["import_picture_counter"] += 1
    image_key = f"{IMPORT_PICTURE_PREFIX}{APP_STATE['import_picture_counter']}"

    picture_dir = Path(tempfile.gettempdir()) / "combined_experiment_tool_import_pictures"
    picture_dir.mkdir(parents=True, exist_ok=True)

    target_path = picture_dir / f"{image_key}{image_suffix}"
    shutil.copy2(str(image_file), str(target_path))

    APP_STATE["import_pictures"][image_key] = str(target_path)
    APP_STATE["import_picture_sources"][image_key] = image_file.name

    return image_key


def plot_current_experiment(state):
    try:
        experiment_name = state["current_experiment"]
        if experiment_name is None:
            raise ValueError("还没有选择实验类型。")

        config = state["experiments"][experiment_name]
        plot_func = config["plot_func"]

        plot_func(state)
        save_generated_plot_to_picture_store(state)

    except Exception as e:
        QMessageBox.critical(state["window"], "错误", str(e))


# =========================================================
# 保存图片
# =========================================================
def save_current_figure(state):
    file_path, _ = QFileDialog.getSaveFileName(
        state["window"],
        "导出图片",
        "plot.png",
        "PNG Files (*.png);;JPG Files (*.jpg);;SVG Files (*.svg);;PDF Files (*.pdf)"
    )

    if not file_path:
        return

    state["figure"].savefig(file_path, dpi=300, bbox_inches="tight")
    QMessageBox.information(state["window"], "成功", "图片已导出。")


# =========================================================
# 自动加行
# =========================================================
def auto_add_rows(table, item):
    # 只有当用户改动的是最后一行时，才检查是否扩展
    if item is None:
        return

    if item.row() != table.rowCount() - 1:
        return

    # 最后一行只要有任意一个格子非空，补 5 行
    last_row = table.rowCount() - 1
    for col in range(table.columnCount()):
        cell = table.item(last_row, col)
        if cell and cell.text().strip():
            old_row_count = table.rowCount()
            previous_block_state = table.blockSignals(True)
            try:
                table.setRowCount(old_row_count + 5)

                for row in range(old_row_count, old_row_count + 5):
                    for c in range(table.columnCount()):
                        if table.item(row, c) is None:
                            table.setItem(row, c, QTableWidgetItem(""))
            finally:
                table.blockSignals(previous_block_state)
            break


# =========================================================
# Enter 换行，Tab 换列功能
# =========================================================
def enable_excel_navigation(table):
    old_key_press_event = table.keyPressEvent

    def move_to_cell(row, col):
        # Tab 到最后一列后，跳到下一行第一列
        if col >= table.columnCount():
            col = 0
            row += 1

        # Shift + Tab 到第一列前，跳到上一行最后一列
        if col < 0:
            row -= 1
            col = table.columnCount() - 1

        # 不允许跳到第 0 行之前
        if row < 0:
            row = 0

        # 如果超过最后一行，根据表格属性决定是否自动加行
        if row >= table.rowCount():
            allow_expand_rows = getattr(table, "report_allow_expand_rows", True)

            if allow_expand_rows:
                old_row_count = table.rowCount()
                previous_block_state = table.blockSignals(True)
                try:
                    table.setRowCount(old_row_count + 5)

                    for r in range(old_row_count, old_row_count + 5):
                        for c in range(table.columnCount()):
                            if table.item(r, c) is None:
                                table.setItem(r, c, QTableWidgetItem(""))
                finally:
                    table.blockSignals(previous_block_state)
            else:
                row = table.rowCount() - 1

        # 如果目标格子没有 item，就补一个空 item
        if table.item(row, col) is None:
            table.setItem(row, col, QTableWidgetItem(""))

        table.setCurrentCell(row, col)

    def new_key_press_event(_self, event):
        key = event.key()
        row = table.currentRow()
        col = table.currentColumn()

        # 回车：跳到下一行同一列
        if key in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            if event.modifiers() & Qt.KeyboardModifier.ShiftModifier:
                move_to_cell(row - 1, col)
            else:
                move_to_cell(row + 1, col)

            event.accept()
            return

        # Tab：跳到下一列
        if key == Qt.Key.Key_Tab:
            move_to_cell(row, col + 1)
            event.accept()
            return

        # Shift + Tab：跳到上一列
        if key == Qt.Key.Key_Backtab:
            move_to_cell(row, col - 1)
            event.accept()
            return

        old_key_press_event(event)

    table.keyPressEvent = MethodType(new_key_press_event, table)
    table.setTabKeyNavigation(False)


# =========================================================
# 1：离子选择性电极测定氟离子
# =========================================================
def plot_calibration1(state):
    table = state["table"]
    figure = state["figure"]
    canvas = state["canvas"]

    arr = read_numeric_columns(table, 2)
    x = arr[:, 0]
    y = arr[:, 1]

    if len(x) < 2:
        raise ValueError("标准曲线至少需要 2 个数据点。")

    # 一次线性拟合
    slope, intercept = np.polyfit(x, y, 1)
    y_fit = slope * x + intercept

    # 计算 R^2
    ss_res = np.sum((y - y_fit) ** 2)
    ss_tot = np.sum((y - np.mean(y)) ** 2)
    r2 = 1 - ss_res / ss_tot if ss_tot != 0 else 1.0

    sort_idx = np.argsort(x)
    x_sorted = x[sort_idx]
    y_fit_sorted = y_fit[sort_idx]

    figure.clear()
    ax = figure.add_subplot(111)

    ax.scatter(x, y, s=40, label="Data")
    ax.plot(x_sorted, y_fit_sorted, linewidth=1.5, label="Fit", linestyle="--")

    ax.set_title("Calibration Curve")
    ax.set_xlabel("lgC F-")
    ax.set_ylabel("Voltage")
    ax.grid(True, alpha=0.3)
    ax.legend()

    eq_text = f"y = {slope:.4f}x + {intercept:.4f}\nR² = {r2:.4f}"
    ax.text(
        0.05, 0.95,
        eq_text,
        transform=ax.transAxes,
        verticalalignment="top",
        bbox=dict(boxstyle="round", facecolor="white", alpha=0.8)
    )

    canvas.draw()

# =========================================================
# 2: 甲苯，萘的高效液相色谱分析
# =========================================================
def plot_calibration2(state):
    table = state["table"]
    figure = state["figure"]
    canvas = state["canvas"]

    arr = read_numeric_columns(table, 2)
    x = arr[:, 0]
    y = arr[:, 1]

    if len(x) < 2:
        raise ValueError("标准曲线至少需要 2 个数据点。")

    # 一次线性拟合
    slope, intercept = np.polyfit(x, y, 1)
    y_fit = slope * x + intercept

    # 计算 R^2
    ss_res = np.sum((y - y_fit) ** 2)
    ss_tot = np.sum((y - np.mean(y)) ** 2)
    r2 = 1 - ss_res / ss_tot if ss_tot != 0 else 1.0

    sort_idx = np.argsort(x)
    x_sorted = x[sort_idx]
    y_fit_sorted = y_fit[sort_idx]

    figure.clear()
    ax = figure.add_subplot(111)

    ax.scatter(x, y, s=40, label="Data")
    ax.plot(x_sorted, y_fit_sorted, linewidth=1.5, label="Fit", linestyle="--")

    ax.set_title("Calibration Curve")
    ax.set_xlabel("Concentration")
    ax.set_ylabel("Response")
    ax.grid(True, alpha=0.3)
    ax.legend()

    eq_text = f"y = {slope:.4f}x + {intercept:.4f}\nR² = {r2:.4f}"
    ax.text(
        0.05, 0.95,
        eq_text,
        transform=ax.transAxes,
        verticalalignment="top",
        bbox=dict(boxstyle="round", facecolor="white", alpha=0.8)
    )

    canvas.draw()


# =========================================================
# 3: 气相色谱流动相速度对柱效的影响
# =========================================================
def plot_point(state):
    table = state["table"]
    figure = state["figure"]
    canvas = state["canvas"]

    arr = read_numeric_columns(table, 2)
    x = arr[:, 0]
    y = arr[:, 1]

    if len(x) < 3:
        raise ValueError("数据点不足")

    if np.any(x <= 0):
        raise ValueError("流速 u 不能为 0 或负数")

    def model_func(u, A, B, C):
        return A + B / u + C * u

    popt, pcov= curve_fit(model_func, x, y)
    A_fit, B_fit, C_fit = popt

    X_fit = np.linspace(min(x), max(x), 200)
    Y_fit = model_func(X_fit, *popt)

    y_pred = model_func(x, *popt)
    ss_res = np.sum((y - y_pred) ** 2)
    ss_tot = np.sum((y - np.mean(y)) ** 2)
    r2 = 1 - ss_res / ss_tot if ss_tot != 0 else 1.0

    figure.clear()
    ax = figure.add_subplot(111)

    ax.scatter(x, y, color="blue", label="Data")

    ax.plot(X_fit, Y_fit, label="Fit", color='black', linewidth=1, linestyle="--")

    ax.text(
        0.10, 0.95,
        f"H = {A_fit:.3f} + {B_fit:.3f}/u + {C_fit:.3f}u\nR² = {r2:.4f}",
        transform=ax.transAxes,
        verticalalignment="top",
        bbox=dict(boxstyle="round", facecolor="white", alpha=0.8)
    )

    ax.set_title("Point Plot")
    ax.set_xlabel("u")
    ax.set_ylabel("H")
    ax.grid(True, alpha=0.3)

    canvas.draw()


# =========================================================
# 4: ICP-OES的多元素同时测定, 5: 火焰原子吸收测定水样中的钾, 6: 原子荧光法测定天然水中砷和汞, 7: 吖啶橙荧光探针法测定DNA
# =========================================================
def plot_calibration3(state):
    table = state["table"]
    figure = state["figure"]
    canvas = state["canvas"]

    arr = read_numeric_columns(table, 2)

    x = arr[:, 0]
    y = arr[:, 1]

    if len(x) < 2:
        raise ValueError("标准曲线至少需要 2 个数据点。")

    # 一次线性拟合
    slope, intercept = np.polyfit(x, y, 1)
    y_fit = slope * x + intercept

    # 计算 R^2
    ss_res = np.sum((y - y_fit) ** 2)
    ss_tot = np.sum((y - np.mean(y)) ** 2)
    r2 = 1 - ss_res / ss_tot if ss_tot != 0 else 1.0

    sort_idx = np.argsort(x)
    x_sorted = x[sort_idx]
    y_fit_sorted = y_fit[sort_idx]

    figure.clear()
    ax = figure.add_subplot(111)

    ax.scatter(x, y, s=40, label="Data")
    ax.plot(x_sorted, y_fit_sorted, linewidth=1.5, label="Fit", linestyle="--")

    ax.set_title("Calibration Curve")
    ax.set_xlabel("Concentration")
    ax.set_ylabel("Signal")
    ax.grid(True, alpha=0.3)
    ax.legend()

    eq_text = f"y = {slope:.4f}x + {intercept:.4f}\nR² = {r2:.4f}"
    ax.text(
        0.05, 0.95,
        eq_text,
        transform=ax.transAxes,
        verticalalignment="top",
        bbox=dict(boxstyle="round", facecolor="white", alpha=0.8)
    )

    canvas.draw()


# =========================================================
# 7: 吖啶橙荧光探针法测定DNA-光谱图
# =========================================================
def plot_line(state):
    table = state["table"]
    figure = state["figure"]
    canvas = state["canvas"]

    arr = read_numeric_columns(table, 2)

    x = arr[:, 0]
    y = arr[:, 1]

    figure.clear()
    ax = figure.add_subplot(111)
    ax.plot(x, y, color="black", linewidth=1, )

    ax.set_title("Spectrum Plot")
    ax.set_xlabel("Wavelength")
    ax.set_ylabel("Intensity")

    canvas.draw()


PLOT_FUNCTIONS = {
    "calibration1": plot_calibration1,
    "calibration2": plot_calibration2,
    "point": plot_point,
    "calibration3": plot_calibration3,
    "line": plot_line
}


def build_qtable_style(style_config):
    """
    根据 JSON 中的 style 配置生成 QTableWidget 样式。
    """

    gridline_color = style_config.get("gridline_color", "#CFCFCF")
    font_size_px = int(style_config.get("font_size_px", 13))
    alternate_background_color = style_config.get("alternate_background_color", "#F7F9FC")
    background = style_config.get("background", "white")
    item_color = style_config.get("item_color", "black")
    header_background_color = style_config.get("header_background_color", "#006CBF")

    return f"""
        QTableWidget {{
            gridline-color: {gridline_color};
            font-size: {font_size_px}px;
            alternate-background-color: {alternate_background_color};
            background: {background};
        }}
        QTableWidget::item {{
            color: {item_color};
        }}
        QHeaderView::section {{
            background-color: {header_background_color};
            padding: 4px;
            border: 1px solid #D5DDE8;
            font-weight: bold;
        }}
    """


def load_single_plot_config(json_path):
    """
    读取单个绘图配置 JSON。
    """

    with open(json_path, "r", encoding="utf-8") as file:
        return json.load(file)


def load_plot_experiment_configs(config_dir):
    """
    从 templates/plot_configs 文件夹读取绘图实验配置。
    返回格式与原来的 experiments 字典一致。
    """

    config_dir = Path(config_dir)

    if not config_dir.exists():
        raise FileNotFoundError(f"绘图配置文件夹不存在：{config_dir}")

    json_files = sorted(config_dir.glob("*.json"))

    if not json_files:
        raise FileNotFoundError(f"绘图配置文件夹中没有找到 JSON 文件：{config_dir}")

    experiments = {}

    for json_path in json_files:
        config_data = load_single_plot_config(json_path)
        styles = config_data.get("styles", {})

        for experiment_config in config_data.get("experiments", []):
            name = str(experiment_config.get("name", "")).strip()

            if not name:
                raise ValueError(f"{json_path} 中存在没有 name 的实验配置。")

            if name in experiments:
                raise ValueError(f"绘图实验名称重复：{name}")

            plot_type = str(experiment_config.get("plot_type", "")).strip()

            if plot_type not in PLOT_FUNCTIONS:
                raise ValueError(
                    f"未知 plot_type：{plot_type}。"
                    f"可用类型：{', '.join(PLOT_FUNCTIONS.keys())}"
                )

            headers = experiment_config.get("headers", [])
            rows = int(experiment_config.get("rows", 20))
            cols = int(experiment_config.get("cols", len(headers)))

            if not headers:
                raise ValueError(f"{name} 缺少 headers。")

            if len(headers) != cols:
                raise ValueError(
                    f"{name} 的 headers 数量与 cols 不一致："
                    f"headers={len(headers)}, cols={cols}"
                )

            style_ref = experiment_config.get("style", {})

            if isinstance(style_ref, str):
                style_config = styles.get(style_ref)

                if style_config is None:
                    raise ValueError(f"{name} 引用了不存在的 style：{style_ref}")

            elif isinstance(style_ref, dict):
                style_config = style_ref

            else:
                style_config = {}

            experiments[name] = {
                "headers": headers,
                "rows": rows,
                "cols": cols,
                "style": build_qtable_style(style_config),
                "plot_func": PLOT_FUNCTIONS[plot_type]
            }

    return experiments


# =========================================================
# 主函数：创建整个界面
# =========================================================
def build_plot_window():

    # 创建主窗口
    window = QMainWindow()
    window.setWindowTitle("实验数据绘图工具")
    window.resize(1200, 700)

    # 中央区域
    central = QWidget()
    window.setCentralWidget(central)

    # 总布局：上下分栏
    root_layout = QVBoxLayout(central)

    # 顶部控制区
    control_panel = QWidget()
    control_panel.setMinimumHeight(60)
    control_layout = QHBoxLayout(control_panel)
    root_layout.addWidget(control_panel, 0)

    label = QLabel("实验类型：")
    control_layout.addWidget(label)

    combo = QComboBox()
    control_layout.addWidget(combo)

    btn_plot = QPushButton("绘图")
    control_layout.addWidget(btn_plot)

    btn_clear = QPushButton("清空表格")
    control_layout.addWidget(btn_clear)

    btn_save = QPushButton("导出图片")
    control_layout.addWidget(btn_save)

    control_layout.addStretch()

    # 下方表格区和图片区
    splitter = QSplitter(Qt.Orientation.Horizontal)
    root_layout.addWidget(splitter, 1)

    # 左侧表格区
    left_panel = QWidget()
    left_layout = QVBoxLayout(left_panel)

    # 可选：限制左边不要太宽
    left_panel.setMaximumWidth(420)

    # 右侧图像区
    right_panel = QWidget()
    right_layout = QVBoxLayout(right_panel)

    splitter.addWidget(left_panel)
    splitter.addWidget(right_panel)

    # 初始宽度比例
    splitter.setStretchFactor(0, 1)
    splitter.setStretchFactor(1, 3)

    # 初始像素宽度
    splitter.setSizes([320, 880])

    # 左侧表格
    table = QTableWidget()

    # 绘图区表格：允许粘贴时增加行，不允许增加列
    table.report_allow_expand_rows = True
    table.report_allow_expand_cols = False

    enable_excel_navigation(table)

    # 允许表格被压窄一点
    table.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Expanding)

    left_layout.addWidget(table)

    # 给表格设置快捷键：Ctrl+V 粘贴，Delete 删除
    shortcut_paste = QShortcut(QKeySequence("Ctrl+V"), table)
    shortcut_delete = QShortcut(QKeySequence("Delete"), table)

    shortcut_paste.activated.connect(lambda: paste_from_clipboard(table))
    shortcut_delete.activated.connect(lambda: clear_selected_cells(table))

    # 右侧图像区域
    right_layout.addWidget(QLabel("图像预览："))

    figure = Figure(figsize=(8, 6), dpi=100)
    canvas = FigureCanvas(figure)
    canvas.setSizePolicy(
        QSizePolicy.Policy.Expanding,
        QSizePolicy.Policy.Expanding
    )
    right_layout.addWidget(canvas)

    # 放置空坐标轴
    figure.add_subplot(111)

    # 实验的配置
    experiments = load_plot_experiment_configs(PLOT_CONFIG_DIR)

    # 状态字典
    state = {
        "window": window,
        "table": table,
        "combo": combo,
        "figure": figure,
        "canvas": canvas,
        "experiments": experiments,
        "current_experiment": None,
    }

    # 下拉框加入实验名称
    combo.addItems(experiments.keys())

    # 连接信号
    combo.currentTextChanged.connect(lambda text: change_experiment(state, text))
    btn_plot.clicked.connect(lambda: plot_current_experiment(state))
    btn_clear.clicked.connect(lambda: clear_all_cells(table))
    btn_save.clicked.connect(lambda: save_current_figure(state))
    table.itemChanged.connect(lambda item: auto_add_rows(table, item))

    # 启动时，默认加载第一个实验
    first_experiment_name = combo.currentText()
    change_experiment(state, first_experiment_name)

    return window


# ============================================================
# ============================================================
# 报告生成模块
# ============================================================
# 1. 实验报告模板
# ============================================================
# 设计逻辑：
# - fields：定义 GUI 中需要用户填写的内容
# - sections：定义 PDF 中的预设正文
# - 正文中的 {{ 字段名 }} 会被替换成用户填写的内容
# - 用户未填写的内容会显示为 ________
# ============================================================

def load_single_template(json_path):
    """
    读取单个 JSON 模板文件。
    参数：
    json_path：模板文件路径
    返回：
    template_name：模板显示名称
    template_data：模板完整数据
    """

    with open(json_path, "r", encoding="utf-8") as file:
        template_data = json.load(file)

    template_name = template_data.get("template_name")

    if not template_name:
        template_name = json_path.stem

    return template_name, template_data


def load_report_templates(template_dir: str | Path = "templates") -> dict:
    """
    从 templates 文件夹读取所有 JSON 报告模板。
    返回格式：
    {
        "通用实验报告": {...},
        "质谱实验报告": {...}
    }
    """

    template_dir = Path(template_dir)

    if not template_dir.exists():
        raise FileNotFoundError(f"模板文件夹不存在：{template_dir}")

    templates = {}

    json_files = sorted(template_dir.glob("*.json"))

    if not json_files:
        raise FileNotFoundError(f"模板文件夹中没有找到 JSON 模板：{template_dir}")

    for json_path in json_files:
        template_name, template_data = load_single_template(json_path)
        templates[template_name] = template_data

    return templates


def is_result_section_heading(heading):
    """
    判断当前章节是否为实验结果章节。
    """

    heading_text = str(heading)
    return "实验结果" in heading_text or "结果" in heading_text


TEMPLATE_DIR = Path(__file__).resolve().parent / "templates"
PRESET_IMAGE_DIR = TEMPLATE_DIR / "image"
PLOT_CONFIG_DIR = TEMPLATE_DIR / "plot_configs"

REPORT_TEMPLATES = load_report_templates(TEMPLATE_DIR)


# ============================================================
# 2. 字体与文本处理函数
# ============================================================

def register_chinese_font():
    """
    注册 PDF 所需字体：
    1. 中文默认字体：STSong-Light；
    2. 西文、数字、单位符号、µ：Times New Roman。
    """

    pdfmetrics.registerFont(UnicodeCIDFont(CHINESE_FONT_NAME))

    if not TIMES_NEW_ROMAN_PATH.exists():
        raise FileNotFoundError(
            f"没有找到 Times New Roman 字体文件：{TIMES_NEW_ROMAN_PATH}\n"
            "请检查该路径是否存在，或把 TIMES_NEW_ROMAN_PATH 改为你电脑上的 times.ttf 实际路径。"
        )

    pdfmetrics.registerFont(
        TTFont(WESTERN_FONT_NAME, str(TIMES_NEW_ROMAN_PATH))
    )

    return CHINESE_FONT_NAME


def is_western_text_char(char):
    """
    判断字符是否应使用 Times New Roman。

    包括：
    1. 英文字母；
    2. 数字；
    3. 常见 ASCII 符号；
    4. µ；
    5. 常见单位/数学符号。
    """

    if char in "µμ×°℃℉±−‰λΛΔαβγδθπσΩω":
        return True

    code_point = ord(char)

    # ASCII：英文字母、数字、英文标点等
    if 0x0021 <= code_point <= 0x007E:
        return True

    # Latin-1 Supplement：包括 µ、°、± 等
    if 0x00A0 <= code_point <= 0x00FF:
        return True

    return False


def escape_text_with_mixed_fonts(text):
    """
    转义文本，同时把西文、数字和 µ 等字符包进 Times New Roman。
    中文字符保持段落默认字体 STSong-Light。
    """

    result = []
    western_run = []

    def flush_western_run():
        if western_run:
            western_text = "".join(western_run)
            result.append(
                f'<font name="{WESTERN_FONT_NAME}">{escape(western_text)}</font>'
            )
            western_run.clear()

    for char in str(text):
        if is_western_text_char(char):
            western_run.append(char)
        else:
            flush_western_run()
            result.append(escape(char))

    flush_western_run()

    return "".join(result)


def safe_paragraph_text(text):
    """
    将普通文本转换为 ReportLab Paragraph 可识别的安全文本。
    支持：
    1. 中文默认使用 STSong-Light；
    2. 西文、数字、µ、λ、Δ 等自动切换为 Times New Roman；
    3. 上标：[sup=-3]；
    4. 下标：[sub=pc]；
    5. 换行符转换为 <br/>。
    """

    if text is None:
        text = ""

    text = str(text)

    inline_pattern = r"\[(sup|sub)=([^\]]+)]"

    parts = []
    last_index = 0

    for match in re.finditer(inline_pattern, text):
        before_text = text[last_index:match.start()]
        marker_type = match.group(1)
        marker_text = match.group(2)

        parts.append(escape_text_with_mixed_fonts(before_text))

        if marker_type == "sup":
            parts.append(
                f"<super>{escape_text_with_mixed_fonts(marker_text)}</super>"
            )

        elif marker_type == "sub":
            parts.append(
                f"<sub>{escape_text_with_mixed_fonts(marker_text)}</sub>"
            )

        last_index = match.end()

    parts.append(escape_text_with_mixed_fonts(text[last_index:]))

    rendered_text = "".join(parts)
    rendered_text = rendered_text.replace("\n", "<br/>")

    return rendered_text


def remove_extra_blank_lines(text):
    """
    清理模板文本中多余的空行。
    作用：
    让写在三引号中的模板文本排版更自然，避免 PDF 里出现太多空白。
    """

    if text is None:
        return ""

    lines = text.strip().splitlines()
    cleaned_lines = []

    for line in lines:
        cleaned_lines.append(line.strip())

    return "\n".join(cleaned_lines)


def render_template_text(template_text, data, blank_text="________"):
    """
    将模板正文中的 {{ key }} 替换成用户输入的内容。
    """

    template_text = remove_extra_blank_lines(template_text)

    # 匹配 {{}} 占位符
    pattern = r"\{\{\s*([a-zA-Z_][a-zA-Z0-9_]*)\s*\}\}"

    def replace_match(match):
        key = match.group(1)

        if is_picture_placeholder_key(key):
            return match.group(0)

        value = data.get(key, "")

        if value is None:
            return blank_text

        value = str(value).strip()

        if value == "":
            return blank_text

        return value

    rendered_text = re.sub(pattern, replace_match, template_text)

    return rendered_text


def get_field_label_by_key(template, key):
    """
    根据字段 key 获取用户界面中的中文标签。
    """

    for field in template.get("fields", []):
        if field.get("key") == key:
            return field.get("label", key)

    return key


# ============================================================
# 3. PDF 生成函数
# ============================================================

def create_pdf_styles(font_name):
    """
    创建 PDF 中用到的文字样式。
    返回：
    title_style：标题样式
    heading_style：章节标题样式
    body_style：正文样式
    normal_style：表格普通文字样式
    """

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        name="ChineseTitle",
        parent=styles["Title"],
        fontName=font_name,
        fontSize=20,
        leading=28,
        alignment=1,
        spaceAfter=12
    )

    heading_style = ParagraphStyle(
        name="ChineseHeading",
        parent=styles["Heading2"],
        fontName=font_name,
        fontSize=13,
        leading=20,
        spaceBefore=10,
        spaceAfter=6
    )

    # noinspection PyTypeChecker
    body_style = ParagraphStyle(
        name="ChineseBody",
        parent=styles["BodyText"],
        fontName=font_name,
        fontSize=10.5,
        leading=18,
        firstLineIndent=0,
        alignment=TA_LEFT,
        wordWrap="CJK",
        spaceAfter=8
    )

    normal_style = ParagraphStyle(
        name="ChineseNormal",
        parent=styles["BodyText"],
        fontName=font_name,
        fontSize=10.5,
        leading=16
    )

    return title_style, heading_style, body_style, normal_style

def create_body_style_with_indent(base_style, indent_value):
    """
    基于正文样式复制一个新的样式，只修改首行缩进。
    0  = 不缩进
    21 = 约两个中文字符
    42 = 约四个中文字符
    """
    return ParagraphStyle(
        name=f"{base_style.name}_indent_{indent_value}",
        parent=base_style,
        firstLineIndent=indent_value,
        alignment=base_style.alignment
    )


def build_basic_info_table(template, data, normal_style, font_name):
    """
    生成报告顶部的基本信息表。
    表格字段由模板中的 basic_info_keys 控制。
    """

    table_rows = []

    for key in template.get("basic_info_keys", []):
        label = get_field_label_by_key(template, key)
        value = data.get(key, "")

        if not str(value).strip():
            value = "________"

        table_rows.append([
            Paragraph(safe_paragraph_text(label), normal_style),
            Paragraph(safe_paragraph_text(value), normal_style)
        ])

    info_table = Table(
        table_rows,
        colWidths=[35 * mm, 115 * mm]
    )

    info_table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), font_name),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("BACKGROUND", (0, 0), (0, -1), colors.whitesmoke),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))

    return info_table


def is_picture_placeholder_key(key):
    """
    判断一个占位符 key 是否是图片占位符。
    """

    key = str(key).strip()

    if re.fullmatch(r"picture[0-9]+", key) is not None:
        return True

    if is_import_picture_placeholder_key(key):
        return True

    return False


def get_image_path_by_placeholder_key(key):
    """
    根据图片占位符 key 获取图片路径。
    """

    key = str(key).strip()

    if key in APP_STATE["generated_pictures"]:
        return APP_STATE["generated_pictures"][key]

    if key in APP_STATE["import_pictures"]:
        return APP_STATE["import_pictures"][key]

    return ""


def get_preset_image_path(file_name):
    """
    从 templates/image 文件夹读取预设图片。
    只允许写文件名，不允许写 ../ 或子路径，避免路径混乱。
    """

    file_name = str(file_name).strip()

    if not file_name:
        raise ValueError("preset_image 缺少 file。")

    image_name = Path(file_name).name

    if image_name != file_name:
        raise ValueError("preset_image 的 file 只能写图片文件名，不能写路径。")

    image_path = PRESET_IMAGE_DIR / image_name

    if not image_path.exists():
        raise FileNotFoundError(f"预设图片不存在：{image_path}")

    return str(image_path)


def append_image_to_story(story, image_path, placeholder_text="图片占位符"):
    """
    将图片作为独立对象插入内容流，并设置为居中。
    """

    if not str(image_path).strip():
        raise ValueError(f"{placeholder_text} 没有对应图片。")

    image_file = Path(image_path)

    if not image_file.exists():
        raise FileNotFoundError(f"图片不存在：{image_file}")

    result_image = Image(str(image_file))
    result_image.hAlign = "CENTER"

    max_width = 150 * mm
    max_height = 95 * mm

    width_ratio = max_width / result_image.imageWidth
    height_ratio = max_height / result_image.imageHeight
    scale_ratio = min(width_ratio, height_ratio, 1)

    result_image.drawWidth = result_image.imageWidth * scale_ratio
    result_image.drawHeight = result_image.imageHeight * scale_ratio

    story.append(Spacer(1, 6))
    story.append(result_image)
    story.append(Spacer(1, 8))


def append_preset_image_to_story(story, image_config, normal_style):
    """
    将 templates/image 中的预设图片插入 PDF。
    """

    file_name = image_config.get("file", "")
    image_path = get_preset_image_path(file_name)

    image_title = str(image_config.get("title", "")).strip()
    max_width_mm = float(image_config.get("max_width_mm", 150))
    max_height_mm = float(image_config.get("max_height_mm", 95))
    keep_together = bool(image_config.get("keep_together", False))

    image_elements = []

    image_file = Path(image_path)

    result_image = Image(str(image_file))
    result_image.hAlign = "CENTER"

    max_width = max_width_mm * mm
    max_height = max_height_mm * mm

    width_ratio = max_width / result_image.imageWidth
    height_ratio = max_height / result_image.imageHeight
    scale_ratio = min(width_ratio, height_ratio, 1)

    result_image.drawWidth = result_image.imageWidth * scale_ratio
    result_image.drawHeight = result_image.imageHeight * scale_ratio

    image_elements.append(Spacer(1, 6))
    image_elements.append(result_image)

    if image_title:
        # noinspection PyTypeChecker
        image_title_style = ParagraphStyle(
            name=f"{normal_style.name}_preset_image_title_center",
            parent=normal_style,
            alignment=TA_CENTER
        )

        image_elements.append(Spacer(1, 4))
        image_elements.append(
            Paragraph(
                safe_paragraph_text(image_title),
                image_title_style
            )
        )

    image_elements.append(Spacer(1, 8))

    if keep_together:
        story.append(KeepTogether(image_elements))
    else:
        story.extend(image_elements)


def add_text_with_manual_indent_to_story(story, text, body_style):
    """
    支持在正文中用 [indent=数字] 手动控制段落首行缩进。
    [indent=0]这一段不缩进。
    [indent=21]这一段缩进两个中文字符。
    [indent=42]这一段缩进四个中文字符。
    """

    if not text or not text.strip():
        return

    paragraphs = text.split("\n")

    for paragraph in paragraphs:
        paragraph = paragraph.strip()

        if not paragraph:
            continue

        indent_match = re.match(r"^\[indent=(\d+)](.*)$", paragraph)

        if indent_match:
            indent_value = int(indent_match.group(1))
            paragraph_text = indent_match.group(2).strip()
            paragraph_style = create_body_style_with_indent(body_style, indent_value)
        else:
            paragraph_text = paragraph
            paragraph_style = body_style

        story.append(Paragraph(safe_paragraph_text(paragraph_text), paragraph_style))


def build_three_line_table(table_config, data, normal_style, font_name):
    """
    根据 JSON 中的 three_line_table 配置生成三线表。
    """

    headers = table_config.get("headers", [])
    rows = table_config.get("rows", [])
    col_widths_mm = table_config.get("col_widths_mm", [])

    if not headers:
        raise ValueError("three_line_table 缺少 headers。")

    table_data = [headers] + rows
    col_count = len(headers)

    for row_index, row in enumerate(table_data):
        if len(row) != col_count:
            raise ValueError(
                f"three_line_table 第 {row_index + 1} 行列数不一致。"
            )

    rendered_table_data = []

    for row in table_data:
        rendered_row = []

        for cell in row:
            cell_text = render_template_text(str(cell), data)
            rendered_row.append(
                Paragraph(
                    safe_paragraph_text(cell_text),
                    normal_style
                )
            )

        rendered_table_data.append(rendered_row)

    if col_widths_mm:
        if len(col_widths_mm) != col_count:
            raise ValueError("three_line_table 的 col_widths_mm 数量必须与 headers 列数一致。")

        col_widths = [float(width) * mm for width in col_widths_mm]
    else:
        col_widths = [150 * mm / col_count] * col_count

    pdf_table = LongTable(
        rendered_table_data,
        colWidths=col_widths,
        hAlign="CENTER",
        repeatRows=1,
        splitByRow=1
    )

    pdf_table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), font_name),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),

        # 正常三线表：顶线、表头下线、底线
        ("LINEABOVE", (0, 0), (-1, 0), 1.0, colors.black),
        ("LINEBELOW", (0, 0), (-1, 0), 0.75, colors.black),
        ("LINEBELOW", (0, -1), (-1, -1), 1.0, colors.black),

        # 跨页时：续页顶部补顶线，当前页底部补底线
        ("LINEABOVE", (0, "splitfirst"), (-1, "splitfirst"), 1.0, colors.black),
        ("LINEBELOW", (0, "splitlast"), (-1, "splitlast"), 1.0, colors.black),

        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))

    return pdf_table


def append_three_line_table_to_story(story, table_config, data, normal_style, font_name):
    """
    将三线表加入 PDF 内容流。
    keep_together=True 时，短表格尽量不拆页。
    """

    table_elements = []

    table_title = table_config.get("title", "").strip()

    if table_title:
        table_title_style = ParagraphStyle(
            name=f"{normal_style.name}_table_title_center",
            parent=normal_style,
            alignment=1
        )

        title_paragraph = Paragraph(
            safe_paragraph_text(table_title),
            table_title_style
        )

        table_elements.append(Spacer(1, 6))
        table_elements.append(title_paragraph)
        table_elements.append(Spacer(1, 4))

    pdf_table = build_three_line_table(
        table_config=table_config,
        data=data,
        normal_style=normal_style,
        font_name=font_name
    )

    table_elements.append(pdf_table)
    table_elements.append(Spacer(1, 8))

    keep_together = bool(table_config.get("keep_together", False))

    if keep_together:
        story.append(KeepTogether(table_elements))
    else:
        story.extend(table_elements)


def get_field_config_by_key(template, field_key):
    """
    根据 fields 中的 key 找到完整字段配置。
    """

    for field in template.get("fields", []):
        if field.get("key") == field_key:
            return field

    raise ValueError(f"没有找到字段配置：{field_key}")


def append_grid_table_to_story(story, table_config, template, data, normal_style, font_name):
    """
    将填写区 grid_table 字段插入 PDF。

    支持：
    1. 自定义列宽；
    2. 自定义行高；
    3. 合并单元格；
    4. 固定单元格和可填写单元格混排。
    """

    field_key = str(table_config.get("field_key", "")).strip()

    if not field_key:
        raise ValueError("grid_table 缺少 field_key。")

    field_config = get_field_config_by_key(template, field_key)

    table_data = data.get(field_key)

    if not isinstance(table_data, list) or not table_data:
        raise ValueError(f"grid_table 找不到有效表格数据：{field_key}")

    row_count = len(table_data)
    col_count = len(table_data[0])

    for row_index, row in enumerate(table_data):
        if len(row) != col_count:
            raise ValueError(
                f"grid_table 的第 {row_index + 1} 行列数不一致。"
            )

    elements = []

    title = str(table_config.get("title", "")).strip()

    if title:
        # noinspection PyTypeChecker
        title_style = ParagraphStyle(
            name=f"{normal_style.name}_{field_key}_title_center",
            parent=normal_style,
            alignment=TA_CENTER
        )

        elements.append(Spacer(1, 6))
        elements.append(Paragraph(safe_paragraph_text(title), title_style))
        elements.append(Spacer(1, 4))

    rendered_table_data = []

    for row in table_data:
        rendered_row = []

        for cell in row:
            rendered_row.append(
                Paragraph(
                    safe_paragraph_text(str(cell)),
                    normal_style
                )
            )

        rendered_table_data.append(rendered_row)

    col_widths_mm = expand_dimension_values(
        table_config.get("col_widths_mm"),
        col_count,
        150 / col_count
    )

    col_widths = [width * mm for width in col_widths_mm]

    row_heights_config = table_config.get("row_heights_mm")

    if row_heights_config is None:
        row_heights = None
    else:
        row_heights_mm = expand_dimension_values(
            row_heights_config,
            row_count,
            10
        )
        row_heights = [height * mm for height in row_heights_mm]

    pdf_table = Table(
        rendered_table_data,
        colWidths=col_widths,
        rowHeights=row_heights,
        hAlign="CENTER"
    )

    table_style_commands = [
        ("FONTNAME", (0, 0), (-1, -1), font_name),
        ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]

    for span in field_config.get("spans", []):
        row = int(span["row"])
        col = int(span["col"])
        rowspan = int(span.get("rowspan", 1))
        colspan = int(span.get("colspan", 1))

        start = (col, row)
        end = (col + colspan - 1, row + rowspan - 1)

        # noinspection PyTypeChecker
        table_style_commands.append(("SPAN", start, end))

    pdf_table.setStyle(TableStyle(table_style_commands))

    elements.append(pdf_table)
    elements.append(Spacer(1, 8))

    keep_together = bool(table_config.get("keep_together", False))

    if keep_together:
        story.append(KeepTogether(elements))
    else:
        story.extend(elements)


def add_content_with_picture_placeholders_to_story(story, content, body_style):
    """
    按正文中的图片占位符顺序加入文字和图片。
    支持：
    1. {picture1}、{picture2} ...：绘图页生成的图片；
    2. {import_picture1}、{import_picture2} ...：报告页导入的图片；
    3. {{ picture1 }}、{{ import_picture1 }}：双大括号兼容写法。
    """

    image_key_pattern = rf"(picture[0-9]+|{re.escape(IMPORT_PICTURE_PREFIX)}[0-9]+)"
    placeholder_pattern = (
        rf"\{{\{{\s*{image_key_pattern}\s*\}}\}}"
        rf"|\{{\s*{image_key_pattern}\s*\}}"
    )

    last_index = 0
    has_match = False

    for match in re.finditer(placeholder_pattern, content):
        has_match = True
        before_text = content[last_index:match.start()]

        if before_text.strip():
            add_text_with_manual_indent_to_story(story, before_text, body_style)

        key = match.group(1) or match.group(2)
        image_path = get_image_path_by_placeholder_key(key)

        if not image_path:
            raise ValueError(f"没有找到图片变量：{{{key}}}。请先生成或导入对应图片。")

        append_image_to_story(story, image_path, f"{{{key}}}")
        last_index = match.end()

    remaining_text = content[last_index:]

    if remaining_text.strip():
        add_text_with_manual_indent_to_story(story, remaining_text, body_style)

    if not has_match and content.strip():
        return


def add_report_sections_to_story(story, template, data, heading_style, body_style, normal_style, font_name):
    """
    将模板中的预设章节加入 PDF 内容流。
    支持：
    1. 普通字符串段落；
    2. 图片占位符 {picture1}、{import_picture1}；
    3. 三线表对象；
    4. 预设图片对象；
    5. 填写区 grid_table 对象。
    """

    for section in template.get("sections", []):
        heading = section.get("heading", "")
        content = section.get("content", "")

        story.append(Paragraph(safe_paragraph_text(heading), heading_style))

        if isinstance(content, list):
            for content_item in content:
                if isinstance(content_item, dict):
                    content_type = content_item.get("type", "")

                    if content_type == "three_line_table":
                        append_three_line_table_to_story(
                            story=story,
                            table_config=content_item,
                            data=data,
                            normal_style=normal_style,
                            font_name=font_name
                        )

                    elif content_type == "preset_image":
                        append_preset_image_to_story(
                            story=story,
                            image_config=content_item,
                            normal_style=normal_style
                        )

                    elif content_type == "grid_table":
                        append_grid_table_to_story(
                            story=story,
                            table_config=content_item,
                            template=template,
                            data=data,
                            normal_style=normal_style,
                            font_name=font_name
                        )

                    else:
                        raise ValueError(f"未知的 sections.content 对象类型：{content_type}")

                else:
                    rendered_content = render_template_text(str(content_item), data)

                    add_content_with_picture_placeholders_to_story(
                        story=story,
                        content=rendered_content,
                        body_style=body_style
                    )

        else:
            rendered_content = render_template_text(str(content), data)

            add_content_with_picture_placeholders_to_story(
                story=story,
                content=rendered_content,
                body_style=body_style
            )

        story.append(Spacer(1, 4))


def generate_pdf(output_path, template, data):
    """
    根据模板和用户输入数据生成 A4 PDF。
    参数：
    output_path：PDF 保存路径
    template：当前选择的报告模板
    data：用户从 GUI 中输入的数据
    """

    font_name = register_chinese_font()

    title_style, heading_style, body_style, normal_style = create_pdf_styles(font_name)

    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        rightMargin=20 * mm,
        leftMargin=20 * mm,
        topMargin=18 * mm,
        bottomMargin=18 * mm
    )

    story = []

    # 报告标题
    report_title = template.get("title", "实验报告")
    story.append(Paragraph(safe_paragraph_text(report_title), title_style))
    story.append(Spacer(1, 8))

    # 基本信息表
    info_table = build_basic_info_table(template, data, normal_style, font_name)
    story.append(info_table)
    story.append(Spacer(1, 12))

    # 正文章节
    add_report_sections_to_story(
        story=story,
        template=template,
        data=data,
        heading_style=heading_style,
        body_style=body_style,
        normal_style=normal_style,
        font_name=font_name
    )

    # 生成 PDF
    doc.build(story)


def expand_dimension_values(raw_values, expected_count, default_value):
    """
    展开行高、列宽配置。

    支持：
    1. 直接写数字；
    2. 使用 {"repeat": 数量, "value": 数值} 批量重复。
    """

    if raw_values is None:
        return [default_value] * expected_count

    result = []

    for item in raw_values:
        if isinstance(item, dict):
            repeat_count = int(item["repeat"])
            value = float(item["value"])

            for _ in range(repeat_count):
                result.append(value)
        else:
            result.append(float(item))

    if len(result) != expected_count:
        raise ValueError(
            f"尺寸配置数量不正确：需要 {expected_count} 个，实际得到 {len(result)} 个。"
        )

    return result


def plain_inline_markup_for_gui(text):
    """
    仅用于 GUI 显示：
    将 [sub=...] / [sup=...] 去掉格式标记，显示为普通文本。
    例如：
    E[sub=pc]       -> Epc
    I[sub=pa]       -> Ipa
    10[sup=-3]      -> 10⁻³
    """

    if text is None:
        return ""

    text = str(text)

    superscript_map = str.maketrans({
        "0": "⁰",
        "1": "¹",
        "2": "²",
        "3": "³",
        "4": "⁴",
        "5": "⁵",
        "6": "⁶",
        "7": "⁷",
        "8": "⁸",
        "9": "⁹",
        "+": "⁺",
        "-": "⁻",
        "=": "⁼",
        "(": "⁽",
        ")": "⁾"
    })

    def replace_match(match):
        marker_type = match.group(1)
        marker_text = match.group(2)

        if marker_type == "sup":
            return marker_text.translate(superscript_map)

        if marker_type == "sub":
            return marker_text

        return marker_text

    return re.sub(r"\[(sup|sub)=([^]]+)]", replace_match, text)


def create_grid_table_input_widget(field):
    """
    创建可编辑网格表格输入控件。

    支持：
    1. 固定行列；
    2. 默认单元格文本；
    3. 只读单元格；
    4. 合并单元格；
    5. 自定义填写区列宽、行高；
    6. 从 Excel 粘贴。
    """

    rows = int(field.get("rows", 2))
    cols = int(field.get("cols", 2))

    container = QWidget()
    layout = QVBoxLayout(container)
    layout.setContentsMargins(0, 0, 0, 0)

    table = QTableWidget()
    table.setRowCount(rows)
    table.setColumnCount(cols)

    table.horizontalHeader().setVisible(False)
    table.verticalHeader().setVisible(False)

    for row in range(rows):
        for col in range(cols):
            table.setItem(row, col, QTableWidgetItem(""))

    # 设置默认单元格
    readonly_cells = set()

    for cell in field.get("cells", []):
        row = int(cell["row"])
        col = int(cell["col"])
        text = str(cell.get("text", ""))

        if row < 0 or row >= rows or col < 0 or col >= cols:
            raise ValueError(f"grid_table 默认单元格超出范围：row={row}, col={col}")

        display_text = plain_inline_markup_for_gui(text)
        item = QTableWidgetItem(display_text)

        if bool(cell.get("readonly", False)):
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            item.setData(Qt.ItemDataRole.UserRole, text)
            readonly_cells.add((row, col))

        table.setItem(row, col, item)

    # 设置合并单元格
    for span in field.get("spans", []):
        row = int(span["row"])
        col = int(span["col"])
        rowspan = int(span.get("rowspan", 1))
        colspan = int(span.get("colspan", 1))

        table.setSpan(row, col, rowspan, colspan)

        # 合并区域内，除左上角外，其余隐藏单元格都禁止编辑和粘贴
        for rr in range(row, row + rowspan):
            for cc in range(col, col + colspan):
                if (rr, cc) == (row, col):
                    continue

                readonly_cells.add((rr, cc))

                item = table.item(rr, cc)
                if item is None:
                    item = QTableWidgetItem("")
                    table.setItem(rr, cc, item)

                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)

    # 设置填写区列宽
    col_widths_px = expand_dimension_values(
        field.get("col_widths_px"),
        cols,
        70
    )

    for col, width in enumerate(col_widths_px):
        table.setColumnWidth(col, int(width))

    # 设置填写区行高
    row_heights_px = expand_dimension_values(
        field.get("row_heights_px"),
        rows,
        34
    )

    for row, height in enumerate(row_heights_px):
        table.setRowHeight(row, int(height))

    table.setMinimumHeight(min(360, sum(int(x) for x in row_heights_px) + 40))
    table.setSizePolicy(
        QSizePolicy.Policy.Expanding,
        QSizePolicy.Policy.Fixed
    )

    enable_excel_navigation(table)

    button_layout = QHBoxLayout()

    paste_button = QPushButton("粘贴 Excel 数据")
    clear_button = QPushButton("清空可编辑单元格")

    def clear_editable_cells():
        for row_t in range(table.rowCount()):
            for col_t in range(table.columnCount()):
                if (row_t, col_t) in readonly_cells:
                    continue

                item_t = table.item(row_t, col_t)

                if item_t is None:
                    table.setItem(row_t, col_t, QTableWidgetItem(""))
                else:
                    item_t.setText("")

    paste_button.clicked.connect(lambda: paste_from_clipboard(table))
    clear_button.clicked.connect(clear_editable_cells)

    button_layout.addWidget(paste_button)
    button_layout.addWidget(clear_button)
    button_layout.addStretch()

    layout.addLayout(button_layout)
    layout.addWidget(table)

    shortcut_paste = QShortcut(QKeySequence("Ctrl+V"), table)
    shortcut_delete = QShortcut(QKeySequence("Delete"), table)

    shortcut_paste.activated.connect(lambda: paste_from_clipboard(table))
    shortcut_delete.activated.connect(lambda: clear_selected_cells(table))

    table.report_readonly_cells = readonly_cells
    table.report_allow_expand_rows = False
    table.report_allow_expand_cols = False

    container.report_grid_table_widget = table
    container.report_grid_table_spans = field.get("spans", [])

    return container


# ============================================================
# 4. GUI 输入控件相关函数
# ============================================================

def create_input_widget(field):
    """
    根据字段类型创建对应的 PyQt6 输入控件。
    支持类型：
    line：单行输入框
    text：多行输入框
    date：日期选择框
    """

    field_type = field.get("type", "line")
    default_value = field.get("default", "")

    if field_type == "text":
        widget = QTextEdit()
        widget.setPlainText(default_value)
        widget.setMinimumHeight(90)
        return widget

    if field_type == "date":
        widget = QDateEdit()
        widget.setCalendarPopup(True)
        widget.setDisplayFormat("yyyy-MM-dd")

        if default_value:
            date = QDate.fromString(default_value, "yyyy-MM-dd")
            if date.isValid():
                widget.setDate(date)
            else:
                widget.setDate(QDate.currentDate())
        else:
            widget.setDate(QDate.currentDate())

        return widget

    if field_type == "grid_table":
        return create_grid_table_input_widget(field)

    if field_type == "image":
        container = QWidget()
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)

        line_edit = QLineEdit()
        line_edit.setText(default_value)
        browse_button = QPushButton("选择图片")

        def choose_image_file():
            image_path, _ = QFileDialog.getOpenFileName(
                container,
                "选择实验结果图片",
                "",
                "Image Files (*.png *.jpg *.jpeg *.bmp)"
            )

            if image_path:
                line_edit.setText(image_path)

        browse_button.clicked.connect(choose_image_file)
        container.line_edit = line_edit
        layout.addWidget(line_edit)
        layout.addWidget(browse_button)
        return container

    widget = QLineEdit()
    widget.setText(default_value)
    return widget


def get_grid_table_widget_data(table):
    """
    读取 QTableWidget 的所有单元格内容。
    如果单元格有 UserRole 原始文本，则优先使用原始文本；
    否则使用 GUI 中显示的文本。
    """

    data = []

    for row in range(table.rowCount()):
        row_data = []

        for col in range(table.columnCount()):
            item = table.item(row, col)

            if item is None:
                row_data.append("")
            else:
                raw_text = item.data(Qt.ItemDataRole.UserRole)

                if raw_text is not None:
                    row_data.append(str(raw_text).strip())
                else:
                    row_data.append(item.text().strip())

        data.append(row_data)

    return data


def get_widget_value(widget):
    """
    从不同类型的输入控件中读取用户输入内容。
    """

    if isinstance(widget, QTextEdit):
        return widget.toPlainText().strip()

    if isinstance(widget, QDateEdit):
        return widget.date().toString("yyyy-MM-dd")

    if isinstance(widget, QLineEdit):
        return widget.text().strip()

    if hasattr(widget, "line_edit"):
        return widget.line_edit.text().strip()

    if hasattr(widget, "report_grid_table_widget"):
        return get_grid_table_widget_data(widget.report_grid_table_widget)

    return ""


def clear_layout(layout):
    """
    清空布局中的所有控件。
    """

    while layout.count():
        item = layout.takeAt(0)

        widget = item.widget()
        child_layout = item.layout()

        if widget is not None:
            widget.deleteLater()

        if child_layout is not None:
            clear_layout(child_layout)


def rebuild_form(form_layout, input_widgets, template):
    """
    根据当前模板重新生成 GUI 表单。
    参数：
    form_layout：表单布局
    input_widgets：用于保存 key 与控件的对应关系
    template：当前报告模板
    """

    clear_layout(form_layout)
    input_widgets.clear()

    for field in template.get("fields", []):
        key = field.get("key")
        label = field.get("label", key)
        required = field.get("required", False)

        widget = create_input_widget(field)
        input_widgets[key] = widget

        display_label = plain_inline_markup_for_gui(label)

        if required:
            label_text = f"{display_label} *"
        else:
            label_text = display_label

        form_layout.addRow(QLabel(label_text), widget)


def is_empty_form_value(value):
    """
    判断表单值是否为空。
    支持：
    1. 普通字符串；
    2. 二维表格。
    """

    if isinstance(value, list):
        for row in value:
            for cell in row:
                if str(cell).strip():
                    return False
        return True

    return str(value).strip() == ""


def collect_form_data(template, input_widgets, check_required=True):
    """
    收集用户在 GUI 中输入的数据。
    参数：
    check_required：
        True  表示检查必填项，没有填写就报错；
        False 表示不检查必填项，可用于预览。
    """

    data = {}
    missing_fields = []

    for field in template.get("fields", []):
        key = field.get("key")
        label = field.get("label", key)
        required = field.get("required", False)

        widget = input_widgets.get(key)

        if widget is None:
            value = ""
        else:
            value = get_widget_value(widget)

        data[key] = value

        if check_required and required and is_empty_form_value(value):
            missing_fields.append(label)

    if missing_fields:
        error_message = "以下必填项尚未填写：\n" + "\n".join(missing_fields)
        raise ValueError(error_message)

    return data


def strip_manual_indent_markers(text):
    """
    仅用于预览文本：
    去掉 [indent=数字] 标记，但不影响 PDF 正文中的真实缩进。
    """

    if text is None:
        return ""

    cleaned_lines = []

    for line in str(text).splitlines():
        indent_match = re.match(r"^\s*\[indent=\d+](.*)$", line)

        if indent_match:
            cleaned_lines.append(indent_match.group(1).strip())
        else:
            cleaned_lines.append(line)

    return "\n".join(cleaned_lines)


# ============================================================
# 5. 预览文本生成函数
# ============================================================
def build_three_line_table_preview(table_config, data):
    """
    生成三线表的纯文本预览。
    这里只用于右侧预览，不影响 PDF 排版。
    """

    lines = []

    table_title = table_config.get("title", "").strip()
    headers = table_config.get("headers", [])
    rows = table_config.get("rows", [])

    if table_title:
        lines.append(f"[三线表] {plain_inline_markup_for_gui(table_title)}")

    if headers:
        lines.append(" | ".join(plain_inline_markup_for_gui(str(header)) for header in headers))
        lines.append(" | ".join("---" for _ in headers))

    for row in rows:
        rendered_cells = []

        for cell in row:
            rendered_cell = render_template_text(str(cell), data)
            rendered_cell = strip_manual_indent_markers(rendered_cell)
            rendered_cells.append(plain_inline_markup_for_gui(rendered_cell))

        lines.append(" | ".join(rendered_cells))

    return "\n".join(lines)


def get_picture_key_sort_number(picture_key):
    """
    从 picture1、picture2 这类变量名中提取数字，用于正确排序。
    """

    match = re.search(r"\d+", str(picture_key))

    if match:
        return int(match.group())

    return 0


def build_report_hint_text():
    """
    生成报告生成页右下角的提示文本。
    包括：
    1. 当前可用绘图图片变量；
    2. 当前可用导入图片变量；
    3. 常用正文标记；
    4. 图片插入提示。
    """

    lines = ["【绘图图片变量】"]

    generated_pictures = APP_STATE["generated_pictures"]
    picture_sources = APP_STATE.get("picture_sources", {})

    picture_names = sorted(
        (
            key for key in generated_pictures.keys()
            if re.fullmatch(r"picture[0-9]+", key) is not None
        ),
        key=get_picture_key_sort_number
    )

    if picture_names:
        for key in picture_names:
            source_name = str(picture_sources.get(key, "")).strip()

            if source_name:
                lines.append(f"{{{key}}}：来自“{source_name}”")
            else:
                lines.append(f"{{{key}}}：来源未知")
    else:
        lines.append("当前没有绘图图片变量。")
        lines.append("在“实验数据绘图”页点击“绘图”后，会生成 {picture1}、{picture2} 等变量。")

    lines.append("")
    lines.append("【导入图片变量】")

    import_pictures = APP_STATE["import_pictures"]
    import_picture_sources = APP_STATE.get("import_picture_sources", {})

    import_picture_names = sorted(
        (
            key for key in import_pictures.keys()
            if is_import_picture_placeholder_key(key)
        ),
        key=get_picture_key_sort_number
    )

    if import_picture_names:
        for key in import_picture_names:
            source_name = str(import_picture_sources.get(key, "")).strip()

            if source_name:
                lines.append(f"{{{key}}}：来自文件“{source_name}”")
            else:
                lines.append(f"{{{key}}}：来源未知")
    else:
        lines.append("当前没有导入图片变量。")
        lines.append("{import_picture1}、{import_picture2} 等变量会在点击“导入图片”后生成。")

    lines.append("")
    lines.append("【常用标记】")
    lines.append("下标：E[sub=pc]，PDF 中 pc 会显示为下标。")
    lines.append("上标：10[sup=-3]，PDF 中 -3 会显示为上标。")

    lines.append("")
    lines.append("【图片插入】")
    lines.append("插入绘图图片：{picture1}、{picture2} ...")
    lines.append("插入导入图片：{import_picture1}、{import_picture2} ...")

    return "\n".join(lines)


def update_report_hint(hint_edit):
    """
    刷新提示框内容。
    """

    hint_edit.setPlainText(build_report_hint_text())


def build_preview_text(template, data):
    """
    生成纯文本预览内容。
    支持：
    1. 普通文本；
    2. [indent=数字] 缩进标记隐藏；
    3. 图片占位符文本预览；
    4. three_line_table 三线表对象预览。
    """

    lines = [
        template.get("title", "实验报告"),
        "=" * 30,
        "",
        "【基本信息】"
    ]

    for key in template.get("basic_info_keys", []):
        label = get_field_label_by_key(template, key)
        value = data.get(key, "").strip()

        if not value:
            value = "________"

        lines.append(f"{label}：{value}")

    lines.append("")

    for section in template.get("sections", []):
        heading = section.get("heading", "")
        content = section.get("content", "")

        lines.append(heading)

        if isinstance(content, list):
            for content_item in content:
                if isinstance(content_item, dict):
                    content_type = content_item.get("type", "")

                    if content_type == "three_line_table":
                        lines.append(build_three_line_table_preview(content_item, data))

                    elif content_type == "preset_image":
                        image_title = str(content_item.get("title", "")).strip()
                        file_name = str(content_item.get("file", "")).strip()

                        if image_title:
                            lines.append(f"[预设图片] {image_title}：templates/image/{file_name}")
                        else:
                            lines.append(f"[预设图片] templates/image/{file_name}")

                    elif content_type == "grid_table":
                        field_key = str(content_item.get("field_key", "")).strip()
                        title = str(content_item.get("title", "")).strip()

                        if title:
                            lines.append(f"[填写表格] {title}：{field_key}")
                        else:
                            lines.append(f"[填写表格] {field_key}")

                    else:
                        lines.append(f"[未知内容类型] {content_type}")

                else:
                    rendered_content = render_template_text(str(content_item), data)
                    preview_content = strip_manual_indent_markers(rendered_content)
                    preview_content = plain_inline_markup_for_gui(preview_content)
                    lines.append(preview_content)

        else:
            rendered_content = render_template_text(str(content), data)
            preview_content = strip_manual_indent_markers(rendered_content)
            preview_content = plain_inline_markup_for_gui(preview_content)
            lines.append(preview_content)

        lines.append("")

    return "\n".join(lines)


def update_preview(template_combo, input_widgets, preview_edit):
    """
    更新右侧预览框内容。
    """

    template_name = template_combo.currentText()
    template = REPORT_TEMPLATES[template_name]

    data = collect_form_data(
        template=template,
        input_widgets=input_widgets,
        check_required=False
    )

    preview_text = build_preview_text(template, data)
    preview_edit.setPlainText(preview_text)


# ============================================================
# 6. 文件保存与 PDF 导出函数
# ============================================================

def get_default_pdf_name(data):
    """
    根据实验名称生成默认 PDF 文件名。
    """

    experiment_name = data.get("experiment_name", "").strip()

    if not experiment_name:
        experiment_name = "实验报告"

    # 去掉 Windows 文件名中不允许使用的特殊字符
    experiment_name = re.sub(r'[\\/:*?"<>|]', "_", experiment_name)

    return f"{experiment_name}.pdf"


def choose_output_pdf_path(parent_window, default_filename):
    """
    弹出文件保存对话框，让用户选择 PDF 保存位置。
    """

    output_path, _ = QFileDialog.getSaveFileName(
        parent_window,
        "保存实验报告 PDF",
        default_filename,
        "PDF Files (*.pdf)"
    )

    if not output_path:
        return ""

    if not output_path.lower().endswith(".pdf"):
        output_path += ".pdf"

    return output_path


def export_pdf(parent_window, template_combo, input_widgets):
    """
    GUI 中点击“生成 PDF”按钮后执行的函数。
    """

    try:
        template_name = template_combo.currentText()
        template = REPORT_TEMPLATES[template_name]

        # 导出 PDF 时检查必填项
        data = collect_form_data(
            template=template,
            input_widgets=input_widgets,
            check_required=True
        )

        default_filename = get_default_pdf_name(data)

        output_path = choose_output_pdf_path(parent_window, default_filename)

        if not output_path:
            return

        generate_pdf(output_path, template, data)

        QMessageBox.information(
            parent_window,
            "生成成功",
            f"PDF 已成功生成：\n{output_path}"
        )

    except Exception as error:
        QMessageBox.critical(
            parent_window,
            "生成失败",
            str(error)
        )


# ============================================================
# 7. 主窗口构建函数
# ============================================================

def build_main_window():
    """
    构建主窗口。
    整体界面结构：
    上方：模板选择 + 按钮
    左侧：用户输入表单
    右侧：报告文本预览
    """

    window = QWidget()
    window.setWindowTitle("实验报告 PDF 生成工具")
    window.resize(1100, 720)

    main_layout = QVBoxLayout(window)

    # -------------------------------
    # 顶部工具栏
    # -------------------------------
    top_layout = QHBoxLayout()
    main_layout.addLayout(top_layout)

    top_layout.addWidget(QLabel("报告模板："))

    template_combo = QComboBox()
    template_combo.addItems(REPORT_TEMPLATES.keys())
    top_layout.addWidget(template_combo)

    import_picture_button = QPushButton("导入外部图片")
    top_layout.addWidget(import_picture_button)

    preview_button = QPushButton("刷新预览")
    top_layout.addWidget(preview_button)

    export_button = QPushButton("生成 PDF")
    top_layout.addWidget(export_button)

    # -------------------------------
    # 中间主体区域：左侧表单 + 右侧预览
    # -------------------------------
    content_layout = QHBoxLayout()
    main_layout.addLayout(content_layout)

    # 左侧输入区
    form_group = QGroupBox("需要填写的内容")
    form_group_layout = QVBoxLayout(form_group)

    scroll_area = QScrollArea()
    scroll_area.setWidgetResizable(True)
    form_group_layout.addWidget(scroll_area)

    form_container = QWidget()
    form_layout = QFormLayout(form_container)
    form_layout.setSpacing(12)

    scroll_area.setWidget(form_container)

    content_layout.addWidget(form_group, stretch=1)

    # 右侧区域：上方报告预览，下方提示框
    right_splitter = QSplitter(Qt.Orientation.Vertical)

    # 右上：报告预览
    preview_group = QGroupBox("报告预览")
    preview_group_layout = QVBoxLayout(preview_group)

    preview_edit = QTextEdit()
    preview_edit.setReadOnly(True)
    preview_group_layout.addWidget(preview_edit)

    # 右下：使用提示
    hint_group = QGroupBox("使用提示")
    hint_group_layout = QVBoxLayout(hint_group)

    hint_edit = QTextEdit()
    hint_edit.setReadOnly(True)
    hint_edit.setMinimumHeight(130)
    hint_edit.setMaximumHeight(220)
    hint_group_layout.addWidget(hint_edit)

    right_splitter.addWidget(preview_group)
    right_splitter.addWidget(hint_group)

    right_splitter.setStretchFactor(0, 4)
    right_splitter.setStretchFactor(1, 1)

    content_layout.addWidget(right_splitter, stretch=1)

    # 用字典保存输入控件
    input_widgets = {}

    def refresh_current_form():
        # 刷新当前模板对应的输入表单。

        template_name = template_combo.currentText()
        template = REPORT_TEMPLATES[template_name]

        rebuild_form(form_layout, input_widgets, template)
        update_preview(template_combo, input_widgets, preview_edit)
        update_report_hint(hint_edit)

    def refresh_current_preview():
        # 刷新右侧文本预览和提示信息。

        update_preview(template_combo, input_widgets, preview_edit)
        update_report_hint(hint_edit)

    def import_pictures_from_files():
        # 从本地选择一张或多张图片，并生成 import_picture 数字变量。

        image_paths, _ = QFileDialog.getOpenFileNames(
            window,
            "导入图片",
            "",
            "Image Files (*.png *.jpg *.jpeg *.bmp)"
        )

        if not image_paths:
            return

        new_keys = []

        try:
            for image_path in image_paths:
                new_key = save_import_picture_to_picture_store(image_path)
                new_keys.append(new_key)

                # 每导入一张图片，就刷新一次提示区文本。
                update_report_hint(hint_edit)

            update_preview(template_combo, input_widgets, preview_edit)

            if new_keys:
                QMessageBox.information(
                    window,
                    "导入成功",
                    "已生成图片变量：\n" + "\n".join(f"{{{key}}}" for key in new_keys)
                )

        except Exception as error:
            QMessageBox.critical(
                window,
                "导入失败",
                str(error)
            )

    def export_current_pdf():
        # 导出当前模板对应的 PDF。

        export_pdf(window, template_combo, input_widgets)

    # 让外部窗口也能触发报告页刷新
    def refresh_report_page_from_outside():
        # 从其他标签页切换回来时，同时刷新报告预览和提示框。
        update_preview(template_combo, input_widgets, preview_edit)
        update_report_hint(hint_edit)

    window.refresh_report_page = refresh_report_page_from_outside

    # -------------------------------
    # 绑定按钮事件
    # -------------------------------
    template_combo.currentTextChanged.connect(refresh_current_form)
    import_picture_button.clicked.connect(import_pictures_from_files)
    preview_button.clicked.connect(refresh_current_preview)
    export_button.clicked.connect(export_current_pdf)

    # 初始加载表单
    refresh_current_form()

    return window


def build_integrated_window():
    """
    构建总窗口，用标签页整合绘图工具和报告生成工具。
    """

    window = QMainWindow()
    window.setWindowTitle("实验数据绘图与报告生成工具")
    window.resize(1250, 760)

    tabs = QTabWidget()
    window.setCentralWidget(tabs)

    plot_window = build_plot_window()
    report_window = build_main_window()

    plot_tab_index = tabs.addTab(plot_window, "实验数据绘图")
    report_tab_index = tabs.addTab(report_window, "实验报告生成")

    def handle_tab_changed(index):
        # 切换到“实验报告生成”页时，刷新报告预览和图片变量提示。
        if index == report_tab_index and hasattr(report_window, "refresh_report_page"):
            report_window.refresh_report_page()

    tabs.currentChanged.connect(handle_tab_changed)

    window.plot_window = plot_window
    window.report_window = report_window
    window.plot_tab_index = plot_tab_index
    window.report_tab_index = report_tab_index

    return window


def main():
    app = QApplication(sys.argv)
    window = build_integrated_window()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
