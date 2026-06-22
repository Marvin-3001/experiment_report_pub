import re
import sys
import json
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
from matplotlib.figure import Figure
from reportlab.lib.pagesizes import A4
from reportlab.lib.enums import TA_LEFT
from reportlab.pdfbase import pdfmetrics
from PyQt6.QtGui import QShortcut, QKeySequence
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



APP_STATE = {
    "picture_counter": 0,
    "generated_pictures": {}
}


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
    支持从 Excel 复制多行多列后，直接粘贴到 QTableWidget
    """
    clipboard = QApplication.clipboard()
    text = clipboard.text()

    if not text.strip():
        return

    start_row = table.currentRow()
    start_col = table.currentColumn()

    # 如果当前没有选中单元格，就默认从左上角开始
    if start_row < 0:
        start_row = 0
    if start_col < 0:
        start_col = 0

    rows = text.strip().split("\n")

    for r, row_text in enumerate(rows):
        cols = row_text.split("\t")

        for c, value in enumerate(cols):
            row = start_row + r
            col = start_col + c

            # 如果粘贴的数据超出当前表格范围，就自动扩展
            if row >= table.rowCount():
                table.setRowCount(row + 1)
            if col >= table.columnCount():
                table.setColumnCount(col + 1)

            table.setItem(row, col, QTableWidgetItem(value.strip()))


# =========================================================
# 清空选中的单元格
# =========================================================
def clear_selected_cells(table):
    for item in table.selectedItems():
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
    2. {picture} 永远指向最近一次点击“绘图”后保存的图片。
    """

    APP_STATE["picture_counter"] += 1
    picture_key = f"picture{APP_STATE['picture_counter']}"

    picture_dir = Path(tempfile.gettempdir()) / "combined_experiment_tool_pictures"
    picture_dir.mkdir(parents=True, exist_ok=True)

    picture_path = picture_dir / f"{picture_key}.png"
    state["figure"].savefig(str(picture_path), dpi=300, bbox_inches="tight")

    APP_STATE["generated_pictures"][picture_key] = str(picture_path)
    APP_STATE["generated_pictures"]["picture"] = str(picture_path)

    picture_label = state.get("picture_label")
    if picture_label is not None:
        picture_label.setText(
            f"当前图片变量：{{{picture_key}}}；最新图片也可用 {{picture}}"
        )


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

        # 如果超过最后一行，自动加 5 行
        if row >= table.rowCount():
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

    if np.any(x < 0):
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

    picture_label = QLabel("当前图片变量：无")
    control_layout.addWidget(picture_label)

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
    experiments = {
        "离子选择性电极测定氟离子": {
            "headers": ["lgC F-", "Voltage"],
            "rows": 20,
            "cols": 2,
            "style": """
                QTableWidget {
                    gridline-color: #CFCFCF;
                    font-size: 13px;
                    alternate-background-color: #F7F9FC;
                    background: white;
                }
                QTableWidget::item {
                    color: black;
                }
                QHeaderView::section {
                    background-color: #006CBF;
                    padding: 4px;
                    border: 1px solid #D5DDE8;
                    font-weight: bold;
                }
            """,
            "plot_func": plot_calibration1
        },

        "甲苯，萘的高效液相色谱分析": {
            "headers": ["Concentration", "Response"],
            "rows": 20,
            "cols": 2,
            "style": """
                QTableWidget {
                    gridline-color: #D0D0D0;
                    font-size: 13px;
                    alternate-background-color: #FFF8E8;
                    background: white;
                }
                QTableWidget::item {
                    color: black;
                }
                QHeaderView::section {
                    background-color: #127840;
                    padding: 4px;
                    border: 1px solid #D5DDE8;
                    font-weight: bold;
                }
            """,
            "plot_func": plot_calibration2
        },

        "气相色谱流动相速度对柱效的影响": {
            "headers": ["u", "H"],
            "rows": 20,
            "cols": 2,
            "style": """
                QTableWidget {
                    gridline-color: #CFCFCF;
                    font-size: 13px;
                    alternate-background-color: #F4F4F4;
                    background: white;
                }
                QTableWidget::item {
                    color: black;
                }
                QHeaderView::section {
                    background-color: #595959;
                    padding: 4px;
                    border: 1px solid #D5DDE8;
                    font-weight: bold;
                }
            """,
            "plot_func": plot_point
        },

        "ICP-OES多元素的同时测定": {
            "headers": ["Concentration", "Signal"],
            "rows": 20,
            "cols": 2,
            "style": """
                QTableWidget {
                    gridline-color: #CFCFCF;
                    font-size: 13px;
                    alternate-background-color: #F7F9FC;
                    background: white;
                }
                QTableWidget::item {
                    color: black;
                }
                QHeaderView::section {
                    background-color: #006CBF;
                    padding: 4px;
                    border: 1px solid #D5DDE8;
                    font-weight: bold;
                }
            """,
            "plot_func": plot_calibration3
        },

        "火焰原子吸收测定水样中的钾": {
            "headers": ["Concentration", "Signal"],
            "rows": 20,
            "cols": 2,
            "style": """
            QTableWidget {
                gridline-color: #D0D0D0;
                font-size: 13px;
                alternate-background-color: #FFF8E8;
                background: white;
            }
            QTableWidget::item {
                color: black;
            }
            QHeaderView::section {
                background-color: #127840;
                padding: 4px;
                border: 1px solid #D5DDE8;
                font-weight: bold;
            }
        """,
            "plot_func": plot_calibration3
        },

        "原子荧光测定天然水中的砷和汞": {
            "headers": ["Concentration", "Signal"],
            "rows": 20,
            "cols": 2,
            "style": """
                QTableWidget {
                    gridline-color: #CFCFCF;
                    font-size: 13px;
                    alternate-background-color: #F4F4F4;
                    background: white;
                }
                QTableWidget::item {
                    color: black;
                }
                QHeaderView::section {
                    background-color: #595959;
                    padding: 4px;
                    border: 1px solid #D5DDE8;
                    font-weight: bold;
                }
                """,
            "plot_func": plot_calibration3
        },

        "吖啶橙荧光探针法测定DNA-标准曲线图": {
            "headers": ["Concentration", "Signal"],
            "rows": 20,
            "cols": 2,
            "style": """
                QTableWidget {
                    gridline-color: #CFCFCF;
                    font-size: 13px;
                    alternate-background-color: #F7F9FC;
                    background: white;
                }
                QTableWidget::item {
                    color: black;
                }
                QHeaderView::section {
                    background-color: #006CBF;
                    padding: 4px;
                    border: 1px solid #D5DDE8;
                    font-weight: bold;
                    }
                    """,
            "plot_func": plot_calibration3
        },

        "吖啶橙荧光探针法测定DNA-光谱图": {
            "headers": ["Wavelength", "Intensity"],
            "rows": 20,
            "cols": 2,
            "style": """
                QTableWidget {
                    gridline-color: #D0D0D0;
                    font-size: 13px;
                    alternate-background-color: #FFF8E8;
                    background: white;
                }
                QTableWidget::item {
                    color: black;
                }
                QHeaderView::section {
                    background-color: #127840;
                    padding: 4px;
                    border: 1px solid #D5DDE8;
                    font-weight: bold;
                        }
                        """,
            "plot_func": plot_line
        },
    }

    # 状态字典
    state = {
        "window": window,
        "table": table,
        "combo": combo,
        "figure": figure,
        "canvas": canvas,
        "experiments": experiments,
        "current_experiment": None,
        "picture_label": picture_label
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


def normalize_content(content):
    """
    统一处理章节正文。

    支持两种写法：
    1. content 是字符串
    2. content 是列表，每个元素是一段文字

    最终统一返回字符串。
    """

    if content is None:
        return ""

    if isinstance(content, list):
        return "\n".join(str(paragraph) for paragraph in content)

    return str(content)


def ensure_result_image_field(template):
    """
    给报告模板补充一个结果图片字段。
    该字段不写入基本信息表，只用于在 PDF 的实验结果章节中插入图片。
    如果模板里已经存在相同 key，则不重复添加。
    """

    fields = template.setdefault("fields", [])

    for field in fields:
        if field.get("key") == "result_image":
            return

    fields.append({
        "key": "result_image",
        "label": "实验结果图片",
        "type": "image",
        "required": False,
        "default": ""
    })


def is_result_section_heading(heading):
    """
    判断当前章节是否为实验结果章节。
    """

    heading_text = str(heading)
    return "实验结果" in heading_text or "结果" in heading_text


REPORT_TEMPLATES = load_report_templates(Path(__file__).resolve().parent / "templates")


# ============================================================
# 2. 字体与文本处理函数
# ============================================================

def register_chinese_font():
    """
    注册 ReportLab 内置中文字体。

    说明：
    1. ReportLab 默认字体不支持中文。
    2. STSong-Light 是 ReportLab 可直接使用的中文 CID 字体。
    3. 这种方式不需要额外提供字体文件，适合先快速跑通程序。
    """

    font_name = "STSong-Light"
    pdfmetrics.registerFont(UnicodeCIDFont(font_name))
    return font_name


def safe_paragraph_text(text):
    """
    将普通文本转换为 ReportLab Paragraph 可以识别的安全文本。
    主要处理：
    1. 转义特殊符号，避免 <、>、& 等符号影响 PDF 生成；
    2. 将换行符转换为 <br/>，保证 PDF 中能正常换行。
    """

    if text is None:
        text = ""

    text = str(text)
    text = escape(text)
    text = text.replace("\n", "<br/>")
    return text


def safe_body_paragraph_text(text):
    """
    只把中文/英文/数字混排附近的空格变成不换行空格。
    避免中文实验报告中出现“用 ICP / 程序”这种异常断行。
    """
    safe_text = safe_paragraph_text(text)

    safe_text = re.sub(
        r"(?<=[\u4e00-\u9fffA-Za-z0-9]) (?=[\u4e00-\u9fffA-Za-z0-9])",
        "\u00A0",
        safe_text
    )

    return safe_text


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
    return key == "picture" or key == "result_image" or re.fullmatch(r"picture[0-9]+", key) is not None


def get_image_path_by_placeholder_key(key, data):
    """
    根据图片占位符 key 获取图片路径。
    """

    key = str(key).strip()

    if key == "result_image":
        return str(data.get("result_image", "")).strip()

    if key in APP_STATE["generated_pictures"]:
        return APP_STATE["generated_pictures"][key]

    return ""


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


def add_result_image_to_story(story, data):
    """
    兼容旧逻辑：如果用户选择了实验结果图片，但正文中没有手动放置图片占位符，
    仍可将 result_image 插入实验结果章节。
    """

    image_path = data.get("result_image", "")

    if not str(image_path).strip():
        return

    append_image_to_story(story, image_path, "{result_image}")


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

        story.append(Paragraph(safe_body_paragraph_text(paragraph_text), paragraph_style))


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


def add_content_with_picture_placeholders_to_story(story, content, data, body_style):
    """
    按正文中的图片占位符顺序加入文字和图片。
    支持：
    1. {picture}：最近一次绘图；
    2. {picture1}、{picture2} ...：第 1、2 ... 次绘图；
    3. {result_image}：手动选择的实验结果图片；
    4. {{ picture }}、{{ picture1 }}、{{ result_image }}：双大括号兼容写法。
    """

    placeholder_pattern = r"\{\{\s*(picture[0-9]*|result_image)\s*\}\}|\{\s*(picture[0-9]*|result_image)\s*\}"

    last_index = 0
    has_match = False

    for match in re.finditer(placeholder_pattern, content):
        has_match = True
        before_text = content[last_index:match.start()]

        if before_text.strip():
            add_text_with_manual_indent_to_story(story, before_text, body_style)

        key = match.group(1) or match.group(2)
        image_path = get_image_path_by_placeholder_key(key, data)

        if not image_path:
            raise ValueError(f"没有找到图片变量：{{{key}}}。请先在绘图页生成对应图片，或选择对应图片。")

        append_image_to_story(story, image_path, f"{{{key}}}")
        last_index = match.end()

    remaining_text = content[last_index:]

    if remaining_text.strip():
        add_text_with_manual_indent_to_story(story, remaining_text, body_style)

    if not has_match and content.strip():
        return


def content_has_result_image_placeholder(content):
    """
    判断正文是否已经手动放置了 result_image 图片占位符。
    """

    return re.search(r"\{\{\s*result_image\s*}}|\{\s*result_image\s*}", content) is not None


def add_report_sections_to_story(story, template, data, heading_style, body_style, normal_style, font_name):
    """
    将模板中的预设章节加入 PDF 内容流。
    支持：
    1. 普通字符串段落；
    2. 图片占位符 {picture}、{picture1}、{result_image}；
    3. 三线表对象：
       {
         "type": "three_line_table",
         "title": "...",
         "headers": [...],
         "rows": [...],
         "col_widths_mm": [...]
       }
    """

    for section in template.get("sections", []):
        heading = section.get("heading", "")
        content = section.get("content", "")

        story.append(Paragraph(safe_paragraph_text(heading), heading_style))

        has_result_image_placeholder = False

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
                    else:
                        raise ValueError(f"未知的 sections.content 对象类型：{content_type}")

                else:
                    rendered_content = render_template_text(str(content_item), data)

                    if content_has_result_image_placeholder(rendered_content):
                        has_result_image_placeholder = True

                    add_content_with_picture_placeholders_to_story(
                        story=story,
                        content=rendered_content,
                        data=data,
                        body_style=body_style
                    )

        else:
            rendered_content = render_template_text(str(content), data)

            if content_has_result_image_placeholder(rendered_content):
                has_result_image_placeholder = True

            add_content_with_picture_placeholders_to_story(
                story=story,
                content=rendered_content,
                data=data,
                body_style=body_style
            )

        if is_result_section_heading(heading) and not has_result_image_placeholder:
            add_result_image_to_story(story, data)

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
    ensure_result_image_field(template)

    for field in template.get("fields", []):
        key = field.get("key")
        label = field.get("label", key)
        required = field.get("required", False)

        widget = create_input_widget(field)
        input_widgets[key] = widget

        if required:
            label_text = f"{label} *"
        else:
            label_text = label

        form_layout.addRow(QLabel(label_text), widget)


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

        if check_required and required and value.strip() == "":
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
        lines.append(f"[三线表] {table_title}")

    if headers:
        lines.append(" | ".join(str(header) for header in headers))
        lines.append(" | ".join("---" for _ in headers))

    for row in rows:
        rendered_cells = []

        for cell in row:
            rendered_cell = render_template_text(str(cell), data)
            rendered_cell = strip_manual_indent_markers(rendered_cell)
            rendered_cells.append(rendered_cell)

        lines.append(" | ".join(rendered_cells))

    return "\n".join(lines)


def build_preview_text(template, data):
    """
    生成纯文本预览内容。
    支持：
    1. 普通文本；
    2. [indent=数字] 缩进标记隐藏；
    3. 图片占位符文本预览；
    4. three_line_table 三线表对象预览。
    """

    lines = []

    lines.append(template.get("title", "实验报告"))
    lines.append("=" * 30)
    lines.append("")

    lines.append("【基本信息】")
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

        has_result_image_placeholder = False

        if isinstance(content, list):
            for content_item in content:
                if isinstance(content_item, dict):
                    content_type = content_item.get("type", "")

                    if content_type == "three_line_table":
                        lines.append(build_three_line_table_preview(content_item, data))
                    else:
                        lines.append(f"[未知内容类型] {content_type}")

                else:
                    rendered_content = render_template_text(str(content_item), data)

                    if content_has_result_image_placeholder(rendered_content):
                        has_result_image_placeholder = True

                    preview_content = strip_manual_indent_markers(rendered_content)
                    lines.append(preview_content)

        else:
            rendered_content = render_template_text(str(content), data)

            if content_has_result_image_placeholder(rendered_content):
                has_result_image_placeholder = True

            preview_content = strip_manual_indent_markers(rendered_content)
            lines.append(preview_content)

        if is_result_section_heading(heading):
            image_path = data.get("result_image", "").strip()
            if image_path and not has_result_image_placeholder:
                lines.append(f"[实验结果图片] {image_path}")

        generated_pictures = APP_STATE["generated_pictures"]
        if generated_pictures:
            picture_names = sorted(
                key for key in generated_pictures.keys()
                if re.fullmatch(r"picture[0-9]+", key) is not None
            )
            if picture_names:
                lines.append("[已生成图片变量] " + "，".join(f"{{{key}}}" for key in picture_names))
                lines.append("[最新图片变量] {picture}")

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

    # 右侧预览区
    preview_group = QGroupBox("报告预览")
    preview_group_layout = QVBoxLayout(preview_group)

    preview_edit = QTextEdit()
    preview_edit.setReadOnly(True)
    preview_group_layout.addWidget(preview_edit)

    content_layout.addWidget(preview_group, stretch=1)

    # 用字典保存输入控件
    input_widgets = {}

    def refresh_current_form():
        # 刷新当前模板对应的输入表单。

        template_name = template_combo.currentText()
        template = REPORT_TEMPLATES[template_name]

        rebuild_form(form_layout, input_widgets, template)
        update_preview(template_combo, input_widgets, preview_edit)

    def refresh_current_preview():
        # 刷新右侧文本预览。

        update_preview(template_combo, input_widgets, preview_edit)

    def export_current_pdf():
        # 导出当前模板对应的 PDF。

        export_pdf(window, template_combo, input_widgets)

    # -------------------------------
    # 绑定按钮事件
    # -------------------------------
    template_combo.currentTextChanged.connect(refresh_current_form)
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

    tabs.addTab(plot_window, "实验数据绘图")
    tabs.addTab(report_window, "实验报告生成")

    window.plot_window = plot_window
    window.report_window = report_window

    return window


def main():
    app = QApplication(sys.argv)
    window = build_integrated_window()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
