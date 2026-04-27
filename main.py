import sys
import numpy as np
from scipy.optimize import curve_fit
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget,
    QVBoxLayout, QHBoxLayout,
    QLabel, QComboBox, QPushButton,
    QTableWidget, QTableWidgetItem,
    QMessageBox, QHeaderView, QSizePolicy, QFileDialog)
from PyQt6.QtGui import QShortcut, QKeySequence
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QSplitter
from types import MethodType


# =========================================================
# 读取表格中的前 n 列数字
# =========================================================
def read_numeric_columns(table, required_cols):
    """
    从表格中读取前 required_cols 列数据，并转成 numpy 数组
    自动跳过整行为空的行
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
def plot_current_experiment(state):
    try:
        experiment_name = state["current_experiment"]
        if experiment_name is None:
            raise ValueError("还没有选择实验类型。")

        config = state["experiments"][experiment_name]
        plot_func = config["plot_func"]

        plot_func(state)

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
def auto_add_rows(table, item):
    # 只有当用户改动的是最后一行时，才检查是否扩展
    if item.row() != table.rowCount() - 1:
        return

    # 最后一行只要有任意一个格子非空，补 5 行
    last_row = table.rowCount() - 1
    for col in range(table.columnCount()):
        cell = table.item(last_row, col)
        if cell and cell.text().strip():
            table.setRowCount(table.rowCount() + 5)

            for row in range(table.rowCount() - 5, table.rowCount()):
                for c in range(table.columnCount()):
                    if table.item(row, c) is None:
                        table.setItem(row, c, QTableWidgetItem(""))
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
            table.setRowCount(table.rowCount() + 5)

            for r in range(old_row_count, table.rowCount()):
                for c in range(table.columnCount()):
                    if table.item(r, c) is None:
                        table.setItem(r, c, QTableWidgetItem(""))

        # 如果目标格子没有 item，就补一个空 item
        if table.item(row, col) is None:
            table.setItem(row, col, QTableWidgetItem(""))

        table.setCurrentCell(row, col)

    def new_key_press_event(event):
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
def main():
    app = QApplication(sys.argv)

    # -------------------------
    # 创建主窗口
    # -------------------------
    window = QMainWindow()
    window.setWindowTitle("实验数据绘图工具")
    window.resize(1200, 700)

    # 中央区域
    central = QWidget()
    window.setCentralWidget(central)

    # =========================
    # 总布局：上下分栏
    # =========================
    root_layout = QVBoxLayout(central)

    # -------------------------
    # 顶部控制区
    # -------------------------
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

    # -------------------------
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

    # -------------------------
    # 左侧表格
    # -------------------------
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

    # -------------------------
    # 右侧图像区域
    # -------------------------
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
        "current_experiment": None
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

    # 显示窗口
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()