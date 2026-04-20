import sys
import numpy as np
from scipy.optimize import curve_fit

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget,
    QVBoxLayout, QHBoxLayout,
    QLabel, QComboBox, QPushButton,
    QTableWidget, QTableWidgetItem,
    QMessageBox, QHeaderView, QSizePolicy, QFileDialog
)
from PyQt6.QtGui import QShortcut, QKeySequence
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure


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
            raise ValueError(f"第 {row + 1} 行存在非数字内容，请检查。")

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
# 实验1：离子选择性电极测定氟离子
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

    # 为了让拟合线更顺眼，按 x 排序后再画
    sort_idx = np.argsort(x)
    x_sorted = x[sort_idx]
    y_fit_sorted = y_fit[sort_idx]

    figure.clear()
    ax = figure.add_subplot(111)

    ax.scatter(x, y, s=40, label="Data")
    ax.plot(x_sorted, y_fit_sorted, linewidth=1.5, label="Fit")

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
# 甲苯，萘的高效液相色谱分析
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

    # 为了让拟合线更顺眼，按 x 排序后再画
    sort_idx = np.argsort(x)
    x_sorted = x[sort_idx]
    y_fit_sorted = y_fit[sort_idx]

    figure.clear()
    ax = figure.add_subplot(111)

    ax.scatter(x, y, s=40, label="Data")
    ax.plot(x_sorted, y_fit_sorted, linewidth=1.5, label="Fit")

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
# 气象色谱流动相速度对柱效的影响
# =========================================================
def plot_point(state):
    table = state["table"]
    figure = state["figure"]
    canvas = state["canvas"]

    arr = read_numeric_columns(table, 2)
    x = arr[:, 0]
    y = arr[:, 1]

    def model_func(u, A, B, C):
        return A + B / u + C * u

    popt, pcov= curve_fit(model_func, x, y)
    A_fit, B_fit, C_fit = popt

    X_fit = np.linspace(min(x), max(x), 200)
    Y_fit = model_func(X_fit, *popt)


    figure.clear()
    ax = figure.add_subplot(111)

    ax.scatter(x, y, color="blue", label="Data")

    ax.plot(X_fit, Y_fit, label="Calibration",
             color='black', linewidth=0.5)

    ax.text(
        0.10, 0.95,
        f"Y = {A_fit:.3f} + {B_fit:.3f}/X + {C_fit:.3f}X",
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
# 主函数：创建整个界面
# =========================================================
def main():
    app = QApplication(sys.argv)

    # -------------------------
    # 创建主窗口
    # -------------------------
    window = QMainWindow()
    window.setWindowTitle("实验数据绘图工具（函数版）")
    window.resize(1200, 700)

    # 中央区域
    central = QWidget()
    window.setCentralWidget(central)

    # 总布局：左右分栏
    main_layout = QHBoxLayout(central)

    # 左边布局：上方控制区 + 下方表格
    left_layout = QVBoxLayout()
    main_layout.addLayout(left_layout, 3)

    # 右边布局：图像显示
    right_layout = QVBoxLayout()
    main_layout.addLayout(right_layout, 6)

    # -------------------------
    # 左上控制区
    # -------------------------
    top_layout = QHBoxLayout()
    left_layout.addLayout(top_layout)

    label = QLabel("实验类型：")
    top_layout.addWidget(label)

    combo = QComboBox()
    top_layout.addWidget(combo)

    btn_plot = QPushButton("绘图")
    top_layout.addWidget(btn_plot)

    btn_clear = QPushButton("清空表格")
    top_layout.addWidget(btn_clear)

    btn_save = QPushButton("导出图片")
    top_layout.addWidget(btn_save)

    top_layout.addStretch()

    # -------------------------
    # 左下表格
    # -------------------------
    table = QTableWidget()
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

    # 先放一个空坐标轴，避免一开始空白过于突兀
    figure.add_subplot(111)

    # -------------------------
    # 每种实验的配置
    # -------------------------
    experiments = {
        "离子选择性电极测定氟离子": {
            "headers": ["X", "Y"],
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
            "rows": 15,
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

        "气象色谱流动相速度对柱效的影响": {
            "headers": ["u", "H"],
            "rows": 10,
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
        }
    }

    # -------------------------
    # 状态字典：把界面里的重要对象都存起来
    # -------------------------
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

    # 启动时，默认加载第一个实验
    first_experiment_name = combo.currentText()
    change_experiment(state, first_experiment_name)

    # 显示窗口
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()