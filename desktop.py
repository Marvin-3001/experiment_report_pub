import sys
from types import MethodType
from PyQt6.QtCore import QDate, QTimer, Qt
from matplotlib.figure import Figure
from PyQt6.QtGui import QKeySequence, QShortcut
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from core import (APP_STATE,CUSTOM_TEMPLATE_DIR,PLOT_CONFIG_DIR,TEMPLATE_DIR,build_report_hint_text,
                  expand_dimension_values,get_app_state,is_empty_form_value,
                  get_template_dir_fingerprint,load_external_report_templates,
                  load_report_templates,plain_inline_markup_for_gui,
                  save_generated_plot_to_picture_store,save_import_picture_to_picture_store,)
from PyQt6.QtWidgets import (QApplication,QComboBox,QDateEdit,QFileDialog,
                            QFormLayout,QGroupBox,QHBoxLayout,QHeaderView,
                            QLabel,QLineEdit,QMainWindow,QMessageBox,QPushButton,
                            QScrollArea,QSizePolicy,QSplitter,QTabWidget,QTableWidget,
                             QTableWidgetItem,QTextEdit,QVBoxLayout,QWidget,)
from report import build_preview_text, generate_pdf, get_default_pdf_name
from plotting import load_plot_experiment_configs, plot_experiment, read_numeric_rows, save_figure


BUILTIN_REPORT_TEMPLATES = load_report_templates(TEMPLATE_DIR)


def table_to_rows(table, required_cols=None):
    cols = table.columnCount() if required_cols is None else required_cols
    rows = []

    for row in range(table.rowCount()):
        row_values = []

        for col in range(cols):
            item = table.item(row, col)
            row_values.append(item.text().strip() if item else "")

        rows.append(row_values)

    return rows


def read_numeric_columns_from_table(table, required_cols):
    return read_numeric_rows(table_to_rows(table, required_cols), required_cols)


def get_plot_labels_from_state(state):
    labels = {}

    widget_map = {
        "title": "plot_title_edit",
        "xlabel": "plot_xlabel_edit",
        "ylabel": "plot_ylabel_edit",
    }

    for label_key, widget_key in widget_map.items():
        widget = state.get(widget_key)

        if widget is not None:
            labels[label_key] = widget.text().strip()

    return labels


def plot_current_experiment(state):
    try:
        experiment_name = state["current_experiment"]

        if experiment_name is None:
            raise ValueError('还没有选择实验类型。')

        config = state["experiments"][experiment_name]
        required_cols = int(config.get("required_cols", 2))
        numeric_data = read_numeric_columns_from_table(state["table"], required_cols)
        labels = get_plot_labels_from_state(state)

        plot_experiment(config, numeric_data, state["figure"], labels=labels)
        state["has_current_plot"] = True
        state["canvas"].draw()

    except Exception as error:
        state["has_current_plot"] = False
        QMessageBox.critical(state["window"], '错误', str(error))


def add_current_figure_to_report(state):
    if not state.get("has_current_plot", False):
        QMessageBox.warning(
            state["window"],
            '还没有可加入的图片',
            '请先点击“绘图”生成图片。'
        )
        return

    picture_key = save_generated_plot_to_picture_store(
        state["figure"],
        state.get("app_state"),
        source_name=state.get("current_experiment", "")
    )
    QMessageBox.information(
        state["window"],
        '已加入报告图片',
        '已生成图片变量：' + "{" + picture_key + "}"
    )


def save_current_figure(state):
    file_path, _ = QFileDialog.getSaveFileName(
        state["window"],
        '导出图片',
        "plot.png",
        "PNG Files (*.png);;JPG Files (*.jpg);;SVG Files (*.svg);;PDF Files (*.pdf)"
    )

    if not file_path:
        return

    save_figure(state["figure"], file_path)
    QMessageBox.information(state["window"], '成功', '图片已导出。')


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


def clear_selected_cells(table):

    readonly_cells = getattr(table, "report_readonly_cells", set())
    for item in table.selectedItems():
        row = item.row()
        col = item.column()

        if (row, col) in readonly_cells:
            continue

        item.setText("")


def clear_all_cells(table):
    for row in range(table.rowCount()):
        for col in range(table.columnCount()):
            item = table.item(row, col)
            if item:
                item.setText("")
            else:
                table.setItem(row, col, QTableWidgetItem(""))


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


def clear_plot(state):
    figure = state["figure"]
    canvas = state["canvas"]

    figure.clear()
    figure.add_subplot(111)
    state["has_current_plot"] = False
    canvas.draw()


def change_experiment(state, experiment_name):
    config = state["experiments"][experiment_name]

    state["current_experiment"] = experiment_name
    setup_table(state["table"], config)
    plot_labels = config.get("plot_labels", {})

    label_widgets = [
        ("plot_title_edit", "title"),
        ("plot_xlabel_edit", "xlabel"),
        ("plot_ylabel_edit", "ylabel")
    ]

    for widget_key, label_key in label_widgets:
        widget = state.get(widget_key)

        if widget is not None:
            widget.setText(str(plot_labels.get(label_key, "")))

    clear_plot(state)


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


def update_report_hint(hint_edit, app_state=None):
    """
    刷新提示框内容。
    """

    hint_edit.setPlainText(build_report_hint_text(app_state))


def update_preview(template, input_widgets, preview_edit):
    """
    更新右侧预览框内容。
    """

    data = collect_form_data(
        template=template,
        input_widgets=input_widgets,
        check_required=False
    )

    preview_text = build_preview_text(template, data)
    preview_edit.setPlainText(preview_text)


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


def export_pdf(parent_window, template, input_widgets, app_state=None):
    """
    GUI 中点击“生成 PDF”按钮后执行的函数。
    """

    try:
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

        generate_pdf(output_path, template, data, app_state=app_state)

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


def build_plot_window(app_state=None):

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

    btn_add_to_report = QPushButton("加入报告图片")
    control_layout.addWidget(btn_add_to_report)

    btn_clear = QPushButton("清空表格")
    control_layout.addWidget(btn_clear)

    btn_save = QPushButton("导出图片")
    control_layout.addWidget(btn_save)

    control_layout.addStretch()

    # 图标题和坐标轴标题：同一行，留空时使用当前实验默认值
    label_panel = QWidget()
    label_layout = QHBoxLayout(label_panel)
    label_layout.setContentsMargins(0, 0, 0, 0)
    root_layout.addWidget(label_panel, 0)

    plot_title_edit = QLineEdit()
    plot_xlabel_edit = QLineEdit()
    plot_ylabel_edit = QLineEdit()

    plot_title_edit.setPlaceholderText("默认图标题")
    plot_xlabel_edit.setPlaceholderText("默认 X 轴标题")
    plot_ylabel_edit.setPlaceholderText("默认 Y 轴标题")

    label_layout.addWidget(QLabel("图标题："))
    label_layout.addWidget(plot_title_edit, stretch=2)
    label_layout.addWidget(QLabel("X轴标题："))
    label_layout.addWidget(plot_xlabel_edit, stretch=1)
    label_layout.addWidget(QLabel("Y轴标题："))
    label_layout.addWidget(plot_ylabel_edit, stretch=1)

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
    experiments = load_plot_experiment_configs(
        PLOT_CONFIG_DIR,
        table_style_builder=build_qtable_style
    )

    # 状态字典
    state = {
        "window": window,
        "table": table,
        "combo": combo,
        "figure": figure,
        "canvas": canvas,
        "experiments": experiments,
        "current_experiment": None,
        "has_current_plot": False,
        "plot_title_edit": plot_title_edit,
        "plot_xlabel_edit": plot_xlabel_edit,
        "plot_ylabel_edit": plot_ylabel_edit,
        "app_state": app_state,
    }

    # 下拉框加入实验名称
    combo.addItems(experiments.keys())

    # 连接信号
    combo.currentTextChanged.connect(lambda text: change_experiment(state, text))
    btn_plot.clicked.connect(lambda: plot_current_experiment(state))
    btn_add_to_report.clicked.connect(lambda: add_current_figure_to_report(state))
    btn_clear.clicked.connect(lambda: clear_all_cells(table))
    btn_save.clicked.connect(lambda: save_current_figure(state))
    table.itemChanged.connect(lambda item: auto_add_rows(table, item))

    # 启动时，默认加载第一个实验
    first_experiment_name = combo.currentText()
    change_experiment(state, first_experiment_name)

    return window


def build_main_window(app_state=None):
    """
    构建主窗口。
    整体界面结构：
    上方：模板选择 + 按钮
    左侧：用户输入表单
    右侧：报告文本预览
    """

    app_state = get_app_state(app_state)

    window = QWidget()
    window.setWindowTitle("实验报告 PDF 生成工具")
    window.resize(1100, 720)

    main_layout = QVBoxLayout(window)

    # -------------------------------
    # 顶部工具栏
    # -------------------------------
    top_layout = QHBoxLayout()
    main_layout.addLayout(top_layout)

    top_layout.addWidget(QLabel("内置模板："))

    template_combo = QComboBox()
    template_combo.addItems(BUILTIN_REPORT_TEMPLATES.keys())
    template_combo.setMinimumWidth(260)
    top_layout.addWidget(template_combo, stretch=2)

    top_layout.addWidget(QLabel("外部模板："))

    external_template_combo = QComboBox()
    external_template_combo.addItem("不使用外部模板")
    external_template_combo.setMinimumWidth(210)
    external_template_combo.setEnabled(False)
    top_layout.addWidget(external_template_combo, stretch=2)

    import_picture_button = QPushButton("导入外部图片")
    top_layout.addWidget(import_picture_button, stretch=1)

    preview_button = QPushButton("刷新预览")
    top_layout.addWidget(preview_button, stretch=1)

    export_button = QPushButton("生成 PDF")
    top_layout.addWidget(export_button, stretch=1)

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

    # 外部模板使用独立字典；活动模板保留快照，避免文件修改时清空正在填写的表单。
    external_templates = {}
    current_selection = {
        "source": "builtin",
        "name": template_combo.currentText(),
        "template": BUILTIN_REPORT_TEMPLATES[template_combo.currentText()]
    }
    external_template_fingerprint: tuple | None = None
    reported_external_template_errors: tuple | None = None

    def get_current_template():
        return current_selection["template"]

    def refresh_current_form():
        # 刷新当前模板对应的输入表单。

        template = get_current_template()

        rebuild_form(form_layout, input_widgets, template)
        update_preview(template, input_widgets, preview_edit)
        update_report_hint(hint_edit, app_state)

    def refresh_current_preview():
        # 刷新右侧文本预览和提示信息。

        update_preview(get_current_template(), input_widgets, preview_edit)
        update_report_hint(hint_edit, app_state)

    def select_builtin_template(template_name):
        if not template_name or template_name not in BUILTIN_REPORT_TEMPLATES:
            return

        external_template_combo.blockSignals(True)
        external_template_combo.setCurrentIndex(0)
        external_template_combo.blockSignals(False)

        current_selection.update({
            "source": "builtin",
            "name": template_name,
            "template": BUILTIN_REPORT_TEMPLATES[template_name]
        })
        refresh_current_form()

    def select_external_template(template_name):
        if template_name not in external_templates:
            if current_selection["source"] == "external":
                select_builtin_template(template_combo.currentText())
            return

        current_selection.update({
            "source": "external",
            "name": template_name,
            "template": external_templates[template_name]
        })
        refresh_current_form()

    def show_external_template_errors(errors):
        if not errors:
            return

        QMessageBox.warning(
            window,
            "外部模板加载失败",
            "以下外部模板未能加载：\n" + "\n".join(errors)
        )

    def check_external_templates():
        nonlocal external_templates
        nonlocal external_template_fingerprint
        nonlocal reported_external_template_errors

        fingerprint = get_template_dir_fingerprint(CUSTOM_TEMPLATE_DIR)

        if fingerprint == external_template_fingerprint:
            return

        external_template_fingerprint = fingerprint
        loaded_templates, errors = load_external_report_templates(CUSTOM_TEMPLATE_DIR)
        external_templates = loaded_templates

        selected_external_name = None
        if current_selection["source"] == "external":
            selected_external_name = current_selection["name"]

        external_template_combo.blockSignals(True)
        external_template_combo.clear()
        external_template_combo.addItem("不使用外部模板")
        external_template_combo.addItems(external_templates.keys())
        external_template_combo.setEnabled(bool(external_templates))

        if selected_external_name in external_templates:
            external_template_combo.setCurrentText(selected_external_name)
        else:
            external_template_combo.setCurrentIndex(0)

        external_template_combo.blockSignals(False)

        if selected_external_name and selected_external_name not in external_templates:
            builtin_name = template_combo.currentText()
            current_selection.update({
                "source": "builtin",
                "name": builtin_name,
                "template": BUILTIN_REPORT_TEMPLATES[builtin_name]
            })
            refresh_current_form()

        error_key = (fingerprint, tuple(errors))
        if errors and error_key != reported_external_template_errors:
            reported_external_template_errors = error_key
            show_external_template_errors(errors)

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
                new_key = save_import_picture_to_picture_store(image_path, app_state)
                new_keys.append(new_key)

                # 每导入一张图片，就刷新一次提示区文本。
                update_report_hint(hint_edit, app_state)

            update_preview(get_current_template(), input_widgets, preview_edit)

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

        export_pdf(window, get_current_template(), input_widgets, app_state)

    # 让外部窗口也能触发报告页刷新
    def refresh_report_page_from_outside():
        # 从其他标签页切换回来时，同时刷新报告预览和提示框。
        update_preview(get_current_template(), input_widgets, preview_edit)
        update_report_hint(hint_edit, app_state)

    window.refresh_report_page = refresh_report_page_from_outside

    # -------------------------------
    # 绑定按钮事件
    # -------------------------------
    template_combo.currentTextChanged.connect(select_builtin_template)
    external_template_combo.currentTextChanged.connect(select_external_template)
    import_picture_button.clicked.connect(import_pictures_from_files)
    preview_button.clicked.connect(refresh_current_preview)
    export_button.clicked.connect(export_current_pdf)

    # 初始加载表单
    refresh_current_form()

    # 运行期间自动发现新增、删除或修改的外部 JSON 模板。
    external_template_timer = QTimer(window)
    external_template_timer.setInterval(1000)
    external_template_timer.timeout.connect(check_external_templates)
    external_template_timer.start()
    QTimer.singleShot(0, check_external_templates)

    window.external_template_timer = external_template_timer

    return window


def build_integrated_window(app_state=None):
    """
    构建总窗口，用标签页整合绘图工具和报告生成工具。
    """

    app_state = get_app_state(app_state)

    window = QMainWindow()
    window.setWindowTitle("实验数据绘图与报告生成工具")
    window.resize(1250, 760)

    tabs = QTabWidget()
    window.setCentralWidget(tabs)

    plot_window = build_plot_window(app_state)
    report_window = build_main_window(app_state)

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


def run_desktop_app(app_state=None):
    app = QApplication(sys.argv)
    window = build_integrated_window(app_state or APP_STATE)
    window.show()
    sys.exit(app.exec())


def main():
    run_desktop_app()

