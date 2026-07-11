import re
from pathlib import Path
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Image, KeepTogether, LongTable, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from core import (CHINESE_FONT_NAME,IMPORT_PICTURE_PREFIX,TIMES_NEW_ROMAN_PATH,WESTERN_FONT_NAME,
                  expand_dimension_values,get_field_label_by_key,get_image_path_by_placeholder_key,
                  get_preset_image_path,plain_inline_markup_for_gui,render_template_text,
                  safe_paragraph_text,strip_manual_indent_markers,)


def is_result_section_heading(heading):
    """
    判断当前章节是否为实验结果章节。
    """

    heading_text = str(heading)
    return "实验结果" in heading_text or "结果" in heading_text


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


def add_content_with_picture_placeholders_to_story(story, content, body_style, app_state=None):
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
        image_path = get_image_path_by_placeholder_key(key, app_state)

        if not image_path:
            raise ValueError(f"没有找到图片变量：{{{key}}}。请先生成或导入对应图片。")

        append_image_to_story(story, image_path, f"{{{key}}}")
        last_index = match.end()

    remaining_text = content[last_index:]

    if remaining_text.strip():
        add_text_with_manual_indent_to_story(story, remaining_text, body_style)

    if not has_match and content.strip():
        return


def add_report_sections_to_story(story, template, data, heading_style, body_style, normal_style, font_name, app_state=None):
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
                        body_style=body_style,
                        app_state=app_state
                    )

        else:
            rendered_content = render_template_text(str(content), data)

            add_content_with_picture_placeholders_to_story(
                story=story,
                content=rendered_content,
                body_style=body_style,
                app_state=app_state
            )

        story.append(Spacer(1, 4))


def generate_pdf(output_path, template, data, app_state=None):
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
        font_name=font_name,
        app_state=app_state
    )

    # 生成 PDF
    doc.build(story)


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

