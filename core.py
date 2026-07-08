import json
import re
import shutil
import tempfile
from pathlib import Path
from xml.sax.saxutils import escape

from matplotlib import font_manager, rcParams


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
MATPLOTLIB_WESTERN_FONT_NAME = "Times New Roman"
MATPLOTLIB_CHINESE_FONT_CANDIDATES = [
    "SimSun",
    "宋体",
    "Microsoft YaHei",
    "SimHei",
    "Noto Sans CJK SC"
]

def find_available_matplotlib_font(font_names):
    for font_name in font_names:
        try:
            font_path = font_manager.findfont(
                font_name,
                fallback_to_default=False
            )
        except (ValueError, RuntimeError, OSError):
            continue

        if Path(font_path).exists():
            return font_name

    return None

MATPLOTLIB_CHINESE_FONT_NAME = find_available_matplotlib_font(
    MATPLOTLIB_CHINESE_FONT_CANDIDATES
)


def setup_matplotlib_fonts():
    font_families = [MATPLOTLIB_WESTERN_FONT_NAME]

    if MATPLOTLIB_CHINESE_FONT_NAME:
        font_families.append(MATPLOTLIB_CHINESE_FONT_NAME)

    rcParams["font.family"] = font_families
    rcParams["font.sans-serif"] = font_families
    rcParams["axes.unicode_minus"] = False

setup_matplotlib_fonts()


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

TEMPLATE_DIR = Path(__file__).resolve().parent / "templates"
PRESET_IMAGE_DIR = TEMPLATE_DIR / "image"
PLOT_CONFIG_DIR = TEMPLATE_DIR / "plot_configs"


def get_app_state(app_state=None):
    if app_state is None:
        return APP_STATE
    return app_state


def create_app_state():
    return {
        "picture_counter": 0,
        "generated_pictures": {},
        "picture_sources": {},
        "import_picture_counter": 0,
        "import_pictures": {},
        "import_picture_sources": {}
    }


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

def is_import_picture_placeholder_key(key):
    """
    判断 key 是否是手动导入图片变量。
    支持：
    import_picture1、import_picture2、import_picture3 ...
    """

    key = str(key).strip()
    pattern = rf"{re.escape(IMPORT_PICTURE_PREFIX)}[0-9]+"

    return re.fullmatch(pattern, key) is not None

def save_import_picture_to_picture_store(image_path, app_state=None):
    """
    将手动导入的单张图片复制到临时目录，并生成图片变量。
    每导入一张图片，生成一个新变量：
    {import_picture1}、{import_picture2}、{import_picture3} ...
    """

    app_state = get_app_state(app_state)
    image_path = str(image_path).strip()

    if not image_path:
        raise ValueError("导入图片路径为空。")

    image_file = Path(image_path)

    if not image_file.exists():
        raise FileNotFoundError(f"导入图片不存在：{image_file}")

    image_suffix = image_file.suffix.lower()

    if image_suffix not in [".png", ".jpg", ".jpeg", ".bmp"]:
        raise ValueError(f"不支持的图片格式：{image_suffix}")

    app_state["import_picture_counter"] += 1
    image_key = f"{IMPORT_PICTURE_PREFIX}{app_state['import_picture_counter']}"

    picture_dir = Path(tempfile.gettempdir()) / "combined_experiment_tool_import_pictures"
    picture_dir.mkdir(parents=True, exist_ok=True)

    target_path = picture_dir / f"{image_key}{image_suffix}"
    shutil.copy2(str(image_file), str(target_path))

    app_state["import_pictures"][image_key] = str(target_path)
    app_state["import_picture_sources"][image_key] = image_file.name

    return image_key

def save_generated_plot_to_picture_store(figure, app_state=None, source_name=""):
    """
    将当前 Matplotlib 图保存为一个可在报告正文中调用的图片变量。
    规则：
    1. {picture1}、{picture2}、{picture3} ... 分别对应每次加入报告的图片；
    """

    app_state = get_app_state(app_state)
    app_state["picture_counter"] += 1
    picture_key = f"picture{app_state['picture_counter']}"

    picture_dir = Path(tempfile.gettempdir()) / "combined_experiment_tool_pictures"
    picture_dir.mkdir(parents=True, exist_ok=True)

    picture_path = picture_dir / f"{picture_key}.png"
    figure.savefig(str(picture_path), dpi=300, bbox_inches="tight")

    app_state["generated_pictures"][picture_key] = str(picture_path)
    app_state["picture_sources"][picture_key] = str(source_name).strip()

    return picture_key

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

def get_image_path_by_placeholder_key(key, app_state=None):
    """
    根据图片占位符 key 获取图片路径。
    """

    app_state = get_app_state(app_state)
    key = str(key).strip()

    if key in app_state["generated_pictures"]:
        return app_state["generated_pictures"][key]

    if key in app_state["import_pictures"]:
        return app_state["import_pictures"][key]

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

def get_picture_key_sort_number(picture_key):
    """
    从 picture1、picture2 这类变量名中提取数字，用于正确排序。
    """

    match = re.search(r"\d+", str(picture_key))

    if match:
        return int(match.group())

    return 0

def build_report_hint_text(app_state=None):
    """
    生成报告生成页右下角的提示文本。
    包括：
    1. 当前可用绘图图片变量；
    2. 当前可用导入图片变量；
    3. 常用正文标记；
    4. 图片插入提示。
    """

    app_state = get_app_state(app_state)

    lines = ["【绘图图片变量】"]

    generated_pictures = app_state["generated_pictures"]
    picture_sources = app_state.get("picture_sources", {})

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

    import_pictures = app_state["import_pictures"]
    import_picture_sources = app_state.get("import_picture_sources", {})

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
        lines.append("{import_picture1}、{import_picture2} 等变量会在点击“导入外部图片”后生成。")

    lines.append("")
    lines.append("【常用标记】")
    lines.append("下标：E[sub=pc]，PDF 中 pc 会显示为下标。")
    lines.append("上标：10[sup=-3]，PDF 中 -3 会显示为上标。")

    lines.append("")
    lines.append("【图片插入】")
    lines.append("插入绘图图片：{picture1}、{picture2} ...")
    lines.append("插入导入图片：{import_picture1}、{import_picture2} ...")

    return "\n".join(lines)
