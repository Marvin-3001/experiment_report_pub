import json
import numpy as np
from pathlib import Path
from scipy.optimize import curve_fit
from core import MATPLOTLIB_CHINESE_FONT_NAME, MATPLOTLIB_WESTERN_FONT_NAME


def read_numeric_rows(rows, required_cols):
    data = []

    for row_index, row in enumerate(rows):
        row_data = []
        is_empty_row = True

        for col in range(required_cols):
            text = ""

            if col < len(row):
                text = str(row[col]).strip()

            if text != "":
                is_empty_row = False

            row_data.append(text)

        if is_empty_row:
            continue

        try:
            numeric_row = [float(x) for x in row_data]
        except ValueError:
            raise ValueError('第 ' + str(row_index + 1) + ' 行存在问题，请检查。')

        data.append(numeric_row)

    if not data:
        raise ValueError('表格中没有有效数据。')

    return np.array(data, dtype=float)


def get_plot_label_text(labels, label_key, default_text):
    labels = labels or {}
    text = str(labels.get(label_key, "")).strip()

    if text:
        return text

    return default_text


def get_plot_font_families():
    font_families = [MATPLOTLIB_WESTERN_FONT_NAME]

    if MATPLOTLIB_CHINESE_FONT_NAME:
        font_families.append(MATPLOTLIB_CHINESE_FONT_NAME)

    return font_families


def apply_plot_text_fonts(ax, extra_text_items=None, legend=None):
    font_families = get_plot_font_families()

    ax.title.set_fontfamily(font_families)
    ax.xaxis.label.set_fontfamily(font_families)
    ax.yaxis.label.set_fontfamily(font_families)

    if legend is not None:
        for text_item in legend.get_texts():
            text_item.set_fontfamily(font_families)

    for text_item in extra_text_items or []:
        text_item.set_fontfamily(font_families)


def plot_linear_calibration(numeric_data, figure, labels=None, default_title="Calibration Curve", default_xlabel="Concentration", default_ylabel="Response"):
    title = get_plot_label_text(labels, "title", default_title)
    xlabel = get_plot_label_text(labels, "xlabel", default_xlabel)
    ylabel = get_plot_label_text(labels, "ylabel", default_ylabel)

    x = numeric_data[:, 0]
    y = numeric_data[:, 1]

    if len(x) < 2:
        raise ValueError('标准曲线至少需要 2 个数据点。')

    slope, intercept = np.polyfit(x, y, 1)
    y_fit = slope * x + intercept

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

    ax.set_title(title)
    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    ax.grid(True, alpha=0.3)
    legend = ax.legend()

    eq_text = f"y = {slope:.4f}x + {intercept:.4f}\nR² = {r2:.4f}"
    equation_text = ax.text(
        0.05, 0.95,
        eq_text,
        transform=ax.transAxes,
        verticalalignment="top",
        bbox=dict(boxstyle="round", facecolor="white", alpha=0.8)
    )
    apply_plot_text_fonts(ax, extra_text_items=[equation_text], legend=legend)


def plot_calibration1(numeric_data, figure, labels=None):
    plot_linear_calibration(numeric_data, figure, labels=labels, default_title="Calibration Curve", default_xlabel="lgC F-", default_ylabel="Voltage")


def plot_calibration2(numeric_data, figure, labels=None):
    plot_linear_calibration(numeric_data, figure, labels=labels, default_title="Calibration Curve", default_xlabel="Concentration", default_ylabel="Response")


def model_func(u, A, B, C):
    return A + B / u + C * u


def plot_point(numeric_data, figure, labels=None):
    title = get_plot_label_text(labels, "title", "Point Plot")
    xlabel = get_plot_label_text(labels, "xlabel", "u")
    ylabel = get_plot_label_text(labels, "ylabel", "H")

    x = numeric_data[:, 0]
    y = numeric_data[:, 1]

    if len(x) < 3:
        raise ValueError('数据点不足')

    if np.any(x <= 0):
        raise ValueError('流速 u 不能为 0 或负数')

    popt, pcov = curve_fit(model_func, x, y)
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
    ax.plot(X_fit, Y_fit, label="Fit", color="black", linewidth=1, linestyle="--")

    equation_text = ax.text(
        0.10, 0.95,
        f"H = {A_fit:.3f} + {B_fit:.3f}/u + {C_fit:.3f}u\nR² = {r2:.4f}",
        transform=ax.transAxes,
        verticalalignment="top",
        bbox=dict(boxstyle="round", facecolor="white", alpha=0.8)
    )

    ax.set_title(title)
    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    ax.grid(True, alpha=0.3)
    apply_plot_text_fonts(ax, extra_text_items=[equation_text])


def plot_calibration3(numeric_data, figure, labels=None):
    plot_linear_calibration(numeric_data, figure, labels=labels, default_title="Calibration Curve", default_xlabel="Concentration", default_ylabel="Signal")


def plot_line(numeric_data, figure, labels=None):
    title = get_plot_label_text(labels, "title", "Spectrum Plot")
    xlabel = get_plot_label_text(labels, "xlabel", "Wavelength")
    ylabel = get_plot_label_text(labels, "ylabel", "Intensity")

    x = numeric_data[:, 0]
    y = numeric_data[:, 1]

    figure.clear()
    ax = figure.add_subplot(111)
    ax.plot(x, y, color="black", linewidth=1)

    ax.set_title(title)
    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    apply_plot_text_fonts(ax)


PLOT_FUNCTIONS = {
    "calibration1": plot_calibration1,
    "calibration2": plot_calibration2,
    "point": plot_point,
    "calibration3": plot_calibration3,
    "line": plot_line
}

PLOT_DEFAULT_LABELS = {
    "calibration1": {"title": "Calibration Curve", "xlabel": "lgC F-", "ylabel": "Voltage"},
    "calibration2": {"title": "Calibration Curve", "xlabel": "Concentration", "ylabel": "Response"},
    "point": {"title": "Point Plot", "xlabel": "u", "ylabel": "H"},
    "calibration3": {"title": "Calibration Curve", "xlabel": "Concentration", "ylabel": "Signal"},
    "line": {"title": "Spectrum Plot", "xlabel": "Wavelength", "ylabel": "Intensity"}
}


def get_plot_default_labels(plot_type):
    return dict(PLOT_DEFAULT_LABELS.get(plot_type, {}))


def load_single_plot_config(json_path):
    """
    读取单个绘图配置 JSON。
    """

    with open(json_path, "r", encoding="utf-8") as file:
        return json.load(file)


def load_plot_experiment_configs(config_dir, table_style_builder=None):
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

            plot_func = PLOT_FUNCTIONS[plot_type]

            if table_style_builder is None:
                table_style = style_config
            else:
                table_style = table_style_builder(style_config)

            experiments[name] = {
                "headers": headers,
                "rows": rows,
                "cols": cols,
                "style": table_style,
                "plot_type": plot_type,
                "plot_func": plot_func,
                "plot_labels": get_plot_default_labels(plot_type),
                "required_cols": int(experiment_config.get("required_cols", 2))
            }

    return experiments


def plot_experiment(config, numeric_data, figure, labels=None):
    plot_func = config["plot_func"]
    plot_func(numeric_data, figure, labels=labels)


def save_figure(figure, file_path):
    figure.savefig(file_path, dpi=300, bbox_inches="tight")

