from __future__ import annotations

from dataclasses import dataclass, field
from math import floor, isfinite, log10, sqrt

from .models import BSource, BSourceType, ProjectData, b_source_display_name


@dataclass
class BSourceComponent:
    name: str
    source_type: str
    input_value: float
    divisor: float
    standard_uncertainty: float
    note: str


@dataclass
class CalculationResult:
    sample_count: int = 0
    mean: float = 0.0
    sample_std: float = 0.0
    type_a_uncertainty: float = 0.0
    type_b_uncertainty: float = 0.0
    combined_uncertainty: float = 0.0
    expanded_uncertainty: float = 0.0
    coverage_factor: float = 2.0
    b_components: list[BSourceComponent] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    process_rows: list[dict[str, str]] = field(default_factory=list)
    summary_text: str = ""


def calculate_project(project: ProjectData) -> CalculationResult:
    warnings: list[str] = []
    numeric_values = [float(value) for value in project.measured_values if isfinite(float(value))]
    sample_count = len(numeric_values)

    if sample_count == 0:
        mean = 0.0
        sample_std = 0.0
        type_a_uncertainty = 0.0
        warnings.append("A 类评定需要至少 1 个有效测量值；当前仅计算 B 类及其合成结果。")
    else:
        mean = sum(numeric_values) / sample_count
        if sample_count == 1:
            sample_std = 0.0
            type_a_uncertainty = 0.0
            warnings.append("A 类标准不确定度通常需要至少 2 次重复测量；当前按 0 处理。")
        else:
            squared_sum = sum((value - mean) ** 2 for value in numeric_values)
            sample_std = sqrt(squared_sum / (sample_count - 1))
            type_a_uncertainty = sample_std / sqrt(sample_count)

    b_components = [calculate_b_component(source) for source in project.b_sources]
    type_b_uncertainty = sqrt(sum(component.standard_uncertainty ** 2 for component in b_components))
    combined_uncertainty = sqrt(type_a_uncertainty ** 2 + type_b_uncertainty ** 2)
    coverage_factor = project.coverage_factor if project.coverage_factor > 0 else 2.0
    expanded_uncertainty = coverage_factor * combined_uncertainty

    result = CalculationResult(
        sample_count=sample_count,
        mean=mean,
        sample_std=sample_std,
        type_a_uncertainty=type_a_uncertainty,
        type_b_uncertainty=type_b_uncertainty,
        combined_uncertainty=combined_uncertainty,
        expanded_uncertainty=expanded_uncertainty,
        coverage_factor=coverage_factor,
        b_components=b_components,
        warnings=warnings,
    )
    result.process_rows = build_process_rows(project, result)
    result.summary_text = build_text_report(project, result)
    return result


def calculate_b_component(source: BSource) -> BSourceComponent:
    source_type = BSourceType.from_value(source.source_type)
    input_value = abs(float(source.value))
    divisor = source.normalized_divisor()

    if source_type == BSourceType.GIVEN_STDDEV:
        standard_uncertainty = input_value
        note = "厂家直接给出标准差，按标准不确定度使用。"
    else:
        standard_uncertainty = input_value / divisor if divisor > 0 else 0.0
        if source_type == BSourceType.RESOLUTION:
            note = "分度值半宽按均匀分布处理，u = a / 分布因子。"
        elif source_type == BSourceType.TOLERANCE:
            note = "准确度或允差按均匀分布处理，u = a / 分布因子。"
        else:
            note = "使用自定义分布因子计算标准不确定度。"

    return BSourceComponent(
        name=source.name or b_source_display_name(source.source_type),
        source_type=source.source_type,
        input_value=input_value,
        divisor=divisor,
        standard_uncertainty=standard_uncertainty,
        note=note,
    )


def build_process_rows(project: ProjectData, result: CalculationResult) -> list[dict[str, str]]:
    decimal_places = normalize_decimal_places(project.result_decimal_places)
    rows = [
        {
            "类别": "A类评定",
            "项目": "测量次数 n",
            "数值": str(result.sample_count),
            "说明": "重复测量有效数据点数量。",
        },
        {
            "类别": "A类评定",
            "项目": "平均值 x̄",
            "数值": format_number(result.mean, decimal_places),
            "说明": "x̄ = Σxi / n",
        },
        {
            "类别": "A类评定",
            "项目": "样本标准差 s",
            "数值": format_number(result.sample_std, decimal_places),
            "说明": "s = sqrt(Σ(xi - x̄)^2 / (n - 1))",
        },
        {
            "类别": "A类评定",
            "项目": "A类标准不确定度 uA",
            "数值": format_number(result.type_a_uncertainty, decimal_places),
            "说明": "uA = s / sqrt(n)",
        },
    ]

    for index, component in enumerate(result.b_components, start=1):
        rows.append(
            {
                "类别": "B类评定",
                "项目": f"{index}. {component.name}",
                "数值": format_number(component.standard_uncertainty, decimal_places),
                "说明": component.note,
            }
        )

    rows.extend(
        [
            {
                "类别": "结果",
                "项目": "B类合成标准不确定度 uB",
                "数值": format_number(result.type_b_uncertainty, decimal_places),
                "说明": "uB = sqrt(ΣuBi^2)",
            },
            {
                "类别": "结果",
                "项目": "合成标准不确定度 uc",
                "数值": format_number(result.combined_uncertainty, decimal_places),
                "说明": "uc = sqrt(uA^2 + uB^2)",
            },
            {
                "类别": "结果",
                "项目": "扩展不确定度 U",
                "数值": format_number(result.expanded_uncertainty, decimal_places),
                "说明": f"U = k * uc, k = {format_number(result.coverage_factor)}",
            },
        ]
    )
    return rows


def build_text_report(project: ProjectData, result: CalculationResult) -> str:
    decimal_places = normalize_decimal_places(project.result_decimal_places)
    quantity_name = project.quantity_name or "测量量"
    unit_suffix = f" {project.unit}" if project.unit else ""
    rounded_value, rounded_uncertainty = rounded_measurement(
        result.mean,
        result.expanded_uncertainty,
        decimal_places,
    )
    lines = [
        "物理实验不确定度计算报告",
        "=" * 32,
        f"测量量: {quantity_name}",
        f"单位: {project.unit or '-'}",
        f"最终结果: {quantity_name} = ({rounded_value} ± {rounded_uncertainty}){unit_suffix}, k = {format_number(result.coverage_factor)}",
        "",
        "A类评定",
        f"- 测量次数 n = {result.sample_count}",
        f"- 平均值 x̄ = {format_number(result.mean, decimal_places)}{unit_suffix}",
        f"- 样本标准差 s = {format_number(result.sample_std, decimal_places)}{unit_suffix}",
        f"- A类标准不确定度 uA = {format_number(result.type_a_uncertainty, decimal_places)}{unit_suffix}",
        "",
        "B类评定",
    ]

    if result.b_components:
        for component in result.b_components:
            lines.append(
                f"- {component.name}: u = {format_number(component.standard_uncertainty, decimal_places)}{unit_suffix} ({component.note})"
            )
    else:
        lines.append("- 当前没有 B 类分量。")

    lines.extend(
        [
            "",
            "合成结果",
            f"- B类合成标准不确定度 uB = {format_number(result.type_b_uncertainty, decimal_places)}{unit_suffix}",
            f"- 合成标准不确定度 uc = {format_number(result.combined_uncertainty, decimal_places)}{unit_suffix}",
            f"- 扩展不确定度 U = {format_number(result.expanded_uncertainty, decimal_places)}{unit_suffix}",
        ]
    )

    if project.notes.strip():
        lines.extend(["", "备注", project.notes.strip()])

    if result.warnings:
        lines.extend(["", "提示"])
        lines.extend(f"- {warning}" for warning in result.warnings)

    return "\n".join(lines)
def rounded_measurement_with_decimals(value: float, uncertainty: float, decimal_places: int | None) -> tuple[str, str]:
    decimal_places = normalize_decimal_places(decimal_places)
    if decimal_places is not None:
        return format_rounded(value, decimal_places), format_rounded(uncertainty, decimal_places)

    if uncertainty <= 0 or not isfinite(uncertainty):
        return format_number(value), format_number(uncertainty)

    exponent = floor(log10(abs(uncertainty)))
    leading_digit = int(abs(uncertainty) / (10 ** exponent))
    significant_digits = 2 if leading_digit in (1, 2) else 1
    decimals = significant_digits - exponent - 1

    rounded_uncertainty = round(uncertainty, decimals)
    rounded_value = round(value, decimals)
    return format_rounded(rounded_value, decimals), format_rounded(rounded_uncertainty, decimals)


def rounded_measurement(value: float, uncertainty: float, decimal_places: int | None = None) -> tuple[str, str]:
    return rounded_measurement_with_decimals(value, uncertainty, decimal_places)


def format_rounded(value: float, decimals: int) -> str:
    normalized = 0.0 if abs(value) < 1e-15 else value
    if decimals >= 0:
        return f"{normalized:.{decimals}f}"
    return f"{round(normalized, decimals):.0f}"


def normalize_decimal_places(decimal_places: int | None) -> int | None:
    if decimal_places is None:
        return None
    return max(0, int(decimal_places))


def format_number(value: float, decimal_places: int | None = None) -> str:
    decimal_places = normalize_decimal_places(decimal_places)
    if decimal_places is not None:
        return format_rounded(value, decimal_places)
    if not isfinite(value):
        return "0"
    if abs(value) >= 10000 or (0 < abs(value) < 0.001):
        return f"{value:.4e}"
    return f"{value:.6f}".rstrip("0").rstrip(".") or "0"