from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from math import floor, isfinite, log10, sqrt

from .formula_engine import (
    FormulaExpressionError,
    derive_expression_unit,
    estimate_partial_derivative,
    evaluate_expression,
    list_expression_symbols,
    split_expression_assignment,
)
from .models import BSource, BSourceType, FormulaVariable, ProjectData, ProjectMode, b_source_display_name
from .persistence import load_project_file


@dataclass
class BSourceComponent:
    name: str
    source_type: str
    input_value: float
    divisor: float
    standard_uncertainty: float
    note: str


@dataclass
class FormulaVariableResult:
    symbol: str
    quantity_name: str
    unit: str
    value: float
    standard_uncertainty: float
    expanded_uncertainty: float
    sensitivity_coefficient: float
    contribution: float
    source_path: str | None
    source_label: str
    status: str


@dataclass
class CalculationResult:
    project_mode: str = ProjectMode.MEASUREMENT.value
    sample_count: int = 0
    mean: float = 0.0
    sample_std: float = 0.0
    type_a_uncertainty: float = 0.0
    type_b_uncertainty: float = 0.0
    combined_uncertainty: float = 0.0
    expanded_uncertainty: float = 0.0
    coverage_factor: float = 2.0
    b_components: list[BSourceComponent] = field(default_factory=list)
    formula_variables: list[FormulaVariableResult] = field(default_factory=list)
    expression: str = ""
    resolved_quantity_name: str = ""
    resolved_unit: str = ""
    dominant_component: str = ""
    warnings: list[str] = field(default_factory=list)
    process_rows: list[dict[str, str]] = field(default_factory=list)
    summary_text: str = ""


def calculate_project(project: ProjectData, visited_project_paths: set[str] | None = None) -> CalculationResult:
    project.ensure_defaults()
    visited = set(visited_project_paths or set())
    project_path = _normalized_project_path(project.project_path)
    if project_path:
        visited.add(project_path)

    if ProjectMode.from_value(project.project_mode) == ProjectMode.FORMULA:
        return _calculate_formula_project(project, visited)
    return _calculate_measurement_project(project)


def _calculate_measurement_project(project: ProjectData) -> CalculationResult:
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
        project_mode=ProjectMode.MEASUREMENT.value,
        sample_count=sample_count,
        mean=mean,
        sample_std=sample_std,
        type_a_uncertainty=type_a_uncertainty,
        type_b_uncertainty=type_b_uncertainty,
        combined_uncertainty=combined_uncertainty,
        expanded_uncertainty=expanded_uncertainty,
        coverage_factor=coverage_factor,
        b_components=b_components,
        resolved_quantity_name=project.quantity_name.strip() or "测量量",
        resolved_unit=project.unit.strip(),
        warnings=warnings,
    )
    result.process_rows = build_process_rows(project, result)
    result.summary_text = build_text_report(project, result)
    return result


def _calculate_formula_project(project: ProjectData, visited_project_paths: set[str]) -> CalculationResult:
    warnings: list[str] = []
    coverage_factor = project.coverage_factor if project.coverage_factor > 0 else 2.0
    assignment_name, _ = split_expression_assignment(project.formula_expression)
    resolved_quantity_name = project.quantity_name.strip() or assignment_name or "结果量"
    result = CalculationResult(
        project_mode=ProjectMode.FORMULA.value,
        coverage_factor=coverage_factor,
        expression=project.formula_expression.strip(),
        resolved_quantity_name=resolved_quantity_name,
        resolved_unit=project.unit.strip(),
        warnings=warnings,
    )

    if not project.formula_expression.strip():
        warnings.append("公式项目需要输入表达式。")
        result.process_rows = build_process_rows(project, result)
        result.summary_text = build_text_report(project, result)
        return result

    variable_values: dict[str, float] = {}
    variable_uncertainties: dict[str, float] = {}
    variable_units: dict[str, str] = {}
    resolved_variables: list[FormulaVariableResult] = []
    used_symbols: set[str] = set()

    for variable in project.formula_variables:
        symbol = variable.symbol.strip()
        if not symbol:
            warnings.append("协同工作区中存在空白变量符号，已忽略。")
            continue
        if symbol in used_symbols:
            warnings.append(f"变量符号 {symbol} 重复，后续同名变量已忽略。")
            continue

        resolved_variable, resolve_warnings = _resolve_formula_variable(variable, visited_project_paths)
        warnings.extend(resolve_warnings)
        if resolved_variable is None:
            continue

        used_symbols.add(symbol)
        resolved_variables.append(resolved_variable)
        variable_values[symbol] = resolved_variable.value
        variable_uncertainties[symbol] = resolved_variable.standard_uncertainty
        variable_units[symbol] = resolved_variable.unit

    result.formula_variables = resolved_variables
    result.sample_count = len(resolved_variables)

    try:
        expression_symbols = list_expression_symbols(project.formula_expression)
    except FormulaExpressionError as exc:
        warnings.append(str(exc))
        result.process_rows = build_process_rows(project, result)
        result.summary_text = build_text_report(project, result)
        return result

    missing_symbols = [symbol for symbol in expression_symbols if symbol not in variable_values]
    if missing_symbols:
        warnings.append(f"表达式中存在未导入的变量：{', '.join(missing_symbols)}。")

    unused_symbols = [variable.symbol for variable in resolved_variables if variable.symbol not in expression_symbols]
    if unused_symbols:
        warnings.append(f"协同工作区中有未参与当前表达式的变量：{', '.join(unused_symbols)}。")

    try:
        result.mean = evaluate_expression(project.formula_expression, variable_values)
    except FormulaExpressionError as exc:
        warnings.append(str(exc))
        result.process_rows = build_process_rows(project, result)
        result.summary_text = build_text_report(project, result)
        return result

    derived_unit, unit_warnings = derive_expression_unit(project.formula_expression, variable_units)
    warnings.extend(unit_warnings)
    if not result.resolved_unit:
        result.resolved_unit = derived_unit

    variance_sum = 0.0
    for index, variable in enumerate(result.formula_variables):
        try:
            sensitivity = estimate_partial_derivative(project.formula_expression, variable_values, variable.symbol, variable.standard_uncertainty)
        except FormulaExpressionError as exc:
            warnings.append(str(exc))
            sensitivity = 0.0

        contribution = abs(sensitivity) * variable.standard_uncertainty
        variance_sum += (sensitivity * variable.standard_uncertainty) ** 2
        result.formula_variables[index] = FormulaVariableResult(
            symbol=variable.symbol,
            quantity_name=variable.quantity_name,
            unit=variable.unit,
            value=variable.value,
            standard_uncertainty=variable.standard_uncertainty,
            expanded_uncertainty=variable.expanded_uncertainty,
            sensitivity_coefficient=sensitivity,
            contribution=contribution,
            source_path=variable.source_path,
            source_label=variable.source_label,
            status=variable.status,
        )

    result.combined_uncertainty = sqrt(variance_sum)
    result.type_b_uncertainty = result.combined_uncertainty
    result.expanded_uncertainty = coverage_factor * result.combined_uncertainty

    if result.formula_variables and result.combined_uncertainty > 0:
        dominant = max(result.formula_variables, key=lambda item: item.contribution)
        share = (dominant.contribution**2) / (result.combined_uncertainty**2) * 100 if result.combined_uncertainty > 0 else 0.0
        result.dominant_component = f"{dominant.symbol} ({share:.1f}%)"
    elif result.formula_variables:
        result.dominant_component = result.formula_variables[0].symbol

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


def _resolve_formula_variable(
    variable: FormulaVariable,
    visited_project_paths: set[str],
) -> tuple[FormulaVariableResult | None, list[str]]:
    warnings: list[str] = []
    project_path = _normalized_project_path(variable.project_path)
    source_project: ProjectData | None = None
    status = "使用保存快照"

    if project_path:
        if project_path in visited_project_paths:
            warnings.append(f"变量 {variable.symbol} 检测到循环引用：{Path(project_path).name}。")
            return None, warnings
        try:
            source_project = load_project_file(project_path)
            status = "已加载实时项目"
        except Exception as exc:
            if variable.project_snapshot:
                source_project = ProjectData.from_dict(variable.project_snapshot)
                source_project.project_path = variable.project_path
                status = "源项目不可用，已回退到保存快照"
                warnings.append(f"变量 {variable.symbol} 无法读取 {Path(project_path).name}，已回退到保存快照。")
            else:
                warnings.append(f"变量 {variable.symbol} 无法读取来源项目：{exc}")
                return None, warnings
    elif variable.project_snapshot:
        source_project = ProjectData.from_dict(variable.project_snapshot)
    else:
        warnings.append(f"变量 {variable.symbol} 尚未关联有效项目来源。")
        return None, warnings

    if source_project is None:
        warnings.append(f"变量 {variable.symbol} 无法解析来源项目。")
        return None, warnings

    nested_result = calculate_project(source_project, visited_project_paths)
    quantity_name = nested_result.resolved_quantity_name or source_project.quantity_name.strip() or variable.quantity_name or variable.symbol
    resolved_unit = nested_result.resolved_unit or source_project.unit.strip() or variable.unit
    source_label = variable.source_label.strip() or (Path(variable.project_path).name if variable.project_path else quantity_name)
    return (
        FormulaVariableResult(
            symbol=variable.symbol.strip(),
            quantity_name=quantity_name,
            unit=resolved_unit,
            value=nested_result.mean,
            standard_uncertainty=nested_result.combined_uncertainty,
            expanded_uncertainty=nested_result.expanded_uncertainty,
            sensitivity_coefficient=0.0,
            contribution=0.0,
            source_path=variable.project_path,
            source_label=source_label,
            status=status,
        ),
        warnings,
    )


def build_process_rows(project: ProjectData, result: CalculationResult) -> list[dict[str, str]]:
    if ProjectMode.from_value(project.project_mode) == ProjectMode.FORMULA:
        return _build_formula_process_rows(project, result)

    return _build_measurement_process_rows(project, result)


def _build_measurement_process_rows(project: ProjectData, result: CalculationResult) -> list[dict[str, str]]:
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


def _build_formula_process_rows(project: ProjectData, result: CalculationResult) -> list[dict[str, str]]:
    decimal_places = normalize_decimal_places(project.result_decimal_places)
    rows = [
        {
            "类别": "公式项目",
            "项目": "表达式",
            "数值": result.expression or "-",
            "说明": "按输入量相互独立的一阶不确定度传播模型计算。",
        },
        {
            "类别": "公式项目",
            "项目": "输入量数量 n",
            "数值": str(len(result.formula_variables)),
            "说明": "当前协同工作区中成功参与表达式计算的项目数量。",
        },
    ]

    for variable in result.formula_variables:
        unit_suffix = f" {variable.unit}" if variable.unit else ""
        rows.append(
            {
                "类别": "输入量",
                "项目": f"{variable.symbol}: {variable.quantity_name}",
                "数值": format_number(variable.value, decimal_places),
                "说明": (
                    f"uc = {format_number(variable.standard_uncertainty, decimal_places)}{unit_suffix}，"
                    f"c = {format_number(variable.sensitivity_coefficient, decimal_places)}，"
                    f"|c|u = {format_number(variable.contribution, decimal_places)}，"
                    f"来源：{variable.source_label}，状态：{variable.status}"
                ),
            }
        )

    result_unit_suffix = f" {result.resolved_unit}" if result.resolved_unit else ""
    rows.extend(
        [
            {
                "类别": "结果",
                "项目": "结果值 y",
                "数值": format_number(result.mean, decimal_places),
                "说明": "将协同工作区变量代入表达式得到的结果值。",
            },
            {
                "类别": "结果",
                "项目": "合成标准不确定度 uc",
                "数值": format_number(result.combined_uncertainty, decimal_places),
                "说明": f"uc = sqrt(Σ((∂f/∂xi)ui)^2){result_unit_suffix}",
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
    if ProjectMode.from_value(project.project_mode) == ProjectMode.FORMULA:
        return _build_formula_text_report(project, result)

    return _build_measurement_text_report(project, result)


def _build_measurement_text_report(project: ProjectData, result: CalculationResult) -> str:
    decimal_places = normalize_decimal_places(project.result_decimal_places)
    quantity_name = result.resolved_quantity_name or project.quantity_name or "测量量"
    unit_suffix = f" {result.resolved_unit}" if result.resolved_unit else ""
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


def _build_formula_text_report(project: ProjectData, result: CalculationResult) -> str:
    decimal_places = normalize_decimal_places(project.result_decimal_places)
    quantity_name = result.resolved_quantity_name or project.quantity_name or "结果量"
    unit_suffix = f" {result.resolved_unit}" if result.resolved_unit else ""
    rounded_value, rounded_uncertainty = rounded_measurement(
        result.mean,
        result.expanded_uncertainty,
        decimal_places,
    )
    lines = [
        "物理实验不确定度计算报告",
        "=" * 32,
        "项目类型: 公式项目",
        f"结果量: {quantity_name}",
        f"表达式: {result.expression or '-'}",
        f"单位: {result.resolved_unit or '-'}",
        f"最终结果: {quantity_name} = ({rounded_value} ± {rounded_uncertainty}){unit_suffix}, k = {format_number(result.coverage_factor)}",
        "",
        "协同工作区",
    ]

    if result.formula_variables:
        for variable in result.formula_variables:
            variable_unit_suffix = f" {variable.unit}" if variable.unit else ""
            lines.append(
                "- "
                f"{variable.symbol}: {variable.quantity_name}, 值 = {format_number(variable.value, decimal_places)}{variable_unit_suffix}, "
                f"uc = {format_number(variable.standard_uncertainty, decimal_places)}{variable_unit_suffix}, "
                f"c = {format_number(variable.sensitivity_coefficient, decimal_places)}, "
                f"来源 = {variable.source_label} [{variable.status}]"
            )
    else:
        lines.append("- 当前协同工作区没有可参与计算的项目变量。")

    lines.extend(
        [
            "",
            "合成结果",
            f"- 结果值 y = {format_number(result.mean, decimal_places)}{unit_suffix}",
            f"- 合成标准不确定度 uc = {format_number(result.combined_uncertainty, decimal_places)}{unit_suffix}",
            f"- 扩展不确定度 U = {format_number(result.expanded_uncertainty, decimal_places)}{unit_suffix}",
        ]
    )

    if result.dominant_component:
        lines.append(f"- 主导输入量 = {result.dominant_component}")

    if project.notes.strip():
        lines.extend(["", "备注", project.notes.strip()])

    if result.warnings:
        lines.extend(["", "提示"])
        lines.extend(f"- {warning}" for warning in result.warnings)

    return "\n".join(lines)


def _normalized_project_path(path: str | None) -> str | None:
    if not path:
        return None
    try:
        return str(Path(path).resolve())
    except OSError:
        return str(Path(path))
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