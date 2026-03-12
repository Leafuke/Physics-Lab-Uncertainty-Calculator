from __future__ import annotations

import ast
import math
from dataclasses import dataclass, field


class FormulaExpressionError(ValueError):
    pass


ALLOWED_FUNCTIONS = {
    "sqrt": math.sqrt,
    "sin": math.sin,
    "cos": math.cos,
    "tan": math.tan,
    "asin": math.asin,
    "acos": math.acos,
    "atan": math.atan,
    "exp": math.exp,
    "log": math.log,
    "ln": math.log,
    "log10": math.log10,
    "abs": abs,
    "pow": pow,
}

ALLOWED_CONSTANTS = {
    "pi": math.pi,
    "e": math.e,
}


@dataclass(frozen=True)
class UnitExpression:
    factors: dict[str, float] = field(default_factory=dict)

    @classmethod
    def from_unit_text(cls, unit_text: str) -> "UnitExpression":
        normalized = unit_text.strip()
        if not normalized:
            return cls()
        return cls({normalized: 1.0})

    def is_dimensionless(self) -> bool:
        return not self.factors

    def multiply(self, other: "UnitExpression") -> "UnitExpression":
        merged = dict(self.factors)
        for key, exponent in other.factors.items():
            merged[key] = merged.get(key, 0.0) + exponent
            if abs(merged[key]) < 1e-12:
                merged.pop(key)
        return UnitExpression(merged)

    def divide(self, other: "UnitExpression") -> "UnitExpression":
        merged = dict(self.factors)
        for key, exponent in other.factors.items():
            merged[key] = merged.get(key, 0.0) - exponent
            if abs(merged[key]) < 1e-12:
                merged.pop(key)
        return UnitExpression(merged)

    def power(self, exponent: float) -> "UnitExpression":
        powered = {key: value * exponent for key, value in self.factors.items() if abs(value * exponent) >= 1e-12}
        return UnitExpression(powered)

    def format(self) -> str:
        numerator: list[str] = []
        denominator: list[str] = []
        for token, exponent in sorted(self.factors.items()):
            formatted = _format_unit_token(token, abs(exponent))
            if exponent > 0:
                numerator.append(formatted)
            elif exponent < 0:
                denominator.append(formatted)

        if not numerator and not denominator:
            return ""

        numerator_text = "·".join(numerator) if numerator else "1"
        if not denominator:
            return numerator_text

        denominator_text = "·".join(denominator)
        return f"{numerator_text}/{denominator_text}"


def split_expression_assignment(expression_text: str) -> tuple[str, str]:
    cleaned = expression_text.strip()
    if "=" not in cleaned:
        return "", cleaned
    left, right = cleaned.split("=", 1)
    return left.strip(), right.strip()


def list_expression_symbols(expression_text: str) -> list[str]:
    expression = _expression_only(expression_text)
    tree = _parse_expression(expression)
    names: list[str] = []

    for node in ast.walk(tree):
        if isinstance(node, ast.Name):
            if node.id in ALLOWED_FUNCTIONS or node.id in ALLOWED_CONSTANTS:
                continue
            if node.id not in names:
                names.append(node.id)

    return names


def evaluate_expression(expression_text: str, variables: dict[str, float]) -> float:
    expression = _expression_only(expression_text)
    tree = _parse_expression(expression)
    try:
        return float(_evaluate_node(tree.body, variables))
    except FormulaExpressionError:
        raise
    except Exception as exc:
        raise FormulaExpressionError(f"表达式计算失败：{exc}") from exc


def estimate_partial_derivative(expression_text: str, variables: dict[str, float], symbol: str, uncertainty: float) -> float:
    if symbol not in variables:
        raise FormulaExpressionError(f"表达式中缺少变量 {symbol} 的数值。")

    base_value = evaluate_expression(expression_text, variables)
    value = variables[symbol]
    step = max(abs(value) * 1e-8, abs(uncertainty) * 1e-4, 1e-8)

    for _ in range(8):
        plus_variables = dict(variables)
        plus_variables[symbol] = value + step
        minus_variables = dict(variables)
        minus_variables[symbol] = value - step

        plus_value = _try_evaluate_expression(expression_text, plus_variables)
        minus_value = _try_evaluate_expression(expression_text, minus_variables)
        if plus_value is not None and minus_value is not None:
            return (plus_value - minus_value) / (2 * step)
        if plus_value is not None:
            return (plus_value - base_value) / step
        if minus_value is not None:
            return (base_value - minus_value) / step
        step *= 0.5

    raise FormulaExpressionError(f"无法为变量 {symbol} 计算灵敏系数，请检查表达式定义域。")


def derive_expression_unit(expression_text: str, variable_units: dict[str, str]) -> tuple[str, list[str]]:
    expression = _expression_only(expression_text)
    tree = _parse_expression(expression)
    warnings: list[str] = []
    unit_expression = _derive_unit(tree.body, variable_units, warnings)
    return unit_expression.format(), warnings


def _expression_only(expression_text: str) -> str:
    _, expression = split_expression_assignment(expression_text)
    if not expression:
        raise FormulaExpressionError("公式项目需要输入表达式。")
    return expression


def _parse_expression(expression_text: str) -> ast.Expression:
    try:
        tree = ast.parse(expression_text, mode="eval")
    except SyntaxError as exc:
        raise FormulaExpressionError(f"表达式语法错误：{exc.msg}") from exc
    return tree


def _evaluate_node(node: ast.AST, variables: dict[str, float]) -> float:
    if isinstance(node, ast.Constant):
        if isinstance(node.value, (int, float)):
            return float(node.value)
        raise FormulaExpressionError("表达式中包含不支持的常量类型。")

    if isinstance(node, ast.Name):
        if node.id in ALLOWED_CONSTANTS:
            return ALLOWED_CONSTANTS[node.id]
        if node.id in variables:
            return float(variables[node.id])
        raise FormulaExpressionError(f"表达式中引用了未导入的变量 {node.id}。")

    if isinstance(node, ast.UnaryOp):
        operand = _evaluate_node(node.operand, variables)
        if isinstance(node.op, ast.UAdd):
            return operand
        if isinstance(node.op, ast.USub):
            return -operand
        raise FormulaExpressionError("表达式中包含不支持的一元运算。")

    if isinstance(node, ast.BinOp):
        left = _evaluate_node(node.left, variables)
        right = _evaluate_node(node.right, variables)
        if isinstance(node.op, ast.Add):
            return left + right
        if isinstance(node.op, ast.Sub):
            return left - right
        if isinstance(node.op, ast.Mult):
            return left * right
        if isinstance(node.op, ast.Div):
            return left / right
        if isinstance(node.op, ast.Pow):
            return left**right
        if isinstance(node.op, ast.BitXor):
            raise FormulaExpressionError("请使用 ** 表示乘方，而不是 ^。")
        raise FormulaExpressionError("表达式中包含不支持的二元运算。")

    if isinstance(node, ast.Call):
        if not isinstance(node.func, ast.Name) or node.func.id not in ALLOWED_FUNCTIONS:
            raise FormulaExpressionError("表达式中调用了不支持的函数。")
        function = ALLOWED_FUNCTIONS[node.func.id]
        arguments = [_evaluate_node(argument, variables) for argument in node.args]
        try:
            return float(function(*arguments))
        except TypeError as exc:
            raise FormulaExpressionError(f"函数 {node.func.id} 的参数数量不正确。") from exc

    raise FormulaExpressionError("表达式中包含不支持的语法。")


def _derive_unit(node: ast.AST, variable_units: dict[str, str], warnings: list[str]) -> UnitExpression:
    if isinstance(node, ast.Constant):
        return UnitExpression()

    if isinstance(node, ast.Name):
        if node.id in ALLOWED_CONSTANTS:
            return UnitExpression()
        return UnitExpression.from_unit_text(variable_units.get(node.id, ""))

    if isinstance(node, ast.UnaryOp):
        return _derive_unit(node.operand, variable_units, warnings)

    if isinstance(node, ast.BinOp):
        left = _derive_unit(node.left, variable_units, warnings)
        right = _derive_unit(node.right, variable_units, warnings)
        if isinstance(node.op, (ast.Add, ast.Sub)):
            if left.factors != right.factors:
                warnings.append("表达式中存在单位不一致的加减项，结果单位可能不可靠。")
                return UnitExpression()
            return left
        if isinstance(node.op, ast.Mult):
            return left.multiply(right)
        if isinstance(node.op, ast.Div):
            return left.divide(right)
        if isinstance(node.op, ast.Pow):
            exponent = _constant_value(node.right)
            if exponent is None:
                warnings.append("乘方的指数不是常量，结果单位无法自动推导。")
                return UnitExpression()
            return left.power(exponent)
        if isinstance(node.op, ast.BitXor):
            warnings.append("检测到 ^ 运算，请改用 ** 表示乘方。")
            return UnitExpression()
        warnings.append("表达式中包含无法识别的单位运算。")
        return UnitExpression()

    if isinstance(node, ast.Call) and isinstance(node.func, ast.Name):
        function_name = node.func.id
        arguments = [_derive_unit(argument, variable_units, warnings) for argument in node.args]
        if function_name == "sqrt" and arguments:
            return arguments[0].power(0.5)
        if function_name == "abs" and arguments:
            return arguments[0]
        if function_name == "pow" and len(node.args) == 2:
            exponent = _constant_value(node.args[1])
            if exponent is None:
                warnings.append("pow 的指数不是常量，结果单位无法自动推导。")
                return UnitExpression()
            return arguments[0].power(exponent)
        if function_name in {"sin", "cos", "tan", "asin", "acos", "atan", "exp", "log", "ln", "log10"}:
            if arguments and not arguments[0].is_dimensionless():
                warnings.append(f"函数 {function_name} 通常要求无量纲输入，已按无量纲结果处理。")
            return UnitExpression()
        warnings.append(f"函数 {function_name} 的单位推导暂不支持，已按无量纲结果处理。")
        return UnitExpression()

    warnings.append("表达式中存在无法自动推导单位的语法。")
    return UnitExpression()


def _constant_value(node: ast.AST) -> float | None:
    if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
        return float(node.value)
    if isinstance(node, ast.Name) and node.id in ALLOWED_CONSTANTS:
        return ALLOWED_CONSTANTS[node.id]
    if isinstance(node, ast.UnaryOp) and isinstance(node.op, ast.USub):
        operand = _constant_value(node.operand)
        return -operand if operand is not None else None
    if isinstance(node, ast.UnaryOp) and isinstance(node.op, ast.UAdd):
        return _constant_value(node.operand)
    if isinstance(node, ast.BinOp):
        left = _constant_value(node.left)
        right = _constant_value(node.right)
        if left is None or right is None:
            return None
        if isinstance(node.op, ast.Add):
            return left + right
        if isinstance(node.op, ast.Sub):
            return left - right
        if isinstance(node.op, ast.Mult):
            return left * right
        if isinstance(node.op, ast.Div):
            return left / right
        if isinstance(node.op, ast.Pow):
            return left**right
    return None


def _try_evaluate_expression(expression_text: str, variables: dict[str, float]) -> float | None:
    try:
        return evaluate_expression(expression_text, variables)
    except FormulaExpressionError:
        return None


def _format_unit_token(token: str, exponent: float) -> str:
    needs_parentheses = any(character in token for character in " /·^*")
    rendered = f"({token})" if needs_parentheses else token
    if abs(exponent - 1.0) < 1e-12:
        return rendered
    rounded = int(exponent) if abs(exponent - round(exponent)) < 1e-12 else round(exponent, 4)
    return f"{rendered}^{rounded}"