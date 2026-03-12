from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from math import sqrt
from typing import Any


class ProjectMode(str, Enum):
    MEASUREMENT = "measurement"
    FORMULA = "formula"

    @classmethod
    def from_value(cls, value: str) -> "ProjectMode":
        try:
            return cls(value)
        except ValueError:
            return cls.MEASUREMENT


class BSourceType(str, Enum):
    RESOLUTION = "resolution"
    TOLERANCE = "tolerance"
    GIVEN_STDDEV = "given_stddev"
    CUSTOM = "custom"

    @classmethod
    def from_value(cls, value: str) -> "BSourceType":
        try:
            return cls(value)
        except ValueError:
            return cls.RESOLUTION


B_SOURCE_DISPLAY_NAMES: dict[BSourceType, str] = {
    BSourceType.RESOLUTION: "仪器分度值",
    BSourceType.TOLERANCE: "仪器准确度/允差",
    BSourceType.GIVEN_STDDEV: "厂家给定标准差",
    BSourceType.CUSTOM: "自定义分布因子",
}

DEFAULT_DIVISORS: dict[BSourceType, float] = {
    BSourceType.RESOLUTION: sqrt(3),
    BSourceType.TOLERANCE: sqrt(3),
    BSourceType.GIVEN_STDDEV: 1.0,
    BSourceType.CUSTOM: sqrt(3),
}

PROJECT_MODE_DISPLAY_NAMES: dict[ProjectMode, str] = {
    ProjectMode.MEASUREMENT: "单一测量项目",
    ProjectMode.FORMULA: "公式项目",
}


def b_source_display_name(source_type: str) -> str:
    return B_SOURCE_DISPLAY_NAMES[BSourceType.from_value(source_type)]


def default_divisor_for(source_type: str) -> float:
    return DEFAULT_DIVISORS[BSourceType.from_value(source_type)]


def project_mode_display_name(project_mode: str) -> str:
    return PROJECT_MODE_DISPLAY_NAMES[ProjectMode.from_value(project_mode)]


def _safe_float(value: Any, default: float = 0.0) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def _safe_optional_int(value: Any) -> int | None:
    if value in (None, "", "null"):
        return None
    try:
        parsed = int(value)
    except (TypeError, ValueError):
        return None
    return max(0, parsed)


def _safe_dict(value: Any) -> dict[str, Any]:
    return value if isinstance(value, dict) else {}


@dataclass
class FormulaVariable:
    symbol: str = "X1"
    project_path: str | None = None
    source_label: str = ""
    quantity_name: str = ""
    unit: str = ""
    project_snapshot: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {
            "symbol": self.symbol,
            "project_path": self.project_path,
            "source_label": self.source_label,
            "quantity_name": self.quantity_name,
            "unit": self.unit,
            "project_snapshot": self.project_snapshot,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "FormulaVariable":
        return cls(
            symbol=str(data.get("symbol", "X1")).strip() or "X1",
            project_path=str(data.get("project_path", "")).strip() or None,
            source_label=str(data.get("source_label", "")).strip(),
            quantity_name=str(data.get("quantity_name", "")).strip(),
            unit=str(data.get("unit", "")).strip(),
            project_snapshot=_safe_dict(data.get("project_snapshot")),
        )


@dataclass
class BSource:
    source_type: str = BSourceType.RESOLUTION.value
    name: str = "B类分量"
    value: float = 0.0
    divisor: float = sqrt(3)
    notes: str = ""

    def normalized_divisor(self) -> float:
        source_type = BSourceType.from_value(self.source_type)
        if source_type == BSourceType.GIVEN_STDDEV:
            return 1.0
        return self.divisor if self.divisor > 0 else DEFAULT_DIVISORS[source_type]

    def to_dict(self) -> dict[str, Any]:
        return {
            "source_type": self.source_type,
            "name": self.name,
            "value": self.value,
            "divisor": self.normalized_divisor(),
            "notes": self.notes,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "BSource":
        source_type = BSourceType.from_value(str(data.get("source_type", BSourceType.RESOLUTION.value)))
        return cls(
            source_type=source_type.value,
            name=str(data.get("name", "B类分量")),
            value=_safe_float(data.get("value", 0.0)),
            divisor=_safe_float(data.get("divisor", DEFAULT_DIVISORS[source_type]), DEFAULT_DIVISORS[source_type]),
            notes=str(data.get("notes", "")),
        )


@dataclass
class ProjectData:
    project_mode: str = ProjectMode.MEASUREMENT.value
    quantity_name: str = ""
    unit: str = ""
    measured_values: list[float] = field(default_factory=list)
    b_sources: list[BSource] = field(default_factory=list)
    formula_expression: str = ""
    formula_variables: list[FormulaVariable] = field(default_factory=list)
    coverage_factor: float = 2.0
    result_decimal_places: int | None = None
    notes: str = ""
    project_path: str | None = None
    last_import_path: str | None = None
    created_at: str = ""
    updated_at: str = ""

    def ensure_defaults(self) -> None:
        if ProjectMode.from_value(self.project_mode) != ProjectMode.MEASUREMENT:
            return
        if not self.b_sources:
            self.b_sources = [
                BSource(
                    source_type=BSourceType.RESOLUTION.value,
                    name="仪器分度值",
                    value=0.0,
                    divisor=DEFAULT_DIVISORS[BSourceType.RESOLUTION],
                )
            ]

    def to_dict(self) -> dict[str, Any]:
        return {
            "project_mode": self.project_mode,
            "quantity_name": self.quantity_name,
            "unit": self.unit,
            "measured_values": self.measured_values,
            "b_sources": [source.to_dict() for source in self.b_sources],
            "formula_expression": self.formula_expression,
            "formula_variables": [variable.to_dict() for variable in self.formula_variables],
            "coverage_factor": self.coverage_factor,
            "result_decimal_places": self.result_decimal_places,
            "notes": self.notes,
            "last_import_path": self.last_import_path,
            "created_at": self.created_at,
            "updated_at": self.updated_at,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "ProjectData":
        project = cls(
            project_mode=ProjectMode.from_value(str(data.get("project_mode", ProjectMode.MEASUREMENT.value))).value,
            quantity_name=str(data.get("quantity_name", "")),
            unit=str(data.get("unit", "")),
            measured_values=[_safe_float(value) for value in data.get("measured_values", [])],
            b_sources=[BSource.from_dict(item) for item in data.get("b_sources", [])],
            formula_expression=str(data.get("formula_expression", "")),
            formula_variables=[FormulaVariable.from_dict(item) for item in data.get("formula_variables", [])],
            coverage_factor=_safe_float(data.get("coverage_factor", 2.0), 2.0),
            result_decimal_places=_safe_optional_int(data.get("result_decimal_places")),
            notes=str(data.get("notes", "")),
            last_import_path=data.get("last_import_path"),
            created_at=str(data.get("created_at", "")),
            updated_at=str(data.get("updated_at", "")),
        )
        project.ensure_defaults()
        return project