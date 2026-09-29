from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any


@dataclass
class ValidationIssue:
    kind: str
    message: str
    severity: str = "warning"
    expected: Any = None
    actual: Any = None
    location: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class InspectionReport:
    path: str
    format: str
    paragraphs_count: int
    tables_count: int
    suggested_preset: str
    detected_fonts: list[str] = field(default_factory=list)
    issues: list[ValidationIssue] = field(default_factory=list)
    samples: list[dict[str, Any]] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        data["issues"] = [issue.to_dict() if hasattr(issue, "to_dict") else issue for issue in self.issues]
        return data


@dataclass
class StandardizationResult:
    source_path: str
    output_path: str
    preset: str
    changed_paragraphs: int
    changed_tables: int
    margins_adjusted: bool
    issues_fixed: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)
