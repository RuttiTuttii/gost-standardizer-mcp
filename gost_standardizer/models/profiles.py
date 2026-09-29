from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any, Literal

PresetName = Literal["report", "office", "technical", "legacy-college"]


@dataclass
class MarginsProfile:
    left_mm: float = 30.0
    right_mm: float = 15.0
    top_mm: float = 20.0
    bottom_mm: float = 20.0


@dataclass
class FontProfile:
    name: str = "Times New Roman"
    body_size_pt: float = 14.0
    heading_size_pt: float = 14.0
    title_size_pt: float = 16.0
    table_size_pt: float = 12.0
    caption_size_pt: float = 12.0


@dataclass
class SpacingProfile:
    line_spacing: float = 1.5
    space_after_pt: float = 0.0
    space_before_pt: float = 0.0
    first_line_indent_mm: float = 12.5


@dataclass
class DocumentProfile:
    name: str
    preset: str = "report"
    title: str = ""
    description: str = ""
    kind: str = "standard"
    margins: MarginsProfile = field(default_factory=MarginsProfile)
    fonts: FontProfile = field(default_factory=FontProfile)
    spacing: SpacingProfile = field(default_factory=SpacingProfile)
    notes: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> DocumentProfile:
        margins_data = data.get("margins", {})
        fonts_data = data.get("fonts", {})
        spacing_data = data.get("spacing", {})
        return cls(
            name=data.get("name", "custom"),
            preset=data.get("preset", "report"),
            title=data.get("title", ""),
            description=data.get("description", ""),
            kind=data.get("kind", "organization"),
            margins=MarginsProfile(**margins_data) if margins_data else MarginsProfile(),
            fonts=FontProfile(**fonts_data) if fonts_data else FontProfile(),
            spacing=SpacingProfile(**spacing_data) if spacing_data else SpacingProfile(),
            notes=list(data.get("notes", [])),
        )
