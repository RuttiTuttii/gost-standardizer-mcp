from __future__ import annotations

from dataclasses import asdict, dataclass, field
from enum import Enum
import re
from typing import Any


class DocumentStatus(str, Enum):
    ACTIVE = "действующий"
    REPLACED = "заменён"
    CANCELLED = "отменён"
    NOT_YET_ACTIVE = "не вступил в силу"
    UNKNOWN = "не указан"

    @classmethod
    def from_str(cls, value: str | None) -> DocumentStatus:
        if not value:
            return cls.UNKNOWN
        v = value.strip().lower()
        if "действ" in v and "не" not in v:
            return cls.ACTIVE
        if "замен" in v:
            return cls.REPLACED
        if "отмен" in v or "утратил" in v:
            return cls.CANCELLED
        if "не вступил" in v:
            return cls.NOT_YET_ACTIVE
        return cls.UNKNOWN


def normalize_gost_number(query: str) -> str:
    """Normalize a GOST query into a comparable designation and clean number.

    Examples:
        'ГОСТ Р 7.0.97-2025' -> '7.0.97-2025'
        'гост  7.32 - 2017' -> '7.32-2017'
        'Р 7.0.97' -> '7.0.97'
    """
    raw = query.strip()
    # Strip common prefixes
    raw = re.sub(r"(?i)^(?:гост\s+р|гост|gost\s+r|gost)\s*", "", raw)
    # Normalize spaces around hyphens and dots
    raw = re.sub(r"\s*-\s*", "-", raw)
    raw = re.sub(r"\s*\.\s*", ".", raw)
    return raw.strip()


@dataclass
class NormDocument:
    id: str
    designation: str
    clean_number: str
    title: str
    status: str = DocumentStatus.ACTIVE.value
    is_active: bool = True
    date_intro: str | None = None
    date_published: str | None = None
    date_expired: str | None = None
    replaces: str | None = None
    replaced_by: str | None = None
    normative_refs: list[str] = field(default_factory=list)
    url: str = ""
    source: str = "meganorm"
    category: str = "ГОСТ"

    def __post_init__(self) -> None:
        if not self.clean_number:
            self.clean_number = normalize_gost_number(self.designation or self.title)
        parsed_status = DocumentStatus.from_str(self.status)
        self.is_active = (parsed_status == DocumentStatus.ACTIVE)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> NormDocument:
        clean_data = dict(data)
        return cls(
            id=str(clean_data.get("id") or ""),
            designation=str(clean_data.get("designation") or ""),
            clean_number=str(clean_data.get("clean_number") or ""),
            title=str(clean_data.get("title") or ""),
            status=str(clean_data.get("status") or DocumentStatus.ACTIVE.value),
            is_active=bool(clean_data.get("is_active", True)),
            date_intro=clean_data.get("date_intro"),
            date_published=clean_data.get("date_published"),
            date_expired=clean_data.get("date_expired"),
            replaces=clean_data.get("replaces"),
            replaced_by=clean_data.get("replaced_by"),
            normative_refs=list(clean_data.get("normative_refs") or []),
            url=str(clean_data.get("url") or clean_data.get("href") or ""),
            source=str(clean_data.get("source") or "meganorm"),
            category=str(clean_data.get("category") or "ГОСТ"),
        )
