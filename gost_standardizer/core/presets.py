from __future__ import annotations

from dataclasses import asdict, dataclass
import json
from pathlib import Path
import re
from typing import Any

from gost_standardizer.models.profiles import DocumentProfile, FontProfile, MarginsProfile, SpacingProfile

ROOT_DIR = Path(__file__).resolve().parent.parent.parent
PROFILES_DIR = ROOT_DIR / "profiles"
DEFAULT_PROFILE_KIND = "gost"
DEFAULT_PROFILE_SOURCE = "builtin"


@dataclass(frozen=True)
class Preset:
    key: str
    title: str
    description: str
    page_width_mm: float
    page_height_mm: float
    margin_left_mm: float
    margin_right_mm: float
    margin_top_mm: float
    margin_bottom_mm: float
    header_mm: float
    footer_mm: float
    body_font_name: str = "Times New Roman"
    body_font_size_pt: int = 14
    heading_font_size_pt: int = 14
    title_font_size_pt: int = 16
    table_font_size_pt: int = 12
    body_line_spacing: float = 1.5
    body_first_line_indent_mm: float = 12.5
    body_space_before_pt: float = 0.0
    body_space_after_pt: float = 0.0
    heading_space_before_pt: float = 12.0
    heading_space_after_pt: float = 6.0
    title_space_before_pt: float = 0.0
    title_space_after_pt: float = 12.0
    caption_space_before_pt: float = 6.0
    caption_space_after_pt: float = 6.0


PRESETS: dict[str, Preset] = {
    "report": Preset(
        key="report",
        title="GOST report (ГОСТ 7.32-2017)",
        description="Balanced preset for reports, coursework, explanatory notes, and formal documents.",
        page_width_mm=210,
        page_height_mm=297,
        margin_left_mm=30,
        margin_right_mm=10,
        margin_top_mm=20,
        margin_bottom_mm=20,
        header_mm=10,
        footer_mm=10,
    ),
    "office": Preset(
        key="office",
        title="GOST office (ГОСТ Р 7.0.97-2025)",
        description="Preset for office / administrative documents with tighter line spacing.",
        page_width_mm=210,
        page_height_mm=297,
        margin_left_mm=20,
        margin_right_mm=10,
        margin_top_mm=20,
        margin_bottom_mm=20,
        header_mm=10,
        footer_mm=10,
        body_line_spacing=1.0,
        body_first_line_indent_mm=0.0,
    ),
    "technical": Preset(
        key="technical",
        title="GOST technical (ГОСТ 2.105-2019)",
        description="Preset for technical docs, specs, and engineering notes.",
        page_width_mm=210,
        page_height_mm=297,
        margin_left_mm=30,
        margin_right_mm=10,
        margin_top_mm=20,
        margin_bottom_mm=20,
        header_mm=10,
        footer_mm=10,
    ),
    "legacy-college": Preset(
        key="legacy-college",
        title="Legacy college sample",
        description="A looser baseline inspired by archive college sample documents.",
        page_width_mm=210,
        page_height_mm=297,
        margin_left_mm=28,
        margin_right_mm=6,
        margin_top_mm=18,
        margin_bottom_mm=19,
        header_mm=10,
        footer_mm=10,
    ),
}


def preset_to_dict(preset: Preset) -> dict[str, Any]:
    return asdict(preset)


def profile_payload(
    *,
    key: str,
    preset: Preset,
    title: str | None = None,
    description: str | None = None,
    kind: str = DEFAULT_PROFILE_KIND,
    source: str = DEFAULT_PROFILE_SOURCE,
    notes: list[str] | None = None,
) -> dict[str, Any]:
    return {
        "key": key,
        "title": title or preset.title,
        "description": description or preset.description,
        "kind": kind,
        "source": source,
        "preset": preset_to_dict(preset),
        "notes": notes or [],
    }


BUILTIN_PROFILES: dict[str, dict[str, Any]] = {
    key: profile_payload(key=key, preset=preset) for key, preset in PRESETS.items()
}


def coerce_preset(data: dict[str, Any]) -> Preset:
    allowed = {field.name for field in Preset.__dataclass_fields__.values()}
    values = {key: value for key, value in data.items() if key in allowed}
    missing = allowed - values.keys()
    if missing:
        raise ValueError(f"Profile preset is missing fields: {', '.join(sorted(missing))}")
    return Preset(**values)


def profile_from_payload(payload: dict[str, Any]) -> dict[str, Any]:
    preset_value = payload.get("preset")
    if isinstance(preset_value, Preset):
        preset = preset_value
    elif isinstance(preset_value, dict):
        preset = coerce_preset(preset_value)
    else:
        raise ValueError("Profile payload must include a preset mapping")

    return {
        "key": payload.get("key") or preset.key,
        "title": payload.get("title") or preset.title,
        "description": payload.get("description") or preset.description,
        "kind": payload.get("kind") or DEFAULT_PROFILE_KIND,
        "source": payload.get("source") or DEFAULT_PROFILE_SOURCE,
        "preset": preset_to_dict(preset),
        "notes": list(payload.get("notes") or []),
    }


def load_profile_file(path: Path) -> dict[str, Any]:
    data = json.loads(path.read_text(encoding="utf-8"))
    payload = profile_from_payload(data)
    payload["source"] = "file"
    payload["path"] = str(path)
    return payload


def get_profiles_dir() -> Path:
    try:
        import gost_standardizer
        return getattr(gost_standardizer, "PROFILES_DIR", PROFILES_DIR)
    except ImportError:
        return PROFILES_DIR



def profile_path(name: str) -> Path:
    safe = re.sub(r"[^A-Za-z0-9._-]+", "_", name.strip()) or "profile"
    return get_profiles_dir() / f"{safe}.json"


def save_profile_file(name: str, payload: dict[str, Any]) -> Path:
    p_dir = get_profiles_dir()
    p_dir.mkdir(parents=True, exist_ok=True)
    path = profile_path(name)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path



def resolve_preset(name: str | None) -> Preset:
    key = (name or "report").strip().lower()
    if key not in PRESETS:
        available = ", ".join(sorted(PRESETS))
        raise ValueError(f"Unknown preset '{name}'. Available presets: {available}")
    return PRESETS[key]


def load_profile(name: str) -> dict[str, Any]:
    candidate = Path(name).expanduser()
    if candidate.exists() and candidate.is_file():
        return load_profile_file(candidate.resolve())

    key = name.strip().lower()
    if key in BUILTIN_PROFILES:
        payload = dict(BUILTIN_PROFILES[key])
        payload["source"] = "builtin"
        return payload

    profile_f = profile_path(name)
    if profile_f.exists():
        return load_profile_file(profile_f)

    available = ", ".join(sorted(BUILTIN_PROFILES))
    raise ValueError(f"Unknown profile '{name}'. Available built-ins: {available}")


def resolve_profile(name: str | None) -> dict[str, Any]:
    if not name:
        return dict(BUILTIN_PROFILES["report"])
    return load_profile(name)


def save_profile(
    name: str,
    preset_name: str | None = None,
    title: str | None = None,
    description: str | None = None,
    kind: str = "organization",
    notes: list[str] | None = None,
) -> dict[str, Any]:
    preset = resolve_preset(preset_name)
    payload = profile_payload(
        key=name,
        preset=preset,
        title=title or name,
        description=description,
        kind=kind,
        source="file",
        notes=notes,
    )
    path = save_profile_file(name, payload)
    payload["path"] = str(path)
    return payload


def list_presets() -> list[dict[str, Any]]:
    return [dict(profile_payload(key=key, preset=preset)) for key, preset in PRESETS.items()]


def list_profiles() -> list[dict[str, Any]]:
    profiles = [dict(profile) for profile in BUILTIN_PROFILES.values()]
    if PROFILES_DIR.exists():
        for file in sorted(PROFILES_DIR.glob("*.json")):
            try:
                profiles.append(load_profile_file(file))
            except Exception:
                continue
    return profiles


def explain_preset(name: str | None = None) -> dict[str, Any]:
    preset = resolve_preset(name)
    return {
        "key": preset.key,
        "title": preset.title,
        "description": preset.description,
        "preset": preset_to_dict(preset),
    }
