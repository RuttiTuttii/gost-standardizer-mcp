from __future__ import annotations

from gost_standardizer.core.classifier import classify_paragraph, normalize_text
from gost_standardizer.core.engine import (
    analyze_document,
    compare_to_preset,
    detect_preset,
    inspect_document,
    make_output_path,
    open_document_source,
    resolve_input_path,
    standardize_document,
    validate_document,
    explain_preset,
)
from gost_standardizer.core.presets import (
    BUILTIN_PROFILES,
    PRESETS,
    PROFILES_DIR,
    Preset,
    list_presets,
    list_profiles,
    load_profile,
    resolve_preset,
    resolve_profile,
    save_profile,
)



__all__ = [
    "BUILTIN_PROFILES",
    "PRESETS",
    "PROFILES_DIR",
    "Preset",

    "analyze_document",
    "classify_paragraph",
    "compare_to_preset",
    "detect_preset",
    "explain_preset",
    "inspect_document",
    "list_presets",
    "list_profiles",
    "load_profile",
    "make_output_path",
    "normalize_text",
    "open_document_source",
    "resolve_input_path",
    "resolve_preset",
    "resolve_profile",
    "save_profile",
    "standardize_document",
    "validate_document",
]
