from __future__ import annotations

import shutil

from gost_standardizer.catalog import (
    MeganormClient,
    cache_manager,
    fetch_norm_markdown,
    find_current_gost,
    get_current_topics,
    search_catalog,
)
from gost_standardizer.converter import (
    compile_typst,
    convert_html_to_markdown,
    find_typst_binary,
    generate_gost_typst,
    markdown_to_gost_typst,
)
from gost_standardizer.core import (
    BUILTIN_PROFILES,
    PRESETS,
    PROFILES_DIR,
    Preset,
    analyze_document,
    classify_paragraph,
    compare_to_preset,
    detect_preset,
    explain_preset,
    inspect_document,
    list_presets,
    list_profiles,
    load_profile,
    make_output_path,
    open_document_source,
    resolve_input_path,
    resolve_preset,
    resolve_profile,
    save_profile,
    standardize_document,
    validate_document,
)
from gost_standardizer.cli.main import main
from gost_standardizer.core.engine import convert_legacy_doc
from gost_standardizer.core.rules import set_r_fonts

from gost_standardizer.models import (
    DocumentProfile,
    DocumentStatus,
    FontProfile,
    InspectionReport,
    MarginsProfile,
    NormDocument,
    SpacingProfile,
    StandardizationResult,
    ValidationIssue,
    normalize_gost_number,
)

# Compatibility aliases
_classify_paragraph = classify_paragraph
_convert_legacy_doc = convert_legacy_doc
_set_r_fonts = set_r_fonts

__version__ = "0.3.0"

__all__ = [
    "BUILTIN_PROFILES",
    "PRESETS",
    "PROFILES_DIR",
    "DocumentProfile",
    "DocumentStatus",
    "FontProfile",
    "InspectionReport",
    "MarginsProfile",
    "MeganormClient",
    "NormDocument",
    "Preset",
    "SpacingProfile",
    "StandardizationResult",
    "ValidationIssue",
    "__version__",
    "_classify_paragraph",
    "_convert_legacy_doc",
    "_set_r_fonts",
    "analyze_document",
    "cache_manager",
    "classify_paragraph",
    "compare_to_preset",
    "compile_typst",
    "convert_html_to_markdown",
    "convert_legacy_doc",
    "detect_preset",
    "explain_preset",
    "fetch_norm_markdown",
    "find_current_gost",
    "find_typst_binary",
    "generate_gost_typst",
    "get_current_topics",
    "inspect_document",
    "list_presets",
    "list_profiles",
    "load_profile",
    "main",
    "make_output_path",
    "markdown_to_gost_typst",

    "normalize_gost_number",
    "open_document_source",
    "resolve_input_path",
    "resolve_preset",
    "resolve_profile",
    "save_profile",
    "search_catalog",
    "set_r_fonts",
    "shutil",
    "standardize_document",
    "validate_document",
]
