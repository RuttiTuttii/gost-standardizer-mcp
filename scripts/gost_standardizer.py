from __future__ import annotations

import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

import gost_standardizer as _pkg
from gost_standardizer.cli.main import main
from gost_standardizer.core.classifier import (
    HEADING_PATTERNS,
    LIST_PATTERNS,
    SECTION_KEYWORDS,
    classify_paragraph,
    classify_paragraph as _classify_paragraph,
    has_numbering as _has_numbering,
    normalize_text as _normalize_text,
)
from gost_standardizer.core.engine import (
    analyze_document,
    collect_text as _collect_text,
    compare_to_preset,
    convert_legacy_doc as _convert_legacy_doc,
    detect_preset,
    document_statistics as _document_statistics,
    inspect_document,
    make_output_path,
    open_document_source,
    preset_scores as _preset_scores,
    resolve_input_path,
    sample_paragraphs as _sample_paragraphs,
    standardize_document,
    validate_document,
    validate_page_setup as _validate_page_setup,
    validate_paragraphs as _validate_paragraphs,
)
from gost_standardizer.core.presets import (
    BUILTIN_PROFILES,
    PRESETS,
    PROFILES_DIR,
    Preset,
    coerce_preset as _coerce_preset,
    explain_preset,
    list_presets,
    list_profiles,
    load_profile,
    profile_path as _profile_path,
    profile_payload as _profile_payload,
    resolve_preset,
    resolve_profile,
    save_profile,
)
from gost_standardizer.core.rules import (
    RULE_LIBRARY,
    approx_equal as _approx_equal,
    clear_theme_fonts as _clear_theme_fonts,
    effective_alignment as _effective_alignment,
    effective_paragraph_metrics as _effective_paragraph_metrics,
    effective_run_metrics as _effective_run_metrics,
    is_in_table as _is_in_table,
    make_issue as _issue,
    paragraph_expected_metrics as _paragraph_expected_metrics,
    set_page_setup as _set_page_setup,
    set_r_fonts as _set_r_fonts,
    set_style_font as _set_style_font,
)

if __name__ == "__main__":
    raise SystemExit(main())
