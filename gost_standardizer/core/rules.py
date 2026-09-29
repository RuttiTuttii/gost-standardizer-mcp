from __future__ import annotations

import re
from typing import Any

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Mm, Pt

from gost_standardizer.core.presets import Preset

DEFAULT_PAGE_TOLERANCE_MM = 1.0
DEFAULT_FONT_TOLERANCE_PT = 0.5

RULE_LIBRARY: dict[str, dict[str, Any]] = {
    "document.macro_enabled": {
        "severity": "info",
        "confidence": 0.95,
        "auto_fixable": False,
        "recommendation": "Macros are preserved, but review macro-enabled documents separately.",
    },
    "document.empty_body": {
        "severity": "warning",
        "confidence": 0.98,
        "auto_fixable": False,
        "recommendation": "Add visible body text before standardizing.",
    },
    "document.title_suspect": {
        "severity": "warning",
        "confidence": 0.76,
        "auto_fixable": False,
        "recommendation": "Check whether the opening lines are a title block or just body text.",
    },
    "page.size.mismatch": {
        "severity": "error",
        "confidence": 0.99,
        "auto_fixable": True,
        "recommendation": "Set the page size to the selected profile.",
    },
    "page.margin.left": {
        "severity": "error",
        "confidence": 0.99,
        "auto_fixable": True,
        "recommendation": "Set the left margin to the selected profile.",
    },
    "page.margin.right": {
        "severity": "error",
        "confidence": 0.99,
        "auto_fixable": True,
        "recommendation": "Set the right margin to the selected profile.",
    },
    "page.margin.top": {
        "severity": "error",
        "confidence": 0.99,
        "auto_fixable": True,
        "recommendation": "Set the top margin to the selected profile.",
    },
    "page.margin.bottom": {
        "severity": "error",
        "confidence": 0.99,
        "auto_fixable": True,
        "recommendation": "Set the bottom margin to the selected profile.",
    },
    "paragraph.kind.mismatch": {
        "severity": "warning",
        "confidence": 0.8,
        "auto_fixable": True,
        "recommendation": "Apply the profile paragraph formatting to the detected structure.",
    },
    "paragraph.alignment.mismatch": {
        "severity": "warning",
        "confidence": 0.9,
        "auto_fixable": True,
        "recommendation": "Align the paragraph to the profile expectation.",
    },
    "paragraph.indent.mismatch": {
        "severity": "warning",
        "confidence": 0.9,
        "auto_fixable": True,
        "recommendation": "Adjust the paragraph indentation to the selected profile.",
    },
    "paragraph.spacing.mismatch": {
        "severity": "warning",
        "confidence": 0.88,
        "auto_fixable": True,
        "recommendation": "Adjust the paragraph spacing and line spacing to the selected profile.",
    },
    "run.font.family.mismatch": {
        "severity": "warning",
        "confidence": 0.9,
        "auto_fixable": True,
        "recommendation": "Apply the profile font family to the text runs.",
    },
    "run.font.size.mismatch": {
        "severity": "warning",
        "confidence": 0.9,
        "auto_fixable": True,
        "recommendation": "Apply the profile font size to the text runs.",
    },
    "run.font.style.mismatch": {
        "severity": "warning",
        "confidence": 0.85,
        "auto_fixable": True,
        "recommendation": "Normalize bold and italic styling to the profile expectation.",
    },
    "table.paragraph.format": {
        "severity": "warning",
        "confidence": 0.9,
        "auto_fixable": True,
        "recommendation": "Normalize table paragraph formatting to the profile expectation.",
    },
}


def clear_theme_fonts(r_fonts: Any) -> None:
    for key in list(r_fonts.attrib.keys()):
        if "theme" in key.lower():
            del r_fonts.attrib[key]


def set_r_fonts(run: Any, font_name: str) -> None:
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.get_or_add_rFonts()
    r_fonts.set(qn("w:ascii"), font_name)
    r_fonts.set(qn("w:hAnsi"), font_name)
    r_fonts.set(qn("w:cs"), font_name)
    r_fonts.set(qn("w:eastAsia"), font_name)
    clear_theme_fonts(r_fonts)


def set_style_font(
    style: Any,
    font_name: str,
    font_size_pt: int | None = None,
    bold: bool | None = None,
) -> None:
    font = style.font
    font.name = font_name
    if font_size_pt is not None:
        font.size = Pt(font_size_pt)
    if bold is not None:
        font.bold = bold
    if style._element.rPr is not None:
        r_fonts = style._element.rPr.rFonts
        if r_fonts is None:
            r_fonts = style._element.rPr._add_rFonts()
        r_fonts.set(qn("w:ascii"), font_name)
        r_fonts.set(qn("w:hAnsi"), font_name)
        r_fonts.set(qn("w:cs"), font_name)
        r_fonts.set(qn("w:eastAsia"), font_name)
        clear_theme_fonts(r_fonts)


def is_in_table(paragraph: Any) -> bool:
    element = paragraph._element
    parent = element.getparent()
    while parent is not None:
        if parent.tag.endswith("}tc"):
            return True
        parent = parent.getparent()
    return False


def length_mm(value: Any) -> float | None:
    if value is None:
        return None
    return float(value.mm)


def length_pt(value: Any) -> float | None:
    if value is None:
        return None
    return float(value.pt)


def effective_alignment(paragraph: Any) -> Any:
    if paragraph.alignment is not None:
        return paragraph.alignment
    if paragraph.style and paragraph.style.paragraph_format.alignment is not None:
        return paragraph.style.paragraph_format.alignment
    return None


def effective_paragraph_metrics(paragraph: Any) -> dict[str, Any]:
    fmt = paragraph.paragraph_format
    style_fmt = paragraph.style.paragraph_format if paragraph.style else None

    def _choose(attr: str) -> Any:
        value = getattr(fmt, attr)
        if value is not None:
            return value
        if style_fmt is not None:
            return getattr(style_fmt, attr)
        return None

    return {
        "alignment": effective_alignment(paragraph),
        "first_line_indent_mm": length_mm(_choose("first_line_indent")),
        "left_indent_mm": length_mm(_choose("left_indent")),
        "right_indent_mm": length_mm(_choose("right_indent")),
        "space_before_pt": length_pt(_choose("space_before")),
        "space_after_pt": length_pt(_choose("space_after")),
        "line_spacing": _choose("line_spacing"),
    }


def effective_run_metrics(run: Any, paragraph: Any) -> dict[str, Any]:
    font = run.font
    style_font = paragraph.style.font if paragraph.style else None
    return {
        "name": font.name or (style_font.name if style_font else None),
        "size_pt": length_pt(font.size or (style_font.size if style_font else None)),
        "bold": font.bold if font.bold is not None else (style_font.bold if style_font else None),
        "italic": font.italic if font.italic is not None else (style_font.italic if style_font else None),
    }


def paragraph_expected_metrics(kind: str, preset: Preset) -> dict[str, Any]:
    if kind == "title":
        return {
            "alignment": WD_ALIGN_PARAGRAPH.CENTER,
            "first_line_indent_mm": 0.0,
            "left_indent_mm": 0.0,
            "right_indent_mm": 0.0,
            "space_before_pt": preset.title_space_before_pt,
            "space_after_pt": preset.title_space_after_pt,
            "line_spacing": 1.0,
            "font_size_pt": preset.title_font_size_pt,
            "bold": True,
            "italic": False,
        }
    if kind == "heading":
        return {
            "alignment": WD_ALIGN_PARAGRAPH.LEFT,
            "first_line_indent_mm": 0.0,
            "left_indent_mm": 0.0,
            "right_indent_mm": 0.0,
            "space_before_pt": preset.heading_space_before_pt,
            "space_after_pt": preset.heading_space_after_pt,
            "line_spacing": 1.0,
            "font_size_pt": preset.heading_font_size_pt,
            "bold": True,
            "italic": False,
        }
    if kind == "caption":
        return {
            "alignment": WD_ALIGN_PARAGRAPH.CENTER,
            "first_line_indent_mm": 0.0,
            "left_indent_mm": 0.0,
            "right_indent_mm": 0.0,
            "space_before_pt": preset.caption_space_before_pt,
            "space_after_pt": preset.caption_space_after_pt,
            "line_spacing": 1.0,
            "font_size_pt": preset.table_font_size_pt,
            "bold": False,
            "italic": True,
        }
    if kind == "list":
        return {
            "alignment": WD_ALIGN_PARAGRAPH.LEFT,
            "first_line_indent_mm": 0.0,
            "left_indent_mm": 8.0,
            "right_indent_mm": 0.0,
            "space_before_pt": 0.0,
            "space_after_pt": 0.0,
            "line_spacing": preset.body_line_spacing,
            "font_size_pt": preset.table_font_size_pt,
            "bold": False,
            "italic": False,
        }
    return {
        "alignment": WD_ALIGN_PARAGRAPH.JUSTIFY,
        "first_line_indent_mm": preset.body_first_line_indent_mm,
        "left_indent_mm": 0.0,
        "right_indent_mm": 0.0,
        "space_before_pt": preset.body_space_before_pt,
        "space_after_pt": preset.body_space_after_pt,
        "line_spacing": preset.body_line_spacing,
        "font_size_pt": preset.body_font_size_pt,
        "bold": False,
        "italic": False,
    }


def make_issue(rule_id: str, message: str, **overrides: Any) -> dict[str, Any]:
    rule = dict(RULE_LIBRARY.get(rule_id, {}))
    payload = {
        "severity": overrides.pop("severity", rule.get("severity", "warning")),
        "rule_id": rule_id,
        "message": message,
        "confidence": float(overrides.pop("confidence", rule.get("confidence", 0.8))),
        "auto_fixable": bool(overrides.pop("auto_fixable", rule.get("auto_fixable", False))),
        "evidence": overrides.pop("evidence", {}),
        "recommendation": overrides.pop("recommendation", rule.get("recommendation", "")),
    }
    payload.update(overrides)
    return payload


def approx_equal(actual: float | None, expected: float | None, tolerance: float) -> bool:
    if actual is None or expected is None:
        return True
    return abs(actual - expected) <= tolerance


def set_page_setup(document: Any, preset: Preset) -> None:
    for section in document.sections:
        section.page_width = Mm(preset.page_width_mm)
        section.page_height = Mm(preset.page_height_mm)
        section.left_margin = Mm(preset.margin_left_mm)
        section.right_margin = Mm(preset.margin_right_mm)
        section.top_margin = Mm(preset.margin_top_mm)
        section.bottom_margin = Mm(preset.margin_bottom_mm)
        section.header_distance = Mm(preset.header_mm)
        section.footer_distance = Mm(preset.footer_mm)
        section.different_first_page_header_footer = False

