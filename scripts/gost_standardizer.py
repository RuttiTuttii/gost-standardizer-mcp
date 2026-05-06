from __future__ import annotations

from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any
import json
import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Mm, Pt


SECTION_KEYWORDS = {
    "report": {
        "аннотация",
        "содержание",
        "введение",
        "заключение",
        "список литературы",
        "список использованных источников",
        "приложение",
    },
    "office": {
        "приказ",
        "распоряжение",
        "положение",
        "регламент",
        "инструкция",
    },
    "technical": {
        "техническое задание",
        "пояснительная записка",
        "описание",
        "требования",
        "архитектура",
    },
}


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
        title="GOST report",
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
        title="GOST office",
        description="Preset for office / administrative documents with a tighter page block.",
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
        title="GOST technical",
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
        description="A looser baseline inspired by the archive sample documents.",
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


HEADING_PATTERNS = [
    re.compile(r"^\d+(?:\.\d+)*\s+\S"),
    re.compile(r"^(аннотация|содержание|введение|заключение|список литературы|список использованных источников|приложения?)$", re.I),
    re.compile(r"^(техническое задание|пояснительная записка)$", re.I),
]

LIST_PATTERNS = [
    re.compile(r"^[•\-–—]\s+\S"),
    re.compile(r"^\(?\d+[.)]\s+\S"),
    re.compile(r"^\d+\)\s+\S"),
]


def list_presets() -> list[dict[str, Any]]:
    return [asdict(preset) for preset in PRESETS.values()]


def resolve_preset(name: str | None) -> Preset:
    key = (name or "report").strip().lower()
    if key not in PRESETS:
        available = ", ".join(sorted(PRESETS))
        raise ValueError(f"Unknown preset '{name}'. Available presets: {available}")
    return PRESETS[key]


def resolve_input_path(raw_path: str, base_dir: Path | None = None) -> Path:
    candidate = Path(raw_path).expanduser()
    if not candidate.is_absolute():
        candidate = (base_dir or Path.cwd()) / candidate
    candidate = candidate.resolve()
    if not candidate.exists():
        raise FileNotFoundError(f"Input file not found: {candidate}")
    if candidate.suffix.lower() not in {".docx", ".docm"}:
        raise ValueError("Only .docx and .docm files are supported by the current implementation")
    return candidate


def make_output_path(input_path: Path, output_path: str | None = None) -> Path:
    if output_path:
        candidate = Path(output_path).expanduser()
        if not candidate.is_absolute():
            candidate = (Path.cwd() / candidate).resolve()
        if candidate.suffix.lower() not in {".docx", ".docm"}:
            candidate = candidate.with_suffix(".docx")
        return candidate
    return input_path.with_name(f"{input_path.stem}_gost.docx")


def _set_r_fonts(run, font_name: str) -> None:
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.get_or_add_rFonts()
    r_fonts.set(qn("w:ascii"), font_name)
    r_fonts.set(qn("w:hAnsi"), font_name)
    r_fonts.set(qn("w:cs"), font_name)
    r_fonts.set(qn("w:eastAsia"), font_name)


def _set_style_font(style, font_name: str, font_size_pt: int | None = None, bold: bool | None = None) -> None:
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


def _has_numbering(paragraph) -> bool:
    p_pr = paragraph._p.pPr
    return bool(p_pr is not None and p_pr.numPr is not None)


def _normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text.replace("\xa0", " ")).strip()


def _is_in_table(paragraph) -> bool:
    element = paragraph._element
    parent = element.getparent()
    while parent is not None:
        if parent.tag.endswith("}tc"):
            return True
        parent = parent.getparent()
    return False


def _classify_paragraph(paragraph, index: int, non_empty_index: int) -> str:
    text = _normalize_text(paragraph.text)
    if not text:
        return "empty"

    style_name = (paragraph.style.name if paragraph.style else "").lower()
    lowered = text.lower()

    if "title" in style_name:
        return "title"
    if "heading" in style_name:
        return "heading"
    if "caption" in style_name:
        return "caption"
    if _has_numbering(paragraph):
        return "list"
    if any(pattern.match(text) for pattern in LIST_PATTERNS):
        return "list"
    if any(pattern.match(lowered) for pattern in HEADING_PATTERNS):
        return "heading"
    if non_empty_index == 1 and len(text) <= 180 and not text.endswith((".", "!", "?")):
        return "title"
    if paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER and len(text) <= 180:
        return "title"
    if len(text) <= 72 and lowered.isupper():
        return "heading"
    if lowered.startswith(("рисунок ", "таблица ")):
        return "caption"
    if len(text) <= 96 and text[:1].isupper() and lowered.count(" ") <= 6 and not text.endswith("."):
        return "heading"
    return "body"


def _apply_paragraph_format(paragraph, kind: str, preset: Preset, inside_table: bool = False) -> None:
    fmt = paragraph.paragraph_format
    if kind == "empty":
        fmt.space_before = Pt(0)
        fmt.space_after = Pt(0)
        return

    if inside_table:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        fmt.first_line_indent = Mm(0)
        fmt.left_indent = Mm(0)
        fmt.right_indent = Mm(0)
        fmt.space_before = Pt(0)
        fmt.space_after = Pt(0)
        fmt.line_spacing = 1.0
        return

    if kind == "title":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fmt.first_line_indent = Mm(0)
        fmt.left_indent = Mm(0)
        fmt.right_indent = Mm(0)
        fmt.space_before = Pt(preset.title_space_before_pt)
        fmt.space_after = Pt(preset.title_space_after_pt)
        fmt.line_spacing = 1.0
        return

    if kind == "heading":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        fmt.first_line_indent = Mm(0)
        fmt.left_indent = Mm(0)
        fmt.right_indent = Mm(0)
        fmt.space_before = Pt(preset.heading_space_before_pt)
        fmt.space_after = Pt(preset.heading_space_after_pt)
        fmt.line_spacing = 1.0
        return

    if kind == "caption":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fmt.first_line_indent = Mm(0)
        fmt.left_indent = Mm(0)
        fmt.right_indent = Mm(0)
        fmt.space_before = Pt(preset.caption_space_before_pt)
        fmt.space_after = Pt(preset.caption_space_after_pt)
        fmt.line_spacing = 1.0
        return

    if kind == "list":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        fmt.left_indent = Mm(8)
        fmt.first_line_indent = Mm(0)
        fmt.right_indent = Mm(0)
        fmt.space_before = Pt(0)
        fmt.space_after = Pt(0)
        fmt.line_spacing = preset.body_line_spacing
        return

    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    fmt.first_line_indent = Mm(preset.body_first_line_indent_mm)
    fmt.left_indent = Mm(0)
    fmt.right_indent = Mm(0)
    fmt.space_before = Pt(preset.body_space_before_pt)
    fmt.space_after = Pt(preset.body_space_after_pt)
    fmt.line_spacing = preset.body_line_spacing


def _apply_run_format(run, kind: str, preset: Preset) -> None:
    if not run.text:
        return

    if kind == "title":
        run.font.name = preset.body_font_name
        run.font.size = Pt(preset.title_font_size_pt)
        run.font.bold = True
        _set_r_fonts(run, preset.body_font_name)
        return

    if kind == "heading":
        run.font.name = preset.body_font_name
        run.font.size = Pt(preset.heading_font_size_pt)
        run.font.bold = True
        _set_r_fonts(run, preset.body_font_name)
        return

    if kind == "caption":
        run.font.name = preset.body_font_name
        run.font.size = Pt(preset.table_font_size_pt)
        run.font.italic = True
        _set_r_fonts(run, preset.body_font_name)
        return

    size = preset.table_font_size_pt if kind == "list" else preset.body_font_size_pt
    run.font.name = preset.body_font_name
    run.font.size = Pt(size)
    _set_r_fonts(run, preset.body_font_name)


def _apply_style_defaults(document: Document, preset: Preset) -> None:
    for style_name in [
        "Normal",
        "Body Text",
        "List Paragraph",
        "Caption",
        "Title",
        "Subtitle",
        "Heading 1",
        "Heading 2",
        "Heading 3",
    ]:
        try:
            style = document.styles[style_name]
        except KeyError:
            continue

        if style_name in {"Heading 1", "Heading 2", "Heading 3"}:
            _set_style_font(style, preset.body_font_name, preset.heading_font_size_pt, True)
            continue
        if style_name == "Title":
            _set_style_font(style, preset.body_font_name, preset.title_font_size_pt, True)
            continue
        if style_name == "Caption":
            _set_style_font(style, preset.body_font_name, preset.table_font_size_pt, False)
            continue
        _set_style_font(style, preset.body_font_name, preset.body_font_size_pt, False)

        if style.paragraph_format is not None:
            fmt = style.paragraph_format
            fmt.space_before = Pt(preset.body_space_before_pt)
            fmt.space_after = Pt(preset.body_space_after_pt)
            fmt.line_spacing = preset.body_line_spacing
            if style_name == "List Paragraph":
                fmt.first_line_indent = Mm(0)
                fmt.left_indent = Mm(8)
            else:
                fmt.first_line_indent = Mm(preset.body_first_line_indent_mm)
                fmt.left_indent = Mm(0)


def _set_page_setup(document: Document, preset: Preset) -> None:
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


def _collect_text(document: Document) -> str:
    chunks: list[str] = []
    for paragraph in document.paragraphs:
        text = _normalize_text(paragraph.text)
        if text:
            chunks.append(text)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    text = _normalize_text(paragraph.text)
                    if text:
                        chunks.append(text)
    return "\n".join(chunks)


def guess_preset(document: Document, source_path: Path | None = None) -> str:
    text = _collect_text(document).lower()
    scores = {key: 0 for key in PRESETS}

    if source_path is not None:
        stem = source_path.stem.lower()
        if any(token in stem for token in ("tz", "тз", "technical", "spec")):
            scores["technical"] += 2
        if any(token in stem for token in ("report", "отчет", "otchet", "пз", "poyasnit")):
            scores["report"] += 2
        if any(token in stem for token in ("order", "приказ", "reglament", "instruction")):
            scores["office"] += 2

    for key, keywords in SECTION_KEYWORDS.items():
        for keyword in keywords:
            if keyword in text:
                scores[key] += 1

    return max(scores.items(), key=lambda item: (item[1], item[0] == "report"))[0]


def inspect_document(path: str, sample_size: int = 8) -> dict[str, Any]:
    source = resolve_input_path(path)
    document = Document(str(source))
    preset_key = guess_preset(document, source)
    preset = resolve_preset(preset_key)

    non_empty = [
        {
            "index": index,
            "text": _normalize_text(paragraph.text),
            "style": paragraph.style.name if paragraph.style else None,
            "alignment": str(paragraph.alignment) if paragraph.alignment is not None else None,
        }
        for index, paragraph in enumerate(document.paragraphs)
        if _normalize_text(paragraph.text)
    ]

    issues: list[str] = []
    if source.suffix.lower() == ".docm":
        issues.append("Document is macro-enabled; macros are preserved but not interpreted.")

    if not any(_normalize_text(p.text) for p in document.paragraphs):
        issues.append("No body text detected in the main document paragraphs.")

    if len(document.paragraphs) > 0:
        first_text = _normalize_text(document.paragraphs[0].text)
        if first_text and not any(pattern.match(first_text.lower()) for pattern in HEADING_PATTERNS):
            if not first_text.endswith((".", "!", "?")) and len(first_text) > 120:
                issues.append("The opening text looks like raw body text rather than a clear title block.")

    deviations: list[str] = []
    if document.sections:
        section = document.sections[0]
        current = {
            "left_mm": round(section.left_margin.mm, 2),
            "right_mm": round(section.right_margin.mm, 2),
            "top_mm": round(section.top_margin.mm, 2),
            "bottom_mm": round(section.bottom_margin.mm, 2),
        }
        target = {
            "left_mm": preset.margin_left_mm,
            "right_mm": preset.margin_right_mm,
            "top_mm": preset.margin_top_mm,
            "bottom_mm": preset.margin_bottom_mm,
        }
        for key, value in current.items():
            if abs(value - target[key]) > 1.0:
                deviations.append(f"Section margin {key} is {value} mm, target is {target[key]} mm")

    samples = non_empty[:sample_size]
    return {
        "path": str(source),
        "preset_guess": preset_key,
        "preset": asdict(preset),
        "statistics": {
            "paragraphs": len(document.paragraphs),
            "tables": len(document.tables),
            "inline_shapes": len(document.inline_shapes),
            "sections": len(document.sections),
            "non_empty_paragraphs": len(non_empty),
        },
        "deviations": deviations,
        "issues": issues,
        "sample_paragraphs": samples,
    }


def standardize_document(
    path: str,
    output_path: str | None = None,
    preset_name: str | None = None,
    overwrite: bool = False,
    aggressive: bool = False,
) -> dict[str, Any]:
    source = resolve_input_path(path)
    preset = resolve_preset(preset_name)
    document = Document(str(source))

    _set_page_setup(document, preset)
    _apply_style_defaults(document, preset)

    non_empty_seen = 0
    paragraph_actions: dict[str, int] = {
        "title": 0,
        "heading": 0,
        "caption": 0,
        "list": 0,
        "body": 0,
        "empty": 0,
    }

    for index, paragraph in enumerate(document.paragraphs):
        text = _normalize_text(paragraph.text)
        if text:
            non_empty_seen += 1
        kind = _classify_paragraph(paragraph, index, non_empty_seen)
        if aggressive and kind == "body" and len(text) <= 96 and not text.endswith("."):
            kind = "heading"
        paragraph_actions[kind] = paragraph_actions.get(kind, 0) + 1
        _apply_paragraph_format(paragraph, kind, preset, inside_table=False)
        for run in paragraph.runs:
            _apply_run_format(run, kind, preset)

    table_paragraphs = 0
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    table_paragraphs += 1
                    kind = "table"
                    _apply_paragraph_format(paragraph, kind, preset, inside_table=True)
                    for run in paragraph.runs:
                        run.font.name = preset.body_font_name
                        run.font.size = Pt(preset.table_font_size_pt)
                        _set_r_fonts(run, preset.body_font_name)

    if not source.suffix.lower() == ".docm":
        for section in document.sections:
            section.different_first_page_header_footer = False

    output = make_output_path(source, output_path)
    if output.exists() and not overwrite:
        output = output.with_name(f"{output.stem}_v2{output.suffix}")
    output.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(output))

    return {
        "source_path": str(source),
        "output_path": str(output),
        "preset": asdict(preset),
        "paragraph_actions": paragraph_actions,
        "table_paragraphs_touched": table_paragraphs,
        "statistics": {
            "paragraphs": len(document.paragraphs),
            "tables": len(document.tables),
            "inline_shapes": len(document.inline_shapes),
            "sections": len(document.sections),
        },
        "notes": [
            "The document was reformatted in-place in a copy, preserving the original source file.",
            "If the source had custom table layouts or complex section breaks, review the result visually.",
        ],
    }


def render_report(result: dict[str, Any]) -> str:
    return json.dumps(result, ensure_ascii=False, indent=2)
