from __future__ import annotations

import re
from typing import Any

from docx.enum.text import WD_ALIGN_PARAGRAPH

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

HEADING_PATTERNS = [
    re.compile(r"^(?:раздел|глава|часть|приложение)\s+[0-9a-zа-яёivx]+(?:\.[0-9a-zа-яё]+)*\.?\s*", re.I),
    re.compile(r"^\d+\.\d+(?:\.\d+)*\.?\s+\S"),
    re.compile(r"^\d+\.?\s+[А-ЯA-ZЁ]"),
    re.compile(r"^\d+(?:\.\d+)*\s+\S"),
    re.compile(
        r"^(?:аннотация|содержание|введение|заключение|список литературы|список использованных источников|приложения?)$",
        re.I,
    ),
    re.compile(r"^(?:техническое задание|пояснительная записка)$", re.I),
]

LIST_PATTERNS = [
    re.compile(r"^[•\-–—]\s+\S"),
    re.compile(r"^\(?\d+\)\s+\S"),
    re.compile(r"^\d+\.\s+[а-яa-zё]"),
    re.compile(r"^[а-яa-zё]\)\s+\S"),
    re.compile(r"^\(?\d+[.)]\s+\S"),
    re.compile(r"^\d+\)\s+\S"),
]


def normalize_text(text: str) -> str:
    return " ".join(text.replace("\xa0", " ").split()).strip()


def has_numbering(paragraph: Any) -> bool:
    p_pr = paragraph._p.pPr
    return bool(p_pr is not None and p_pr.numPr is not None)


def classify_paragraph(paragraph: Any, index: int, non_empty_index: int) -> str:
    text = normalize_text(paragraph.text)
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

    # Prioritize heading patterns over list patterns for dotted section numbers (e.g. 1. Введение)
    if any(pattern.match(text) or pattern.match(lowered) for pattern in HEADING_PATTERNS):
        if len(text) <= 160 and not text.rstrip().endswith((";", ",")):
            return "heading"

    if has_numbering(paragraph):
        return "list"
    if any(pattern.match(text) for pattern in LIST_PATTERNS):
        return "list"

    if non_empty_index == 1 and len(text) <= 180 and not text.endswith((".", "!", "?")):
        return "title"
    if (
        non_empty_index <= 3
        and paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER
        and len(text) <= 180
        and not text.endswith((".", "!", "?"))
    ):
        return "title"
    if len(text) <= 72 and lowered.isupper() and not text.endswith((".", ";", ",")):
        return "heading"
    if lowered.startswith(("рисунок ", "таблица ")):
        return "caption"
    if len(text) <= 96 and text[:1].isupper() and lowered.count(" ") <= 6 and not text.endswith("."):
        return "heading"
    return "body"
