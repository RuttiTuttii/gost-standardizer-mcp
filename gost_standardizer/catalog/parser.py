from __future__ import annotations

import re
from urllib.parse import urljoin

from gost_standardizer.models.document import DocumentStatus, NormDocument, normalize_gost_number


def parse_index_card(html: str, url: str) -> NormDocument | None:
    """Parse a Meganorm document card (e.g.

    Index/85/85280.htm).
    """
    if not html:
        return None

    # Title
    title = ""
    m_title = re.search(r"<title>(.*?)</title>", html, re.I | re.S)
    if m_title:
        title = " ".join(m_title.group(1).split())
        title = re.sub(r"^Скачать\s+", "", title, flags=re.I).strip()

    # Extract ID from URL
    m_id = re.search(r"/(\d+)\.htm", url)
    doc_id = m_id.group(1) if m_id else ""

    designation = ""
    status = DocumentStatus.ACTIVE.value
    date_intro = None
    date_published = None
    date_expired = None
    replaces = None
    replaced_by = None
    normative_refs: list[str] = []

    # Parse rows
    rows = re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.I | re.S)
    for row in rows:
        clean = re.sub(r"<[^>]+>", " ", row)
        clean = " ".join(clean.split())

        if "Обозначение:" in clean:
            m = re.search(r"Обозначение:\s*([^;,\n]+)", clean)
            if m and not designation:
                designation = m.group(1).strip()
        if "Статус:" in clean:
            m = re.search(r"Статус:\s*([^\s;,\n]+)", clean)
            if m:
                status = m.group(1).strip().lower()
        if "Дата введения:" in clean:
            m = re.search(r"Дата введения:\s*([0-9]{2}\.[0-9]{2}\.[0-9]{4})", clean)
            if m:
                date_intro = m.group(1).strip()
        if "Дата издания:" in clean:
            m = re.search(r"Дата издания:\s*([0-9]{2}\.[0-9]{2}\.[0-9]{4})", clean)
            if m:
                date_published = m.group(1).strip()
        if "Дата окончания срока действия:" in clean:
            m = re.search(r"Дата окончания срока действия:\s*([0-9]{2}\.[0-9]{2}\.[0-9]{4})", clean)
            if m:
                date_expired = m.group(1).strip()
        if "Взамен:" in clean:
            m = re.search(r"Взамен:\s*([^;<\n]+)", clean)
            if m:
                replaces = m.group(1).strip()
        if "Заменяющий:" in clean:
            m = re.search(r"Заменяющий:\s*([^;<\n]+)", clean)
            if m:
                replaced_by = m.group(1).strip()
        if "Нормативные ссылки:" in clean:
            m = re.search(r"Нормативные ссылки:\s*([^<\n]+)", clean)
            if m:
                refs_raw = m.group(1).strip()
                normative_refs = [r.strip() for r in re.split(r"[;,]", refs_raw) if r.strip()]

    if not designation:
        m_desig = re.search(r"(ГОСТ(?:\s+Р)?\s+[0-9]+(?:\.[0-9]+)*(?:-[0-9]+)?)", title, re.I)
        if m_desig:
            designation = m_desig.group(1).strip()
        else:
            designation = title

    clean_number = normalize_gost_number(designation)

    return NormDocument(
        id=doc_id,
        designation=designation,
        clean_number=clean_number,
        title=title,
        status=status,
        date_intro=date_intro,
        date_published=date_published,
        date_expired=date_expired,
        replaces=replaces,
        replaced_by=replaced_by,
        normative_refs=normative_refs,
        url=url,
        source="meganorm",
        category="ГОСТ Р" if "Р" in designation.upper() else "ГОСТ",
    )


def parse_list_links(html: str, base_url: str) -> list[dict[str, str]]:
    """Parse links to documents from a list page (e.g. list/8-0.htm)."""
    links: list[dict[str, str]] = []
    matches = re.findall(r"<a\s+[^>]*href=[\"\x27]([^\"\x27]+)[\"\x27][^>]*>(.*?)</a>", html, re.I | re.S)
    for href, text in matches:
        clean_text = " ".join(re.sub(r"<[^>]+>", " ", text).split())
        full_url = urljoin(base_url, href)
        if "/Index/" in full_url:
            links.append({"href": full_url, "title": clean_text})
    return links
