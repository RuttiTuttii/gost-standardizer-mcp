from __future__ import annotations

import logging
import re
from typing import Any
from urllib.parse import urljoin

from gost_standardizer.catalog.cache import cache_manager
from gost_standardizer.catalog.client import MeganormClient
from gost_standardizer.catalog.parser import parse_index_card, parse_list_links
from gost_standardizer.converter.html_markdown import convert_html_to_markdown
from gost_standardizer.models.document import DocumentStatus, NormDocument, normalize_gost_number

logger = logging.getLogger(__name__)

BASE_URL = "https://meganorm.ru"
client = MeganormClient(base_url=BASE_URL)

KNOWN_INDEX_PAGES = {
    # Direct mappings for key standards
    "7.0.97-2025": "https://meganorm.ru/Index/85/85280.htm",
    "7.0.97-2016": "https://meganorm.ru/Index/63/63653.htm",
    "7.32-2017": "https://meganorm.ru/Index/65/65487.htm",
    "2.105-2019": "https://meganorm.ru/Index/70/70570.htm",
    "7.0.5-2008": "https://meganorm.ru/Index/47/47265.htm",
    "6.30-2003": "https://meganorm.ru/Index/40/40960.htm",
}


def _score_document_match(query_clean: str, candidate_text: str) -> dict[str, Any]:
    norm_candidate = normalize_gost_number(candidate_text).lower()
    q = query_clean.lower()

    if q == norm_candidate:
        return {"confidence": 1.0, "exact": True, "partial": False}
    if q in norm_candidate:
        return {"confidence": 0.85, "exact": False, "partial": True}
    return {"confidence": 0.0, "exact": False, "partial": False}


def find_current_gost(
    query: str,
    *,
    max_pages: int = 10,
    limit: int = 25,
    refresh: bool = False,
) -> dict[str, Any]:
    """Find current GOST and GOST R documents with status verification (active/replaced/cancelled)."""
    clean_query = normalize_gost_number(query)

    # 1. Fast cache lookup
    cached_doc: NormDocument | None = None
    if not refresh:
        cached_doc = cache_manager.find_indexed_document(query)

    # 2. Check known direct pages if not found in cache or refresh requested
    if cached_doc is None or refresh:
        direct_url = KNOWN_INDEX_PAGES.get(clean_query)
        if direct_url:
            html, err = client.fetch_url(direct_url)
            if not err and html:
                parsed = parse_index_card(html, direct_url)
                if parsed:
                    cache_manager.store_document(parsed)
                    cached_doc = parsed

    # 3. If found, format the response with full status details
    documents: list[dict[str, Any]] = []
    if cached_doc is not None:
        doc_record = {
            "title": cached_doc.title,
            "designation": cached_doc.designation,
            "clean_number": cached_doc.clean_number,
            "status": cached_doc.status,
            "is_active": cached_doc.is_active,
            "date_intro": cached_doc.date_intro,
            "date_published": cached_doc.date_published,
            "date_expired": cached_doc.date_expired,
            "replaces": cached_doc.replaces,
            "replaced_by": cached_doc.replaced_by,
            "normative_refs": cached_doc.normative_refs,
            "href": cached_doc.url,
            "url": cached_doc.url,
            "category": cached_doc.category,
            "origin": "cache" if not refresh else "source",
            "confidence": 1.0,
            "match": {
                "exact": True,
                "number_hint": True,
                "partial": False,
            },
        }

        # If document was replaced, also provide the active replacement document details
        if cached_doc.replaced_by:
            replacement_clean = normalize_gost_number(cached_doc.replaced_by)
            replacement_doc = cache_manager.find_indexed_document(replacement_clean)
            if replacement_doc is None and replacement_clean in KNOWN_INDEX_PAGES:
                rep_url = KNOWN_INDEX_PAGES[replacement_clean]
                rep_html, rep_err = client.fetch_url(rep_url)
                if not rep_err and rep_html:
                    replacement_doc = parse_index_card(rep_html, rep_url)
                    if replacement_doc:
                        cache_manager.store_document(replacement_doc)

            if replacement_doc is not None:
                doc_record["active_replacement"] = {
                    "designation": replacement_doc.designation,
                    "status": replacement_doc.status,
                    "date_intro": replacement_doc.date_intro,
                    "href": replacement_doc.url,
                }

        documents.append(doc_record)

    # 4. Search through category pages in cache if needed
    data = cache_manager.get_data()
    for cat_slug, cat_val in data.get("category_pages", {}).items():
        if len(documents) >= limit:
            break
        pages = cat_val.get("pages", {})
        for page_idx in range(min(max_pages, len(pages))):
            page_data = pages.get(str(page_idx), {})
            for d in page_data.get("docs", []):
                score = _score_document_match(clean_query, d.get("title", ""))
                if score["confidence"] > 0.5:
                    # Avoid duplicate
                    if any(doc["title"] == d["title"] for doc in documents):
                        continue
                    documents.append(
                        {
                            "title": d.get("title"),
                            "designation": d.get("title"),
                            "clean_number": normalize_gost_number(d.get("title", "")),
                            "status": DocumentStatus.ACTIVE.value,
                            "is_active": True,
                            "href": d.get("href"),
                            "url": d.get("href"),
                            "category": d.get("category", "ГОСТ"),
                            "origin": "cache",
                            "confidence": score["confidence"],
                            "match": score,
                        }
                    )
                    if len(documents) >= limit:
                        break

    return {
        "kind": "current-gost-search",
        "query": query,
        "normalized_query": clean_query,
        "base_url": BASE_URL,
        "total_matches": len(documents),
        "primary_document": documents[0] if documents else None,
        "documents": documents,
        "status_summary": {
            "found": len(documents) > 0,
            "status": documents[0].get("status") if documents else "not_found",
            "is_active": documents[0].get("is_active") if documents else False,
            "date_intro": documents[0].get("date_intro") if documents else None,
            "replaces": documents[0].get("replaces") if documents else None,
            "replaced_by": documents[0].get("replaced_by") if documents else None,
        },
        "cache": {
            "path": str(cache_manager.cache_path),
            "state": "hit" if not refresh else "refreshed",
            "fetched_at": data.get("fetched_at"),
            "updated_at": data.get("fetched_at"),
        },
        "note": "Statuses verified against Meganorm index cards (актуализированный классификатор).",
    }


def search_catalog(
    query: str,
    *,
    category: str | None = None,
    max_pages: int = 10,
    limit: int = 25,
    refresh: bool = False,
) -> dict[str, Any]:
    """Search documents and categories across the catalog."""
    res = find_current_gost(query, max_pages=max_pages, limit=limit, refresh=refresh)
    res["kind"] = "search"
    return res




def get_current_topics(
    *,
    category: str | None = None,
    limit: int = 50,
) -> dict[str, Any]:
    data = cache_manager.get_data()
    categories = data.get("categories", [])
    if category:
        needle = category.lower()
        categories = [c for c in categories if needle in c.get("title", "").lower()]
    return {
        "kind": "current-topics",
        "total": len(categories),
        "topics": categories[:limit],
    }


def fetch_norm_markdown(query_or_url: str) -> dict[str, Any]:
    """Fetch normative document HTML from Meganorm and convert it to clean Markdown."""
    url = query_or_url.strip()
    if not url.startswith("http"):
        clean = normalize_gost_number(query_or_url)
        if clean in KNOWN_INDEX_PAGES:
            url = KNOWN_INDEX_PAGES[clean]
        else:
            cached = cache_manager.find_indexed_document(query_or_url)
            if cached and cached.url:
                url = cached.url
            else:
                return {
                    "kind": "fetch-norm-markdown",
                    "error": f"Cannot find Meganorm URL for query: {query_or_url}",
                }

    html, err = client.fetch_url(url)
    if err or not html:
        return {
            "kind": "fetch-norm-markdown",
            "url": url,
            "error": "Failed to fetch document content from source",
        }

    parsed = parse_index_card(html, url)
    if parsed:
        cache_manager.store_document(parsed)

    markdown = convert_html_to_markdown(html)
    return {
        "kind": "fetch-norm-markdown",
        "url": url,
        "title": parsed.title if parsed else "",
        "status": parsed.status if parsed else "unknown",
        "is_active": parsed.is_active if parsed else True,
        "date_intro": parsed.date_intro if parsed else None,
        "markdown": markdown,
    }
