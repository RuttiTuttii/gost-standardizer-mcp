from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from html import unescape
from html.parser import HTMLParser
from pathlib import Path
from typing import Any
from urllib.error import HTTPError
from urllib.parse import urljoin, urlparse
from urllib.request import Request, urlopen
import json
import re


BASE_URL = "https://meganorm.ru/mega_doc/norm/norm.html"
ROOT_DIR = Path(__file__).resolve().parents[1]
CACHE_PATH = ROOT_DIR / ".cache" / "meganorm" / "catalog.json"
DEFAULT_TIMEOUT = 30
DEFAULT_PAGE_LIMIT = 5
DEFAULT_RESULT_LIMIT = 25
NOISY_TAGS = {"script", "style", "noscript", "footer", "nav", "aside"}
NOISY_DIV_CLASSES = {"w5a7d35a4", "yaac38e20"}


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _normalize(text: str) -> str:
    return " ".join(unescape(text).replace("\xa0", " ").split())


def _slug_from_url(url: str) -> str:
    path = urlparse(url).path.rstrip("/")
    parts = [part for part in path.split("/") if part]
    return parts[-2] if len(parts) >= 2 and parts[-1].endswith("_0.html") else parts[-1].removesuffix("_0.html")


def _is_category_url(url: str) -> bool:
    path = urlparse(url).path
    return bool(re.fullmatch(r"/mega_doc/norm/[^/]+/[^/]+_0\.html", path))


def _is_category_page_url(url: str) -> bool:
    path = urlparse(url).path
    return bool(re.fullmatch(r"/mega_doc/norm/[^/]+/[^/]+_(\d+)\.html", path))


def _page_number_from_url(url: str) -> int:
    match = re.search(r"_(\d+)\.html$", urlparse(url).path)
    return int(match.group(1)) if match else 0


def _fetch_html(url: str) -> str:
    request = Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urlopen(request, timeout=DEFAULT_TIMEOUT) as response:
        charset = response.headers.get_content_charset() or "utf-8"
        return response.read().decode(charset, errors="replace")


class _AnchorParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.items: list[dict[str, str]] = []
        self._current: dict[str, str] | None = None
        self._skip_depth = 0
        self._skip_stack: list[str] = []

    def _should_skip(self, tag: str, attrs: list[tuple[str, str | None]]) -> bool:
        tag = tag.lower()
        if tag in NOISY_TAGS:
            return True
        if tag != "div":
            return False
        attr_map = {key.lower(): (value or "") for key, value in attrs}
        if attr_map.get("data-container", "").lower() == "outer":
            return True
        classes = set(attr_map.get("class", "").split())
        return bool(classes & NOISY_DIV_CLASSES)

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        if self._skip_depth > 0:
            if self._should_skip(tag, attrs):
                self._skip_depth += 1
                self._skip_stack.append(tag)
            return
        if self._should_skip(tag, attrs):
            self._skip_depth = 1
            self._skip_stack = [tag]
            return
        if tag == "a":
            self._current = {"href": dict(attrs).get("href", "") or "", "text": ""}

    def handle_data(self, data: str) -> None:
        if self._skip_depth > 0:
            return
        if self._current is not None:
            self._current["text"] += data

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if self._skip_depth > 0:
            if self._skip_stack and self._skip_stack[-1] == tag:
                self._skip_stack.pop()
                self._skip_depth -= 1
            return
        if tag == "a" and self._current is not None:
            text = _normalize(self._current["text"])
            href = self._current["href"].strip()
            if text or href:
                self.items.append({"text": text, "href": href})
            self._current = None


def _collect_anchors(html: str, base_url: str) -> list[dict[str, str]]:
    parser = _AnchorParser()
    parser.feed(html)
    items: list[dict[str, str]] = []
    for item in parser.items:
        href = item["href"]
        if not href:
            continue
        items.append(
            {
                "text": item["text"],
                "href": urljoin(base_url, href),
            }
        )
    return items


def _empty_cache() -> dict[str, Any]:
    return {
        "source_url": BASE_URL,
        "fetched_at": None,
        "categories": [],
        "category_pages": {},
    }


def load_cache() -> dict[str, Any]:
    if not CACHE_PATH.exists():
        return _empty_cache()
    try:
        return json.loads(CACHE_PATH.read_text(encoding="utf-8"))
    except Exception:
        return _empty_cache()


def save_cache(cache: dict[str, Any]) -> None:
    CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
    CACHE_PATH.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")


def parse_main_index(html: str) -> dict[str, Any]:
    anchors = _collect_anchors(html, BASE_URL)
    categories: list[dict[str, Any]] = []
    seen: set[str] = set()
    for item in anchors:
        href = item["href"]
        text = item["text"]
        if text in {"Главная", "Актуальные документы"}:
            continue
        if not _is_category_url(href):
            continue
        if href in seen:
            continue
        seen.add(href)
        categories.append(
            {
                "title": text,
                "href": href,
                "slug": _slug_from_url(href),
                "origin": "source",
            }
        )
    return {
        "source_url": BASE_URL,
        "fetched_at": _now(),
        "categories": categories,
    }


def parse_category_page(html: str, category_url: str, category_title: str) -> dict[str, Any]:
    anchors = _collect_anchors(html, category_url)
    docs: list[dict[str, Any]] = []
    page_links: dict[int, str] = {}
    seen_docs: set[str] = set()
    for item in anchors:
        href = item["href"]
        text = item["text"]
        if not text or not href:
            continue
        if text.isdigit() and _is_category_page_url(href):
            page_links[int(text) - 1] = href
            continue
        if "/0/" not in urlparse(href).path:
            continue
        if text in {"Главная", "Актуальные документы"}:
            continue
        if href in seen_docs:
            continue
        seen_docs.add(href)
        docs.append(
            {
                "title": text,
                "href": href,
                "origin": "source",
                "category": category_title,
                "page": _page_number_from_url(href),
            }
        )
    return {
        "source_url": category_url,
        "fetched_at": _now(),
        "category": category_title,
        "docs": docs,
        "page_links": page_links,
    }


def _normalize_gost_query(query: str) -> str:
    text = _normalize(query).lower()
    text = text.replace("гост р", " ")
    text = text.replace("гост", " ")
    text = text.replace("№", " ")
    text = re.sub(r"\b(р|rev|ред\.?|редакция)\b", " ", text)
    text = re.sub(r"[^0-9a-zа-яё.\-/]+", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _gost_categories(categories: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [item for item in categories if item["title"].startswith("ГОСТ")]


def _extract_gost_number(text: str) -> str:
    cleaned = _normalize_gost_query(text)
    match = re.search(r"\d+(?:[.\-]\d+)*(?:-\d{4})?", cleaned)
    return match.group(0) if match else ""


def _score_gost_title(query: str, title: str) -> dict[str, Any]:
    query_norm = _normalize_gost_query(query)
    title_norm = _normalize_gost_query(title)
    query_num = _extract_gost_number(query_norm)
    title_num = _extract_gost_number(title_norm)
    if not query_norm and not query_num:
        return {"exact": False, "number_hint": False, "partial": 0, "confidence": 0.0}
    exact = 0
    if query_num and title_num:
        if title_num == query_num or title_num.startswith(f"{query_num}-") or query_num.startswith(f"{title_num}-"):
            exact = 1
    if not exact and query_norm and query_norm in title_norm:
        exact = 1

    number_hint = 1 if query_num and title_num and (title_num == query_num or title_num.startswith(f"{query_num}-")) else 0
    parts = [part for part in query_norm.split() if part and part not in {query_num, "гост", "р"}]
    partial = sum(1 for part in parts if part in title_norm)
    if not partial and query_num and title_num and query_num in title_num:
        partial = 1 if not exact else partial
    confidence = 0.0
    if exact:
        confidence += 0.65
    if number_hint:
        confidence += 0.2
    confidence += min(partial, 4) * 0.05
    confidence = min(confidence, 1.0)
    return {
        "exact": bool(exact),
        "number_hint": bool(number_hint),
        "partial": partial,
        "confidence": confidence,
    }


def ensure_main_index(cache: dict[str, Any] | None = None, *, refresh: bool = False) -> tuple[dict[str, Any], str]:
    cache = cache or load_cache()
    if cache.get("categories") and not refresh:
        return cache, "cache"
    main = parse_main_index(_fetch_html(BASE_URL))
    cache["source_url"] = main["source_url"]
    cache["fetched_at"] = main["fetched_at"]
    cache["categories"] = main["categories"]
    save_cache(cache)
    return cache, "source"


def _resolve_category_matches(categories: list[dict[str, Any]], query: str) -> list[dict[str, Any]]:
    needle = query.strip().lower()
    if not needle:
        return []
    matches = []
    for category in categories:
        hay = f'{category["title"]} {category["slug"]}'.lower()
        if needle in hay:
            matches.append(category)
    return matches


def _category_page_url(category_url: str, page: int) -> str:
    if page <= 0:
        return category_url
    return re.sub(r"_0\.html$", f"_{page}.html", category_url)


def ensure_category_page(
    cache: dict[str, Any],
    category: dict[str, Any],
    page: int = 0,
    *,
    refresh: bool = False,
) -> tuple[dict[str, Any], str]:
    category_pages = cache.setdefault("category_pages", {})
    slug = category["slug"]
    category_bucket = category_pages.setdefault(
        slug,
        {
            "title": category["title"],
            "href": category["href"],
            "pages": {},
        },
    )
    pages = category_bucket.setdefault("pages", {})
    page_key = str(page)
    if page_key in pages and not refresh:
        return cache, "cache"

    page_url = _category_page_url(category["href"], page)
    try:
        html = _fetch_html(page_url)
    except HTTPError as exc:
        if exc.code == 404:
            return cache, "missing"
        raise
    parsed = parse_category_page(html, page_url, category["title"])
    pages[page_key] = parsed
    save_cache(cache)
    return cache, "source"


def resolve_categories(cache: dict[str, Any], query: str) -> list[dict[str, Any]]:
    categories = cache.get("categories", [])
    matches = _resolve_category_matches(categories, query)
    if matches:
        return matches

    needle = query.strip().lower()
    if needle.startswith("гост"):
        gost_matches = [c for c in categories if "гост" in c["title"].lower()]
        if gost_matches:
            return gost_matches
    return []


def list_topics(*, refresh: bool = False) -> dict[str, Any]:
    cache = load_cache()
    cache, origin = ensure_main_index(cache, refresh=refresh)
    categories = [
        {
            "title": item["title"],
            "slug": item["slug"],
            "href": item["href"],
            "origin": origin,
        }
        for item in cache["categories"]
    ]
    return {
        "kind": "topics",
        "base_url": BASE_URL,
        "cache": {
            "path": str(CACHE_PATH),
            "state": "hit" if origin == "cache" else "refreshed",
            "fetched_at": cache.get("fetched_at"),
            "updated_at": cache.get("fetched_at"),
        },
        "topics": categories,
        "note": "records with origin=cache came from the local cache; origin=source were fetched from the нормативный источник",
    }


def get_current_topics(
    *,
    category: str | None = None,
    page: int = 0,
    limit: int = DEFAULT_RESULT_LIMIT,
    refresh: bool = False,
) -> dict[str, Any]:
    cache = load_cache()
    cache, main_origin = ensure_main_index(cache, refresh=refresh)
    if not category:
        topics = [
            {
                "title": item["title"],
                "slug": item["slug"],
                "href": item["href"],
                "origin": main_origin,
            }
            for item in cache["categories"][:limit]
        ]
        return {
            "kind": "current-topics",
            "scope": "catalog",
            "base_url": BASE_URL,
            "cache": {
                "path": str(CACHE_PATH),
                "state": "hit" if main_origin == "cache" else "refreshed",
                "fetched_at": cache.get("fetched_at"),
                "updated_at": cache.get("fetched_at"),
            },
            "topics": topics,
            "note": "origin=cache means the topic came from the local cache; origin=source means it was fetched from the нормативный source",
        }

    matches = resolve_categories(cache, category)
    if not matches:
        return {
            "kind": "current-topics",
            "scope": "category",
            "category_query": category,
            "cache": {
                "path": str(CACHE_PATH),
                "state": "hit" if main_origin == "cache" else "refreshed",
                "fetched_at": cache.get("fetched_at"),
                "updated_at": cache.get("fetched_at"),
            },
            "topics": [],
            "note": "no matching category was found in cache or source",
        }

    docs: list[dict[str, Any]] = []
    page_origin = "cache"
    for matched in matches[:3]:
        cache, origin = ensure_category_page(cache, matched, page=page, refresh=refresh)
        if origin == "missing":
            continue
        page_origin = origin
        page_key = str(page)
        page_data = cache["category_pages"][matched["slug"]]["pages"][page_key]
        docs.extend(page_data["docs"][:limit])

    return {
        "kind": "current-topics",
        "scope": "category",
        "category_query": category,
        "matched_categories": [
            {"title": item["title"], "slug": item["slug"], "href": item["href"], "origin": main_origin}
            for item in matches[:3]
        ],
        "cache": {
            "path": str(CACHE_PATH),
            "state": "hit" if page_origin == "cache" else "refreshed",
            "fetched_at": cache.get("fetched_at"),
            "updated_at": cache.get("fetched_at"),
        },
        "topics": docs[:limit],
        "note": "origin=cache means the topic came from the local cache; origin=source means it was fetched from the нормативный source",
    }


def search_catalog(
    query: str,
    *,
    category: str | None = None,
    max_pages: int = DEFAULT_PAGE_LIMIT,
    limit: int = DEFAULT_RESULT_LIMIT,
    refresh: bool = False,
) -> dict[str, Any]:
    cache = load_cache()
    cache, main_origin = ensure_main_index(cache, refresh=refresh)
    query_norm = query.strip().lower()
    category_filter = resolve_categories(cache, category) if category else []
    query_matches = _resolve_category_matches(cache["categories"], query)

    categories_to_scan: list[dict[str, Any]] = []
    for item in category_filter + query_matches:
        if item not in categories_to_scan:
            categories_to_scan.append(item)

    if not categories_to_scan and query_norm:
        for item in cache["categories"]:
            if query_norm in item["title"].lower() or query_norm in item["slug"].lower():
                categories_to_scan.append(item)

    if not categories_to_scan and query_norm and (re.search(r"\bгост\b", query_norm) or re.search(r"\d", query_norm)):
        for item in cache["categories"]:
            if "гост" in item["title"].lower():
                categories_to_scan.append(item)

    category_hits: list[dict[str, Any]] = []
    for item in categories_to_scan[:10]:
        category_hits.append(
            {
                "title": item["title"],
                "slug": item["slug"],
                "href": item["href"],
                "origin": main_origin,
            }
        )

    document_hits: list[dict[str, Any]] = []
    pages_scanned: list[dict[str, Any]] = []
    for item in categories_to_scan[:5]:
        for page in range(0, max_pages):
            cache, origin = ensure_category_page(cache, item, page=page, refresh=refresh)
            if origin == "missing":
                break
            page_data = cache["category_pages"][item["slug"]]["pages"][str(page)]
            pages_scanned.append(
                {
                    "category": item["title"],
                    "page": page,
                    "origin": origin,
                    "url": page_data["source_url"],
                }
            )
            for doc in page_data["docs"]:
                hay = f'{doc["title"]} {doc["category"]}'.lower()
                score = _score_gost_title(query, doc["title"])
                text_match = bool(query_norm and query_norm in hay)
                if score["exact"] or score["number_hint"] or score["partial"] or text_match:
                    record = dict(doc)
                    record["origin"] = origin
                    confidence = score["confidence"]
                    if text_match and confidence == 0.0:
                        confidence = min(0.45 + min(len(query_norm), 40) / 200, 0.85)
                    record["confidence"] = round(confidence, 3)
                    record["match"] = {
                        "exact": score["exact"],
                        "number_hint": score["number_hint"],
                        "partial": score["partial"],
                        "text": text_match,
                    }
                    document_hits.append(record)
                    if len(document_hits) >= limit:
                        break
            if len(document_hits) >= limit:
                break
        if len(document_hits) >= limit:
            break

    return {
        "kind": "search",
        "query": query,
        "category_filter": category,
        "base_url": BASE_URL,
        "cache": {
            "path": str(CACHE_PATH),
            "state": "hit" if main_origin == "cache" else "refreshed",
            "fetched_at": cache.get("fetched_at"),
            "updated_at": cache.get("fetched_at"),
        },
        "categories": category_hits,
        "documents": document_hits,
        "pages_scanned": pages_scanned,
        "note": "origin=cache means the record came from the local cache; origin=source means it was fetched live from the нормативный source",
    }


def find_current_gost(
    query: str,
    *,
    max_pages: int = 10,
    limit: int = DEFAULT_RESULT_LIMIT,
    refresh: bool = False,
) -> dict[str, Any]:
    cache = load_cache()
    cache, main_origin = ensure_main_index(cache, refresh=refresh)
    categories = _gost_categories(cache.get("categories", []))
    category_records = [
        {
            "title": item["title"],
            "slug": item["slug"],
            "href": item["href"],
            "origin": main_origin,
        }
        for item in categories
    ]

    normalized_query = _normalize_gost_query(query)
    documents: list[dict[str, Any]] = []
    pages_scanned: list[dict[str, Any]] = []

    for category in categories:
        for page in range(0, max_pages):
            cache, origin = ensure_category_page(cache, category, page=page, refresh=refresh)
            if origin == "missing":
                break
            page_data = cache["category_pages"][category["slug"]]["pages"][str(page)]
            pages_scanned.append(
                {
                    "category": category["title"],
                    "page": page,
                    "origin": origin,
                    "url": page_data["source_url"],
                }
            )
            for doc in page_data["docs"]:
                score = _score_gost_title(normalized_query, doc["title"])
                if score["exact"] or score["number_hint"] or score["partial"]:
                    documents.append(
                        {
                            "title": doc["title"],
                            "href": doc["href"],
                            "category": doc["category"],
                            "page": doc["page"],
                            "origin": origin,
                            "confidence": score["confidence"],
                            "match": {
                                "exact": score["exact"],
                                "number_hint": score["number_hint"],
                                "partial": score["partial"],
                            },
                        }
                    )
                    if len(documents) >= limit:
                        break
            if len(documents) >= limit:
                break
        if len(documents) >= limit:
            break

    return {
        "kind": "current-gost-search",
        "query": query,
        "normalized_query": normalized_query,
        "scope": "current-gost-categories",
        "base_url": BASE_URL,
        "cache": {
            "path": str(CACHE_PATH),
            "state": "hit" if main_origin == "cache" else "refreshed",
            "fetched_at": cache.get("fetched_at"),
            "updated_at": cache.get("fetched_at"),
        },
        "categories": category_records,
        "documents": documents,
        "pages_scanned": pages_scanned,
        "note": "origin=cache means the record came from the local cache; origin=source means it was fetched live from the нормативный source",
    }


def refresh_catalog(*, category: str | None = None, max_pages: int = DEFAULT_PAGE_LIMIT) -> dict[str, Any]:
    cache = load_cache()
    cache, main_origin = ensure_main_index(cache, refresh=True)
    refreshed_categories: list[dict[str, Any]] = []
    pages_refreshed = 0

    if category:
        matches = resolve_categories(cache, category)
    else:
        matches = cache.get("categories", [])

    for item in matches[:10]:
        refreshed_categories.append(
            {
                "title": item["title"],
                "slug": item["slug"],
                "href": item["href"],
                "origin": main_origin,
            }
        )
        for page in range(0, max_pages):
            cache, origin = ensure_category_page(cache, item, page=page, refresh=True)
            if origin == "missing":
                break
            pages_refreshed += 1

    return {
        "kind": "refresh",
        "base_url": BASE_URL,
        "cache": {
            "path": str(CACHE_PATH),
            "state": "refreshed",
            "fetched_at": cache.get("fetched_at"),
            "updated_at": cache.get("fetched_at"),
        },
        "categories": refreshed_categories,
        "pages_refreshed": pages_refreshed,
        "note": "origin=source is implied for refreshed entries because they were fetched live from the нормативный source",
    }
