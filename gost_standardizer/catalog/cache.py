from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any

from gost_standardizer.models.document import NormDocument, normalize_gost_number

logger = logging.getLogger(__name__)

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
SEED_PATH = REPO_ROOT / "data" / "catalog_seed.json"
CACHE_PATH = Path.home() / ".cache" / "gost_standardizer" / "catalog.json"


def _empty_cache() -> dict[str, Any]:
    return {
        "version": "1.0",
        "fetched_at": None,
        "indexed_documents": {},
        "categories": [],
        "category_pages": {},
    }


class CatalogCacheManager:
    def __init__(self, seed_path: Path = SEED_PATH, cache_path: Path = CACHE_PATH) -> None:
        self.seed_path = seed_path
        self.cache_path = cache_path
        self._data: dict[str, Any] | None = None

    def get_data(self) -> dict[str, Any]:
        if self._data is not None:
            return self._data

        cache = _empty_cache()

        # 1. Load seed cache first
        if self.seed_path.exists():
            try:
                seed = json.loads(self.seed_path.read_text(encoding="utf-8"))
                cache["categories"] = seed.get("categories", [])
                cache["category_pages"] = seed.get("category_pages", {})
                cache["fetched_at"] = seed.get("fetched_at")
                cache["indexed_documents"] = seed.get("indexed_documents", {})
            except Exception as exc:
                logger.warning("Failed to load catalog seed: %s", exc)

        # 2. Overlay user cache if it exists
        if self.cache_path.exists():
            try:
                user_cache = json.loads(self.cache_path.read_text(encoding="utf-8"))
                if user_cache.get("categories"):
                    cache["categories"] = user_cache["categories"]
                if user_cache.get("category_pages"):
                    cache["category_pages"].update(user_cache["category_pages"])
                if user_cache.get("indexed_documents"):
                    cache["indexed_documents"].update(user_cache["indexed_documents"])
                if user_cache.get("fetched_at"):
                    cache["fetched_at"] = user_cache["fetched_at"]
            except Exception as exc:
                logger.warning("Failed to load user cache: %s", exc)

        self._data = cache
        return self._data

    def save(self) -> None:
        if self._data is None:
            return
        try:
            self.cache_path.parent.mkdir(parents=True, exist_ok=True)
            self.cache_path.write_text(
                json.dumps(self._data, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception as exc:
            logger.warning("Failed to save cache to %s: %s", self.cache_path, exc)

    def find_indexed_document(self, query: str) -> NormDocument | None:
        data = self.get_data()
        clean = normalize_gost_number(query)
        indexed = data.get("indexed_documents", {})

        # Direct key match
        if clean in indexed:
            return NormDocument.from_dict(indexed[clean])

        # Match by designation, clean_number or title
        clean_lower = clean.lower()
        for key, doc_dict in indexed.items():
            if key.lower() == clean_lower:
                return NormDocument.from_dict(doc_dict)
            doc_clean = str(doc_dict.get("clean_number", "")).lower()
            if doc_clean == clean_lower:
                return NormDocument.from_dict(doc_dict)
            doc_desig = str(doc_dict.get("designation", "")).lower()
            if clean_lower in doc_desig or doc_desig in clean_lower:
                return NormDocument.from_dict(doc_dict)

        return None

    def store_document(self, doc: NormDocument) -> None:
        data = self.get_data()
        indexed = data.setdefault("indexed_documents", {})
        indexed[doc.clean_number] = doc.to_dict()
        self.save()


# Default singleton instance
cache_manager = CatalogCacheManager()
