from __future__ import annotations

from gost_standardizer.catalog.cache import cache_manager
from gost_standardizer.catalog.client import MeganormClient
from gost_standardizer.catalog.parser import parse_index_card, parse_list_links
from gost_standardizer.catalog.service import (
    fetch_norm_markdown,
    find_current_gost,
    get_current_topics,
    search_catalog,
)

__all__ = [
    "MeganormClient",
    "cache_manager",
    "fetch_norm_markdown",
    "find_current_gost",
    "get_current_topics",
    "parse_index_card",
    "parse_list_links",
    "search_catalog",
]
