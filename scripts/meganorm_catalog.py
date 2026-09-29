from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from gost_standardizer.catalog.cache import (
    CACHE_PATH,
    SEED_PATH,
    CatalogCacheManager,
    cache_manager,
)
from gost_standardizer.catalog.client import DEFAULT_TIMEOUT, DEFAULT_USER_AGENT, MeganormClient
from gost_standardizer.catalog.parser import parse_index_card, parse_list_links
from gost_standardizer.catalog.service import (
    BASE_URL,
    fetch_norm_markdown,
    find_current_gost,
    get_current_topics,
    search_catalog,
)
from gost_standardizer.models.document import (
    DocumentStatus,
    NormDocument,
    normalize_gost_number as _normalize_gost_query,
)


def load_cache() -> dict[str, Any]:
    return cache_manager.get_data()


def save_cache(cache: dict[str, Any]) -> None:
    cache_manager._data = cache
    cache_manager.save()


def refresh_catalog() -> dict[str, Any]:
    cache_manager._data = None
    return {"status": "cleared"}


def _fetch_html(url: str) -> str:
    from gost_standardizer.catalog.service import client
    html, err = client.fetch_url(url)
    if err:
        from urllib.error import URLError
        raise URLError(f"Failed to fetch {url}")
    return html


# Allow mocking _fetch_html directly on meganorm_catalog
from gost_standardizer.catalog.service import client as _srv_client

_srv_client.fetch_hook = lambda u: _fetch_html(u)



if __name__ == "__main__":
    from gost_standardizer.cli.main import main
    raise SystemExit(main())
