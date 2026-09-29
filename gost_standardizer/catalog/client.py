from __future__ import annotations

import logging
import urllib.error
import urllib.parse
import urllib.request

logger = logging.getLogger(__name__)

DEFAULT_TIMEOUT = 12
DEFAULT_USER_AGENT = (
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
)


class MeganormClient:
    def __init__(
        self,
        base_url: str = "https://meganorm.ru",
        timeout: int = DEFAULT_TIMEOUT,
        user_agent: str = DEFAULT_USER_AGENT,
    ) -> None:
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self.user_agent = user_agent
        self.fetch_hook = None

    def fetch_url(self, url: str) -> tuple[str, bool]:
        """Fetch URL content as decoded string.

        Returns (content, is_error).
        """
        if self.fetch_hook is not None:
            try:
                res = self.fetch_hook(url)
                if isinstance(res, tuple):
                    return res
                return str(res), False
            except Exception as exc:
                logger.warning("Fetch hook failed for %s: %s", url, exc)
                return "", True

        parsed = urllib.parse.urlparse(url)

        if not parsed.scheme:
            url = f"{self.base_url}/{url.lstrip('/')}"

        headers = {
            "User-Agent": self.user_agent,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
        }
        req = urllib.request.Request(url, headers=headers)

        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                raw = resp.read()
                # Determine encoding
                charset = resp.headers.get_content_charset()
                if not charset:
                    charset = "utf-8"
                try:
                    return raw.decode(charset), False
                except UnicodeDecodeError:
                    try:
                        return raw.decode("windows-1251"), False
                    except UnicodeDecodeError:
                        return raw.decode("utf-8", errors="replace"), False
        except (urllib.error.URLError, TimeoutError, OSError) as exc:
            logger.warning("Failed to fetch %s: %s", url, exc)
            return "", True
