from __future__ import annotations

import json
import logging
import sys
from typing import Any

logger = logging.getLogger(__name__)


class StdioTransport:
    """Universal MCP Stdio Transport.

    Supports:
    1. Standard newline-delimited JSON (MCP stdio default)
    2. Header-framed JSON with Content-Length (LSP-compatible)
    """

    def __init__(self, stdin: Any = None, stdout: Any = None) -> None:
        self.stdin = stdin if stdin is not None else sys.stdin.buffer
        self.stdout = stdout if stdout is not None else sys.stdout.buffer
        self.is_framed = False

    def read_message(self) -> dict[str, Any] | None:
        line = self.stdin.readline()
        if not line:
            return None

        stripped = line.decode("utf-8", errors="replace").strip()

        if not stripped:
            return None

        # Check if client sent newline-delimited JSON directly
        if stripped.startswith("{") or stripped.startswith("["):
            self.is_framed = False
            return json.loads(stripped)

        # Otherwise, parse headers
        headers: dict[str, str] = {}
        if ":" in stripped:
            name, value = stripped.split(":", 1)
            headers[name.lower()] = value.strip()

        while True:
            header_line = self.stdin.readline()
            if not header_line:
                break
            h_stripped = header_line.decode("utf-8", errors="replace").strip()
            if not h_stripped:
                break
            if ":" in h_stripped:
                name, value = h_stripped.split(":", 1)
                headers[name.lower()] = value.strip()

        content_length = int(headers.get("content-length", "0"))
        if content_length <= 0:
            return None

        self.is_framed = True
        raw = self.stdin.read(content_length)
        return json.loads(raw.decode("utf-8"))

    def send_message(self, message: dict[str, Any]) -> None:
        body = json.dumps(message, ensure_ascii=False).encode("utf-8")
        if self.is_framed:
            header = f"Content-Length: {len(body)}\r\n\r\n".encode("ascii")
            self.stdout.write(header)
            self.stdout.write(body)
        else:
            self.stdout.write(body + b"\n")
        self.stdout.flush()

