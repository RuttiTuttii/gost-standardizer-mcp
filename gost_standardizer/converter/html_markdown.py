from __future__ import annotations

from html.parser import HTMLParser
import os
import shutil
import subprocess
from typing import Any

# Try importing the high-performance xberg-io/html-to-markdown engine
try:
    import html_to_markdown as _native_h2m
except ImportError:
    _native_h2m = None


class _SimpleHtmlToMarkdownParser(HTMLParser):
    """Pure-Python fallback for converting HTML to Markdown without external dependencies."""

    def __init__(self) -> None:
        super().__init__()
        self.output: list[str] = []
        self._tag_stack: list[str] = []
        self._current_href: str | None = None
        self._in_table = False
        self._table_rows: list[list[str]] = []
        self._current_row: list[str] = []
        self._current_cell: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attr_dict = dict(attrs)
        tag = tag.lower()
        self._tag_stack.append(tag)

        if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
            level = int(tag[1])
            self._write(f"\n\n{'#' * level} ")
        elif tag == "p":
            self._write("\n\n")
        elif tag == "br":
            self._write("\n")
        elif tag in ("strong", "b"):
            self._write("**")
        elif tag in ("em", "i"):
            self._write("*")
        elif tag == "a":
            self._current_href = attr_dict.get("href")
            self._write("[")
        elif tag == "li":
            self._write("\n- ")
        elif tag == "table":
            self._in_table = True
            self._table_rows = []
        elif tag == "tr":
            self._current_row = []
        elif tag in ("td", "th"):
            self._current_cell = []

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if self._tag_stack and self._tag_stack[-1] == tag:
            self._tag_stack.pop()

        if tag in ("h1", "h2", "h3", "h4", "h5", "h6", "p"):
            self._write("\n")
        elif tag in ("strong", "b"):
            self._write("**")
        elif tag in ("em", "i"):
            self._write("*")
        elif tag == "a":
            href = self._current_href or ""
            self._write(f"]({href})")
            self._current_href = None
        elif tag in ("td", "th"):
            cell_text = "".join(self._current_cell).strip().replace("\n", " ")
            self._current_row.append(cell_text)
            self._current_cell = []
        elif tag == "tr":
            if self._current_row:
                self._table_rows.append(self._current_row)
            self._current_row = []
        elif tag == "table":
            self._in_table = False
            self._format_table()

    def handle_data(self, data: str) -> None:
        if self._in_table and self._tag_stack and self._tag_stack[-1] in ("td", "th"):
            self._current_cell.append(data)
        else:
            self._write(data)

    def _write(self, text: str) -> None:
        self.output.append(text)

    def _format_table(self) -> None:
        if not self._table_rows:
            return
        num_cols = max(len(row) for row in self._table_rows)
        if num_cols == 0:
            return

        lines: list[str] = ["\n\n"]
        header = self._table_rows[0]
        while len(header) < num_cols:
            header.append("")
        lines.append("| " + " | ".join(header) + " |\n")
        lines.append("| " + " | ".join(["---"] * num_cols) + " |\n")

        for row in self._table_rows[1:]:
            padded = list(row)
            while len(padded) < num_cols:
                padded.append("")
            lines.append("| " + " | ".join(padded) + " |\n")
        lines.append("\n")
        self.output.extend(lines)

    def get_markdown(self) -> str:
        raw = "".join(self.output)
        # Normalize consecutive blank lines
        lines = [line.rstrip() for line in raw.split("\n")]
        result = []
        blank_count = 0
        for line in lines:
            if not line:
                blank_count += 1
                if blank_count <= 2:
                    result.append(line)
            else:
                blank_count = 0
                result.append(line)
        return "\n".join(result).strip()


def convert_html_to_markdown(html_text: str, options: dict[str, Any] | None = None) -> str:
    """Convert HTML string to clean Markdown.

    Strategy:
    1. Fast native bindings (`html_to_markdown` from xberg-io)
    2. CLI binary (`html-to-markdown` in PATH or bin/)
    3. Pure Python fallback parser
    """
    if not html_text:
        return ""

    # Strategy 1: Python library
    if _native_h2m is not None:
        try:
            result = _native_h2m.convert(html_text)
            if hasattr(result, "content"):
                return str(result.content).strip()
            return str(result).strip()
        except Exception:
            pass

    # Strategy 2: CLI binary
    binary_path = shutil.which("html-to-markdown") or (
        os.path.join(os.path.dirname(__file__), "..", "..", "bin", "html-to-markdown")
    )
    if binary_path and os.path.isfile(binary_path) and os.access(binary_path, os.X_OK):
        try:
            proc = subprocess.run(
                [binary_path],
                input=html_text.encode("utf-8"),
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=10,
                check=False,
            )
            if proc.returncode == 0:
                return proc.stdout.decode("utf-8", errors="replace").strip()
        except Exception:
            pass

    # Strategy 3: Pure Python fallback
    parser = _SimpleHtmlToMarkdownParser()
    parser.feed(html_text)
    parser.close()
    return parser.get_markdown()
