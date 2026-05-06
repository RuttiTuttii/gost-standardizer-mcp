from __future__ import annotations

import json
import sys
from typing import Any

from gost_standardizer import inspect_document, list_presets, standardize_document
from meganorm_catalog import find_current_gost, get_current_topics, refresh_catalog, search_catalog


SERVER_NAME = "gost-standardizer"
SERVER_VERSION = "0.1.0"
PROTOCOL_VERSION = "2024-11-05"


TOOLS = [
    {
        "name": "list_presets",
        "description": "List the built-in document presets and what they are for.",
        "inputSchema": {
            "type": "object",
            "properties": {},
            "additionalProperties": False,
        },
    },
    {
        "name": "inspect_document",
        "description": "Inspect a DOCX/DOCM file and report formatting issues, samples, and a suggested preset.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "path": {
                    "type": "string",
                    "description": "Path to the source DOCX or DOCM file.",
                },
                "sample_size": {
                    "type": "integer",
                    "minimum": 1,
                    "maximum": 25,
                    "default": 8,
                },
            },
            "required": ["path"],
            "additionalProperties": False,
        },
    },
    {
        "name": "standardize_document",
        "description": "Standardize a DOCX/DOCM file and save a cleaned copy with GOST-style formatting.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "path": {
                    "type": "string",
                    "description": "Path to the source DOCX or DOCM file.",
                },
                "output_path": {
                    "type": "string",
                    "description": "Optional destination path for the standardized copy.",
                },
                "preset": {
                    "type": "string",
                    "enum": ["report", "office", "technical", "legacy-college"],
                    "default": "report",
                },
                "overwrite": {
                    "type": "boolean",
                    "default": False,
                },
                "aggressive": {
                    "type": "boolean",
                    "default": False,
                },
            },
            "required": ["path"],
            "additionalProperties": False,
        },
    },
    {
        "name": "refresh_meganorm_cache",
        "description": "Refresh the local meganorm HTML cache from the live нормативный source.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "category": {
                    "type": "string",
                    "description": "Optional category name or fragment to refresh selectively.",
                },
                "max_pages": {
                    "type": "integer",
                    "minimum": 1,
                    "maximum": 10,
                    "default": 5,
                },
            },
            "additionalProperties": False,
        },
    },
    {
        "name": "search_meganorm_catalog",
        "description": "Search cached and live meganorm categories/documents by category name or document title.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "query": {
                    "type": "string",
                    "description": "Search text for categories or document titles.",
                },
                "category": {
                    "type": "string",
                    "description": "Optional category filter.",
                },
                "max_pages": {
                    "type": "integer",
                    "minimum": 1,
                    "maximum": 10,
                    "default": 5,
                },
                "limit": {
                    "type": "integer",
                    "minimum": 1,
                    "maximum": 100,
                    "default": 25,
                },
                "refresh": {
                    "type": "boolean",
                    "default": False,
                },
            },
            "required": ["query"],
            "additionalProperties": False,
        },
    },
    {
        "name": "get_meganorm_topics",
        "description": "Get current categories or document topics from the live нормативный source, with cache/source origin labels.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "category": {
                    "type": "string",
                    "description": "Optional category name or fragment.",
                },
                "page": {
                    "type": "integer",
                    "minimum": 0,
                    "maximum": 50,
                    "default": 0,
                },
                "limit": {
                    "type": "integer",
                    "minimum": 1,
                    "maximum": 100,
                    "default": 25,
                },
                "refresh": {
                    "type": "boolean",
                    "default": False,
                },
            },
            "additionalProperties": False,
        },
    },
    {
        "name": "find_current_gost",
        "description": "Find current GOST and GOST R documents only, with exact-number matching over the актуализированная база.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "query": {
                    "type": "string",
                    "description": "GOST number or title fragment, for example '7.32-2017' or 'ГОСТ 2.105'.",
                },
                "max_pages": {
                    "type": "integer",
                    "minimum": 1,
                    "maximum": 20,
                    "default": 10,
                },
                "limit": {
                    "type": "integer",
                    "minimum": 1,
                    "maximum": 100,
                    "default": 25,
                },
                "refresh": {
                    "type": "boolean",
                    "default": False,
                },
            },
            "required": ["query"],
            "additionalProperties": False,
        },
    },
]


def _send(payload: dict[str, Any]) -> None:
    body = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    header = f"Content-Length: {len(body)}\r\n\r\n".encode("ascii")
    sys.stdout.buffer.write(header)
    sys.stdout.buffer.write(body)
    sys.stdout.buffer.flush()


def _read() -> dict[str, Any] | None:
    headers: dict[str, str] = {}
    while True:
        line = sys.stdin.buffer.readline()
        if not line:
            return None
        stripped = line.decode("utf-8", errors="replace").strip()
        if not stripped:
            break
        if ":" in stripped:
            name, value = stripped.split(":", 1)
            headers[name.lower()] = value.strip()
    content_length = int(headers.get("content-length", "0"))
    if content_length <= 0:
        return None
    raw = sys.stdin.buffer.read(content_length)
    return json.loads(raw.decode("utf-8"))


def _result(content: str, *, is_error: bool = False) -> dict[str, Any]:
    return {
        "content": [
            {
                "type": "text",
                "text": content,
            }
        ],
        "isError": is_error,
    }


def _handle_tools_call(arguments: dict[str, Any]) -> dict[str, Any]:
    name = arguments.get("name")
    params = arguments.get("arguments") or {}

    if name == "list_presets":
        return _result(json.dumps(list_presets(), ensure_ascii=False, indent=2))

    if name == "inspect_document":
        report = inspect_document(
            path=params["path"],
            sample_size=int(params.get("sample_size", 8)),
        )
        return _result(json.dumps(report, ensure_ascii=False, indent=2))

    if name == "standardize_document":
        report = standardize_document(
            path=params["path"],
            output_path=params.get("output_path"),
            preset_name=params.get("preset"),
            overwrite=bool(params.get("overwrite", False)),
            aggressive=bool(params.get("aggressive", False)),
        )
        return _result(json.dumps(report, ensure_ascii=False, indent=2))

    if name == "refresh_meganorm_cache":
        report = refresh_catalog(
            category=params.get("category"),
            max_pages=int(params.get("max_pages", 5)),
        )
        return _result(json.dumps(report, ensure_ascii=False, indent=2))

    if name == "search_meganorm_catalog":
        report = search_catalog(
            query=params["query"],
            category=params.get("category"),
            max_pages=int(params.get("max_pages", 5)),
            limit=int(params.get("limit", 25)),
            refresh=bool(params.get("refresh", False)),
        )
        return _result(json.dumps(report, ensure_ascii=False, indent=2))

    if name == "get_meganorm_topics":
        report = get_current_topics(
            category=params.get("category"),
            page=int(params.get("page", 0)),
            limit=int(params.get("limit", 25)),
            refresh=bool(params.get("refresh", False)),
        )
        return _result(json.dumps(report, ensure_ascii=False, indent=2))

    if name == "find_current_gost":
        report = find_current_gost(
            query=params["query"],
            max_pages=int(params.get("max_pages", 10)),
            limit=int(params.get("limit", 25)),
            refresh=bool(params.get("refresh", False)),
        )
        return _result(json.dumps(report, ensure_ascii=False, indent=2))

    raise KeyError(f"Unknown tool: {name}")


def _dispatch(request: dict[str, Any]) -> dict[str, Any] | None:
    method = request.get("method")
    request_id = request.get("id")
    params = request.get("params") or {}

    if method == "initialize":
        return {
            "jsonrpc": "2.0",
            "id": request_id,
            "result": {
                "protocolVersion": params.get("protocolVersion", PROTOCOL_VERSION),
                "capabilities": {
                    "tools": {},
                },
                "serverInfo": {
                    "name": SERVER_NAME,
                    "version": SERVER_VERSION,
                },
            },
        }

    if method == "notifications/initialized":
        return None

    if method == "tools/list":
        return {
            "jsonrpc": "2.0",
            "id": request_id,
            "result": {
                "tools": TOOLS,
            },
        }

    if method == "tools/call":
        try:
            result = _handle_tools_call(params)
        except Exception as exc:  # noqa: BLE001
            return {
                "jsonrpc": "2.0",
                "id": request_id,
                "error": {
                    "code": -32603,
                    "message": str(exc),
                },
            }
        return {
            "jsonrpc": "2.0",
            "id": request_id,
            "result": result,
        }

    if request_id is None:
        return None

    return {
        "jsonrpc": "2.0",
        "id": request_id,
        "error": {
            "code": -32601,
            "message": f"Method not found: {method}",
        },
    }


def main() -> int:
    while True:
        request = _read()
        if request is None:
            return 0
        response = _dispatch(request)
        if response is not None and "id" in response:
            _send(response)


if __name__ == "__main__":
    raise SystemExit(main())
