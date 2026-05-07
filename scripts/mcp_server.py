from __future__ import annotations

import json
import sys
from typing import Any, Callable

from gost_standardizer import (
    compare_to_preset,
    explain_preset,
    inspect_document,
    list_presets,
    list_profiles,
    load_profile,
    save_profile,
    standardize_document,
    validate_document,
)
from meganorm_catalog import find_current_gost, get_current_topics, refresh_catalog, search_catalog


SERVER_NAME = "gost-standardizer"
SERVER_VERSION = "0.2.0"
PROTOCOL_VERSION = "2024-11-05"


def _tool_schema(parameters: dict[str, Any]) -> dict[str, Any]:
    return {
        "type": "object",
        "properties": parameters,
        "additionalProperties": False,
    }


TOOLS = [
    {
        "name": "list_presets",
        "description": "List the built-in document presets and what they are for.",
        "inputSchema": _tool_schema({}),
    },
    {
        "name": "list_profiles",
        "description": "List built-in and saved document profiles available to the standardizer.",
        "inputSchema": _tool_schema({}),
    },
    {
        "name": "load_profile",
        "description": "Load a built-in or saved profile by name or file path.",
        "inputSchema": _tool_schema(
            {
                "name": {
                    "type": "string",
                    "description": "Profile key, saved profile name, or JSON file path.",
                }
            }
        ),
    },
    {
        "name": "save_profile",
        "description": "Save a profile JSON file from one of the built-in presets.",
        "inputSchema": _tool_schema(
            {
                "name": {"type": "string", "description": "Name of the saved profile."},
                "preset": {
                    "type": "string",
                    "enum": ["report", "office", "technical", "legacy-college"],
                    "default": "report",
                },
                "title": {
                    "type": "string",
                    "description": "Optional profile title.",
                },
                "description": {
                    "type": "string",
                    "description": "Optional profile description.",
                },
                "kind": {
                    "type": "string",
                    "default": "organization",
                    "description": "Profile kind, such as organization or custom.",
                },
                "notes": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Optional notes stored in the profile file.",
                },
            }
        ),
    },
    {
        "name": "inspect_document",
        "description": "Inspect a DOCX/DOCM/DOC file and report formatting issues, samples, and a suggested preset.",
        "inputSchema": _tool_schema(
            {
                "path": {
                    "type": "string",
                    "description": "Path to the source document.",
                },
                "sample_size": {
                    "type": "integer",
                    "minimum": 1,
                    "maximum": 25,
                    "default": 8,
                },
            }
        ),
    },
    {
        "name": "explain_preset",
        "description": "Explain why a preset was chosen for a document, with keyword and filename signals.",
        "inputSchema": _tool_schema(
            {
                "path": {"type": "string", "description": "Path to the source document."},
                "preset": {
                    "type": "string",
                    "description": "Optional preset to explain instead of the guessed one.",
                },
                "sample_size": {
                    "type": "integer",
                    "minimum": 1,
                    "maximum": 25,
                    "default": 8,
                },
            }
        ),
    },
    {
        "name": "validate_document",
        "description": "Validate a DOCX/DOCM/DOC file against a preset or profile and return structured issues.",
        "inputSchema": _tool_schema(
            {
                "path": {"type": "string", "description": "Path to the source document."},
                "preset": {
                    "type": "string",
                    "enum": ["report", "office", "technical", "legacy-college"],
                    "description": "Optional built-in preset to validate against.",
                },
                "profile": {
                    "type": "string",
                    "description": "Optional profile name or path. Takes priority over preset.",
                },
                "aggressive": {
                    "type": "boolean",
                    "default": False,
                    "description": "Use stricter heuristics for paragraph classification.",
                },
                "sample_size": {
                    "type": "integer",
                    "minimum": 1,
                    "maximum": 25,
                    "default": 8,
                },
            }
        ),
    },
    {
        "name": "compare_to_preset",
        "description": "Compare a document to a preset or profile and return a diff-style summary.",
        "inputSchema": _tool_schema(
            {
                "path": {"type": "string", "description": "Path to the source document."},
                "preset": {
                    "type": "string",
                    "enum": ["report", "office", "technical", "legacy-college"],
                    "description": "Optional built-in preset to compare against.",
                },
                "profile": {
                    "type": "string",
                    "description": "Optional profile name or path. Takes priority over preset.",
                },
                "aggressive": {
                    "type": "boolean",
                    "default": False,
                },
                "sample_size": {
                    "type": "integer",
                    "minimum": 1,
                    "maximum": 25,
                    "default": 8,
                },
            }
        ),
    },
    {
        "name": "standardize_document",
        "description": "Standardize a DOCX/DOCM/DOC file and save a cleaned copy with profile-driven formatting.",
        "inputSchema": _tool_schema(
            {
                "path": {"type": "string", "description": "Path to the source document."},
                "output_path": {
                    "type": "string",
                    "description": "Optional destination path for the standardized copy.",
                },
                "preset": {
                    "type": "string",
                    "enum": ["report", "office", "technical", "legacy-college"],
                    "default": "report",
                },
                "profile": {
                    "type": "string",
                    "description": "Optional profile name or file path. Takes priority over preset.",
                },
                "overwrite": {
                    "type": "boolean",
                    "default": False,
                },
                "aggressive": {
                    "type": "boolean",
                    "default": False,
                },
                "fix_page_setup": {
                    "type": "boolean",
                    "default": True,
                },
                "fix_styles": {
                    "type": "boolean",
                    "default": True,
                },
                "fix_paragraphs": {
                    "type": "boolean",
                    "default": True,
                },
                "fix_tables": {
                    "type": "boolean",
                    "default": True,
                },
            }
        ),
    },
    {
        "name": "refresh_meganorm_cache",
        "description": "Refresh the local meganorm HTML cache from the live нормативный source.",
        "inputSchema": _tool_schema(
            {
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
            }
        ),
    },
    {
        "name": "search_meganorm_catalog",
        "description": "Search cached and live meganorm categories/documents by category name or document title.",
        "inputSchema": _tool_schema(
            {
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
            }
        ),
    },
    {
        "name": "get_meganorm_topics",
        "description": "Get current categories or document topics from the live нормативный source, with cache/source origin labels.",
        "inputSchema": _tool_schema(
            {
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
            }
        ),
    },
    {
        "name": "find_current_gost",
        "description": "Find current GOST and GOST R documents only, with exact-number matching over the актуализированная база.",
        "inputSchema": _tool_schema(
            {
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
            }
        ),
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


def _serialize(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, indent=2)


def _tool_error(name: str, exc: Exception) -> dict[str, Any]:
    return _result(
        _serialize(
            {
                "kind": "tool-error",
                "tool": name,
                "error": type(exc).__name__,
                "message": str(exc),
            }
        ),
        is_error=True,
    )


def _call_tool(name: str, fn: Callable[..., Any], **kwargs: Any) -> dict[str, Any]:
    try:
        return _result(_serialize(fn(**kwargs)))
    except (FileNotFoundError, ValueError, RuntimeError, KeyError) as exc:
        return _tool_error(name, exc)


def _handle_tools_call(arguments: dict[str, Any]) -> dict[str, Any]:
    name = arguments.get("name")
    params = arguments.get("arguments") or {}

    if name == "list_presets":
        return _result(_serialize(list_presets()))
    if name == "list_profiles":
        return _result(_serialize(list_profiles()))
    if name == "load_profile":
        return _call_tool(name, load_profile, name=params["name"])
    if name == "save_profile":
        return _call_tool(
            name,
            save_profile,
            name=params["name"],
            preset_name=params.get("preset"),
            title=params.get("title"),
            description=params.get("description"),
            kind=params.get("kind", "organization"),
            notes=params.get("notes"),
        )
    if name == "inspect_document":
        return _call_tool(name, inspect_document, path=params["path"], sample_size=int(params.get("sample_size", 8)))
    if name == "explain_preset":
        return _call_tool(
            name,
            explain_preset,
            path=params["path"],
            preset_name=params.get("preset"),
            sample_size=int(params.get("sample_size", 8)),
        )
    if name == "validate_document":
        return _call_tool(
            name,
            validate_document,
            path=params["path"],
            preset_name=params.get("preset"),
            profile_name=params.get("profile"),
            aggressive=bool(params.get("aggressive", False)),
            sample_size=int(params.get("sample_size", 8)),
        )
    if name == "compare_to_preset":
        return _call_tool(
            name,
            compare_to_preset,
            path=params["path"],
            preset_name=params.get("preset"),
            profile_name=params.get("profile"),
            aggressive=bool(params.get("aggressive", False)),
            sample_size=int(params.get("sample_size", 8)),
        )
    if name == "standardize_document":
        return _call_tool(
            name,
            standardize_document,
            path=params["path"],
            output_path=params.get("output_path"),
            preset_name=params.get("preset"),
            profile_name=params.get("profile"),
            overwrite=bool(params.get("overwrite", False)),
            aggressive=bool(params.get("aggressive", False)),
            fix_page_setup=bool(params.get("fix_page_setup", True)),
            fix_styles=bool(params.get("fix_styles", True)),
            fix_paragraphs=bool(params.get("fix_paragraphs", True)),
            fix_tables=bool(params.get("fix_tables", True)),
        )
    if name == "refresh_meganorm_cache":
        return _call_tool(
            name,
            refresh_catalog,
            category=params.get("category"),
            max_pages=int(params.get("max_pages", 5)),
        )
    if name == "search_meganorm_catalog":
        return _call_tool(
            name,
            search_catalog,
            query=params["query"],
            category=params.get("category"),
            max_pages=int(params.get("max_pages", 5)),
            limit=int(params.get("limit", 25)),
            refresh=bool(params.get("refresh", False)),
        )
    if name == "get_meganorm_topics":
        return _call_tool(
            name,
            get_current_topics,
            category=params.get("category"),
            page=int(params.get("page", 0)),
            limit=int(params.get("limit", 25)),
            refresh=bool(params.get("refresh", False)),
        )
    if name == "find_current_gost":
        return _call_tool(
            name,
            find_current_gost,
            query=params["query"],
            max_pages=int(params.get("max_pages", 10)),
            limit=int(params.get("limit", 25)),
            refresh=bool(params.get("refresh", False)),
        )

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
