from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any, Callable

from gost_standardizer.catalog import (
    fetch_norm_markdown,
    find_current_gost,
    get_current_topics,
    search_catalog,
)
from gost_standardizer.catalog.cache import cache_manager
from gost_standardizer.converter import (
    compile_typst,
    convert_html_to_markdown,
    generate_gost_typst,
    markdown_to_gost_typst,
)
from gost_standardizer.core import (
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
from gost_standardizer.server.transport import StdioTransport

SERVER_NAME = "gost-standardizer"
SERVER_VERSION = "0.3.0"
PROTOCOL_VERSION = "2024-11-05"

logger = logging.getLogger(__name__)


def _tool_schema(parameters: dict[str, Any], *, required: list[str] | None = None) -> dict[str, Any]:
    schema: dict[str, Any] = {
        "type": "object",
        "properties": parameters,
        "additionalProperties": False,
    }
    if required:
        schema["required"] = required
    return schema


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
                "title": {"type": "string", "description": "Optional profile title."},
                "description": {"type": "string", "description": "Optional profile description."},
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
                "path": {"type": "string", "description": "Path to the source document."},
                "sample_size": {"type": "integer", "minimum": 1, "maximum": 50, "default": 8},
            }
        ),
    },
    {
        "name": "validate_document",
        "description": "Validate a document against a profile/preset without making any changes.",
        "inputSchema": _tool_schema(
            {
                "path": {"type": "string", "description": "Path to the document to validate."},
                "preset": {
                    "type": "string",
                    "enum": ["report", "office", "technical", "legacy-college"],
                },
                "profile": {"type": "string", "description": "Name or path of a saved profile."},
                "aggressive": {
                    "type": "boolean",
                    "default": False,
                    "description": "Whether to enforce strict heading rules for body lines.",
                },
                "sample_size": {"type": "integer", "minimum": 1, "maximum": 50, "default": 8},
            }
        ),
    },
    {
        "name": "standardize_document",
        "description": "Standardize a DOCX/DOCM/DOC file to GOST formatting rules.",
        "inputSchema": _tool_schema(
            {
                "path": {"type": "string", "description": "Path to the source document."},
                "output_path": {"type": "string", "description": "Optional destination path for the standardized file."},
                "preset": {
                    "type": "string",
                    "enum": ["report", "office", "technical", "legacy-college"],
                    "default": "report",
                },
                "profile": {"type": "string", "description": "Name or path of a saved profile."},
                "overwrite": {"type": "boolean", "default": False},
                "aggressive": {"type": "boolean", "default": False},
                "fix_page_setup": {"type": "boolean", "default": True},
                "fix_styles": {"type": "boolean", "default": True},
                "fix_paragraphs": {"type": "boolean", "default": True},
                "fix_tables": {"type": "boolean", "default": True},
            }
        ),
    },
    {
        "name": "compare_to_preset",
        "description": "Compare a document against a preset or profile, highlighting matches and differences.",
        "inputSchema": _tool_schema(
            {
                "path": {"type": "string", "description": "Path to the document to compare."},
                "preset": {
                    "type": "string",
                    "enum": ["report", "office", "technical", "legacy-college"],
                    "default": "report",
                },
                "profile": {"type": "string", "description": "Name or path of a saved profile."},
                "aggressive": {"type": "boolean", "default": False},
                "sample_size": {"type": "integer", "minimum": 1, "maximum": 50, "default": 8},
            }
        ),
    },
    {
        "name": "explain_preset",
        "description": "Explain the rules and values associated with a built-in preset.",
        "inputSchema": _tool_schema(
            {
                "preset": {
                    "type": "string",
                    "enum": ["report", "office", "technical", "legacy-college"],
                    "default": "report",
                }
            }
        ),
    },
    {
        "name": "get_meganorm_topics",
        "description": "List Meganorm topics and categories.",
        "inputSchema": _tool_schema(
            {
                "category": {"type": "string", "description": "Optional category filter."},
                "limit": {"type": "integer", "minimum": 1, "maximum": 100, "default": 50},
            }
        ),
    },
    {
        "name": "refresh_meganorm_cache",
        "description": "Clear or update cached catalog data.",
        "inputSchema": _tool_schema({}),
    },
    {
        "name": "search_meganorm_catalog",
        "description": "Search the Meganorm catalog for standards by keyword or number.",
        "inputSchema": _tool_schema(
            {
                "query": {"type": "string", "description": "Search keyword or document number."},
                "category": {"type": "string", "description": "Optional category filter."},
                "limit": {"type": "integer", "minimum": 1, "maximum": 100, "default": 25},
            }
        ),
    },
    {
        "name": "find_current_gost",
        "description": "Find current GOST and GOST R documents, verifying active/replaced/cancelled status.",
        "inputSchema": _tool_schema(
            {
                "query": {
                    "type": "string",
                    "description": "GOST number or title fragment, for example '7.0.97-2025' or '7.32-2017'.",
                },
                "max_pages": {"type": "integer", "minimum": 1, "maximum": 20, "default": 10},
                "limit": {"type": "integer", "minimum": 1, "maximum": 100, "default": 25},
                "refresh": {"type": "boolean", "default": False},
            }
        ),
    },
    {
        "name": "convert_html_to_markdown",
        "description": "Convert HTML text to clean Markdown using high-performance xberg-io engine with fallback.",
        "inputSchema": _tool_schema(
            {
                "html": {"type": "string", "description": "Raw HTML string to convert."},
            }
        ),
    },
    {
        "name": "fetch_norm_markdown",
        "description": "Fetch a normative document card or text from Meganorm and return it in clean Markdown.",
        "inputSchema": _tool_schema(
            {
                "query_or_url": {
                    "type": "string",
                    "description": "GOST number (e.g. '7.0.97-2025') or Meganorm URL.",
                },
            }
        ),
    },
    {
        "name": "render_gost_typst",
        "description": "Generate a GOST-compliant Typst (.typ) document from Markdown or preset options.",
        "inputSchema": _tool_schema(
            {
                "preset": {
                    "type": "string",
                    "enum": ["report", "office", "technical", "legacy-college"],
                    "default": "report",
                },
                "markdown": {"type": "string", "description": "Optional Markdown text to convert to Typst body."},
                "title": {"type": "string", "description": "Document title."},
                "author": {"type": "string", "description": "Document author."},
                "organization": {"type": "string", "description": "University, agency, or organization."},
                "year": {"type": "integer", "description": "Year of document."},
                "output_path": {"type": "string", "description": "Optional destination path for the .typ file."},
            }
        ),
    },
    {
        "name": "compile_typst",
        "description": "Compile Typst markup code or a .typ file into a PDF document using the local Typst binary.",
        "inputSchema": _tool_schema(
            {
                "input": {"type": "string", "description": "Path to a .typ file or raw Typst code string."},
                "output_path": {"type": "string", "description": "Optional path for the compiled output PDF."},
            },
            required=["input"],
        ),
    },
]


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


def _tool_error(tool_name: str, exc: Exception) -> dict[str, Any]:
    return _result(
        _serialize(
            {
                "kind": "tool-error",
                "tool": tool_name,
                "error": type(exc).__name__,
                "message": str(exc),
            }
        ),
        is_error=True,
    )


def _call_tool(tool_name: str, fn: Callable[..., Any], /, **kwargs: Any) -> dict[str, Any]:
    try:
        return _result(_serialize(fn(**kwargs)))
    except (FileNotFoundError, ValueError, RuntimeError, KeyError, OSError) as exc:
        return _tool_error(tool_name, exc)
    except Exception as exc:  # noqa: BLE001
        logger.exception("Unexpected error in tool %s: %s", tool_name, exc)
        return _tool_error(tool_name, exc)


def handle_tools_call(params: dict[str, Any]) -> dict[str, Any]:
    name = params.get("name")
    arguments = params.get("arguments") or {}

    if name == "list_presets":
        return _call_tool(name, list_presets)
    if name == "list_profiles":
        return _call_tool(name, list_profiles)
    if name == "load_profile":
        if "name" not in arguments:
            return _result("Missing required argument 'name'", is_error=True)
        return _call_tool(name, load_profile, name=arguments["name"])
    if name == "save_profile":
        if "name" not in arguments:
            return _result("Missing required argument 'name'", is_error=True)
        return _call_tool(
            name,
            save_profile,
            name=arguments["name"],
            preset_name=arguments.get("preset"),
            title=arguments.get("title"),
            description=arguments.get("description"),
            kind=arguments.get("kind", "organization"),
            notes=arguments.get("notes"),
        )
    if name == "inspect_document":
        if "path" not in arguments:
            return _result("Missing required argument 'path'", is_error=True)
        return _call_tool(name, inspect_document, path=arguments["path"], sample_size=arguments.get("sample_size", 8))
    if name == "validate_document":
        if "path" not in arguments:
            return _result("Missing required argument 'path'", is_error=True)
        return _call_tool(
            name,
            validate_document,
            path=arguments["path"],
            preset_name=arguments.get("preset"),
            profile_name=arguments.get("profile"),
            aggressive=arguments.get("aggressive", False),
            sample_size=arguments.get("sample_size", 8),
        )
    if name == "standardize_document":
        if "path" not in arguments:
            return _result("Missing required argument 'path'", is_error=True)
        return _call_tool(
            name,
            standardize_document,
            path=arguments["path"],
            output_path=arguments.get("output_path"),
            preset_name=arguments.get("preset"),
            profile_name=arguments.get("profile"),
            overwrite=arguments.get("overwrite", False),
            aggressive=arguments.get("aggressive", False),
            fix_page_setup=arguments.get("fix_page_setup", True),
            fix_styles=arguments.get("fix_styles", True),
            fix_paragraphs=arguments.get("fix_paragraphs", True),
            fix_tables=arguments.get("fix_tables", True),
        )
    if name == "compare_to_preset":
        if "path" not in arguments:
            return _result("Missing required argument 'path'", is_error=True)
        return _call_tool(
            name,
            compare_to_preset,
            path=arguments["path"],
            preset_name=arguments.get("preset"),
            profile_name=arguments.get("profile"),
            aggressive=arguments.get("aggressive", False),
            sample_size=arguments.get("sample_size", 8),
        )
    if name == "explain_preset":
        preset_val = arguments.get("preset") or arguments.get("name") or arguments.get("path_or_preset")
        return _call_tool(name, explain_preset, path_or_preset=preset_val, preset_name=preset_val, name=preset_val)
    if name == "get_meganorm_topics":
        return _call_tool(
            name,
            get_current_topics,
            category=arguments.get("category"),
            limit=arguments.get("limit", 50),
        )
    if name == "refresh_meganorm_cache":
        cache_manager._data = None
        return _result(_serialize({"status": "cleared", "message": "Catalog cache reset"}))
    if name == "search_meganorm_catalog":
        if "query" not in arguments:
            return _result("Missing required argument 'query'", is_error=True)
        return _call_tool(
            name,
            search_catalog,
            query=arguments["query"],
            category=arguments.get("category"),
            limit=arguments.get("limit", 25),
        )
    if name == "find_current_gost":
        if "query" not in arguments:
            return _result("Missing required argument 'query'", is_error=True)
        return _call_tool(
            name,
            find_current_gost,
            query=arguments["query"],
            max_pages=arguments.get("max_pages", 10),
            limit=arguments.get("limit", 25),
            refresh=arguments.get("refresh", False),
        )
    if name == "convert_html_to_markdown":
        if "html" not in arguments:
            return _result("Missing required argument 'html'", is_error=True)
        return _result(convert_html_to_markdown(arguments["html"]))
    if name == "fetch_norm_markdown":
        if "query_or_url" not in arguments:
            return _result("Missing required argument 'query_or_url'", is_error=True)
        return _call_tool(name, fetch_norm_markdown, query_or_url=arguments["query_or_url"])
    if name == "render_gost_typst":
        markdown = arguments.get("markdown")
        if markdown:
            typst_code = markdown_to_gost_typst(
                markdown,
                preset_or_name=arguments.get("preset", "report"),
                title=arguments.get("title"),
                author=arguments.get("author"),
                organization=arguments.get("organization"),
                year=arguments.get("year"),
            )
        else:
            typst_code = generate_gost_typst(
                preset_or_name=arguments.get("preset", "report"),
                title=arguments.get("title"),
                author=arguments.get("author"),
                organization=arguments.get("organization"),
                year=arguments.get("year"),
            )
        out_path = arguments.get("output_path")
        if out_path:
            p = Path(out_path).resolve()
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_text(typst_code, encoding="utf-8")
            return _result(_serialize({"output_path": str(p), "status": "saved", "content": typst_code}))
        return _result(typst_code)
    if name == "compile_typst":
        if "input" not in arguments:
            return _result("Missing required argument 'input'", is_error=True)
        res = compile_typst(
            input_path_or_content=arguments["input"],
            output_path=arguments.get("output_path"),
        )
        return _result(_serialize(res), is_error=not res.get("success", False))

    raise KeyError(f"Unknown tool: {name}")


def dispatch(request: dict[str, Any]) -> dict[str, Any] | None:
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
            result = handle_tools_call(params)
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
    transport = StdioTransport()
    while True:
        request = transport.read_message()
        if request is None:
            return 0
        response = dispatch(request)
        if response is not None and "id" in response:
            transport.send_message(response)


if __name__ == "__main__":
    raise SystemExit(main())
