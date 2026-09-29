from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys
from typing import Any

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from gost_standardizer.catalog import fetch_norm_markdown, find_current_gost

from gost_standardizer.converter import convert_html_to_markdown
from gost_standardizer.core import (
    explain_preset,
    inspect_document,
    list_presets,
    list_profiles,
    load_profile,
    save_profile,
    standardize_document,
    validate_document,
)
from gost_standardizer.server.mcp_server import main as run_mcp_server


def _print_json(data: Any) -> None:
    print(json.dumps(data, ensure_ascii=False, indent=2))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="gost-standardizer",
        description="GOST Document Standardizer & Meganorm Catalog Assistant",
    )
    subparsers = parser.add_subparsers(dest="command")

    # Command: check / search
    p_check = subparsers.add_parser("check", help="Check GOST status (active/replaced/cancelled)")
    p_check.add_argument("query", help="GOST number or title fragment (e.g. 7.0.97-2025)")

    p_search = subparsers.add_parser("search", help="Search standards in Meganorm catalog")
    p_search.add_argument("query", help="Query keyword or number")
    p_search.add_argument("--limit", type=int, default=25, help="Max results")

    # Command: inspect
    p_inspect = subparsers.add_parser("inspect", help="Inspect a document formatting and detect preset")
    p_inspect.add_argument("path", help="Path to DOCX/DOCM/DOC document")
    p_inspect.add_argument("--sample-size", type=int, default=8, help="Number of sample paragraphs")

    # Command: validate
    p_val = subparsers.add_parser("validate", help="Validate document against GOST preset")
    p_val.add_argument("path", help="Path to document")
    p_val.add_argument("--preset", choices=["report", "office", "technical", "legacy-college"], default=None)
    p_val.add_argument("--profile", default=None, help="Name of saved profile")
    p_val.add_argument("--aggressive", action="store_true", help="Aggressive validation")

    # Command: standardize
    p_std = subparsers.add_parser("standardize", help="Standardize document formatting to GOST")
    p_std.add_argument("path", help="Path to document")
    p_std.add_argument("-o", "--output", default=None, help="Output file path")
    p_std.add_argument("--preset", choices=["report", "office", "technical", "legacy-college"], default="report")
    p_std.add_argument("--profile", default=None, help="Name of saved profile")
    p_std.add_argument("-w", "--overwrite", action="store_true", help="Overwrite output file if exists")
    p_std.add_argument("--aggressive", action="store_true", help="Aggressive heading and styling fixes")

    # Command: convert-html
    p_conv = subparsers.add_parser("convert-html", help="Convert HTML to Markdown")
    p_conv.add_argument("file_or_text", help="HTML file path or raw string")

    # Command: presets / profiles
    subparsers.add_parser("presets", help="List built-in presets")
    subparsers.add_parser("profiles", help="List available profiles")

    # Command: mcp
    subparsers.add_parser("mcp", help="Run MCP JSON-RPC server on stdio")

    args = parser.parse_args(argv)

    if not args.command:
        parser.print_help()
        return 0

    if args.command == "check":
        result = find_current_gost(args.query)
        _print_json(result["status_summary"])
        if result["primary_document"]:
            doc = result["primary_document"]
            print(f"\nDocument: {doc['designation']}")
            print(f"Status: {doc['status']}")
            print(f"Active: {doc['is_active']}")
            if doc.get("date_intro"):
                print(f"Date intro: {doc['date_intro']}")
            if doc.get("replaces"):
                print(f"Replaces: {doc['replaces']}")
            if doc.get("replaced_by"):
                print(f"Replaced by: {doc['replaced_by']}")
            print(f"URL: {doc.get('url')}")
        return 0

    if args.command == "search":
        result = find_current_gost(args.query, limit=args.limit)
        _print_json(result)
        return 0

    if args.command == "inspect":
        _print_json(inspect_document(args.path, sample_size=args.sample_size))
        return 0

    if args.command == "validate":
        _print_json(
            validate_document(
                args.path,
                preset_name=args.preset,
                profile_name=args.profile,
                aggressive=args.aggressive,
            )
        )
        return 0

    if args.command == "standardize":
        res = standardize_document(
            args.path,
            output_path=args.output,
            preset_name=args.preset,
            profile_name=args.profile,
            overwrite=args.overwrite,
            aggressive=args.aggressive,
        )
        print(f"Successfully standardized to: {res['output_path']}")
        _print_json(res["changes"])
        return 0

    if args.command == "convert-html":
        content = args.file_or_text
        import os

        if os.path.exists(content):
            with open(content, "r", encoding="utf-8", errors="replace") as f:
                content = f.read()
        print(convert_html_to_markdown(content))
        return 0

    if args.command == "presets":
        _print_json(list_presets())
        return 0

    if args.command == "profiles":
        _print_json(list_profiles())
        return 0

    if args.command == "mcp":
        return run_mcp_server()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
