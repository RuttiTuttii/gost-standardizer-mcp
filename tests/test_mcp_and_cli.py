from __future__ import annotations

import io
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import MagicMock, patch
from urllib.error import URLError

import gost_standardizer
from gost_standardizer.cli.main import main as cli_main
from gost_standardizer.server import mcp_server
from gost_standardizer.server.mcp_server import (
    _call_tool,
    _serialize,
    _tool_error,
    _tool_schema,
    dispatch,
    handle_tools_call,
)
import meganorm_catalog
import scripts.gost_standardizer as scripts_gost
import scripts.mcp_server as scripts_mcp


def _dummy_docx(path: Path) -> None:
    from docx import Document
    doc = Document()
    doc.add_paragraph("Тестовый документ")
    doc.save(path)


class McpAndCliTests(unittest.TestCase):
    def test_mcp_tool_schema_and_helpers(self) -> None:
        schema = _tool_schema({"key": {"type": "string"}})
        self.assertEqual(schema["type"], "object")
        self.assertFalse(schema["additionalProperties"])
        self.assertIn("key", schema["properties"])

        err_res = _tool_error("test_tool", ValueError("Invalid value"))
        self.assertTrue(err_res.get("isError"))
        err_content = json.loads(err_res["content"][0]["text"])
        self.assertEqual(err_content["kind"], "tool-error")
        self.assertEqual(err_content["tool"], "test_tool")
        self.assertEqual(err_content["error"], "ValueError")
        self.assertEqual(err_content["message"], "Invalid value")

    def test_mcp_call_tool_exceptions(self) -> None:
        # Standard handled exceptions
        for exc in (FileNotFoundError("Missing file"), ValueError("Bad val"), RuntimeError("Err"), KeyError("k"), OSError("os")):
            def bad_fn():
                raise exc
            res = _call_tool("tool", bad_fn)
            self.assertTrue(res.get("isError"))

        # Unexpected exception handled
        def unexpected_fn():
            raise ZeroDivisionError("division by zero")

        with patch.object(mcp_server.logger, "exception") as mock_log:
            res_unexpected = _call_tool("tool", unexpected_fn)
            self.assertTrue(res_unexpected.get("isError"))
            mock_log.assert_called()

    def test_mcp_handle_tools_call_all_tools(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            td_path = Path(td)
            doc_path = td_path / "sample.docx"
            _dummy_docx(doc_path)

            # 1. list_presets
            res = handle_tools_call({"name": "list_presets"})
            self.assertFalse(res.get("isError", False))

            # 2. list_profiles
            res = handle_tools_call({"name": "list_profiles"})
            self.assertFalse(res.get("isError", False))

            # 3. load_profile
            res_bad = handle_tools_call({"name": "load_profile", "arguments": {}})
            self.assertTrue(res_bad.get("isError", False))
            res = handle_tools_call({"name": "load_profile", "arguments": {"name": "report"}})
            self.assertFalse(res.get("isError", False))

            # 4. save_profile
            res_bad = handle_tools_call({"name": "save_profile", "arguments": {}})
            self.assertTrue(res_bad.get("isError", False))
            with patch.object(gost_standardizer, "PROFILES_DIR", td_path):
                res = handle_tools_call(
                    {
                        "name": "save_profile",
                        "arguments": {
                            "name": "my_org",
                            "preset": "report",
                            "title": "Org Profile",
                            "description": "Desc",
                            "kind": "organization",
                            "notes": ["Note 1"],
                        },
                    }
                )
                self.assertFalse(res.get("isError", False))

            # 5. inspect_document
            res_bad = handle_tools_call({"name": "inspect_document", "arguments": {}})
            self.assertTrue(res_bad.get("isError", False))
            res = handle_tools_call({"name": "inspect_document", "arguments": {"path": str(doc_path), "sample_size": 4}})
            self.assertFalse(res.get("isError", False))

            # 6. validate_document
            res_bad = handle_tools_call({"name": "validate_document", "arguments": {}})
            self.assertTrue(res_bad.get("isError", False))
            res = handle_tools_call(
                {
                    "name": "validate_document",
                    "arguments": {
                        "path": str(doc_path),
                        "preset": "report",
                        "profile": None,
                        "aggressive": False,
                        "sample_size": 4,
                    },
                }
            )
            self.assertFalse(res.get("isError", False))

            # 7. standardize_document
            res_bad = handle_tools_call({"name": "standardize_document", "arguments": {}})
            self.assertTrue(res_bad.get("isError", False))
            out_file = td_path / "out.docx"
            res = handle_tools_call(
                {
                    "name": "standardize_document",
                    "arguments": {
                        "path": str(doc_path),
                        "output_path": str(out_file),
                        "preset": "report",
                        "overwrite": True,
                        "aggressive": True,
                        "fix_page_setup": True,
                        "fix_styles": True,
                        "fix_paragraphs": True,
                        "fix_tables": True,
                    },
                }
            )
            self.assertFalse(res.get("isError", False))

            # 8. compare_to_preset
            res_bad = handle_tools_call({"name": "compare_to_preset", "arguments": {}})
            self.assertTrue(res_bad.get("isError", False))
            res = handle_tools_call(
                {
                    "name": "compare_to_preset",
                    "arguments": {"path": str(doc_path), "preset": "report", "aggressive": False, "sample_size": 4},
                }
            )
            self.assertFalse(res.get("isError", False))

            # 9. explain_preset with different argument variations
            res_p = handle_tools_call({"name": "explain_preset", "arguments": {"preset": "report"}})
            self.assertFalse(res_p.get("isError", False))
            res_n = handle_tools_call({"name": "explain_preset", "arguments": {"name": "office"}})
            self.assertFalse(res_n.get("isError", False))
            res_pop = handle_tools_call({"name": "explain_preset", "arguments": {"path_or_preset": "technical"}})
            self.assertFalse(res_pop.get("isError", False))

            # 10. get_meganorm_topics
            res = handle_tools_call({"name": "get_meganorm_topics", "arguments": {"limit": 5}})
            self.assertFalse(res.get("isError", False))

            # 11. refresh_meganorm_cache
            res = handle_tools_call({"name": "refresh_meganorm_cache", "arguments": {}})
            self.assertFalse(res.get("isError", False))

            # 12. search_meganorm_catalog
            res_bad = handle_tools_call({"name": "search_meganorm_catalog", "arguments": {}})
            self.assertTrue(res_bad.get("isError", False))
            res = handle_tools_call({"name": "search_meganorm_catalog", "arguments": {"query": "ГОСТ", "limit": 2}})
            self.assertFalse(res.get("isError", False))

            # 13. find_current_gost
            res_bad = handle_tools_call({"name": "find_current_gost", "arguments": {}})
            self.assertTrue(res_bad.get("isError", False))
            res = handle_tools_call({"name": "find_current_gost", "arguments": {"query": "7.32-2017"}})
            self.assertFalse(res.get("isError", False))

            # 14. convert_html_to_markdown
            res_bad = handle_tools_call({"name": "convert_html_to_markdown", "arguments": {}})
            self.assertTrue(res_bad.get("isError", False))
            res = handle_tools_call({"name": "convert_html_to_markdown", "arguments": {"html": "<p>test</p>"}})
            self.assertFalse(res.get("isError", False))

            # 15. fetch_norm_markdown
            res_bad = handle_tools_call({"name": "fetch_norm_markdown", "arguments": {}})
            self.assertTrue(res_bad.get("isError", False))
            with patch("gost_standardizer.server.mcp_server.fetch_norm_markdown", return_value={"kind": "fetch"}):
                res = handle_tools_call({"name": "fetch_norm_markdown", "arguments": {"query_or_url": "7.32-2017"}})
                self.assertFalse(res.get("isError", False))

            # Unknown tool
            with self.assertRaises(KeyError):
                handle_tools_call({"name": "unknown_tool", "arguments": {}})

    def test_mcp_dispatch_notifications_and_errors(self) -> None:
        # notifications/initialized returns None
        self.assertIsNone(dispatch({"jsonrpc": "2.0", "method": "notifications/initialized"}))

        # tools/call with unknown tool returns code -32603
        bad_tool_call = {
            "jsonrpc": "2.0",
            "id": 99,
            "method": "tools/call",
            "params": {"name": "completely_unknown_tool"},
        }
        res_err = dispatch(bad_tool_call)
        self.assertIsNotNone(res_err)
        assert res_err is not None
        self.assertEqual(res_err["error"]["code"], -32603)

        # unknown method with id returns -32601
        res_method_not_found = dispatch({"jsonrpc": "2.0", "id": 100, "method": "unknown/method"})
        self.assertIsNotNone(res_method_not_found)
        assert res_method_not_found is not None
        self.assertEqual(res_method_not_found["error"]["code"], -32601)

        # unknown method without id returns None
        self.assertIsNone(dispatch({"jsonrpc": "2.0", "method": "unknown/notification"}))

    def test_mcp_server_main_loop(self) -> None:
        # Mock StdioTransport with a sequence of messages
        messages = [
            {"jsonrpc": "2.0", "id": 1, "method": "initialize"},
            {"jsonrpc": "2.0", "method": "notifications/initialized"},
            None,  # EOF to exit loop
        ]
        mock_transport = MagicMock()
        mock_transport.read_message.side_effect = messages

        with patch("gost_standardizer.server.mcp_server.StdioTransport", return_value=mock_transport):
            exit_code = mcp_server.main()
            self.assertEqual(exit_code, 0)
            mock_transport.send_message.assert_called_once()
            sent = mock_transport.send_message.call_args[0][0]
            self.assertEqual(sent["id"], 1)

    def test_cli_all_commands(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            td_path = Path(td)
            doc_path = td_path / "sample.docx"
            _dummy_docx(doc_path)
            html_path = td_path / "sample.html"
            html_path.write_text("<h1>Тест</h1>", encoding="utf-8")

            # 1. No arguments -> print_help, return 0
            with patch("sys.stdout", new=io.StringIO()):
                ret = cli_main([])
                self.assertEqual(ret, 0)

            # 2. check command with primary document (active)
            with patch("sys.stdout", new=io.StringIO()) as out:
                ret = cli_main(["check", "ГОСТ Р 7.0.97-2025"])
                self.assertEqual(ret, 0)
                self.assertIn("Document:", out.getvalue())
                self.assertIn("7.0.97-2025", out.getvalue())

            # 2b. check command with replaced document (has replaced_by)
            with patch("sys.stdout", new=io.StringIO()) as out:
                ret = cli_main(["check", "7.0.97-2016"])
                self.assertEqual(ret, 0)
                self.assertIn("Replaced by:", out.getvalue())

            # 3. check command when document not found
            with patch("gost_standardizer.cli.main.find_current_gost", return_value={"status_summary": {}, "primary_document": None}):
                with patch("sys.stdout", new=io.StringIO()):
                    ret = cli_main(["check", "UNKNOWN_999"])
                    self.assertEqual(ret, 0)

            # 4. search command
            with patch("sys.stdout", new=io.StringIO()) as out:
                ret = cli_main(["search", "7.32", "--limit", "3"])
                self.assertEqual(ret, 0)
                self.assertIn("current-gost-search", out.getvalue())

            # 5. inspect command
            with patch("sys.stdout", new=io.StringIO()) as out:
                ret = cli_main(["inspect", str(doc_path), "--sample-size", "2"])
                self.assertEqual(ret, 0)
                self.assertIn("preset_guess", out.getvalue())

            # 6. validate command
            with patch("sys.stdout", new=io.StringIO()) as out:
                ret = cli_main(["validate", str(doc_path), "--preset", "report", "--aggressive"])
                self.assertEqual(ret, 0)
                self.assertIn("validation", out.getvalue())

            # 7. standardize command
            out_docx = td_path / "standardized.docx"
            with patch("sys.stdout", new=io.StringIO()) as out:
                ret = cli_main(["standardize", str(doc_path), "-o", str(out_docx), "--preset", "report", "-w", "--aggressive"])
                self.assertEqual(ret, 0)
                self.assertIn("Successfully standardized", out.getvalue())
                self.assertTrue(out_docx.exists())

            # 8. convert-html command from file
            with patch("sys.stdout", new=io.StringIO()) as out:
                ret = cli_main(["convert-html", str(html_path)])
                self.assertEqual(ret, 0)
                self.assertIn("# Тест", out.getvalue())

            # 9. convert-html command from raw text
            with patch("sys.stdout", new=io.StringIO()) as out:
                ret = cli_main(["convert-html", "<h2>Прямой HTML</h2>"])
                self.assertEqual(ret, 0)
                self.assertIn("## Прямой HTML", out.getvalue())

            # 10. presets command
            with patch("sys.stdout", new=io.StringIO()) as out:
                ret = cli_main(["presets"])
                self.assertEqual(ret, 0)
                self.assertIn("report", out.getvalue())

            # 11. profiles command
            with patch("sys.stdout", new=io.StringIO()) as out:
                ret = cli_main(["profiles"])
                self.assertEqual(ret, 0)
                self.assertIn("report", out.getvalue())

            # 12. mcp command
            with patch("gost_standardizer.cli.main.run_mcp_server", return_value=0):
                ret = cli_main(["mcp"])
                self.assertEqual(ret, 0)

    def test_scripts_meganorm_catalog_functions(self) -> None:
        # load_cache
        data = meganorm_catalog.load_cache()
        self.assertIn("categories", data)

        # save_cache
        with patch.object(meganorm_catalog.cache_manager, "save") as mock_save:
            meganorm_catalog.save_cache({"test": 123})
            self.assertEqual(meganorm_catalog.cache_manager._data, {"test": 123})
            mock_save.assert_called_once()

        # refresh_catalog
        res = meganorm_catalog.refresh_catalog()
        self.assertEqual(res, {"status": "cleared"})
        self.assertIsNone(meganorm_catalog.cache_manager._data)

        # _fetch_html success and failure
        with patch.object(meganorm_catalog._srv_client, "fetch_url", return_value=("<html>ok</html>", False)):
            content = meganorm_catalog._fetch_html("https://meganorm.ru/test")
            self.assertEqual(content, "<html>ok</html>")

        with patch.object(meganorm_catalog._srv_client, "fetch_url", return_value=("", True)):
            with self.assertRaises(URLError):
                meganorm_catalog._fetch_html("https://meganorm.ru/fail")

    def test_scripts_module_declarations(self) -> None:
        self.assertEqual(scripts_mcp.SERVER_NAME, "gost-standardizer")
        self.assertIsNotNone(scripts_gost.PRESETS)


if __name__ == "__main__":
    unittest.main()
