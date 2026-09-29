from __future__ import annotations

import io
import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from urllib.error import URLError

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Mm, Pt


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import gost_standardizer  # noqa: E402
import mcp_server  # noqa: E402
import meganorm_catalog  # noqa: E402


def _make_docx(path: Path) -> None:
    document = Document()
    section = document.sections[0]
    section.left_margin = Mm(25.4)
    section.right_margin = Mm(25.4)
    section.top_margin = Mm(25.4)
    section.bottom_margin = Mm(25.4)

    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    title_run = title.add_run("Пояснительная записка")
    title_run.font.name = "Arial"
    title_run.font.size = Pt(10)

    heading = document.add_paragraph("Введение")
    heading.paragraph_format.first_line_indent = Mm(0)
    heading.paragraph_format.space_before = Pt(0)
    heading.paragraph_format.space_after = Pt(0)

    body = document.add_paragraph("Это тестовый абзац без нужного оформления.")
    body.paragraph_format.first_line_indent = Mm(0)
    body.paragraph_format.space_before = Pt(0)
    body.paragraph_format.space_after = Pt(0)

    document.save(path)


class GostStandardizerTests(unittest.TestCase):
    def test_profile_roundtrip_uses_temp_profiles_dir(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir_name:
            temp_dir = Path(temp_dir_name)
            with patch.object(gost_standardizer, "PROFILES_DIR", temp_dir):
                saved = gost_standardizer.save_profile(
                    name="my-college",
                    preset_name="report",
                    title="My College",
                    description="Custom preset for tests",
                    kind="organization",
                    notes=["note one"],
                )
                self.assertTrue((temp_dir / "my-college.json").exists())
                loaded = gost_standardizer.load_profile("my-college")
                self.assertEqual(loaded["key"], "my-college")
                self.assertEqual(loaded["preset"]["key"], "report")
                self.assertEqual(saved["title"], "My College")

    def test_validate_compare_and_standardize_document(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir_name:
            temp_dir = Path(temp_dir_name)
            source = temp_dir / "input.docx"
            output = temp_dir / "output.docx"
            _make_docx(source)
            before = source.read_bytes()

            validation = gost_standardizer.validate_document(str(source), preset_name="report")
            self.assertEqual(validation["kind"], "validation")
            self.assertGreaterEqual(validation["summary"]["errors"], 1)
            self.assertTrue(validation["issues"])

            comparison = gost_standardizer.compare_to_preset(str(source), preset_name="report")
            self.assertEqual(comparison["kind"], "comparison")
            self.assertGreaterEqual(comparison["summary"]["differences"], 1)
            self.assertTrue(comparison["differences"])

            standardized = gost_standardizer.standardize_document(
                str(source),
                output_path=str(output),
                preset_name="report",
                overwrite=True,
            )
            self.assertEqual(standardized["kind"], "standardization")
            self.assertTrue(output.exists())
            self.assertEqual(before, source.read_bytes())

    def test_inspect_document_and_explain_preset(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir_name:
            temp_dir = Path(temp_dir_name)
            source = temp_dir / "report.docx"
            _make_docx(source)

            inspection = gost_standardizer.inspect_document(str(source))
            self.assertEqual(inspection["preset_guess"], "report")
            self.assertIn("issues", inspection)
            explanation = gost_standardizer.explain_preset(str(source))
            self.assertEqual(explanation["guessed_preset"], "report")
            self.assertIn("scores", explanation)
            self.assertIn("signals", explanation)

    def test_legacy_doc_failure_is_clear_without_converter(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir_name:
            temp_dir = Path(temp_dir_name)
            legacy = temp_dir / "legacy.doc"
            legacy.write_bytes(b"")
            with patch.object(gost_standardizer.shutil, "which", return_value=None):
                with self.assertRaises(ValueError) as ctx:
                    gost_standardizer.inspect_document(str(legacy))
            self.assertIn("LibreOffice", str(ctx.exception))

    def test_legacy_doc_standardize_persists_file(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir_name:
            temp_dir = Path(temp_dir_name)
            legacy = temp_dir / "legacy.doc"
            legacy.write_bytes(b"dummy")

            def fake_convert(src: Path, tmp: Path) -> Path:
                out = tmp / f"{src.stem}.docx"
                _make_docx(out)
                return out

            with patch.object(gost_standardizer, "_convert_legacy_doc", side_effect=fake_convert):
                res = gost_standardizer.standardize_document(str(legacy))
                out_path = Path(res["output_path"])
                self.assertTrue(out_path.exists())
                self.assertEqual(out_path.parent, legacy.parent)
                self.assertTrue(out_path.name.endswith(".docx"))

    def test_heading_and_list_classification(self) -> None:
        doc = Document()
        cases = [
            ("1 Введение", "heading"),
            ("1. Введение", "heading"),
            ("1.1 Подраздел", "heading"),
            ("1.1. Подраздел", "heading"),
            ("1.1.1 Пункт", "heading"),
            ("1.1.1. Пункт", "heading"),
            ("Глава 1. Общие сведения", "heading"),
            ("Раздел 2. Требования", "heading"),
            ("Приложение А. Спецификация", "heading"),
            ("Введение", "heading"),
            ("1) Пункт перечисления;", "list"),
            ("• Маркер списка", "list"),
            ("- Дефис списка", "list"),
            ("1. пункт со строчной", "list"),
            ("Обычный текст документа, содержащий развернутое описание алгоритма.", "body"),
        ]
        for text, expected_kind in cases:
            p = doc.add_paragraph(text)
            actual_kind = gost_standardizer._classify_paragraph(p, 0, 2)
            self.assertEqual(actual_kind, expected_kind, f"Failed for text: '{text}' (got {actual_kind}, expected {expected_kind})")

    def test_openxml_theme_font_cleanup(self) -> None:
        doc = Document()
        p = doc.add_paragraph()
        run = p.add_run("Тест")
        r_pr = run._element.get_or_add_rPr()
        r_fonts = r_pr.get_or_add_rFonts()
        r_fonts.set(qn("w:asciiTheme"), "minorHAnsi")
        r_fonts.set(qn("w:cstheme"), "minorBidi")

        gost_standardizer._set_r_fonts(run, "Times New Roman")
        self.assertEqual(r_fonts.get(qn("w:ascii")), "Times New Roman")
        for key in r_fonts.attrib:
            self.assertNotIn("theme", key.lower())

    def test_find_current_gost_uses_local_cache(self) -> None:
        cache = meganorm_catalog.load_cache()
        self.assertGreater(len(cache.get("categories", [])), 0)

        result = meganorm_catalog.find_current_gost("ГОСТ 7.32-2017", max_pages=2, limit=5)
        self.assertEqual(result["kind"], "current-gost-search")
        self.assertIn("updated_at", result["cache"])
        self.assertIsInstance(result["documents"], list)
        if result["documents"]:
            self.assertIn("confidence", result["documents"][0])
            self.assertIn("match", result["documents"][0])

    def test_search_catalog_returns_confidence(self) -> None:
        result = meganorm_catalog.search_catalog("ГОСТ", max_pages=1, limit=3)
        self.assertEqual(result["kind"], "search")
        self.assertIn("updated_at", result["cache"])
        if result["documents"]:
            self.assertIn("confidence", result["documents"][0])

    def test_meganorm_offline_graceful_fallback(self) -> None:
        with patch.object(meganorm_catalog, "_fetch_html", side_effect=URLError("Offline")):
            res = meganorm_catalog.find_current_gost("7.32", max_pages=1, limit=2)
            self.assertEqual(res["kind"], "current-gost-search")
            self.assertIsInstance(res["documents"], list)

            topics = meganorm_catalog.get_current_topics(limit=3)
            self.assertEqual(topics["kind"], "current-topics")
            self.assertTrue(topics["topics"])

    def test_mcp_server_protocol_dispatch(self) -> None:
        init_req = {"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {"protocolVersion": "2024-11-05"}}
        init_res = mcp_server._dispatch(init_req)
        self.assertEqual(init_res["id"], 1)
        self.assertEqual(init_res["result"]["serverInfo"]["name"], "gost-standardizer")

        list_req = {"jsonrpc": "2.0", "id": 2, "method": "tools/list", "params": {}}
        list_res = mcp_server._dispatch(list_req)
        tools = [t["name"] for t in list_res["result"]["tools"]]
        self.assertEqual(len(tools), 13)

        # test load_profile and save_profile via MCP without argument collisions
        with tempfile.TemporaryDirectory() as td:
            with patch.object(gost_standardizer, "PROFILES_DIR", Path(td)):
                save_req = {
                    "jsonrpc": "2.0",
                    "id": 3,
                    "method": "tools/call",
                    "params": {"name": "save_profile", "arguments": {"name": "test_mcp_org", "preset": "report"}},
                }
                save_res = mcp_server._dispatch(save_req)
                self.assertFalse(save_res["result"].get("isError", False))

                load_req = {
                    "jsonrpc": "2.0",
                    "id": 4,
                    "method": "tools/call",
                    "params": {"name": "load_profile", "arguments": {"name": "test_mcp_org"}},
                }
                load_res = mcp_server._dispatch(load_req)
                self.assertFalse(load_res["result"].get("isError", False))

        # test missing required parameter
        bad_req = {
            "jsonrpc": "2.0",
            "id": 5,
            "method": "tools/call",
            "params": {"name": "inspect_document", "arguments": {}},
        }
        bad_res = mcp_server._dispatch(bad_req)
        self.assertTrue(bad_res["result"].get("isError", False))

    def test_cli_interface(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            doc_path = Path(td) / "sample.docx"
            _make_docx(doc_path)

            old_stdout = sys.stdout
            try:
                sys.stdout = io.StringIO()
                ret = gost_standardizer.main(["inspect", str(doc_path)])
                self.assertEqual(ret, 0)
                out = sys.stdout.getvalue()
                self.assertIn("preset_guess", out)

                sys.stdout = io.StringIO()
                ret = gost_standardizer.main(["presets"])
                self.assertEqual(ret, 0)
                out = sys.stdout.getvalue()
                self.assertIn("report", out)
            finally:
                sys.stdout = old_stdout


if __name__ == "__main__":
    unittest.main()
