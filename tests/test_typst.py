from __future__ import annotations

import os
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from gost_standardizer.converter.typst import (
    _convert_markdown_table_to_typst,
    compile_typst,
    find_typst_binary,
    generate_gost_typst,
    markdown_to_gost_typst,
)
from gost_standardizer.core.presets import PRESETS
from gost_standardizer.cli.main import main as cli_main
from gost_standardizer.server.mcp_server import handle_tools_call


class TypstIntegrationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.work_dir = Path(self.temp_dir.name)

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def test_find_typst_binary(self) -> None:
        # Test finding binary on host
        binary = find_typst_binary()
        self.assertTrue(binary is None or Path(binary).is_file())

        # Test TYPST_PATH environment variable override
        fake_binary = self.work_dir / "fake_typst"
        fake_binary.write_text("#!/bin/sh\necho typst\n", encoding="utf-8")
        fake_binary.chmod(0o755)

        with patch.dict(os.environ, {"TYPST_PATH": str(fake_binary)}):
            self.assertEqual(find_typst_binary(), str(fake_binary))

        # Test fallback when not found anywhere
        with patch.dict(os.environ, {}, clear=True), \
             patch("shutil.which", return_value=None), \
             patch("gost_standardizer.converter.typst.KNOWN_TYPST_PATHS", [self.work_dir / "nonexistent"]):
            self.assertIsNone(find_typst_binary())

    def test_generate_gost_typst_presets(self) -> None:
        # Default report preset
        doc_default = generate_gost_typst()
        self.assertIn("GOST report", doc_default)
        self.assertIn("left: 30mm", doc_default)
        self.assertIn("right: 10mm", doc_default)
        self.assertIn("= ВВЕДЕНИЕ", doc_default)

        # Office preset (tighter margins, 1.0/1.15 spacing)
        doc_office = generate_gost_typst(
            preset_or_name="office",
            title="Приказ о проведении практики",
            author="Ректор",
            organization="МГТУ им. Баумана",
            year=2026,
        )
        self.assertIn("GOST office", doc_office)
        self.assertIn("left: 20mm", doc_office)
        self.assertIn("Приказ о проведении практики", doc_office)
        self.assertIn("Ректор", doc_office)
        self.assertIn("МГТУ им. Баумана", doc_office)
        self.assertIn("2026", doc_office)

        # Technical preset
        doc_tech = generate_gost_typst(PRESETS["technical"], body_content="= Техническое задание\nОписание ТЗ.")
        self.assertIn("GOST technical", doc_tech)
        self.assertIn("Описание ТЗ.", doc_tech)

    def test_convert_markdown_table_to_typst(self) -> None:
        lines = [
            "| Параметр | Значение | Примечание |",
            "| :--- | :---: | ---: |",
            "| **Шрифт** | _Times New Roman_ | 14 пт |",
            "| Отступ | 12.5 мм | По ГОСТу |",
        ]
        result = _convert_markdown_table_to_typst(lines)
        self.assertIn("#table(", result)
        self.assertIn("columns: 3", result)
        self.assertIn("[*Шрифт*]", result)
        self.assertIn("[_Times New Roman_]", result)

        # Empty lines
        self.assertEqual(_convert_markdown_table_to_typst([]), "")

    def test_markdown_to_gost_typst(self) -> None:
        md = """# Введение

Это вводный текст с **жирным** выделением и _курсивом_.
Также есть ссылка: [ГОСТ Р](https://protect.gost.ru).

## Раздел 1

Список требований:
- Требование 1
- Требование 2

Нумерованный список:
1. Шаг один
2. Шаг два

> Важная цитата или примечание

---

```python
def hello():
    print("Hello Typst")
```

| Заголовок 1 | Заголовок 2 |
| ----------- | ----------- |
| Данные 1    | Данные 2    |
"""
        doc = markdown_to_gost_typst(
            md,
            preset_or_name="technical",
            title="Технический отчет",
            author="Инженер Иванов",
        )
        self.assertIn("= Введение", doc)
        self.assertIn("== Раздел 1", doc)
        self.assertIn("*жирным*", doc)
        self.assertIn("#link(\"https://protect.gost.ru\")[ГОСТ Р]", doc)
        self.assertIn("- Требование 1", doc)
        self.assertIn("+ Шаг один", doc)
        self.assertIn("#quote(block: true)[Важная цитата или примечание]", doc)
        self.assertIn("#line(length: 100%, stroke: 0.5pt)", doc)
        self.assertIn("```python", doc)
        self.assertIn("#table(", doc)
        self.assertIn("Технический отчет", doc)
        self.assertIn("Инженер Иванов", doc)

    def test_compile_typst_real_execution(self) -> None:
        binary = find_typst_binary()
        if not binary:
            self.skipTest("Typst binary is not installed")

        sample_code = """
#set page(paper: "a4", margin: 2cm)
#set text(size: 14pt, lang: "ru")
= Тест
Проверка работы Typst из Python.
"""
        out_pdf = self.work_dir / "compiled.pdf"
        res = compile_typst(sample_code, output_path=str(out_pdf))
        self.assertTrue(res["success"])
        self.assertIsNotNone(res["output_path"])
        self.assertTrue(out_pdf.exists())
        self.assertGreater(out_pdf.stat().st_size, 1000)

        # Check compilation from .typ file path
        sample_file = self.work_dir / "doc.typ"
        sample_file.write_text(sample_code, encoding="utf-8")

        res_file = compile_typst(str(sample_file))
        self.assertTrue(res_file["success"])
        expected_pdf = sample_file.with_suffix(".pdf")
        self.assertTrue(expected_pdf.exists())

    def test_compile_typst_errors(self) -> None:
        binary = find_typst_binary()
        if binary:
            # Syntax error in Typst
            bad_code = "#set page(unknown_argument: 123)"
            res = compile_typst(bad_code, output_path=str(self.work_dir / "err.pdf"))
            self.assertFalse(res["success"])
            self.assertIsNotNone(res["error"])

        # Missing binary
        with patch("gost_standardizer.converter.typst.find_typst_binary", return_value=None):
            with self.assertRaises(RuntimeError):
                compile_typst("test")

        # Subprocess exception handling
        with patch("gost_standardizer.converter.typst.find_typst_binary", return_value="/bin/typst"), \
             patch("subprocess.run", side_effect=OSError("Command failed")):
            res = compile_typst("test", output_path=str(self.work_dir / "fail.pdf"), root_dir="/tmp")
            self.assertFalse(res["success"])
            self.assertIn("Command failed", str(res["error"]))

    def test_mcp_render_and_compile_tools(self) -> None:
        # Tool: render_gost_typst (template without markdown)
        res_tpl = handle_tools_call({
            "name": "render_gost_typst",
            "arguments": {
                "preset": "report",
                "title": "Отчет",
                "author": "Студент",
            },
        })
        self.assertFalse(res_tpl.get("isError", False))
        self.assertIn("Отчет", res_tpl["content"][0]["text"])

        # Tool: render_gost_typst with markdown and output_path
        target_typ = self.work_dir / "saved.typ"
        res_md = handle_tools_call({
            "name": "render_gost_typst",
            "arguments": {
                "preset": "office",
                "markdown": "# Заголовок\nТекст параграфа.",
                "output_path": str(target_typ),
            },
        })
        self.assertFalse(res_md.get("isError", False))
        self.assertTrue(target_typ.exists())
        self.assertIn("= Заголовок", target_typ.read_text(encoding="utf-8"))

        # Tool: compile_typst missing required input
        res_missing = handle_tools_call({
            "name": "compile_typst",
            "arguments": {},
        })
        self.assertTrue(res_missing.get("isError", False))
        self.assertIn("Missing required argument 'input'", res_missing["content"][0]["text"])

        # Tool: compile_typst success (mocked or real)
        with patch("gost_standardizer.server.mcp_server.compile_typst") as mock_comp:
            mock_comp.return_value = {"success": True, "output_path": "/path/to/doc.pdf"}
            res_comp = handle_tools_call({
                "name": "compile_typst",
                "arguments": {"input": "test code", "output_path": "/path/to/doc.pdf"},
            })
            self.assertFalse(res_comp.get("isError", False))
            self.assertIn("/path/to/doc.pdf", res_comp["content"][0]["text"])

    def test_cli_typst_commands(self) -> None:
        # CLI: typst-render template to stdout
        with patch("sys.stdout") as mock_stdout:
            exit_code = cli_main(["typst-render", "--preset", "report", "--title", "Док"])
            self.assertEqual(exit_code, 0)

        # CLI: typst-render markdown from file to output file
        md_file = self.work_dir / "input.md"
        md_file.write_text("# Тестовый заголовок\nПараграф.", encoding="utf-8")
        out_typ = self.work_dir / "output.typ"

        exit_code = cli_main(["typst-render", str(md_file), "-o", str(out_typ), "--preset", "technical"])
        self.assertEqual(exit_code, 0)
        self.assertTrue(out_typ.exists())
        self.assertIn("= Тестовый заголовок", out_typ.read_text(encoding="utf-8"))

        # CLI: typst-compile success and failure
        with patch("gost_standardizer.cli.main.compile_typst") as mock_comp:
            mock_comp.return_value = {"success": True, "output_path": "out.pdf"}
            code_ok = cli_main(["typst-compile", str(out_typ), "-o", "out.pdf"])
            self.assertEqual(code_ok, 0)

            mock_comp.return_value = {"success": False, "error": "Syntax error"}
            code_fail = cli_main(["typst-compile", str(out_typ)])
            self.assertEqual(code_fail, 1)


if __name__ == "__main__":
    unittest.main()
