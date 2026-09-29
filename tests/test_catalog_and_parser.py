from __future__ import annotations

import io
import json
import logging
from pathlib import Path
import tempfile
import unittest
from unittest.mock import MagicMock, patch
from urllib.error import URLError

from gost_standardizer.catalog.cache import CatalogCacheManager
from gost_standardizer.catalog.client import MeganormClient
from gost_standardizer.catalog.parser import parse_index_card, parse_list_links
from gost_standardizer.catalog.service import (
    KNOWN_INDEX_PAGES,
    _score_document_match,
    fetch_norm_markdown,
    find_current_gost,
    get_current_topics,
    search_catalog,
)
from gost_standardizer.models.document import DocumentStatus, NormDocument


class CatalogAndParserTests(unittest.TestCase):
    def test_parse_index_card_empty_or_none(self) -> None:
        self.assertIsNone(parse_index_card("", "https://meganorm.ru/Index/1/123.htm"))
        self.assertIsNone(parse_index_card(None, "https://meganorm.ru/Index/1/123.htm"))  # type: ignore

    def test_parse_index_card_full_metadata(self) -> None:
        html = """
        <html>
        <head><title>Скачать ГОСТ Р 7.0.97-2025 Документация</title></head>
        <body>
          <table>
            <tr><td>Обозначение:</td><td>ГОСТ Р 7.0.97-2025</td></tr>
            <tr><td>Статус:</td><td>Действующий</td></tr>
            <tr><td>Дата введения:</td><td>01.07.2025</td></tr>
            <tr><td>Дата издания:</td><td>15.06.2025</td></tr>
            <tr><td>Дата окончания срока действия:</td><td>01.01.2035</td></tr>
            <tr><td>Взамен:</td><td>ГОСТ Р 7.0.97-2016</td></tr>
            <tr><td>Заменяющий:</td><td>ГОСТ Р 7.0.97-2035</td></tr>
            <tr><td>Нормативные ссылки:</td><td>ГОСТ 1.1; ГОСТ 1.2, ГОСТ 1.3</td></tr>
          </table>
        </body>
        </html>
        """
        doc = parse_index_card(html, "https://meganorm.ru/Index/85/85280.htm")
        self.assertIsNotNone(doc)
        assert doc is not None
        self.assertEqual(doc.id, "85280")
        self.assertEqual(doc.designation, "ГОСТ Р 7.0.97-2025")
        self.assertEqual(doc.clean_number, "7.0.97-2025")
        self.assertEqual(doc.title, "ГОСТ Р 7.0.97-2025 Документация")
        self.assertEqual(doc.status, "действующий")
        self.assertTrue(doc.is_active)
        self.assertEqual(doc.date_intro, "01.07.2025")
        self.assertEqual(doc.date_published, "15.06.2025")
        self.assertEqual(doc.date_expired, "01.01.2035")
        self.assertEqual(doc.replaces, "ГОСТ Р 7.0.97-2016")
        self.assertEqual(doc.replaced_by, "ГОСТ Р 7.0.97-2035")
        self.assertEqual(doc.normative_refs, ["ГОСТ 1.1", "ГОСТ 1.2", "ГОСТ 1.3"])
        self.assertEqual(doc.category, "ГОСТ Р")

    def test_parse_index_card_fallback_designation_from_title(self) -> None:
        html = """
        <html>
        <head><title>ГОСТ 7.32-2017 Отчет о НИР</title></head>
        <body>
          <table>
            <tr><td>Статус:</td><td>Отменен</td></tr>
          </table>
        </body>
        </html>
        """
        doc = parse_index_card(html, "https://meganorm.ru/other/page.html")
        self.assertIsNotNone(doc)
        assert doc is not None
        self.assertEqual(doc.id, "")
        self.assertEqual(doc.designation, "ГОСТ 7.32-2017")
        self.assertEqual(doc.clean_number, "7.32-2017")
        self.assertEqual(doc.category, "ГОСТ")
        self.assertEqual(doc.status, "отменен")
        self.assertFalse(doc.is_active)

    def test_parse_index_card_fallback_to_plain_title(self) -> None:
        html = """
        <html>
        <head><title>Правила оформления документации</title></head>
        <body><p>No table info</p></body>
        </html>
        """
        doc = parse_index_card(html, "https://meganorm.ru/doc.htm")
        self.assertIsNotNone(doc)
        assert doc is not None
        self.assertEqual(doc.designation, "Правила оформления документации")
        self.assertEqual(doc.clean_number, "Правила оформления документации")

    def test_parse_list_links(self) -> None:
        html = """
        <div>
          <a href="/Index/10/10001.htm"><b>ГОСТ 10001-2020</b> Документ 1</a>
          <a href="/Index/10/10002.htm"><span>ГОСТ 10002-2021</span></a>
          <a href="/list/catalog.htm">Каталог</a>
          <a href="https://example.com/other.htm">Внешняя ссылка</a>
        </div>
        """
        links = parse_list_links(html, "https://meganorm.ru/list/1.htm")
        self.assertEqual(len(links), 2)
        self.assertEqual(links[0]["href"], "https://meganorm.ru/Index/10/10001.htm")
        self.assertEqual(links[0]["title"], "ГОСТ 10001-2020 Документ 1")
        self.assertEqual(links[1]["href"], "https://meganorm.ru/Index/10/10002.htm")
        self.assertEqual(links[1]["title"], "ГОСТ 10002-2021")

        self.assertEqual(parse_list_links("", "https://meganorm.ru"), [])

    def test_meganorm_client_relative_url_and_fetch_hook(self) -> None:
        client = MeganormClient(base_url="https://meganorm.ru")

        # Test hook returning tuple
        client.fetch_hook = lambda url: ("<html>Hook content</html>", False)
        content, err = client.fetch_url("relative/path.htm")
        self.assertFalse(err)
        self.assertEqual(content, "<html>Hook content</html>")

        # Test hook returning string
        client.fetch_hook = lambda url: "raw string"
        content, err = client.fetch_url("relative/path.htm")
        self.assertFalse(err)
        self.assertEqual(content, "raw string")

        # Test hook raising exception
        def failing_hook(url: str):
            raise RuntimeError("Hook crashed")

        client.fetch_hook = failing_hook
        content, err = client.fetch_url("relative/path.htm")
        self.assertTrue(err)
        self.assertEqual(content, "")

    def test_meganorm_client_urlopen_encodings_and_errors(self) -> None:
        client = MeganormClient(base_url="https://meganorm.ru")
        client.fetch_hook = None

        # 1. Success with header charset
        mock_resp = MagicMock()
        mock_resp.read.return_value = "тестовый текст".encode("utf-8")
        mock_resp.headers.get_content_charset.return_value = "utf-8"
        mock_resp.__enter__.return_value = mock_resp

        with patch("urllib.request.urlopen", return_value=mock_resp):
            text, err = client.fetch_url("/Index/test.htm")
            self.assertFalse(err)
            self.assertEqual(text, "тестовый текст")

        # 2. Success without header charset (defaults to utf-8)
        mock_resp.headers.get_content_charset.return_value = None
        with patch("urllib.request.urlopen", return_value=mock_resp):
            text, err = client.fetch_url("https://meganorm.ru/Index/test.htm")
            self.assertFalse(err)
            self.assertEqual(text, "тестовый текст")

        # 3. Fallback to windows-1251
        cp1251_bytes = "документ".encode("windows-1251")
        mock_resp.read.return_value = cp1251_bytes
        mock_resp.headers.get_content_charset.return_value = "utf-8"  # wrong charset triggers UnicodeDecodeError
        with patch("urllib.request.urlopen", return_value=mock_resp):
            text, err = client.fetch_url("/Index/test.htm")
            self.assertFalse(err)
            self.assertEqual(text, "документ")

        # 4. Fallback to utf-8 with replace
        mock_raw = MagicMock()
        mock_raw.decode.side_effect = [
            UnicodeDecodeError("utf8", b"", 0, 1, "err"),
            UnicodeDecodeError("cp1251", b"", 0, 1, "err"),
            "replaced",
        ]
        mock_resp.read.return_value = mock_raw
        mock_resp.headers.get_content_charset.return_value = "utf-8"
        with patch("urllib.request.urlopen", return_value=mock_resp):
            text, err = client.fetch_url("/Index/test.htm")
            self.assertFalse(err)
            self.assertEqual(text, "replaced")

        # 5. Exception handling: URLError, TimeoutError, OSError
        for exc in (URLError("Connection refused"), TimeoutError("Timed out"), OSError("Disk error")):
            with patch("urllib.request.urlopen", side_effect=exc):
                text, err = client.fetch_url("https://meganorm.ru/Index/test.htm")
                self.assertTrue(err)
                self.assertEqual(text, "")

    def test_catalog_cache_manager(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            temp_path = Path(td)
            seed_file = temp_path / "seed.json"
            user_cache_file = temp_path / "cache.json"

            # 1. Missing seed and missing user cache
            cm = CatalogCacheManager(seed_path=seed_file, cache_path=user_cache_file)
            data = cm.get_data()
            self.assertEqual(data["version"], "1.0")
            self.assertEqual(data["categories"], [])
            # Calling get_data again uses cached _data
            self.assertIs(cm.get_data(), data)

            # 2. Corrupt seed file
            seed_file.write_text("NOT A VALID JSON", encoding="utf-8")
            cm2 = CatalogCacheManager(seed_path=seed_file, cache_path=user_cache_file)
            data2 = cm2.get_data()
            self.assertEqual(data2["categories"], [])

            # 3. Valid seed and corrupt user cache
            seed_data = {
                "categories": [{"slug": "c1", "title": "Category 1"}],
                "category_pages": {"c1": {"pages": {}}},
                "fetched_at": "2025-01-01T00:00:00",
                "indexed_documents": {
                    "7.32-2017": {
                        "id": "1",
                        "designation": "ГОСТ 7.32-2017",
                        "clean_number": "7.32-2017",
                        "title": "Отчет",
                        "status": "действующий",
                    }
                },
            }
            seed_file.write_text(json.dumps(seed_data), encoding="utf-8")
            user_cache_file.write_text("INVALID USER CACHE", encoding="utf-8")
            cm3 = CatalogCacheManager(seed_path=seed_file, cache_path=user_cache_file)
            data3 = cm3.get_data()
            self.assertEqual(len(data3["categories"]), 1)
            self.assertEqual(data3["fetched_at"], "2025-01-01T00:00:00")

            # 4. Valid user cache overlays seed
            user_data = {
                "categories": [{"slug": "c2", "title": "Category 2"}],
                "category_pages": {"c2": {"pages": {}}},
                "fetched_at": "2025-02-01T00:00:00",
                "indexed_documents": {
                    "2.105-2019": {
                        "id": "2",
                        "designation": "ГОСТ 2.105-2019",
                        "clean_number": "2.105-2019",
                        "title": "ЕСКД",
                        "status": "действующий",
                    }
                },
            }
            user_cache_file.write_text(json.dumps(user_data), encoding="utf-8")
            cm4 = CatalogCacheManager(seed_path=seed_file, cache_path=user_cache_file)
            data4 = cm4.get_data()
            self.assertEqual(data4["categories"][0]["slug"], "c2")
            self.assertIn("c1", data4["category_pages"])
            self.assertIn("c2", data4["category_pages"])
            self.assertIn("7.32-2017", data4["indexed_documents"])
            self.assertIn("2.105-2019", data4["indexed_documents"])
            self.assertEqual(data4["fetched_at"], "2025-02-01T00:00:00")

            # 5. find_indexed_document searches
            # Direct match
            doc1 = cm4.find_indexed_document("2.105-2019")
            self.assertIsNotNone(doc1)
            assert doc1 is not None
            self.assertEqual(doc1.designation, "ГОСТ 2.105-2019")

            # Key case insensitive match (line 93)
            cm4.get_data()["indexed_documents"]["UPPERCASE-KEY"] = {
                "id": "10",
                "designation": "ГОСТ 10",
                "clean_number": "10-2020",
                "title": "Case test",
            }
            doc_case = cm4.find_indexed_document("uppercase-key")
            self.assertIsNotNone(doc_case)

            # clean_number match when key is different (line 96)
            cm4.get_data()["indexed_documents"]["different_key"] = {
                "id": "11",
                "designation": "ГОСТ 11",
                "clean_number": "special-clean-number",
                "title": "Clean test",
            }
            doc_clean = cm4.find_indexed_document("special-clean-number")
            self.assertIsNotNone(doc_clean)

            # Lower key match or clean_number match
            doc2 = cm4.find_indexed_document("ГОСТ 7.32-2017")
            self.assertIsNotNone(doc2)
            assert doc2 is not None
            self.assertEqual(doc2.clean_number, "7.32-2017")

            # Substring match in designation
            doc3 = cm4.find_indexed_document("2.105")
            self.assertIsNotNone(doc3)
            assert doc3 is not None
            self.assertEqual(doc3.clean_number, "2.105-2019")

            # Not found
            self.assertIsNone(cm4.find_indexed_document("99.999-9999"))

            # 6. store_document and save
            new_doc = NormDocument(
                id="3",
                designation="ГОСТ 1.2-2020",
                clean_number="1.2-2020",
                title="Стандартизация",
                status="действующий",
            )
            cm4.store_document(new_doc)
            self.assertIn("1.2-2020", cm4.get_data()["indexed_documents"])
            self.assertTrue(user_cache_file.exists())

            # 7. save edge cases
            cm_unsaved = CatalogCacheManager(seed_path=seed_file, cache_path=user_cache_file)
            cm_unsaved._data = None
            cm_unsaved.save()  # Does nothing when _data is None

            cm4._data = {"invalid": "data"}
            with patch.object(Path, "write_text", side_effect=OSError("Permission denied")):
                cm4.save()  # Catches exception and logs warning without re-raising

    def test_service_score_document_match(self) -> None:
        # Exact match
        s1 = _score_document_match("7.32-2017", "ГОСТ 7.32-2017")
        self.assertEqual(s1["confidence"], 1.0)
        self.assertTrue(s1["exact"])
        self.assertFalse(s1["partial"])

        # Partial match
        s2 = _score_document_match("7.32", "ГОСТ 7.32-2017")
        self.assertEqual(s2["confidence"], 0.85)
        self.assertFalse(s2["exact"])
        self.assertTrue(s2["partial"])

        # No match
        s3 = _score_document_match("9.999", "ГОСТ 7.32-2017")
        self.assertEqual(s3["confidence"], 0.0)
        self.assertFalse(s3["exact"])
        self.assertFalse(s3["partial"])

    def test_service_find_current_gost_fetch_known_page(self) -> None:
        from gost_standardizer.catalog import service

        html = """
        <html>
        <head><title>ГОСТ 7.32-2017 Отчет о научно-исследовательской работе</title></head>
        <body>
          <table>
            <tr><td>Обозначение:</td><td>ГОСТ 7.32-2017</td></tr>
            <tr><td>Статус:</td><td>Действующий</td></tr>
          </table>
        </body>
        </html>
        """
        with patch.object(service.cache_manager, "find_indexed_document", return_value=None):
            with patch.object(service.client, "fetch_url", return_value=(html, False)):
                with patch.object(service.cache_manager, "store_document") as mock_store:
                    res = find_current_gost("7.32-2017", refresh=True)
                    self.assertEqual(res["total_matches"], 1)
                    self.assertEqual(res["primary_document"]["clean_number"], "7.32-2017")
                    self.assertTrue(res["status_summary"]["found"])
                    mock_store.assert_called()

    def test_service_find_current_gost_with_replacement_fetch(self) -> None:
        from gost_standardizer.catalog import service

        superseded_doc = NormDocument(
            id="63653",
            designation="ГОСТ Р 7.0.97-2016",
            clean_number="7.0.97-2016",
            title="Организационно-распорядительная документация",
            status="заменён",
            replaced_by="ГОСТ Р 7.0.97-2025",
        )
        replacement_html = """
        <html>
        <head><title>ГОСТ Р 7.0.97-2025 Документация</title></head>
        <body>
          <table>
            <tr><td>Обозначение:</td><td>ГОСТ Р 7.0.97-2025</td></tr>
            <tr><td>Статус:</td><td>Действующий</td></tr>
            <tr><td>Дата введения:</td><td>01.07.2025</td></tr>
          </table>
        </body>
        </html>
        """

        def mock_find(query: str):
            if "2016" in query:
                return superseded_doc
            return None  # Replacement not cached initially

        with patch.object(service.cache_manager, "find_indexed_document", side_effect=mock_find):
            with patch.object(service.client, "fetch_url", return_value=(replacement_html, False)):
                with patch.object(service.cache_manager, "store_document"):
                    res = find_current_gost("7.0.97-2016")
                    primary = res["primary_document"]
                    self.assertIsNotNone(primary)
                    self.assertIn("active_replacement", primary)
                    self.assertEqual(primary["active_replacement"]["designation"], "ГОСТ Р 7.0.97-2025")
                    self.assertEqual(primary["active_replacement"]["status"], "действующий")

    def test_service_search_category_pages_and_duplicate_skip(self) -> None:
        from gost_standardizer.catalog import service

        cache_data = {
            "categories": [],
            "category_pages": {
                "cat1": {
                    "pages": {
                        "0": {
                            "docs": [
                                {"title": "ГОСТ 1234-2020 Общие требования", "href": "https://meganorm.ru/Index/1.htm"},
                                {"title": "ГОСТ 1234-2020 Общие требования", "href": "https://meganorm.ru/Index/1.htm"},  # Duplicate
                                {"title": "ГОСТ 1234-2021 Дополнительно", "href": "https://meganorm.ru/Index/2.htm"},
                            ]
                        }
                    }
                }
            },
        }

        with patch.object(service.cache_manager, "find_indexed_document", return_value=None):
            with patch.object(service.cache_manager, "get_data", return_value=cache_data):
                res = find_current_gost("1234", max_pages=1, limit=5)
                # Should have 2 unique documents, not 3
                self.assertEqual(len(res["documents"]), 2)
                titles = [d["title"] for d in res["documents"]]
                self.assertIn("ГОСТ 1234-2020 Общие требования", titles)
                self.assertIn("ГОСТ 1234-2021 Дополнительно", titles)

    def test_service_get_current_topics_filter(self) -> None:
        from gost_standardizer.catalog import service

        mock_data = {
            "categories": [
                {"title": "Система стандартов безопасности труда", "slug": "ssbt"},
                {"title": "Единая система конструкторской документации", "slug": "eskd"},
                {"title": "Информация и библиотечное дело", "slug": "sibd"},
            ]
        }
        with patch.object(service.cache_manager, "get_data", return_value=mock_data):
            # All topics
            all_topics = get_current_topics(limit=10)
            self.assertEqual(all_topics["total"], 3)

            # Filtered topics
            filtered = get_current_topics(category="конструкторской")
            self.assertEqual(filtered["total"], 1)
            self.assertEqual(filtered["topics"][0]["slug"], "eskd")

            # Non-matching filter
            none_found = get_current_topics(category="сельское хозяйство")
            self.assertEqual(none_found["total"], 0)

    def test_service_fetch_norm_markdown(self) -> None:
        from gost_standardizer.catalog import service

        # 1. Unresolvable query
        with patch.object(service.cache_manager, "find_indexed_document", return_value=None):
            res_err = fetch_norm_markdown("UNKNOWN_QUERY_9999")
            self.assertIn("error", res_err)

        # 2. Query in KNOWN_INDEX_PAGES but fetch fails
        with patch.object(service.client, "fetch_url", return_value=("", True)):
            res_fetch_err = fetch_norm_markdown("7.32-2017")
            self.assertIn("error", res_fetch_err)

        # 3. Cached doc with url
        cached = NormDocument(
            id="123",
            designation="ГОСТ 123",
            clean_number="123",
            title="Тест",
            url="https://meganorm.ru/Index/1/123.htm",
        )
        html = "<html><head><title>ГОСТ 123</title></head><body><h1>Заголовок</h1><p>Текст стандарта</p></body></html>"
        with patch.object(service.cache_manager, "find_indexed_document", return_value=cached):
            with patch.object(service.client, "fetch_url", return_value=(html, False)):
                with patch.object(service.cache_manager, "store_document"):
                    res = fetch_norm_markdown("ГОСТ 123")
                    self.assertEqual(res["kind"], "fetch-norm-markdown")
                    self.assertIn("# Заголовок", res["markdown"])
                    self.assertIn("Текст стандарта", res["markdown"])


if __name__ == "__main__":
    unittest.main()
