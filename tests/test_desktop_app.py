import importlib
import sys
import tempfile
import types
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import desktop_app


class DesktopAppTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.excel_path = self.root / "students.xlsx"
        self.template_path = self.root / "certificate.html"
        self.output_dir = self.root / "output"
        self.template_path.write_text(
            "<html><body><h1>{{ student_name }}</h1>"
            "<p>{{ course_name }}</p><p>{{ score }}</p>"
            "<p>{{ issue_date }}</p>"
            '<img src="{{ signature_path }}"><img src="{{ stamp_path }}">'
            "</body></html>",
            encoding="utf-8",
        )

    def tearDown(self):
        self.temp_dir.cleanup()

    def write_workbook(self, headers, rows=()):
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append(headers)
        for row in rows:
            worksheet.append(row)
        workbook.save(self.excel_path)
        workbook.close()

    def test_validate_and_load_excel_maps_headers_and_formats_dates(self):
        self.write_workbook(
            ("issue_date", "student_name", "course_name", "score", "custom_field"),
            (
                (date(2026, 10, 31), "Анна Иванова", "Python", 95, "value"),
                (None, None, None, None, None),
            ),
        )

        students = desktop_app.validate_and_load_excel(self.excel_path)

        self.assertEqual(len(students), 1)
        self.assertEqual(
            students[0],
            {
                "issue_date": "31.10.2026",
                "student_name": "Анна Иванова",
                "course_name": "Python",
                "score": "95",
                "custom_field": "value",
            },
        )

    def test_validate_and_load_excel_reports_missing_columns(self):
        self.write_workbook(("student_name", "course_name"))

        with self.assertRaises(desktop_app.MissingColumnsError) as context:
            desktop_app.validate_and_load_excel(self.excel_path)

        self.assertEqual(
            context.exception.missing_columns,
            ("score", "issue_date"),
        )

    def test_validate_and_load_excel_rejects_duplicate_headers_and_empty_data(self):
        self.write_workbook(
            ("student_name", "student_name", "course_name", "score", "issue_date")
        )
        with self.assertRaisesRegex(ValueError, "повторяющиеся"):
            desktop_app.validate_and_load_excel(self.excel_path)

        self.write_workbook(desktop_app.REQUIRED_COLUMNS)
        with self.assertRaisesRegex(ValueError, "нет строк"):
            desktop_app.validate_and_load_excel(self.excel_path)

    def test_validate_and_load_excel_rejects_non_xlsx(self):
        with self.assertRaisesRegex(ValueError, r"\.xlsx"):
            desktop_app.validate_and_load_excel(self.root / "students.csv")

    def test_render_certificate_escapes_values_and_omits_empty_images(self):
        rendered = desktop_app.render_certificate(
            self.template_path.read_text(encoding="utf-8"),
            {
                "student_name": "<Анна>",
                "course_name": "Python & HTML",
                "score": "95",
                "issue_date": "31.10.2026",
            },
        )

        self.assertIn("&lt;Анна&gt;", rendered)
        self.assertIn("Python &amp; HTML", rendered)
        self.assertNotIn("<img", rendered)
        self.assertNotIn("signature_path", rendered)

    def test_render_certificate_resolves_images_only_within_assets(self):
        assets_dir = self.root / "assets"
        assets_dir.mkdir()
        signature = assets_dir / "signature.png"
        signature.write_bytes(b"signature")
        student = {
            "student_name": "Анна",
            "course_name": "Python",
            "score": "95",
            "issue_date": "31.10.2026",
            "signature_path": "signature.png",
        }

        with patch.object(desktop_app, "ASSETS_DIR", assets_dir):
            rendered = desktop_app.render_certificate(
                self.template_path.read_text(encoding="utf-8"),
                student,
            )
            self.assertIn(signature.resolve().as_uri(), rendered)
            student["signature_path"] = str(self.root / "outside.png")
            with self.assertRaisesRegex(ValueError, "должен находиться"):
                desktop_app.render_certificate(
                    self.template_path.read_text(encoding="utf-8"),
                    student,
                )

    def test_safe_resource_fetcher_rejects_network_resources(self):
        fetcher = desktop_app._safe_resource_fetcher((self.root,))

        with self.assertRaisesRegex(ValueError, "локальные ресурсы"):
            fetcher("https://example.com/image.png")

    def test_generate_certificates_writes_unique_pdfs_and_reports_progress(self):
        self.write_workbook(
            desktop_app.REQUIRED_COLUMNS,
            [
                ("Анна Иванова", "Python & HTML", 95, date(2026, 10, 31)),
                ("Иван Петров", "Excel", 90, date(2026, 10, 30)),
            ],
        )
        rendered_documents = []

        class FakeHTML:
            def __init__(self, *, string, base_url, url_fetcher):
                rendered_documents.append((string, base_url, url_fetcher))

            def write_pdf(self, *, target):
                Path(target).write_bytes(b"%PDF-test")

        weasyprint = types.ModuleType("weasyprint")
        weasyprint.HTML = FakeHTML
        url_module = importlib.import_module("weasyprint.urls")
        events = []

        with patch.dict(
            sys.modules,
            {"weasyprint": weasyprint, "weasyprint.urls": url_module},
        ):
            first = desktop_app.generate_certificates(
                self.excel_path,
                self.template_path,
                self.output_dir,
                lambda current, total, message: events.append(
                    (current, total, message)
                ),
            )
            second = desktop_app.generate_certificates(
                self.excel_path,
                self.template_path,
                self.output_dir,
            )

        self.assertEqual(len(first), 2)
        self.assertEqual(len(second), 2)
        self.assertNotEqual(first[0], second[0])
        self.assertTrue(all(path.read_bytes() == b"%PDF-test" for path in first))
        self.assertEqual([event[:2] for event in events], [(1, 2), (2, 2)])
        self.assertIn("&amp;", rendered_documents[0][0])

    def test_generate_certificates_rejects_missing_template_columns(self):
        self.write_workbook(
            desktop_app.REQUIRED_COLUMNS,
            [("Анна Иванова", "Python", 95, "31.10.2026")],
        )
        self.template_path.write_text(
            "{{ student_name }} {{ unknown_field }}",
            encoding="utf-8",
        )

        with self.assertRaisesRegex(ValueError, "unknown_field"):
            desktop_app.generate_certificates(
                self.excel_path,
                self.template_path,
                self.output_dir,
            )

    def test_safe_filename_removes_cross_platform_forbidden_characters(self):
        self.assertEqual(
            desktop_app._safe_filename('Иван: "Тест"/?'),
            "Иван___Тест___",
        )


if __name__ == "__main__":
    unittest.main()
