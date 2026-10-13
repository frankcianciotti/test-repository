"""Exercise workbook contents and safe export behavior using only stdlib."""

from datetime import date, datetime, timezone
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest
from unittest.mock import patch
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from belsito_mail.excel_export import export_report


_NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
_REVIEWED = datetime(2026, 10, 12, 14, 30, tzinfo=timezone.utc)


def _item(identifier="item-1", subject="Customer service request", **changes):
    message = SimpleNamespace(
        id=f"mail-{identifier}", subject=subject,
        sender_name="Jones & Sons <Plumbing>", sender_email="customer@example.test",
        received_at=datetime(2026, 10, 11, 10, 15, tzinfo=timezone.utc),
        body="CONFIDENTIAL BODY MUST NEVER BE EXPORTED", source="sample",
    )
    item = SimpleNamespace(
        id=identifier, message=message, categories=("Customer service", "New job"),
        priority="High", action="Call customer & schedule a visit",
        reasons=("Customer requested service", "Response needed"), due_date=date(2026, 10, 13),
    )
    for key, value in changes.items():
        setattr(item, key, value)
    return item


def _read_sheet(path):
    with ZipFile(path) as archive:
        return ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))


def _text(sheet, address):
    return sheet.find(f".//s:c[@r='{address}']/s:is/s:t", _NS).text or ""


class ExcelExportTests(unittest.TestCase):
    def setUp(self):
        self.directory = TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.path = Path(self.directory.name) / "report.xlsx"

    def _export(self, items, completed_ids=None):
        export_report(items, completed_ids or set(), self.path, _REVIEWED, "Sample data")

    def test_valid_package_with_report_metadata_and_readable_layout(self):
        self._export([_item()])
        with ZipFile(self.path) as archive:
            self.assertIsNone(archive.testzip())
            expected = {
                "[Content_Types].xml", "_rels/.rels", "xl/workbook.xml",
                "xl/_rels/workbook.xml.rels", "xl/worksheets/sheet1.xml",
                "xl/styles.xml", "docProps/core.xml",
            }
            self.assertEqual(set(archive.namelist()), expected)
            for name in expected:
                ET.fromstring(archive.read(name))
        sheet = _read_sheet(self.path)
        self.assertIn("Belsito Plumbing", _text(sheet, "A1"))
        self.assertIn("2026-10-12 14:30+00:00", _text(sheet, "A2"))
        self.assertEqual(_text(sheet, "E2"), "Source: Sample data")
        self.assertEqual(sheet.find("s:sheetViews/s:sheetView/s:pane", _NS).attrib["topLeftCell"], "A5")
        self.assertEqual(sheet.find("s:autoFilter", _NS).attrib["ref"], "A4:K5")
        self.assertEqual(len(sheet.findall("s:cols/s:col", _NS)), 11)
        self.assertEqual(_text(sheet, "A4"), "Status")
        self.assertEqual(_text(sheet, "K4"), "Source")

    def test_all_items_keep_priority_order_and_completion_status(self):
        first = _item("urgent", "URGENT sewer backup", priority="Urgent")
        second = _item("estimate", "Estimate follow-up", priority="Normal", due_date=None)
        self._export(iter([first, second]), {"urgent"})
        sheet = _read_sheet(self.path)
        self.assertEqual(_text(sheet, "G5"), first.message.subject)
        self.assertEqual(_text(sheet, "G6"), second.message.subject)
        self.assertEqual(_text(sheet, "A5"), "Completed")
        self.assertEqual(_text(sheet, "A6"), "Open")
        self.assertEqual(_text(sheet, "B5"), "Urgent")
        self.assertEqual(_text(sheet, "C5"), "Customer service, New job")
        self.assertEqual(_text(sheet, "D5"), "2026-10-11 10:15+00:00")
        self.assertEqual(_text(sheet, "I5"), "Customer requested service; Response needed")
        self.assertEqual(_text(sheet, "J5"), "2026-10-13")
        self.assertEqual(_text(sheet, "J6"), "")
        self.assertEqual(_text(sheet, "K5"), "sample")
        self.assertEqual(sheet.find("s:autoFilter", _NS).attrib["ref"], "A4:K6")

    def test_office_parts_use_conventional_default_namespaces_and_bound_prefixes(self):
        self._export([_item()])
        namespaces = {
            "[Content_Types].xml": "http://schemas.openxmlformats.org/package/2006/content-types",
            "_rels/.rels": "http://schemas.openxmlformats.org/package/2006/relationships",
            "xl/_rels/workbook.xml.rels": "http://schemas.openxmlformats.org/package/2006/relationships",
            "xl/workbook.xml": _NS["s"],
            "xl/styles.xml": _NS["s"],
            "xl/worksheets/sheet1.xml": _NS["s"],
        }
        with ZipFile(self.path) as archive:
            for name, namespace in namespaces.items():
                with self.subTest(part=name):
                    data = archive.read(name)
                    declared = dict(event for _, event in ET.iterparse(BytesIO(data), events=("start-ns",)))
                    self.assertEqual(declared.get(""), namespace)
                    self.assertNotIn(b"ns0:", data)
                    self.assertTrue(ET.fromstring(data).tag.startswith("{" + namespace + "}"))
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
            sheet = workbook.find("s:sheets/s:sheet", _NS)
            self.assertEqual(sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"], "rId1")
            core_data = archive.read("docProps/core.xml")
            declared = dict(event for _, event in ET.iterparse(BytesIO(core_data), events=("start-ns",)))
            self.assertEqual(declared["dcterms"], "http://purl.org/dc/terms/")
            created = ET.fromstring(core_data).find("{http://purl.org/dc/terms/}created")
            self.assertEqual(created.attrib["{http://www.w3.org/2001/XMLSchema-instance}type"], "dcterms:W3CDTF")

    def test_fonts_follow_microsoft_office_schema_order(self):
        self._export([_item()])
        with ZipFile(self.path) as archive:
            styles = ET.fromstring(archive.read("xl/styles.xml"))
        fonts = styles.findall("s:fonts/s:font", _NS)
        # Relevant subsequence of CT_Font's Microsoft Open XML SDK particle:
        # b, ..., sz, color, name. This checks the interoperability contract.
        order = {name: i for i, name in enumerate(("b", "sz", "color", "name"))}
        self.assertEqual(len(fonts), 3)
        for font in fonts:
            positions = [order[child.tag.split("}", 1)[1]] for child in font]
            self.assertEqual(positions, sorted(positions))

    def test_email_text_cannot_create_formulas_links_or_external_references(self):
        subjects = (
            '=HYPERLINK("https://example.test", "Click")',
            "+SUM(1,2)", "-1+2", "@SUM(A1:A2)",
            "[remote.xlsx]Sheet1!A1", "\t=WEBSERVICE(\"https://example.test\")",
        )
        self._export([_item(str(index), subject) for index, subject in enumerate(subjects)])
        sheet = _read_sheet(self.path)
        self.assertEqual(sheet.findall(".//s:f", _NS), [])
        self.assertIsNone(sheet.find("s:hyperlinks", _NS))
        for row, subject in enumerate(subjects, start=5):
            cell = sheet.find(f".//s:c[@r='G{row}']", _NS)
            self.assertEqual(cell.attrib["t"], "inlineStr")
            self.assertEqual(_text(sheet, f"G{row}"), subject)
        with ZipFile(self.path) as archive:
            for name in archive.namelist():
                self.assertNotIn("external", name.lower())
            rels = archive.read("xl/_rels/workbook.xml.rels").decode()
            self.assertNotIn('TargetMode="External"', rels)

    def test_xml_special_characters_and_invalid_controls_are_handled(self):
        subject = '  Pipe & valve <repair> "today"\x00\x01\x0b\x0c\ud800\ufffe\uffff\nKeep\twhitespace  '
        self._export([_item(subject=subject)])
        sheet = _read_sheet(self.path)
        self.assertEqual(_text(sheet, "G5"), '  Pipe & valve <repair> "today"\nKeep\twhitespace  ')
        self.assertEqual(_text(sheet, "E5"), "Jones & Sons <Plumbing>")
        self.assertEqual(_text(sheet, "H5"), "Call customer & schedule a visit")
        text_element = sheet.find(".//s:c[@r='G5']/s:is/s:t", _NS)
        self.assertEqual(text_element.attrib["{http://www.w3.org/XML/1998/namespace}space"], "preserve")

    def test_long_cell_text_fits_excel_limit_without_splitting_emoji(self):
        self._export([_item(subject="A" * 32766 + "\U0001f6bf" + "tail")])
        text = _text(_read_sheet(self.path), "G5")
        self.assertEqual(text, "A" * 32766)
        self.assertLessEqual(len(text.encode("utf-16-le")) // 2, 32767)

    def test_email_bodies_are_never_written(self):
        self._export([_item()])
        with ZipFile(self.path) as archive:
            for name in archive.namelist():
                self.assertNotIn(b"CONFIDENTIAL BODY", archive.read(name))

    def test_empty_report_remains_a_valid_filtered_workbook(self):
        self._export([])
        sheet = _read_sheet(self.path)
        self.assertEqual(sheet.find("s:dimension", _NS).attrib["ref"], "A1:K4")
        self.assertEqual(sheet.find("s:autoFilter", _NS).attrib["ref"], "A4:K4")
        self.assertIsNone(sheet.find(".//s:c[@r='A5']", _NS))

    def test_atomic_replacement_preserves_previous_export_if_file_is_locked(self):
        self.path.write_bytes(b"Existing report contents")
        with patch("belsito_mail.excel_export.os.replace", side_effect=PermissionError("file locked")):
            with self.assertRaises(PermissionError):
                self._export([_item()])
        self.assertEqual(self.path.read_bytes(), b"Existing report contents")
        self.assertEqual(list(Path(self.directory.name).iterdir()), [self.path])

    def test_failed_zip_write_preserves_existing_file_and_cleans_temporary_file(self):
        self.path.write_bytes(b"Existing report contents")
        with patch("belsito_mail.excel_export.ZipFile.writestr", side_effect=OSError("disk full")):
            with self.assertRaises(OSError):
                self._export([_item()])
        self.assertEqual(self.path.read_bytes(), b"Existing report contents")
        self.assertEqual(list(Path(self.directory.name).iterdir()), [self.path])

    def test_success_replaces_existing_file_and_leaves_no_temporary_file(self):
        self.path.write_bytes(b"Existing report contents")
        self._export([_item()])
        self.assertTrue(self.path.read_bytes().startswith(b"PK"))
        self.assertEqual(list(Path(self.directory.name).iterdir()), [self.path])

    def test_wrong_extension_is_rejected_before_any_file_is_written(self):
        with self.assertRaisesRegex(ValueError, "xlsx"):
            export_report([_item()], set(), self.path.with_suffix(".csv"), _REVIEWED, "Sample data")
        self.assertEqual(list(Path(self.directory.name).iterdir()), [])


if __name__ == "__main__":
    unittest.main()
