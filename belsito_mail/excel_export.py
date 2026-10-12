"""Create a local, formula-safe Excel report without third-party packages.

Email bodies are deliberately excluded. Text is written as an OOXML inline
string, so even subjects beginning with ``=`` remain text in Excel.
"""

from __future__ import annotations

import os
from pathlib import Path
import tempfile
from datetime import datetime, timezone
from typing import Iterable, TYPE_CHECKING
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

if TYPE_CHECKING:
    from .models import FollowUpItem


_SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CONTENT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_HEADERS = (
    "Status", "Priority", "Categories", "Received", "Sender", "Email",
    "Subject", "Suggested action", "Reason", "Due date", "Source",
)
_WIDTHS = (14, 12, 30, 25, 24, 32, 55, 66, 66, 14, 20)


def _xml_bytes(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _safe_text(value: object) -> str:
    """Remove XML-invalid characters and respect Excel's cell-text limit."""
    text = "" if value is None else str(value)
    cleaned = "".join(
        char for char in text
        if char in "\t\n\r"
        or 0x20 <= ord(char) <= 0xD7FF
        or 0xE000 <= ord(char) <= 0xFFFD
        or 0x10000 <= ord(char) <= 0x10FFFF
    )
    # Count UTF-16 code units as Excel does; never split a surrogate pair.
    return cleaned.encode("utf-16-le")[: 32767 * 2].decode("utf-16-le", "ignore")


def _column_name(index: int) -> str:
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _cell(row: ET.Element, row_number: int, column: int, value: object,
          style: int = 0) -> None:
    cell = ET.SubElement(row, f"{{{_SHEET_NS}}}c", {
        "r": f"{_column_name(column)}{row_number}",
        "t": "inlineStr", "s": str(style),
    })
    inline = ET.SubElement(cell, f"{{{_SHEET_NS}}}is")
    ET.SubElement(inline, f"{{{_SHEET_NS}}}t", {_XML_SPACE: "preserve"}).text = _safe_text(value)


def _timestamp(value: datetime) -> str:
    return value.isoformat(sep=" ", timespec="minutes")


def _worksheet(items: list[FollowUpItem], completed_ids: set[str],
               reviewed_at: datetime, source_label: str) -> bytes:
    sheet = ET.Element(f"{{{_SHEET_NS}}}worksheet")
    last_row = max(4, len(items) + 4)
    ET.SubElement(sheet, f"{{{_SHEET_NS}}}dimension", {"ref": f"A1:K{last_row}"})
    views = ET.SubElement(sheet, f"{{{_SHEET_NS}}}sheetViews")
    view = ET.SubElement(views, f"{{{_SHEET_NS}}}sheetView", {"workbookViewId": "0"})
    ET.SubElement(view, f"{{{_SHEET_NS}}}pane", {
        "ySplit": "4", "topLeftCell": "A5", "activePane": "bottomLeft", "state": "frozen",
    })
    ET.SubElement(view, f"{{{_SHEET_NS}}}selection", {
        "pane": "bottomLeft", "activeCell": "A5", "sqref": "A5",
    })
    ET.SubElement(sheet, f"{{{_SHEET_NS}}}sheetFormatPr", {"defaultRowHeight": "18"})
    columns = ET.SubElement(sheet, f"{{{_SHEET_NS}}}cols")
    for index, width in enumerate(_WIDTHS, start=1):
        ET.SubElement(columns, f"{{{_SHEET_NS}}}col", {
            "min": str(index), "max": str(index), "width": str(width), "customWidth": "1",
        })
    data = ET.SubElement(sheet, f"{{{_SHEET_NS}}}sheetData")
    title = ET.SubElement(data, f"{{{_SHEET_NS}}}row", {"r": "1", "ht": "30", "customHeight": "1"})
    _cell(title, 1, 1, "Belsito Plumbing — Daily Follow-up Report", style=1)
    metadata = ET.SubElement(data, f"{{{_SHEET_NS}}}row", {"r": "2"})
    _cell(metadata, 2, 1, f"Reviewed: {_timestamp(reviewed_at)}", style=2)
    _cell(metadata, 2, 5, f"Source: {source_label}", style=2)
    headers = ET.SubElement(data, f"{{{_SHEET_NS}}}row", {"r": "4", "ht": "30", "customHeight": "1"})
    for column, header in enumerate(_HEADERS, start=1):
        _cell(headers, 4, column, header, style=3)
    for row_number, item in enumerate(items, start=5):
        row = ET.SubElement(data, f"{{{_SHEET_NS}}}row", {
            "r": str(row_number), "ht": "60", "customHeight": "1",
        })
        message = item.message
        values = (
            "Completed" if item.id in completed_ids else "Open",
            item.priority,
            ", ".join(item.categories),
            _timestamp(message.received_at),
            message.sender_name,
            message.sender_email,
            message.subject,
            item.action,
            "; ".join(item.reasons),
            item.due_date.isoformat() if item.due_date is not None else "",
            message.source,
        )
        for column, value in enumerate(values, start=1):
            _cell(row, row_number, column, value)
    ET.SubElement(sheet, f"{{{_SHEET_NS}}}autoFilter", {"ref": f"A4:K{last_row}"})
    merges = ET.SubElement(sheet, f"{{{_SHEET_NS}}}mergeCells", {"count": "3"})
    for area in ("A1:K1", "A2:D2", "E2:K2"):
        ET.SubElement(merges, f"{{{_SHEET_NS}}}mergeCell", {"ref": area})
    ET.SubElement(sheet, f"{{{_SHEET_NS}}}pageMargins", {
        "left": "0.25", "right": "0.25", "top": "0.5", "bottom": "0.5",
        "header": "0.25", "footer": "0.25",
    })
    return _xml_bytes(sheet)


def _styles() -> bytes:
    root = ET.Element(f"{{{_SHEET_NS}}}styleSheet")
    fonts = ET.SubElement(root, f"{{{_SHEET_NS}}}fonts", {"count": "3"})
    for size, bold, color in ((11, False, "FF243047"), (18, True, "FF18344F"), (11, True, "FFFFFFFF")):
        font = ET.SubElement(fonts, f"{{{_SHEET_NS}}}font")
        ET.SubElement(font, f"{{{_SHEET_NS}}}sz", {"val": str(size)})
        ET.SubElement(font, f"{{{_SHEET_NS}}}name", {"val": "Calibri"})
        ET.SubElement(font, f"{{{_SHEET_NS}}}color", {"rgb": color})
        if bold:
            ET.SubElement(font, f"{{{_SHEET_NS}}}b")
    fills = ET.SubElement(root, f"{{{_SHEET_NS}}}fills", {"count": "3"})
    for pattern in ("none", "gray125", "solid"):
        fill = ET.SubElement(fills, f"{{{_SHEET_NS}}}fill")
        pattern_fill = ET.SubElement(fill, f"{{{_SHEET_NS}}}patternFill", {"patternType": pattern})
        if pattern == "solid":
            ET.SubElement(pattern_fill, f"{{{_SHEET_NS}}}fgColor", {"rgb": "FF18344F"})
            ET.SubElement(pattern_fill, f"{{{_SHEET_NS}}}bgColor", {"indexed": "64"})
    borders = ET.SubElement(root, f"{{{_SHEET_NS}}}borders", {"count": "1"})
    border = ET.SubElement(borders, f"{{{_SHEET_NS}}}border")
    for edge in ("left", "right", "top", "bottom", "diagonal"):
        ET.SubElement(border, f"{{{_SHEET_NS}}}{edge}")
    base_styles = ET.SubElement(root, f"{{{_SHEET_NS}}}cellStyleXfs", {"count": "1"})
    ET.SubElement(base_styles, f"{{{_SHEET_NS}}}xf", {
        "numFmtId": "0", "fontId": "0", "fillId": "0", "borderId": "0",
    })
    styles = ET.SubElement(root, f"{{{_SHEET_NS}}}cellXfs", {"count": "4"})
    for font_id, fill_id in ((0, 0), (1, 0), (0, 0), (2, 2)):
        style = ET.SubElement(styles, f"{{{_SHEET_NS}}}xf", {
            "numFmtId": "0", "fontId": str(font_id), "fillId": str(fill_id),
            "borderId": "0", "xfId": "0", "applyAlignment": "1",
            "applyFont": "1", "applyFill": "1",
        })
        ET.SubElement(style, f"{{{_SHEET_NS}}}alignment", {"vertical": "top", "wrapText": "1"})
    named_styles = ET.SubElement(root, f"{{{_SHEET_NS}}}cellStyles", {"count": "1"})
    ET.SubElement(named_styles, f"{{{_SHEET_NS}}}cellStyle", {"name": "Normal", "xfId": "0", "builtinId": "0"})
    return _xml_bytes(root)


def _parts(sheet: bytes, reviewed_at: datetime) -> dict[str, bytes]:
    content = ET.Element(f"{{{_CONTENT_NS}}}Types")
    ET.SubElement(content, f"{{{_CONTENT_NS}}}Default", {"Extension": "rels", "ContentType": "application/vnd.openxmlformats-package.relationships+xml"})
    ET.SubElement(content, f"{{{_CONTENT_NS}}}Default", {"Extension": "xml", "ContentType": "application/xml"})
    for name, content_type in (
        ("/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"),
        ("/xl/worksheets/sheet1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"),
        ("/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"),
        ("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
    ):
        ET.SubElement(content, f"{{{_CONTENT_NS}}}Override", {"PartName": name, "ContentType": content_type})

    relationships = ET.Element(f"{{{_PACKAGE_REL_NS}}}Relationships")
    ET.SubElement(relationships, f"{{{_PACKAGE_REL_NS}}}Relationship", {
        "Id": "rId1", "Type": f"{_REL_NS}/officeDocument", "Target": "xl/workbook.xml",
    })
    ET.SubElement(relationships, f"{{{_PACKAGE_REL_NS}}}Relationship", {
        "Id": "rId2", "Type": "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "Target": "docProps/core.xml",
    })
    workbook = ET.Element(f"{{{_SHEET_NS}}}workbook")
    views = ET.SubElement(workbook, f"{{{_SHEET_NS}}}bookViews")
    ET.SubElement(views, f"{{{_SHEET_NS}}}workbookView")
    sheets = ET.SubElement(workbook, f"{{{_SHEET_NS}}}sheets")
    ET.SubElement(sheets, f"{{{_SHEET_NS}}}sheet", {
        "name": "Daily follow-up", "sheetId": "1", f"{{{_REL_NS}}}id": "rId1",
    })
    workbook_rels = ET.Element(f"{{{_PACKAGE_REL_NS}}}Relationships")
    for identifier, kind, target in (("rId1", "worksheet", "worksheets/sheet1.xml"), ("rId2", "styles", "styles.xml")):
        ET.SubElement(workbook_rels, f"{{{_PACKAGE_REL_NS}}}Relationship", {
            "Id": identifier, "Type": f"{_REL_NS}/{kind}", "Target": target,
        })
    core_ns = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    dc_ns = "http://purl.org/dc/elements/1.1/"
    dcterms_ns = "http://purl.org/dc/terms/"
    xsi_ns = "http://www.w3.org/2001/XMLSchema-instance"
    core = ET.Element(f"{{{core_ns}}}coreProperties", {"xmlns:dcterms": dcterms_ns})
    ET.SubElement(core, f"{{{dc_ns}}}title").text = "Belsito Plumbing Daily Follow-up Report"
    ET.SubElement(core, f"{{{dc_ns}}}creator").text = "Belsito Plumbing Email Assistant"
    utc_time = reviewed_at.replace(tzinfo=timezone.utc) if reviewed_at.tzinfo is None else reviewed_at.astimezone(timezone.utc)
    # The QName in xsi:type must resolve even though ElementTree chooses its
    # own prefix for the created element.
    ET.SubElement(core, f"{{{dcterms_ns}}}created", {f"{{{xsi_ns}}}type": "dcterms:W3CDTF"}).text = utc_time.isoformat(timespec="seconds").replace("+00:00", "Z")
    return {
        "[Content_Types].xml": _xml_bytes(content),
        "_rels/.rels": _xml_bytes(relationships),
        "xl/workbook.xml": _xml_bytes(workbook),
        "xl/_rels/workbook.xml.rels": _xml_bytes(workbook_rels),
        "xl/worksheets/sheet1.xml": sheet,
        "xl/styles.xml": _styles(),
        "docProps/core.xml": _xml_bytes(core),
    }


def export_report(items: Iterable[FollowUpItem], completed_ids: set[str],
                  path: Path, reviewed_at: datetime, source_label: str) -> None:
    """Atomically write an .xlsx report in the supplied prioritized order.

    The caller is responsible for obtaining permission to overwrite an
    existing export. If preparation or writing fails, that existing file is
    preserved. The destination's parent directory must already exist.
    """
    path = Path(path)
    if path.suffix.lower() != ".xlsx":
        raise ValueError("Choose an Excel report filename ending in .xlsx.")
    rows = list(items)
    if len(rows) > 1_048_572:
        raise ValueError("This report exceeds Excel's worksheet row limit.")
    parts = _parts(_worksheet(rows, completed_ids, reviewed_at, source_label), reviewed_at)
    temporary_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(mode="w+b", prefix=f".{path.stem[:80]}-",
                                         suffix=".xlsx.tmp", dir=path.parent,
                                         delete=False) as temporary:
            temporary_path = Path(temporary.name)
            with ZipFile(temporary, "w", compression=ZIP_DEFLATED) as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
            temporary.flush()
            os.fsync(temporary.fileno())
        os.replace(temporary_path, path)
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)
