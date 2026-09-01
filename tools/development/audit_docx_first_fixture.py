import json
from pathlib import Path
import sys
import zipfile

from lxml import etree
from pypdf import PdfReader


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def integer_attr(node, name):
    return int(node.get(W + name))


def audit(docx_path, pdf_path):
    failures = []
    with zipfile.ZipFile(docx_path) as archive:
        document = etree.fromstring(archive.read("word/document.xml"))
        styles = etree.fromstring(archive.read("word/styles.xml"))

    section = document.find(".//w:sectPr", NS)
    page_size = section.find("w:pgSz", NS)
    page_margin = section.find("w:pgMar", NS)
    geometry = {
        "pageWidthDxa": integer_attr(page_size, "w"),
        "pageHeightDxa": integer_attr(page_size, "h"),
        "topDxa": integer_attr(page_margin, "top"),
        "rightDxa": integer_attr(page_margin, "right"),
        "bottomDxa": integer_attr(page_margin, "bottom"),
        "leftDxa": integer_attr(page_margin, "left"),
        "headerDxa": integer_attr(page_margin, "header"),
        "footerDxa": integer_attr(page_margin, "footer"),
    }
    expected_geometry = {
        "pageWidthDxa": 11906,
        "pageHeightDxa": 16838,
        "topDxa": 794,
        "rightDxa": 1020,
        "bottomDxa": 794,
        "leftDxa": 1020,
        "headerDxa": 397,
        "footerDxa": 397,
    }
    for key, expected in expected_geometry.items():
        if abs(geometry[key] - expected) > 2:
            failures.append(f"page geometry {key}: {geometry[key]} != {expected}")

    normal = styles.find(".//w:style[@w:styleId='Normal']", NS)
    fonts = normal.find("w:rPr/w:rFonts", NS)
    size = normal.find("w:rPr/w:sz", NS)
    spacing = normal.find("w:pPr/w:spacing", NS)
    normal_audit = {
        "ascii": fonts.get(W + "ascii"),
        "eastAsia": fonts.get(W + "eastAsia"),
        "sizeHalfPoints": integer_attr(size, "val"),
        "afterTwips": integer_attr(spacing, "after"),
        "lineTwips": integer_attr(spacing, "line"),
    }
    if normal_audit["eastAsia"] != "Malgun Gothic":
        failures.append("Normal eastAsia font is not Malgun Gothic")
    if normal_audit["sizeHalfPoints"] != 20:
        failures.append("Normal size is not 10 pt")
    if normal_audit["afterTwips"] != 80 or normal_audit["lineTwips"] != 264:
        failures.append("Normal paragraph spacing differs from the token map")

    tables = document.findall(".//w:tbl", NS)
    table_audit = []
    for index, table in enumerate(tables):
        width = integer_attr(table.find("w:tblPr/w:tblW", NS), "w")
        indent = integer_attr(table.find("w:tblPr/w:tblInd", NS), "w")
        grid = [integer_attr(node, "w") for node in table.findall("w:tblGrid/w:gridCol", NS)]
        row_count = len(table.findall("w:tr", NS))
        table_audit.append({
            "index": index,
            "widthDxa": width,
            "indentDxa": indent,
            "gridDxa": grid,
            "gridTotalDxa": sum(grid),
            "rowCount": row_count,
        })
        if width != 9865 or indent != 120 or sum(grid) != 9865:
            failures.append(f"table {index} geometry mismatch")
    if len(tables) != 6:
        failures.append(f"expected 6 tables, found {len(tables)}")

    text = "".join(document.itertext())
    required = ["현장기술인 변경계", "홍길동", "품질관리", "발주처 담당자"]
    missing = [value for value in required if value not in text]
    if missing:
        failures.append("missing required DOCX text: " + ", ".join(missing))

    reader = PdfReader(str(pdf_path))
    pdf_page_count = len(reader.pages)
    media = reader.pages[0].mediabox
    pdf_size_points = [float(media.width), float(media.height)]
    if pdf_page_count != 1:
        failures.append(f"DOCX PDF render has {pdf_page_count} pages")
    if abs(pdf_size_points[0] - 595.3) > 1 or abs(pdf_size_points[1] - 841.9) > 1:
        failures.append(f"DOCX PDF render is not A4: {pdf_size_points}")

    return {
        "ok": not failures,
        "preset": "standard_business_brief",
        "namedOverrides": {
            "page": "A4 portrait, margins 14/18/14/18 mm",
            "baseFont": "Malgun Gothic 10 pt",
            "contentWidthDxa": 9865,
            "bodySpacing": "after 4 pt, line 1.10",
            "headerTemplate": "memo_masthead without header border",
        },
        "pageGeometry": geometry,
        "normalStyle": normal_audit,
        "tables": table_audit,
        "requiredTextMissing": missing,
        "pdfPageCount": pdf_page_count,
        "pdfPageSizePoints": pdf_size_points,
        "failures": failures,
    }


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit("usage: audit_docx_first_fixture.py INPUT.docx RENDER.pdf")
    result = audit(Path(sys.argv[1]).resolve(), Path(sys.argv[2]).resolve())
    print(json.dumps(result, ensure_ascii=False, indent=2))
    raise SystemExit(0 if result["ok"] else 1)
