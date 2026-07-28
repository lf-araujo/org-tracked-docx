#!/usr/bin/env python3
"""Patch every table in a docx's word/document.xml with direct (not
style-level) three-line borders: a single top border, a single bottom
border, and a single bottom border under the header row only -- no
vertical lines, no lines between body rows.

Used as a post-processing stage in otd-export (org-tracked-docx.el).
Direct formatting is used instead of relying on the "Table" style's own
tblBorders / firstRow conditional formatting, because table-STYLE-level
borders and conditional row formatting are applied inconsistently across
Word/LibreOffice versions; direct per-table, per-cell borders always
render the same way everywhere.

Usage: otd-table-borders.py DOCX_PATH
Patches DOCX_PATH in place.
"""
import sys
import re
import zipfile
import shutil
import tempfile
import os

TBL_BORDERS = (
    '<w:tblBorders>'
    '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    '<w:left w:val="nil"/>'
    '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    '<w:right w:val="nil"/>'
    '<w:insideH w:val="nil"/>'
    '<w:insideV w:val="nil"/>'
    '</w:tblBorders>'
)
HEADER_CELL_BORDER = (
    '<w:tcBorders><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tcBorders>'
)


def patch_tblpr(m):
    tblpr = m.group(0)
    if "<w:tblBorders>" in tblpr:
        return tblpr
    if "<w:tblLook" in tblpr:
        return tblpr.replace("<w:tblLook", TBL_BORDERS + "<w:tblLook", 1)
    return tblpr[: -len("</w:tblPr>")] + TBL_BORDERS + "</w:tblPr>"


def patch_header_row(row_xml):
    def add_border(m):
        inner = m.group(1)
        if "<w:tcBorders>" in inner:
            return m.group(0)
        return "<w:tcPr>" + inner + HEADER_CELL_BORDER + "</w:tcPr>"

    row_xml = re.sub(r"<w:tcPr\s*/>", "<w:tcPr>" + HEADER_CELL_BORDER + "</w:tcPr>", row_xml)
    row_xml = re.sub(r"<w:tcPr>(.*?)</w:tcPr>", add_border, row_xml, flags=re.DOTALL)
    return row_xml


def patch_table(m):
    table = m.group(0)
    table = re.sub(r"<w:tblPr>.*?</w:tblPr>", patch_tblpr, table, count=1, flags=re.DOTALL)
    rows = list(re.finditer(r"<w:tr>.*?</w:tr>", table, re.DOTALL))
    if rows and "w:tblHeader" in rows[0].group(0):
        patched = patch_header_row(rows[0].group(0))
        table = table[: rows[0].start()] + patched + table[rows[0].end():]
    return table


def patch_docx(path):
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".docx")
    os.close(tmp_fd)
    with zipfile.ZipFile(path) as zin:
        doc = zin.read("word/document.xml").decode("utf-8")
        new_doc = re.sub(r"<w:tbl>.*?</w:tbl>", patch_table, doc, flags=re.DOTALL)
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    data = new_doc.encode("utf-8")
                zout.writestr(item, data)
    shutil.move(tmp_path, path)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("usage: otd-table-borders.py DOCX_PATH", file=sys.stderr)
        sys.exit(1)
    patch_docx(sys.argv[1])
