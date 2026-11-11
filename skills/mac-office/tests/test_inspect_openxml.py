"""Unit tests: headless OpenXML inspection (``--inspect``) for .docx/.xlsx/.pptx.

No Microsoft Office involved: every package is generated in-memory with ``zipfile`` and
parsed by the runner's zipfile + ElementTree inspectors.
"""

import os
import unittest

import support
from support import MacOfficeTestCase, run_runner

LONG_TEXT = "word " * 60  # 300 characters -> collapsed to 299, trimmed to 200 + ellipsis


class InspectDocxTests(MacOfficeTestCase):
    def build(self, name="report.docx"):
        path = self.path(name)
        support.write_docx(
            path,
            [
                ("p", "Quarterly Report", "Heading1"),
                ("p", "Body paragraph with normal style", None),
                ("tbl", [["Name", "Value"], ["alpha", "1"], ["beta", "2"]]),
                ("p", "After the first table", None),
                ("tbl", [["Only", "Row", "Three"], ["second", "row", "cells"], ["solo"]]),
                ("p", LONG_TEXT, None),
            ],
        )
        return path

    def test_contract_and_counts(self):
        path = self.build()
        result = run_runner(["--inspect", path])
        self.assertEqual(result.returncode, 0, result.stderr)
        payload = result.payload
        self.assertTrue(payload["success"])
        self.assertEqual(payload["mode"], "inspect")
        self.assertEqual(payload["kind"], "docx")
        self.assertEqual(payload["file"], os.path.abspath(path))
        self.assertEqual(payload["inspector"], "headless_openxml")
        self.assertEqual(payload["container"], "openxml")
        self.assertIsNone(payload["hint"])
        self.assertIsNone(payload["error"])
        self.assertIsNone(payload["error_hint"])
        self.assertIsInstance(payload["execution_time_sec"], float)
        self.assertEqual(payload["paragraph_count"], 4)
        self.assertFalse(payload["paragraphs_truncated"])
        self.assertEqual(payload["table_count"], 2)
        self.assertFalse(payload["tables_truncated"])

    def test_paragraph_indices_styles_and_text(self):
        payload = run_runner(["--inspect", self.build()]).payload
        paragraphs = payload["paragraphs"]
        self.assertEqual([item["index"] for item in paragraphs], [1, 2, 3, 4])
        self.assertEqual(paragraphs[0], {"index": 1, "style": "Heading1", "text": "Quarterly Report"})
        self.assertEqual(paragraphs[1]["style"], None)
        self.assertEqual(paragraphs[1]["text"], "Body paragraph with normal style")
        # Paragraphs inside table cells are not document paragraphs.
        self.assertEqual(paragraphs[2]["text"], "After the first table")
        self.assertEqual(len(paragraphs[3]["text"]), 200)
        self.assertTrue(paragraphs[3]["text"].endswith("…"))

    def test_whitespace_is_collapsed(self):
        path = self.path("spaced.docx")
        support.write_docx(path, [("p", "  spaced\n\tout   text  ", None)])
        payload = run_runner(["--inspect", path]).payload
        self.assertEqual(payload["paragraphs"][0]["text"], "spaced out text")

    def test_table_counts_and_preview(self):
        payload = run_runner(["--inspect", self.build()]).payload
        first, second = payload["tables"]
        self.assertEqual(first["index"], 1)
        self.assertEqual(first["rows"], 3)
        self.assertEqual(first["columns"], 2)
        self.assertEqual(first["preview"], [["Name", "Value"], ["alpha", "1"], ["beta", "2"]])
        # Uneven rows: columns is the widest row, missing cells are simply absent.
        self.assertEqual(second["index"], 2)
        self.assertEqual(second["rows"], 3)
        self.assertEqual(second["columns"], 3)
        self.assertEqual(second["preview"][2], ["solo"])

    def test_table_preview_caps_rows_at_five_and_cells_at_80_chars(self):
        path = self.path("wide.docx")
        rows = [["r{}c1".format(i), "r{}c2".format(i)] for i in range(1, 7)]
        rows.append(["long-" + "x" * 120, "tail"])
        support.write_docx(path, [("tbl", rows)])
        table = run_runner(["--inspect", path]).payload["tables"][0]
        self.assertEqual(table["rows"], 7)
        self.assertEqual(len(table["preview"]), 5)
        self.assertEqual(table["preview"][4], ["r5c1", "r5c2"])

        short = self.path("short-cell.docx")
        support.write_docx(short, [("tbl", [["y" * 100]])])
        cell = run_runner(["--inspect", short]).payload["tables"][0]["preview"][0][0]
        self.assertEqual(len(cell), 80)
        self.assertTrue(cell.endswith("…"))

    def test_max_items_truncates_lists_but_keeps_totals(self):
        payload = run_runner(["--inspect", self.build(), "--max-items", "1"]).payload
        self.assertEqual(payload["paragraph_count"], 4)
        self.assertEqual(len(payload["paragraphs"]), 1)
        self.assertTrue(payload["paragraphs_truncated"])
        self.assertEqual(payload["table_count"], 2)
        self.assertEqual(len(payload["tables"]), 1)
        self.assertTrue(payload["tables_truncated"])
        # --max-items does not shrink the per-table row preview (still capped at 5).
        self.assertEqual(len(payload["tables"][0]["preview"]), 3)

    def test_default_max_items_is_two_hundred(self):
        path = self.path("long.docx")
        support.write_docx(path, [("p", "line {}".format(number), None) for number in range(1, 226)])
        payload = run_runner(["--inspect", path]).payload
        self.assertEqual(payload["paragraph_count"], 225)
        self.assertEqual(len(payload["paragraphs"]), 200)
        self.assertTrue(payload["paragraphs_truncated"])

    def test_extension_is_case_insensitive(self):
        source = self.build()
        upper = self.path("REPORT.DOCX")
        os.replace(source, upper)
        result = run_runner(["--inspect", upper])
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(result.payload["kind"], "docx")


class InspectXlsxTests(MacOfficeTestCase):
    def build(self, name="book.xlsx"):
        path = self.path(name)
        support.write_xlsx(
            path,
            [
                {
                    "name": "Data",
                    "dimension": "A1:C3",
                    "rows": [
                        ["Name", "Score", ("runs", ["Rich", "Text"])],
                        ["alpha", None, 42],
                        [("inline", "Inline Cell"), 3.5],
                    ],
                },
                {"name": "Empty", "dimension": "A1", "rows": []},
                {"name": "NoDim", "dimension": None, "rows": [[1, "two"], [3, 4]]},
            ],
        )
        return path

    def test_sheet_names_ranges_and_values(self):
        result = run_runner(["--inspect", self.build()])
        self.assertEqual(result.returncode, 0, result.stderr)
        payload = result.payload
        self.assertTrue(payload["success"])
        self.assertEqual(payload["kind"], "xlsx")
        self.assertEqual(payload["inspector"], "headless_openxml")
        self.assertEqual(payload["sheet_count"], 3)
        self.assertFalse(payload["sheets_truncated"])
        self.assertEqual([sheet["name"] for sheet in payload["worksheets"]], ["Data", "Empty", "NoDim"])

        data, empty, no_dimension = payload["worksheets"]
        self.assertEqual(data["used_range"], "A1:C3")
        self.assertEqual(
            data["preview"],
            [["Name", "Score", "RichText"], ["alpha", None, "42"], ["Inline Cell", "3.5"]],
        )
        self.assertFalse(data["truncated"])
        self.assertEqual(empty["used_range"], "A1")
        self.assertEqual(empty["preview"], [])
        # No <dimension> element: derived from the highest row/column seen.
        self.assertEqual(no_dimension["used_range"], "A1:B2")
        self.assertEqual(no_dimension["preview"], [["1", "two"], ["3", "4"]])
        self.assertFalse(no_dimension["truncated"])

    def test_max_items_truncates_sheets_and_rows(self):
        payload = run_runner(["--inspect", self.build(), "--max-items", "1"]).payload
        self.assertEqual(payload["sheet_count"], 3)
        self.assertEqual(len(payload["worksheets"]), 1)
        self.assertTrue(payload["sheets_truncated"])
        self.assertEqual(payload["worksheets"][0]["name"], "Data")
        self.assertEqual(payload["worksheets"][0]["preview"], [["Name", "Score", "RichText"]])
        self.assertTrue(payload["worksheets"][0]["truncated"])

    def test_row_preview_is_capped_at_twenty_rows(self):
        path = self.path("long.xlsx")
        rows = [[i, "row-{}".format(i)] for i in range(1, 26)]
        support.write_xlsx(path, [{"name": "Sheet1", "dimension": "A1:B25", "rows": rows}])
        sheet = run_runner(["--inspect", path]).payload["worksheets"][0]
        self.assertEqual(len(sheet["preview"]), 20)
        self.assertTrue(sheet["truncated"])

    def test_absolute_relationship_target_resolves(self):
        path = self.path("absolute.xlsx")
        support.write_xlsx(
            path,
            [
                {
                    "name": "Only",
                    "dimension": "A1:A1",
                    "rows": [["absolute"]],
                    "rel_target": "/xl/worksheets/sheet1.xml",
                }
            ],
        )
        payload = run_runner(["--inspect", path]).payload
        self.assertTrue(payload["success"])
        self.assertEqual(payload["worksheets"][0]["preview"], [["absolute"]])

    def test_missing_worksheet_part_yields_empty_preview(self):
        path = self.path("dangling.xlsx")
        support.write_xlsx(
            path,
            [{"name": "Ghost", "dimension": "A1:B2", "rows": [[1, 2]], "rel_target": "worksheets/missing.xml"}],
        )
        sheet = run_runner(["--inspect", path]).payload["worksheets"][0]
        self.assertEqual(sheet["name"], "Ghost")
        self.assertIsNone(sheet["used_range"])
        self.assertEqual(sheet["preview"], [])
        self.assertFalse(sheet["truncated"])

    def test_shared_strings_absent_falls_back_to_raw_index(self):
        path = self.path("nostrings.xlsx")
        support.write_package(
            path,
            {
                "xl/workbook.xml": (
                    support.XML_DECL
                    + '<workbook xmlns="{s}" xmlns:r="{r}"><sheets>'
                    '<sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'.format(
                        s=support.S_NS, r=support.R_NS
                    )
                ),
                "xl/_rels/workbook.xml.rels": (
                    support.XML_DECL
                    + '<Relationships xmlns="{p}"><Relationship Id="rId1" Type="worksheet" '
                    'Target="worksheets/sheet1.xml"/></Relationships>'.format(p=support.PKG_REL_NS)
                ),
                "xl/worksheets/sheet1.xml": (
                    support.XML_DECL
                    + '<worksheet xmlns="{s}"><dimension ref="A1"/><sheetData>'
                    '<row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>'.format(
                        s=support.S_NS
                    )
                ),
            },
        )
        sheet = run_runner(["--inspect", path]).payload["worksheets"][0]
        self.assertEqual(sheet["preview"], [["0"]])


class InspectPptxTests(MacOfficeTestCase):
    TITLE = {
        "kind": "sp",
        "name": "Title 1",
        "id": 2,
        "x": 914400,
        "y": 457200,
        "cx": 8229600,
        "cy": 1143000,
        "text": "Quarterly Deck",
    }
    PICTURE = {
        "kind": "pic",
        "name": "Picture 3",
        "id": 5,
        "x": 127000,
        "y": 254000,
        "cx": 508000,
        "cy": 762000,
    }
    TABLE = {
        "kind": "graphicFrame",
        "name": "Table 4",
        "id": 7,
        "x": 0,
        "y": 0,
        "cx": 914400,
        "cy": 914400,
        "cells": [["Cell A", "Cell B"]],
    }
    BODY = {"kind": "sp", "name": "Body 5", "id": 8, "text": "Only body text"}

    def build(self, name="deck.pptx"):
        path = self.path(name)
        support.write_pptx(path, [[self.TITLE, self.PICTURE, self.TABLE], [self.BODY]])
        return path

    def test_slide_shape_geometry(self):
        result = run_runner(["--inspect", self.build()])
        self.assertEqual(result.returncode, 0, result.stderr)
        payload = result.payload
        self.assertTrue(payload["success"])
        self.assertEqual(payload["kind"], "pptx")
        self.assertEqual(payload["inspector"], "headless_openxml")
        self.assertEqual(payload["slide_count"], 2)
        self.assertEqual([slide["index"] for slide in payload["slides"]], [1, 2])

        self.assertFalse(payload["slides_truncated"])
        first, second = payload["slides"]
        self.assertEqual(first["shape_count"], 3)
        self.assertFalse(first["shapes_truncated"])
        self.assertEqual([shape["kind"] for shape in first["shapes"]], ["sp", "pic", "graphicFrame"])
        title, picture, table = first["shapes"]
        self.assertEqual(
            title,
            {
                "index": 1,
                "id": 2,
                "name": "Title 1",
                "kind": "sp",
                "left_pt": 72.0,
                "top_pt": 36.0,
                "width_pt": 648.0,
                "height_pt": 90.0,
                "text": "Quarterly Deck",
            },
        )
        self.assertEqual(
            (picture["left_pt"], picture["top_pt"], picture["width_pt"], picture["height_pt"]),
            (10.0, 20.0, 40.0, 60.0),
        )
        self.assertEqual(picture["text"], "")
        # <p:graphicFrame> keeps its geometry on <p:xfrm>, which the headless inspector
        # does not read, so table frames report null point geometry.
        self.assertEqual(table["kind"], "graphicFrame")
        self.assertIsNone(table["left_pt"])
        self.assertIsNone(table["top_pt"])
        self.assertIsNone(table["width_pt"])
        self.assertIsNone(table["height_pt"])
        self.assertEqual(table["text"], "Cell A Cell B")

        self.assertEqual(second["shape_count"], 1)
        body = second["shapes"][0]
        self.assertEqual(body["text"], "Only body text")
        self.assertEqual(
            [body["left_pt"], body["top_pt"], body["width_pt"], body["height_pt"]],
            [None, None, None, None],
        )

    def test_shape_text_trims_at_120_chars(self):
        path = self.path("longtext.pptx")
        shape = dict(self.TITLE, text="z" * 300)
        support.write_pptx(path, [[shape]])
        text = run_runner(["--inspect", path]).payload["slides"][0]["shapes"][0]["text"]
        self.assertEqual(len(text), 120)
        self.assertTrue(text.endswith("…"))

    def test_max_items_truncates_slides_and_shapes(self):
        payload = run_runner(["--inspect", self.build(), "--max-items", "1"]).payload
        self.assertEqual(payload["slide_count"], 2)
        self.assertEqual(len(payload["slides"]), 1)
        self.assertEqual(payload["slides"][0]["shape_count"], 3)
        self.assertEqual(len(payload["slides"][0]["shapes"]), 1)
        self.assertTrue(payload["slides"][0]["shapes_truncated"])

    def test_slides_without_id_list_fall_back_to_numeric_order(self):
        path = self.path("unordered.pptx")
        slides = [[dict(self.BODY, text="Slide {}".format(n))] for n in (1, 2, 10, 3)]
        support.write_pptx(
            path,
            slides,
            use_slide_id_list=False,
            slide_file_names=[
                "ppt/slides/slide1.xml",
                "ppt/slides/slide2.xml",
                "ppt/slides/slide10.xml",
                "ppt/slides/slide3.xml",
            ],
        )
        payload = run_runner(["--inspect", path]).payload
        self.assertEqual(payload["slide_count"], 4)
        self.assertEqual(
            [slide["shapes"][0]["text"] for slide in payload["slides"]],
            ["Slide 1", "Slide 2", "Slide 3", "Slide 10"],
        )


class InspectErrorTests(MacOfficeTestCase):
    def test_missing_file(self):
        missing = self.path("nope.docx")
        result = run_runner(["--inspect", missing])
        self.assertEqual(result.returncode, 1)
        payload = result.payload
        self.assertFalse(payload["success"])
        self.assertEqual(payload["mode"], "inspect")
        self.assertIsNone(payload["kind"])
        self.assertEqual(payload["error"], "File not found: {}".format(os.path.abspath(missing)))

    def test_unsupported_extension(self):
        path = self.write(self.path("notes.txt"), "plain text")
        result = run_runner(["--inspect", path])
        self.assertEqual(result.returncode, 1)
        self.assertEqual(
            result.payload["error"], "Unsupported file type. Expected .docx, .xlsx, or .pptx."
        )
        self.assertFalse(result.payload["success"])

    def test_corrupt_zip(self):
        path = self.write(self.path("broken.docx"), b"this is definitely not a zip archive")
        result = run_runner(["--inspect", path])
        self.assertEqual(result.returncode, 1)
        self.assertFalse(result.payload["success"])
        self.assertTrue(
            result.payload["error"].startswith("Failed to parse docx package:"),
            result.payload["error"],
        )

    def test_zip_missing_required_part(self):
        path = self.path("empty.docx")
        support.write_package(path, {"[Content_Types].xml": "<Types/>"})
        result = run_runner(["--inspect", path])
        self.assertEqual(result.returncode, 1)
        self.assertEqual(
            result.payload["error"], "Not a valid docx package: missing word/document.xml"
        )

        book = self.path("empty.xlsx")
        support.write_package(book, {"xl/sharedStrings.xml": "<sst xmlns='x'/>"})
        result = run_runner(["--inspect", book])
        self.assertEqual(
            result.payload["error"], "Not a valid xlsx package: missing xl/workbook.xml"
        )

        deck = self.path("empty.pptx")
        support.write_package(deck, {"ppt/_rels/presentation.xml.rels": "<Relationships/>"})
        result = run_runner(["--inspect", deck])
        self.assertEqual(
            result.payload["error"], "Not a valid pptx package: missing ppt/presentation.xml"
        )

    def test_malformed_required_part(self):
        path = self.path("malformed.docx")
        support.write_package(path, {"word/document.xml": "<w:document><w:body>"})
        result = run_runner(["--inspect", path])
        self.assertEqual(result.returncode, 1)
        self.assertFalse(result.payload["success"])
        self.assertTrue(
            result.payload["error"].startswith("Failed to parse docx package:"),
            result.payload["error"],
        )


if __name__ == "__main__":
    unittest.main()
