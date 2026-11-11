"""Unit tests: container detection and the native-app AppleScript fallback.

The fake ``osascript`` answers in the fallback's exact control-character protocol (field
``\\x1f``, record ``\\x1d``, section ``\\x1e``), so routing, parsing, degradation and the
``truncated`` contract are exercised without Microsoft Office.
"""

import re
import unittest

import support
from support import ShimTestCase, constants, staging_dir

FIELD = constants().FIELD_SEPARATOR
RECORD = constants().RECORD_SEPARATOR
SECTION = constants().SECTION_SEPARATOR


def encode(*sections):
    """Join sections, each a list of records, each a list of fields."""
    encoded = [RECORD.join(FIELD.join(str(value) for value in record) for record in section) for section in sections]
    return SECTION.join(encoded)


# The fake osascript replaces this with the HFS path (as Word reports ``full name``) of the
# staged copy the inspect script opened.
STAGED = "@STAGED_HFS@"


def docx_payload(opened=STAGED):
    return encode(
        [[opened, 5, 1, 1, 1, 612, 792, 72, 72, 72, 72]],
        [
            [1, 0, 11, "align paragraph left", 12, "First paragraph\r"],
            [2, 11, 22, "align paragraph center", 14, "Second paragraph\r"],
            [3, 22, 25, "align paragraph left", 12, "Header cell\r\x07"],
            [4, 25, 28, "align paragraph center", 12, "Second header\r\x07"],
            [5, 34, 40, "align paragraph left", 12, "季度報告 — α β ✓\r"],
        ],
        [
            [1, 1, 1, 22, 25, "Header cell\r\x07"],
            [1, 1, 2, 25, 28, "Second header\r\x07"],
            [1, 2, 1, 28, 31, "Body text\r\x07"],
            [1, 2, 2, 31, 34, "Body tail\r\x07"],
        ],
        [[1, 2, 2]],

        [[1, "Text Box 2", "shape type text box", 128, 303.05, 150, 60, "Boxed Text\r"]],
    )


def xlsx_payload(opened=STAGED, sheets=None, cells=None):
    sheets = sheets if sheets is not None else [
        ["Sheet1", "A1:B2", 2, 2],
        ["Data", "A1:C3", 3, 3],
    ]
    cells = cells if cells is not None else [
        [1, 1, 1, "value A"],
        [1, 1, 2, "value B"],
        [1, 2, 1, "alpha"],
        [1, 2, 2, 42],
        [2, 1, 1, ""],
        [2, 1, 2, "kept"],
    ]
    return encode([[opened, 2]], sheets, cells)


def pptx_payload(opened=STAGED):
    return encode(
        [[opened, 2]],
        [[1, 2], [2, 0]],
        [
            [1, 1, "Title 1", 50, 30, 400, 60, "title text\r"],
            [1, 2, "Box 2", 10, 20, 100, 50, "box text\r"],
        ],
    )


class WordFallbackTests(ShimTestCase):
    def inspect(self, name="encrypted.docx", *extra, stdout=None, exit_code=0, stderr=""):
        stdout = docx_payload() if stdout is None else stdout
        path = support.write_ole2_file(self.path(name))
        result = self.invoke(
            "--inspect",
            path,
            *extra,
            shim_env={
                "MAC_OFFICE_FAKE_INSPECT_STDOUT": stdout,
                "MAC_OFFICE_FAKE_INSPECT_EXIT": str(exit_code),
                "MAC_OFFICE_FAKE_STDERR": stderr,
            },
        )
        self.assertIsNotNone(result.payload, result.stderr)
        return result

    def test_ole2_container_uses_applescript_fallback(self):
        result = self.inspect()
        self.assertEqual(result.returncode, 0, result.stderr)
        payload = result.payload
        self.assertTrue(payload["success"])
        self.assertEqual(payload["kind"], "docx")
        self.assertEqual(payload["container"], "ole2")
        self.assertEqual(payload["inspector"], "applescript_fallback")
        self.assertIn("unzip, LibreOffice and textutil cannot", payload["hint"])
        self.assertIsNone(payload["error"])
        self.assertEqual(payload["paragraph_count"], 5)
        self.assertEqual(payload["table_count"], 1)
        self.assertEqual(payload["shape_count"], 1)
        self.assertEqual(payload["inline_shape_count"], 1)

    def test_non_ascii_text_survives_the_transport(self):
        paragraphs = self.inspect().payload["paragraphs"]
        self.assertEqual(paragraphs[4]["text"], "季度報告 — α β ✓")

    def test_paragraph_fields_include_style_and_position(self):
        payload = self.inspect().payload
        self.assertEqual(
            payload["paragraphs"][0],
            {
                "index": 1,
                "text": "First paragraph",
                "alignment": "align paragraph left",
                "font_size": 12.0,
                "table": None,
                "row": None,
                "column": None,
            },
        )
        self.assertEqual(payload["paragraphs"][1]["alignment"], "align paragraph center")
        self.assertEqual(payload["paragraphs"][1]["font_size"], 14.0)

    def test_paragraph_inside_a_cell_reports_table_row_and_column(self):
        paragraphs = self.inspect().payload["paragraphs"]
        self.assertEqual(
            [(paragraph["table"], paragraph["row"], paragraph["column"]) for paragraph in paragraphs[2:4]],
            [(1, 1, 1), (1, 1, 2)],
        )
        # A paragraph after the table keeps null mapping.
        self.assertEqual(
            (paragraphs[4]["table"], paragraphs[4]["row"], paragraphs[4]["column"]),
            (None, None, None),
        )
        self.assertEqual(paragraphs[2]["text"], "Header cell")

    def test_body_paragraphs_are_not_mapped_to_a_table(self):
        paragraphs = self.inspect().payload["paragraphs"]
        self.assertEqual((paragraphs[0]["table"], paragraphs[0]["row"], paragraphs[0]["column"]), (None, None, None))

    def test_tables_report_grid_and_preview(self):
        table = self.inspect().payload["tables"][0]
        self.assertEqual((table["rows"], table["columns"]), (2, 2))
        self.assertEqual(table["preview"], [["Header cell", "Second header"], ["Body text", "Body tail"]])
        self.assertFalse(table["cells_truncated"])

    def test_shapes_and_page_setup_are_reported(self):
        result = self.inspect().payload
        self.assertEqual(
            result["shapes"][0],
            {
                "index": 1,
                "name": "Text Box 2",
                "type": "shape type text box",
                "left_pt": 128.0,
                "top_pt": 303.05,
                "width_pt": 150.0,
                "height_pt": 60.0,
                "text": "Boxed Text",
            },
        )
        self.assertEqual(
            result["page_setup"],
            {
                "page_width": 612.0,
                "page_height": 792.0,
                "top_margin": 72.0,
                "bottom_margin": 72.0,
                "left_margin": 72.0,
                "right_margin": 72.0,
            },
        )

    def test_max_items_reports_truncation_per_collection(self):
        payload = self.inspect("encrypted.docx", "--max-items", "3").payload
        self.assertEqual(payload["paragraph_count"], 5)
        self.assertEqual(len(payload["paragraphs"]), 3)
        self.assertTrue(payload["paragraphs_truncated"])
        self.assertFalse(payload["tables_truncated"])

    def test_fallback_queries_word_through_osascript(self):
        self.inspect()
        script = self.shim.last_script()
        self.assertIn('tell application "Microsoft Word"', script)
        self.assertIn("start of content of", script)
        # The staged copy inside the Office group container is what Word is asked to open.
        self.assertIn(staging_dir(), script)

    def test_fallback_failure_reports_error(self):
        result = self.inspect(
            "encrypted.docx",
            stdout="",
            exit_code=1,
            stderr="execution error: Word got an error: boom (-1728)",
        )
        self.assertEqual(result.returncode, 1)
        payload = result.payload
        self.assertFalse(payload["success"])
        self.assertEqual(
            payload["error"],
            "File is an ole2 docx container and the Microsoft Word AppleScript fallback failed: "
            "execution error: Word got an error: boom (-1728)",
        )
        self.assertIsNone(payload["inspector"])

    def test_payload_with_the_wrong_document_is_rejected(self):
        """The app can silently open a different document; that must never look successful."""
        wrong = docx_payload("Macintosh HD:Users:someone:Documents:encrypted.docx")
        result = self.inspect("encrypted.docx", stdout=wrong)
        self.assertFalse(result.payload["success"])
        self.assertIn("did not open", result.payload["error"])

    def test_open_swallowed_by_the_app_is_reported(self):
        result = self.inspect("encrypted.docx", stdout="mac-office-open-failed")
        self.assertFalse(result.payload["success"])
        self.assertIn("could not open the file", result.payload["error"])

    def test_mip_container_is_named_in_the_report(self):
        path = support.write_cfb(
            self.path("labelled.docx"), {"EncryptedPackage": b"\x01" * 700, "LabelInfo": b"label"}
        )
        payload = self.invoke(
            "--inspect", path, shim_env={"MAC_OFFICE_FAKE_INSPECT_STDOUT": docx_payload()}
        ).payload
        self.assertEqual(payload["container"], "ole2_mip")
        self.assertIn("EncryptedPackage", payload["hint"])
        self.assertTrue(payload["success"])

    def test_unreadable_container_directory_falls_back_to_a_byte_scan(self):
        path = self.write(
            self.path("opaque.docx"),
            support.OLE2_MAGIC + b"\x00" * 64 + b"EncryptedPackage" + b"\x00" * 64,
        )
        payload = self.invoke(
            "--inspect", path, shim_env={"MAC_OFFICE_FAKE_INSPECT_STDOUT": docx_payload()}
        ).payload
        self.assertEqual(payload["container"], "ole2_mip")

    def test_legacy_payload_is_staged_with_its_real_extension(self):
        """Word refuses a Word 97 payload named .docx, so the staged copy becomes .doc."""
        path = support.write_cfb(
            self.path("legacy.docx"), {"WordDocument": b"\x01" * 700, "1Table": b"\x02" * 300}
        )
        payload = self.invoke(
            "--inspect", path, shim_env={"MAC_OFFICE_FAKE_INSPECT_STDOUT": docx_payload()}
        ).payload
        self.assertEqual(payload["container"], "ole2")
        self.assertTrue(payload["success"], payload["error"])
        script = self.shim.last_script()
        self.assertRegex(script, r'legacy-[0-9a-f]{8}\.doc"')
        self.assertNotRegex(script, r'legacy-[0-9a-f]{8}\.docx"')


class FallbackSafetyTests(ShimTestCase):
    """R1: the scripts may only ever read or close the staged copy they opened."""

    def inspect(self, name, stdout):
        path = support.write_ole2_file(self.path(name))
        return self.invoke("--inspect", path, shim_env={"MAC_OFFICE_FAKE_INSPECT_STDOUT": stdout}).payload

    def test_scripts_look_up_the_staged_copy_by_its_unique_name(self):
        for name, item, payload in (
            ("a.docx", "document", docx_payload()),
            ("b.xlsx", "workbook", xlsx_payload()),
            ("c.pptx", "presentation", pptx_payload()),
        ):
            with self.subTest(name=name):
                self.assertTrue(self.inspect(name, payload)["success"])
                script = self.shim.last_script()
                self.assertNotIn("active " + item, script)
                staged = re.search(r'set \w+ to {} "([^"]+)"'.format(item), script).group(1)
                self.assertRegex(staged, r"^[abc]-[0-9a-f]{8}\.(docx|xlsx|pptx)$")
                self.assertIn('if not (exists {} "{}") then return "mac-office-open-failed"'.format(item, staged), script)

    def test_inspect_opens_read_only_without_touching_recent_files(self):
        self.inspect("a.docx", docx_payload())
        self.assertIn("read only true add to recent files false", self.shim.last_script())
        self.inspect("b.xlsx", xlsx_payload())
        self.assertIn("read only true add to mru false", self.shim.last_script())

    def test_swallowed_open_only_ever_closes_the_staged_copy(self):
        payload = self.inspect("a.docx", "mac-office-open-failed")
        self.assertFalse(payload["success"])
        self.assertIn("could not open the file", payload["error"])
        close_scripts = [script for script in self.shim.e_scripts() if "close document i" in script]
        self.assertEqual(len(close_scripts), 1)
        wanted = re.search(r"set wanted to \{(.*)\}", close_scripts[0]).group(1)
        self.assertRegex(wanted, r'^"{}/[^"]+/a-[0-9a-f]{{8}}\.docx"$'.format(re.escape(staging_dir())))

    def test_fallback_gets_at_least_sixty_seconds_whatever_timeout_says(self):
        """R5: --timeout 5 must not cut an inspect that takes longer than 5 s."""
        path = support.write_ole2_file(self.path("slow.docx"))
        payload = self.invoke(
            "--inspect",
            path,
            "--timeout",
            "5",
            shim_env={"MAC_OFFICE_FAKE_INSPECT_STDOUT": docx_payload(), "MAC_OFFICE_FAKE_SLEEP": "7"},
        ).payload
        self.assertTrue(payload["success"], payload["error"])


class UnreadableTests(ShimTestCase):
    """R6: ``_truncated`` means capped by --max-items; unreadable items are counted apart."""

    def test_failed_reads_are_counted_and_do_not_mark_truncation(self):
        payload_text = encode(
            [[STAGED, 3, 1, 2, 0]],
            [[1, 0, 5, "align paragraph left", 12, "Fine\r"], [2, "", "", "", "", ""], [3, 5, 9, "", 12, "Also\r"]],
            [[1, 1, 1, 9, 12, "Cell\r\x07"]],
            [[1, 1, 2]],
            [[1, "Box", "shape type text box", 1, 2, 3, 4, "t"], [2, "", "", "", "", "", "", ""]],
        )
        path = support.write_ole2_file(self.path("partly.docx"))
        payload = self.invoke("--inspect", path, shim_env={"MAC_OFFICE_FAKE_INSPECT_STDOUT": payload_text}).payload
        self.assertTrue(payload["success"], payload["error"])
        self.assertFalse(payload["paragraphs_truncated"])
        self.assertEqual(payload["paragraphs_unreadable"], 1)
        self.assertFalse(payload["shapes_truncated"])
        self.assertEqual(payload["shapes_unreadable"], 1)
        self.assertFalse(payload["tables"][0]["cells_truncated"])
        self.assertEqual(payload["tables"][0]["cells_unreadable"], 1)

    def test_capped_collection_is_truncated_but_not_unreadable(self):
        path = support.write_ole2_file(self.path("capped.docx"))
        payload = self.invoke(
            "--inspect", path, "--max-items", "3", shim_env={"MAC_OFFICE_FAKE_INSPECT_STDOUT": docx_payload()}
        ).payload
        self.assertTrue(payload["paragraphs_truncated"])
        self.assertEqual(payload["paragraphs_unreadable"], 0)


class ExcelFallbackTests(ShimTestCase):
    def inspect(self, sheets=None, name="encrypted.xlsx", *extra, cells=None):
        path = support.write_ole2_file(self.path(name))
        result = self.invoke(
            "--inspect",
            path,
            *extra,
            shim_env={"MAC_OFFICE_FAKE_INSPECT_STDOUT": xlsx_payload(STAGED, sheets, cells)},
        )
        self.assertIsNotNone(result.payload, result.stderr)
        return result.payload

    def test_ole2_workbook_uses_applescript_fallback(self):
        payload = self.inspect()
        self.assertTrue(payload["success"])
        self.assertEqual(payload["kind"], "xlsx")
        self.assertEqual(payload["inspector"], "applescript_fallback")
        self.assertEqual(payload["sheet_count"], 2)
        self.assertFalse(payload["sheets_truncated"])
        first = payload["worksheets"][0]
        self.assertEqual(first["name"], "Sheet1")
        self.assertEqual(first["used_range"], "A1:B2")
        self.assertEqual(first["preview"], [["value A", "value B"], ["alpha", "42"]])
        self.assertFalse(first["truncated"])
        self.assertEqual(payload["worksheets"][1]["used_range"], "A1:C3")

    def test_sheet_with_an_empty_first_cell_survives(self):
        """A worksheet whose A1 is empty must not disappear from the result."""
        worksheets = self.inspect()["worksheets"]
        self.assertEqual([sheet["name"] for sheet in worksheets], ["Sheet1", "Data"])
        self.assertEqual(worksheets[1]["preview"][0], ["", "kept", ""])

    def test_truncated_means_the_preview_was_capped(self):
        narrow = self.inspect([["One", "A1", 1, 1]], "narrow.xlsx", cells=[[1, 1, 1, "solo"]])
        self.assertEqual(narrow["worksheets"][0]["preview"], [["solo"]])
        self.assertFalse(narrow["worksheets"][0]["truncated"])

        tall = self.inspect([["Tall", "A1:A12", 12, 1]], "tall.xlsx")
        self.assertEqual(len(tall["worksheets"][0]["preview"]), 10)
        self.assertTrue(tall["worksheets"][0]["truncated"])

        wide = self.inspect([["Wide", "A1:Y1", 1, 25]], "wide.xlsx")
        self.assertEqual(len(wide["worksheets"][0]["preview"][0]), 20)
        self.assertTrue(wide["worksheets"][0]["truncated"])

    def test_sheets_truncated_follows_max_items(self):
        payload = self.inspect(None, "trimmed.xlsx", "--max-items", "1")
        self.assertEqual(payload["sheet_count"], 2)
        self.assertEqual(len(payload["worksheets"]), 1)
        self.assertTrue(payload["sheets_truncated"])


class PptxFallbackTests(ShimTestCase):
    def test_ole2_deck_uses_applescript_fallback(self):
        path = support.write_ole2_file(self.path("encrypted.pptx"))
        payload = self.invoke(
            "--inspect", path, shim_env={"MAC_OFFICE_FAKE_INSPECT_STDOUT": pptx_payload()}
        ).payload
        self.assertTrue(payload["success"])
        self.assertEqual(payload["kind"], "pptx")
        self.assertEqual(payload["inspector"], "applescript_fallback")
        self.assertEqual(payload["slide_count"], 2)
        self.assertFalse(payload["slides_truncated"])
        self.assertEqual(
            payload["slides"][0],
            {
                "index": 1,
                "shape_count": 2,
                "shapes_truncated": False,
                "shapes_unreadable": 0,
                "shapes": [
                    {
                        "index": 1,
                        "id": None,
                        "name": "Title 1",
                        "kind": "sp",
                        "left_pt": 50.0,
                        "top_pt": 30.0,
                        "width_pt": 400.0,
                        "height_pt": 60.0,
                        "text": "title text",
                    },
                    {
                        "index": 2,
                        "id": None,
                        "name": "Box 2",
                        "kind": "sp",
                        "left_pt": 10.0,
                        "top_pt": 20.0,
                        "width_pt": 100.0,
                        "height_pt": 50.0,
                        "text": "box text",
                    },
                ],
            },
        )
        self.assertEqual(
            payload["slides"][1],
            {"index": 2, "shape_count": 0, "shapes_truncated": False, "shapes_unreadable": 0, "shapes": []},
        )

    def test_shapes_truncated_follows_max_items(self):
        path = support.write_ole2_file(self.path("trimmed.pptx"))
        payload = self.invoke(
            "--inspect",
            path,
            "--max-items",
            "1",
            shim_env={"MAC_OFFICE_FAKE_INSPECT_STDOUT": pptx_payload()},
        ).payload
        self.assertEqual(payload["slides"][0]["shape_count"], 2)
        self.assertEqual(len(payload["slides"][0]["shapes"]), 1)
        self.assertTrue(payload["slides"][0]["shapes_truncated"])
        self.assertTrue(payload["slides_truncated"])


class FallbackRoutingTests(ShimTestCase):
    def test_ole2_container_with_an_embedded_zip_still_routes_to_the_app(self):
        """A Word 97 / MIP container embeds a readable theme zip.

        ``zipfile`` therefore opens it and the required-part check would report
        "missing word/document.xml"; the OLE2 magic must win instead.
        """
        inner = self.path("inner.zip")
        support.write_package(inner, {"word/theme/theme1.xml": "<a:theme xmlns:a='x'/>"})
        with open(inner, "rb") as handle:
            zip_bytes = handle.read()
        path = self.write(self.path("embedded.docx"), support.OLE2_MAGIC + zip_bytes)

        result = self.invoke(
            "--inspect", path, shim_env={"MAC_OFFICE_FAKE_INSPECT_STDOUT": docx_payload()}
        )
        payload = result.payload
        self.assertTrue(payload["success"], payload.get("error"))
        self.assertEqual(payload["container"], "ole2")
        self.assertEqual(payload["inspector"], "applescript_fallback")
        self.assertIn("Microsoft Word", self.shim.last_script())

    def test_plain_corrupt_zip_never_invokes_osascript(self):
        path = self.write(self.path("broken.docx"), b"not a zip and not ole2")
        result = self.invoke("--inspect", path)
        payload = result.payload
        self.assertFalse(payload["success"])
        self.assertTrue(payload["error"].startswith("Failed to parse docx package:"))
        self.assertEqual(payload["container"], "openxml")
        self.assertIsNone(payload["hint"])
        self.assertEqual(self.shim.calls(), [])

    def test_ole2_with_unsupported_extension_is_rejected(self):
        path = support.write_ole2_file(self.path("legacy.doc"))
        result = self.invoke("--inspect", path)
        self.assertEqual(
            result.payload["error"], "Unsupported file type. Expected .docx, .xlsx, or .pptx."
        )
        self.assertEqual(self.shim.calls(), [])

    def test_headless_success_never_invokes_osascript(self):
        path = self.path("plain.docx")
        support.write_docx(path, [("p", "hello", None)])
        result = self.invoke("--inspect", path)
        self.assertTrue(result.payload["success"])
        self.assertEqual(result.payload["inspector"], "headless_openxml")
        self.assertEqual(self.shim.calls(), [])


if __name__ == "__main__":
    unittest.main()
