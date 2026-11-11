"""Integration tests against real Microsoft Word (gated by MAC_OFFICE_INTEGRATION=1).

Fixtures are generated at test time through ``office_runner.py``; no binary artefacts are
committed and no user document is touched.
"""

import os
import shutil
import unittest

import support
from support import IntegrationCase

CREATE_DOCX = """
tell application "Microsoft Word"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	make new document
	set content of text object of paragraph 1 of active document to "Integration Alpha"
	make new paragraph at end of active document
	set content of text object of last paragraph of active document to "Integration Beta"
	make new paragraph at end of active document
	set theTable to make new table at text object of last paragraph of active document with properties {number of rows:2, number of columns:2}
	set content of text object of cell 1 of row 1 of theTable to "Header A"
	set content of text object of cell 2 of row 1 of theTable to "Header B"
	save as active document file name outPath file format format document
	close active document saving no
end tell
"""

CREATE_RICH_DOCX = """
tell application "Microsoft Word"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set imgPath to system attribute "MAC_OFFICE_ASSET_1"
	make new document
	set content of text object of paragraph 1 of active document to "Body Alpha" & return & "Body Beta" & return
	set theTable to make new table at text object of paragraph (count of paragraphs of active document) of active document with properties {number of rows:2, number of columns:2}
	set content of text object of cell 1 of row 1 of theTable to "H1"
	set content of text object of cell 2 of row 1 of theTable to "H2"
	set content of text object of cell 1 of row 2 of theTable to "B1"
	set content of text object of cell 2 of row 2 of theTable to "B2"
	make new paragraph at end of active document
	set content of text object of last paragraph of active document to "After table" & return
	set thePicture to make new inline picture at end of text object of last paragraph of active document with properties {file name:imgPath}
	set theBox to make new text box at text object of paragraph 1 of active document with properties {left position:200, top:400, width:150, height:60}
	set content of text range of text frame of theBox to "Boxed Text"
	save as active document file name outPath file format format document
	close active document saving no
end tell
"""

FIND_AND_REPLACE = """
tell application "Microsoft Word"
	set targetPath to system attribute "MAC_OFFICE_TARGET_FILE"
	open (targetPath as POSIX file)
	execute find (find object of text object of active document) find text "Integration Alpha" replace with "Edited Alpha" replace replace all
	save active document
	close active document saving no
end tell
"""

PDF_EXPORT = """
tell application "Microsoft Word"
	set targetPath to system attribute "MAC_OFFICE_TARGET_FILE"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	open (targetPath as POSIX file)
	save as active document file name outPath file format format PDF
	close active document saving no
end tell
"""

CREATE_WORD_97 = """
tell application "Microsoft Word"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	make new document
	set content of text object of paragraph 1 of active document to "Legacy Alpha"
	make new paragraph at end of active document
	set content of text object of last paragraph of active document to "Legacy Beta"
	make new paragraph at end of active document
	set theTable to make new table at text object of last paragraph of active document with properties {number of rows:2, number of columns:2}
	set content of text object of cell 1 of row 1 of theTable to "Legacy A"
	save as active document file name outPath file format format document97
	close active document saving no
end tell
"""


class WordIntegrationTests(IntegrationCase):
    app = "word"

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.base = cls.artifact("base.docx")
        cls.create_result = cls.cli(
            "--app", "word", "--output-file", cls.base, stdin_text=CREATE_DOCX, timeout=240
        )
        if not (cls.create_result.payload and cls.create_result.payload["success"]):
            raise AssertionError("Word could not create the fixture: %r" % (cls.create_result,))

    def inspect(self, path, *extra):
        result = self.cli("--inspect", path, *extra, timeout=120)
        self.assertIsNotNone(result.payload, result.stderr)
        self.assertTrue(result.payload["success"], result.payload.get("error"))
        return result.payload

    def render(self, path, name="render"):
        result = self.cli(
            "--render", path, "--output-dir", self.artifact(name), "--timeout", "120", timeout=300
        )
        self.assertIsNotNone(result.payload, result.stderr)
        return result.payload

    def test_created_document_has_expected_structure(self):
        payload = self.create_result.payload
        self.assertTrue(payload["output_exists"])
        self.assertTrue(payload["output_modified"])
        self.assertTrue(os.path.isfile(self.base))

        inspected = self.inspect(self.base)
        self.assertEqual(inspected["kind"], "docx")
        if support.is_ole2_container(self.base):
            # Office on this machine wraps its output in an OLE2 container, so inspection
            # goes through the AppleScript fallback rather than headless zip parsing.
            self.assertEqual(inspected["inspector"], "applescript_fallback")
            self.assertIn(inspected["container"], ("ole2", "ole2_mip"))
            self.assertIsNotNone(inspected["hint"])
        else:
            self.assertEqual(inspected["inspector"], "headless_openxml")
            self.assertEqual(inspected["container"], "openxml")
        self.assertGreaterEqual(inspected["paragraph_count"], 2)
        texts = [paragraph["text"] for paragraph in inspected["paragraphs"]]
        self.assertTrue(any("Integration Alpha" in text for text in texts), texts)
        self.assertTrue(any("Integration Beta" in text for text in texts), texts)
        self.assertEqual(inspected["table_count"], 1)
        self.assertEqual(inspected["tables"][0]["rows"], 2)
        self.assertEqual(inspected["tables"][0]["columns"], 2)

    def test_target_file_edit_creates_revised_copy(self):
        working = self.path("edit.docx")
        shutil.copy2(self.base, working)
        before = self.inspect(working)
        result = self.cli(
            "--app", "word", "--target-file", working, stdin_text=FIND_AND_REPLACE, timeout=240
        )
        payload = result.payload
        self.assertTrue(payload["success"], payload.get("error"))
        revised = self.path("edit_revised.docx")
        self.assertEqual(payload["target_file"], working)
        self.assertEqual(payload["working_file"], revised)
        self.assertTrue(payload["target_modified"])
        self.assertTrue(payload["output_exists"])
        self.assertTrue(payload["output_modified"])
        self.assertFalse(os.path.exists(working + ".bak"))

        edited = self.inspect(revised)
        self.assertEqual(edited["paragraph_count"], before["paragraph_count"])
        self.assertEqual(edited["table_count"], before["table_count"])
        edited_texts = [paragraph["text"] for paragraph in edited["paragraphs"]]
        self.assertIn("Edited Alpha", edited_texts)
        self.assertFalse(any("Integration Alpha" in text for text in edited_texts), edited_texts)

        original = self.inspect(working)
        original_texts = [paragraph["text"] for paragraph in original["paragraphs"]]
        self.assertIn("Integration Alpha", original_texts)
        self.assertFalse(any("Edited Alpha" in text for text in original_texts), original_texts)

    def test_pdf_export(self):
        pdf = self.path("export", "report.pdf")
        result = self.cli(
            "--app",
            "word",
            "--target-file",
            self.base,
            "--output-file",
            pdf,
            stdin_text=PDF_EXPORT,
            timeout=240,
        )
        payload = result.payload
        self.assertTrue(payload["success"], payload.get("error"))
        self.assertTrue(payload["output_exists"])
        self.assertTrue(payload["output_modified"])
        self.assertTrue(os.path.getsize(pdf) > 0)
        with open(pdf, "rb") as handle:
            self.assertEqual(handle.read(5), b"%PDF-")

    def test_word97_document_renamed_to_docx_uses_the_fallback(self):
        """A Word 97 container renamed to .docx must reach Word, not the zip parser.

        The container starts with the OLE2 magic and also carries an embedded theme zip, so
        a zip-first reader would report "Not a valid docx package: missing word/document.xml".
        Word refuses a Word 97 payload named .docx (format/extension mismatch), so the staged
        copy is renamed to .doc before it is opened.
        """
        legacy_doc = self.path("legacy.doc")
        result = self.cli(
            "--app", "word", "--output-file", legacy_doc, stdin_text=CREATE_WORD_97, timeout=240
        )
        self.assertTrue(result.payload["success"], result.payload.get("error"))
        self.assertTrue(support.is_ole2_container(legacy_doc))

        legacy_docx = self.path("legacy.docx")
        shutil.copy2(legacy_doc, legacy_docx)
        self.assertTrue(support.is_ole2_container(legacy_docx))

        inspected = self.cli("--inspect", legacy_docx, timeout=120).payload
        self.assertIsNotNone(inspected)
        self.assertEqual(inspected["container"], "ole2")
        self.assertNotEqual(
            inspected["error"], "Not a valid docx package: missing word/document.xml"
        )
        self.assertTrue(inspected["success"], inspected.get("error"))
        self.assertEqual(inspected["inspector"], "applescript_fallback")
        self.assertEqual(inspected["table_count"], 1)
        self.assertEqual(inspected["tables"][0]["rows"], 2)
        self.assertEqual(inspected["tables"][0]["columns"], 2)
        texts = [paragraph["text"] for paragraph in inspected["paragraphs"]]
        self.assertIn("Legacy Alpha", texts)
        self.assertIn("Legacy Beta", texts)

    def test_word97_document_renamed_to_docx_can_be_edited(self):
        """--app stages a Word 97 payload named .docx as .doc; the result keeps .docx."""
        legacy_doc = self.path("legacy-edit.doc")
        created = self.cli(
            "--app", "word", "--output-file", legacy_doc, stdin_text=CREATE_WORD_97, timeout=240
        ).payload
        self.assertTrue(created["success"], created.get("error"))
        legacy_docx = self.path("legacy-edit.docx")
        shutil.copy2(legacy_doc, legacy_docx)
        payload = self.cli(
            "--app",
            "word",
            "--target-file",
            legacy_docx,
            stdin_text=FIND_AND_REPLACE.replace("Integration Alpha", "Legacy Alpha"),
            timeout=240,
        ).payload
        self.assertTrue(payload["success"], payload.get("error"))
        revised = self.path("legacy-edit_revised.docx")
        self.assertEqual(payload["working_file"], revised)
        texts = [paragraph["text"] for paragraph in self.inspect(revised)["paragraphs"]]
        self.assertIn("Edited Alpha", texts)

    def test_real_document_through_applescript_fallback(self):
        """The fallback path must return the real text and table grid for an OLE2 file."""
        if not support.is_ole2_container(self.base):
            self.skipTest("Office writes plain zip packages on this machine; no OLE2 fixture")
        payload = self.cli("--inspect", self.base, "--max-items", "5", timeout=120).payload
        self.assertTrue(payload["success"], payload.get("error"))
        self.assertEqual(payload["inspector"], "applescript_fallback")
        self.assertEqual(payload["table_count"], 1)
        self.assertEqual(payload["tables"][0]["rows"], 2)
        self.assertEqual(payload["tables"][0]["columns"], 2)
        self.assertEqual(payload["tables"][0]["preview"][0], ["Header A", "Header B"])
        self.assertFalse(payload["tables"][0]["cells_truncated"])
        texts = [paragraph["text"] for paragraph in payload["paragraphs"]]
        self.assertIn("Integration Alpha", texts)
        self.assertIn("Integration Beta", texts)
        self.assertLessEqual(len(texts), 5)

    def test_fallback_maps_paragraphs_into_table_cells(self):
        if not support.is_ole2_container(self.base):
            self.skipTest("Office writes plain zip packages on this machine; no OLE2 fixture")
        payload = self.inspect(self.base)
        by_text = {paragraph["text"]: paragraph for paragraph in payload["paragraphs"]}
        self.assertEqual(
            (by_text["Header A"]["table"], by_text["Header A"]["row"], by_text["Header A"]["column"]),
            (1, 1, 1),
        )
        self.assertEqual(
            (by_text["Header B"]["table"], by_text["Header B"]["row"], by_text["Header B"]["column"]),
            (1, 1, 2),
        )
        body = by_text["Integration Alpha"]
        self.assertEqual((body["table"], body["row"], body["column"]), (None, None, None))
        self.assertIsNotNone(body["alignment"])
        self.assertGreater(body["font_size"], 0)

    def test_render_reports_pages_and_line_coordinates(self):
        payload = self.render(self.base, "render-base")
        self.assertTrue(payload["success"], payload.get("error"))
        self.assertEqual(payload["mode"], "render")
        self.assertTrue(os.path.isfile(payload["pdf"]))
        self.assertGreaterEqual(payload["pages"], 1)
        self.assertEqual(len(payload["png"]), payload["pages"])
        for path in payload["png"]:
            self.assertTrue(os.path.isfile(path), path)
        texts = [line["text"] for line in payload["lines"]]
        self.assertTrue(any("Integration Alpha" in text for text in texts), texts)
        first = payload["lines"][0]
        self.assertEqual(first["page"], 1)
        self.assertGreater(first["x"], 0)
        self.assertGreater(first["y"], 0)
        self.assertGreater(first["w"], 0)

    def test_rendered_pdf_keeps_a_non_ascii_file_name(self):
        working = self.path("收據 範本.docx")
        shutil.copy2(self.base, working)
        payload = self.render(working, "render-cjk")
        self.assertTrue(payload["success"], payload.get("error"))
        self.assertEqual(os.path.basename(payload["pdf"]), "收據 範本.pdf")
        self.assertTrue(os.path.isfile(payload["pdf"]))

    def test_render_is_stable_across_edits(self):
        working = self.path("stable.docx")
        shutil.copy2(self.base, working)
        before = self.render(working, "render-before")
        self.assertTrue(before["success"], before.get("error"))
        result = self.cli(
            "--app", "word", "--target-file", working, "--in-place", stdin_text=FIND_AND_REPLACE, timeout=240
        )
        self.assertTrue(result.payload["success"], result.payload.get("error"))
        self.assertEqual(result.payload["working_file"], working)
        after = self.render(working, "render-after")
        self.assertTrue(after["success"], after.get("error"))
        self.assertEqual(after["pages"], before["pages"])
        # Text changed, so geometry of unchanged lines must stay put: compare the last body
        # line ("Integration Beta") across both renders.
        def beta(payload):
            return [line for line in payload["lines"] if "Integration Beta" in line["text"]]

        self.assertEqual(len(beta(before)), len(beta(after)))
        self.assertAlmostEqual(beta(before)[0]["y"], beta(after)[0]["y"], delta=0.5)
        self.assertAlmostEqual(beta(before)[0]["x"], beta(after)[0]["x"], delta=0.5)


class WordRichFallbackTests(IntegrationCase):
    """A fixture with body text, a table, a text box and an inline picture from --asset."""

    app = "word"

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.base = cls.artifact("rich.docx")
        cls.image = support.write_png(cls.artifact("dot.png"))
        cls.create_result = cls.cli(
            "--app",
            "word",
            "--output-file",
            cls.base,
            "--asset",
            cls.image,
            stdin_text=CREATE_RICH_DOCX,
            timeout=240,
        )
        if not (cls.create_result.payload and cls.create_result.payload["success"]):
            raise AssertionError("Word could not create the fixture: %r" % (cls.create_result,))

    def setUp(self):
        super().setUp()
        reason = "" if support.is_ole2_container(self.base) else "fixture is not an OLE2 container"
        if reason:
            self.skipTest(reason)

    def inspect(self):
        payload = self.cli("--inspect", self.base, timeout=120).payload
        self.assertTrue(payload["success"], payload.get("error"))
        self.assertEqual(payload["inspector"], "applescript_fallback")
        return payload

    def test_asset_reached_word_as_a_staged_file(self):
        self.assertTrue(os.path.isfile(self.base))
        payload = self.cli("--inspect", self.base, timeout=120).payload
        self.assertTrue(payload["success"], payload.get("error"))
        self.assertEqual(payload["inline_shape_count"], 1)

    def test_shapes_include_the_text_box_and_its_text(self):
        payload = self.inspect()
        self.assertGreaterEqual(payload["shape_count"], 1)
        boxes = [shape for shape in payload["shapes"] if shape["text"] == "Boxed Text"]
        self.assertTrue(boxes, payload["shapes"])
        box = boxes[0]
        self.assertIn("text box", box["type"])
        self.assertGreater(box["width_pt"], 0)
        self.assertGreater(box["height_pt"], 0)

    def test_paragraph_mapping_covers_every_cell(self):
        payload = self.inspect()
        mapped = {
            paragraph["text"]: (paragraph["table"], paragraph["row"], paragraph["column"])
            for paragraph in payload["paragraphs"]
        }
        self.assertEqual(mapped["H1"], (1, 1, 1))
        self.assertEqual(mapped["H2"], (1, 1, 2))
        self.assertEqual(mapped["B1"], (1, 2, 1))
        self.assertEqual(mapped["B2"], (1, 2, 2))
        self.assertEqual(mapped["Body Alpha"], (None, None, None))
        self.assertEqual(mapped["After table"], (None, None, None))
        self.assertEqual(payload["tables"][0]["preview"][0], ["H1", "H2"])

    def test_page_setup_is_reported(self):
        payload = self.inspect()
        setup = payload["page_setup"]
        self.assertIsNotNone(setup)
        self.assertGreater(setup["page_width"], 0)
        self.assertGreater(setup["page_height"], 0)
        self.assertGreater(setup["left_margin"], 0)


if __name__ == "__main__":
    unittest.main()
