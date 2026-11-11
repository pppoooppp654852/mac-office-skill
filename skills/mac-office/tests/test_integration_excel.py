"""Integration tests against real Microsoft Excel (gated by MAC_OFFICE_INTEGRATION=1)."""

import os
import shutil
import unittest

import support
from support import IntegrationCase

CREATE_WORKBOOK = """
tell application "Microsoft Excel"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set theWorkbook to make new workbook
	set name of active sheet of theWorkbook to "Data"
	tell active sheet of theWorkbook
		set value of range "A1:B2" to {{"Label", "Score"}, {"alpha", 42}}
	end tell
	save workbook as theWorkbook filename outPath file format Excel XML file format
	close active workbook saving no
end tell
"""

EDIT_CELL = """
tell application "Microsoft Excel"
	set targetPath to system attribute "MAC_OFFICE_TARGET_FILE"
	open (targetPath as POSIX file)
	tell active sheet of active workbook
		set value of range "A1" to "EditedLabel"
	end tell
	save active workbook
	close active workbook saving no
end tell
"""

PDF_EXPORT = """
tell application "Microsoft Excel"
	set targetPath to system attribute "MAC_OFFICE_TARGET_FILE"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	open (targetPath as POSIX file)
	save workbook as active workbook filename outPath file format PDF file format
	close active workbook saving no
end tell
"""


class ExcelIntegrationTests(IntegrationCase):
    app = "excel"

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.base = cls.artifact("base.xlsx")
        cls.create_result = cls.cli(
            "--app", "excel", "--output-file", cls.base, stdin_text=CREATE_WORKBOOK, timeout=240
        )
        if not (cls.create_result.payload and cls.create_result.payload["success"]):
            raise AssertionError("Excel could not create the fixture: %r" % (cls.create_result,))

    def inspect(self, path):
        result = self.cli("--inspect", path, timeout=120)
        self.assertIsNotNone(result.payload, result.stderr)
        self.assertTrue(result.payload["success"], result.payload.get("error"))
        return result.payload

    def first_sheet(self, payload):
        self.assertGreaterEqual(payload["sheet_count"], 1)
        return payload["worksheets"][0]

    def test_created_workbook_inspects(self):
        payload = self.create_result.payload
        self.assertTrue(payload["output_exists"])
        self.assertTrue(payload["output_modified"])
        self.assertTrue(os.path.isfile(self.base))

        inspected = self.inspect(self.base)
        self.assertEqual(inspected["kind"], "xlsx")
        if support.is_ole2_container(self.base):
            self.assertEqual(inspected["inspector"], "applescript_fallback")
        else:
            self.assertEqual(inspected["inspector"], "headless_openxml")
        self.assertEqual(self.first_sheet(inspected)["name"], "Data")
        self.assertEqual(self.first_sheet(inspected)["used_range"].replace("$", ""), "A1:B2")
        self.assertEqual(self.first_sheet(inspected)["preview"][0][0], "Label")

    def test_edit_creates_revised_workbook(self):
        working = self.path("edit.xlsx")
        shutil.copy2(self.base, working)
        result = self.cli(
            "--app", "excel", "--target-file", working, stdin_text=EDIT_CELL, timeout=240
        )
        payload = result.payload
        self.assertTrue(payload["success"], payload.get("error"))
        revised = self.path("edit_revised.xlsx")
        self.assertEqual(payload["target_file"], working)
        self.assertEqual(payload["working_file"], revised)
        self.assertTrue(payload["target_modified"])
        self.assertTrue(payload["output_exists"])
        self.assertTrue(payload["output_modified"])

        edited = self.first_sheet(self.inspect(revised))
        self.assertEqual(edited["preview"][0][0], "EditedLabel")
        original = self.first_sheet(self.inspect(working))
        self.assertEqual(original["preview"][0][0], "Label")

    def test_render_reports_a_sheet_as_one_page(self):
        result = self.cli(
            "--render", self.base, "--output-dir", self.path("render"), "--timeout", "120", timeout=300
        )
        payload = result.payload
        self.assertTrue(payload["success"], payload.get("error"))
        self.assertEqual(payload["pages"], 1)
        self.assertTrue(os.path.isfile(payload["png"][0]))
        texts = [line["text"] for line in payload["lines"]]
        self.assertTrue(any("Label" in text for text in texts), texts)

    def test_pdf_export(self):
        pdf = self.path("export", "sheet.pdf")
        result = self.cli(
            "--app",
            "excel",
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
        with open(pdf, "rb") as handle:
            self.assertEqual(handle.read(5), b"%PDF-")


if __name__ == "__main__":
    unittest.main()
