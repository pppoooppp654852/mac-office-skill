"""Integration tests against real Microsoft PowerPoint (gated by MAC_OFFICE_INTEGRATION=1)."""

import os
import shutil
import unittest

import support
from support import IntegrationCase

CREATE_DECK = """
tell application "Microsoft PowerPoint"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set thePres to make new presentation
	set theSlide to make new slide at end of thePres
	set layout of theSlide to slide layout blank
	set newBox to make new shape at end of theSlide with properties {auto shape type:autoshape rectangle, top:100, left position:50, width:400, height:60}
	set theTextRange to text range of text frame of newBox
	set content of theTextRange to "Integration Slide"
	save thePres in (outPath as POSIX file) as save as Open XML presentation
	close active presentation saving no
end tell
"""

EDIT_SHAPE_TEXT = """
tell application "Microsoft PowerPoint"
	set targetPath to system attribute "MAC_OFFICE_TARGET_FILE"
	open (targetPath as POSIX file)
	repeat 25 times
		if (count of presentations) > 0 then exit repeat
		delay 0.1
	end repeat
	set theTextRange to text range of text frame of shape 1 of slide 1 of active presentation
	set content of theTextRange to "Edited Slide Text"
	save active presentation
	close active presentation saving no
end tell
"""


class PowerPointIntegrationTests(IntegrationCase):
    app = "powerpoint"

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.base = cls.artifact("base.pptx")
        cls.create_result = cls.cli(
            "--app",
            "powerpoint",
            "--output-file",
            cls.base,
            stdin_text=CREATE_DECK,
            timeout=240,
        )
        if not (cls.create_result.payload and cls.create_result.payload["success"]):
            raise AssertionError(
                "PowerPoint could not create the fixture: %r" % (cls.create_result,)
            )

    def inspect(self, path):
        result = self.cli("--inspect", path, timeout=120)
        self.assertIsNotNone(result.payload, result.stderr)
        self.assertTrue(result.payload["success"], result.payload.get("error"))
        return result.payload

    def shapes(self, payload):
        return [shape for slide in payload["slides"] for shape in slide["shapes"]]

    def test_created_deck_inspects(self):
        payload = self.create_result.payload
        self.assertTrue(payload["output_exists"])
        self.assertTrue(payload["output_modified"])
        self.assertTrue(os.path.isfile(self.base))

        inspected = self.inspect(self.base)
        self.assertEqual(inspected["kind"], "pptx")
        self.assertEqual(inspected["slide_count"], 1)
        shapes = self.shapes(inspected)
        self.assertEqual(len(shapes), 1)
        shape = shapes[0]
        self.assertEqual(shape["text"], "Integration Slide")
        self.assertAlmostEqual(shape["left_pt"], 50.0, delta=0.51)
        self.assertAlmostEqual(shape["top_pt"], 100.0, delta=0.51)
        self.assertAlmostEqual(shape["width_pt"], 400.0, delta=0.51)
        self.assertAlmostEqual(shape["height_pt"], 60.0, delta=0.51)

    def test_render_reports_protection_or_a_deck(self):
        """A MIP-protected deck cannot be exported by PowerPoint; report that, never lie."""
        result = self.cli(
            "--render", self.base, "--output-dir", self.path("render"), "--timeout", "120", timeout=300
        )
        payload = result.payload
        if support.is_ole2_container(self.base) and not payload["success"]:
            self.assertIn("MIP/IRM-protected", payload["error"])
            self.assertIsNotNone(payload["error_hint"])
            return
        self.assertTrue(payload["success"], payload.get("error"))
        self.assertEqual(payload["pages"], 1)
        self.assertTrue(os.path.isfile(payload["png"][0]))
        texts = [line["text"] for line in payload["lines"]]
        self.assertTrue(any("Integration Slide" in text for text in texts), texts)

    def test_edit_creates_revised_deck(self):
        working = self.path("edit.pptx")
        shutil.copy2(self.base, working)
        result = self.cli(
            "--app",
            "powerpoint",
            "--target-file",
            working,
            stdin_text=EDIT_SHAPE_TEXT,
            timeout=240,
        )
        payload = result.payload
        self.assertTrue(payload["success"], payload.get("error"))
        revised = self.path("edit_revised.pptx")
        self.assertEqual(payload["target_file"], working)
        self.assertEqual(payload["working_file"], revised)
        self.assertTrue(payload["target_modified"])
        self.assertTrue(payload["output_exists"])
        self.assertTrue(payload["output_modified"])

        edited = self.shapes(self.inspect(revised))
        self.assertEqual([shape["text"] for shape in edited], ["Edited Slide Text"])
        original = self.shapes(self.inspect(working))
        self.assertEqual([shape["text"] for shape in original], ["Integration Slide"])


if __name__ == "__main__":
    unittest.main()
