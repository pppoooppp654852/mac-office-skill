"""Unit tests: ``--render`` (PDF rasterization via PDFKit, Office export via the fake shim).

The ``.pdf`` cases rasterize for real: JXA + PDFKit need no Office app, so page count, PNG
pixel size, line text and coordinates are asserted against a deterministically built PDF.
The Office cases drive the fake osascript, which copies the same PDF to the staged output.
"""

import os
import tempfile
import unittest

import support
from support import MacOfficeTestCase, ShimTestCase, run_runner, staging_dir

DPI = 110
PIXEL_WIDTH = round(support.PAGE_WIDTH * DPI / 72)
PIXEL_HEIGHT = round(support.PAGE_HEIGHT * DPI / 72)


class RenderPdfTests(MacOfficeTestCase):
    def render(self, path, *extra):
        result = run_runner(["--render", path, "--dpi", DPI, *extra])
        self.assertIsNotNone(result.payload, result.stderr)
        return result

    def test_pdf_is_rasterized_and_lines_are_measured(self):
        pdf = support.write_pdf(self.path("source.pdf"), [["Hello Render", "Second line"]])
        output = self.path("out")
        result = self.render(pdf, "--output-dir", output)
        self.assertEqual(result.returncode, 0, result.stderr)
        payload = result.payload
        self.assertTrue(payload["success"])
        self.assertEqual(payload["mode"], "render")
        self.assertEqual(payload["file"], os.path.abspath(pdf))
        self.assertEqual(payload["pdf"], os.path.abspath(pdf))
        self.assertEqual(payload["pages"], 1)
        self.assertIsNone(payload["error"])
        self.assertIsNone(payload["error_hint"])
        self.assertIsInstance(payload["execution_time_sec"], float)

        # The PDF is kept next to the PNGs is only true for Office inputs; a .pdf input is
        # rasterized in place.
        self.assertEqual(len(payload["png"]), 1)
        self.assertEqual(payload["png"][0], os.path.join(os.path.abspath(output), "source-page-1.png"))
        self.assertTrue(os.path.isfile(payload["png"][0]))

        first, second = payload["lines"]
        self.assertEqual(first["page"], 1)
        self.assertEqual(first["text"], "Hello Render")
        # Helvetica 14 pt at (72, 700) on a 612x792 page: the selection bounds start at the
        # left margin and the box sits ~81 pt below the top edge.
        self.assertAlmostEqual(first["x"], 72.0, delta=1.0)
        self.assertAlmostEqual(first["y"], 81.2, delta=2.0)
        self.assertAlmostEqual(first["h"], 14.0, delta=1.0)
        self.assertGreater(first["w"], 60.0)
        self.assertEqual(second["text"], "Second line")
        self.assertAlmostEqual(second["y"] - first["y"], 20.0, delta=0.5)

    def test_png_pixel_size_follows_the_dpi(self):
        pdf = support.write_pdf(self.path("source.pdf"), [["Pixel check"]])
        payload = self.render(pdf, "--output-dir", self.path("out")).payload
        png_path = payload["png"][0]
        with open(png_path, "rb") as handle:
            header = handle.read(33)
        self.assertEqual(header[:8], b"\x89PNG\r\n\x1a\n")
        width = int.from_bytes(header[16:20], "big")
        height = int.from_bytes(header[20:24], "big")
        self.assertEqual((width, height), (PIXEL_WIDTH, PIXEL_HEIGHT))

    def test_every_page_is_written_with_its_page_number(self):
        pdf = support.write_pdf(
            self.path("source.pdf"), [["First page line"], ["Second page line", "more"]]
        )
        payload = self.render(pdf, "--output-dir", self.path("out")).payload
        self.assertEqual(payload["pages"], 2)
        self.assertEqual([os.path.basename(path) for path in payload["png"]], ["source-page-1.png", "source-page-2.png"])
        self.assertEqual([line["page"] for line in payload["lines"]], [1, 2, 2])
        self.assertEqual(payload["lines"][0]["text"], "First page line")
        self.assertEqual(payload["lines"][2]["text"], "more")

    def test_default_output_directory_is_a_render_temp_dir(self):
        pdf = support.write_pdf(self.path("source.pdf"), [["Temp dir line"]])
        before = os.listdir(tempfile.gettempdir())
        payload = self.render(pdf).payload
        created = [
            name
            for name in os.listdir(tempfile.gettempdir())
            if name not in before and name.startswith("mac-office-render-source-")
        ]
        self.assertEqual(len(created), 1, created)
        output_dir = os.path.join(tempfile.gettempdir(), created[0])
        self.addCleanup(lambda: __import__("shutil").rmtree(output_dir, True))
        self.assertEqual(payload["png"], [os.path.join(output_dir, "source-page-1.png")])

    def test_dpi_changes_the_pixel_size(self):
        pdf = support.write_pdf(self.path("source.pdf"), [["Resolution"]])
        payload = run_runner(["--render", pdf, "--dpi", "72", "--output-dir", self.path("out")]).payload
        with open(payload["png"][0], "rb") as handle:
            header = handle.read(24)
        self.assertEqual(int.from_bytes(header[16:20], "big"), support.PAGE_WIDTH)
        self.assertEqual(int.from_bytes(header[20:24], "big"), support.PAGE_HEIGHT)

    def test_missing_file(self):
        result = self.render(self.path("ghost.pdf"))
        self.assertEqual(result.returncode, 1)
        self.assertEqual(
            result.payload["error"],
            "File not found: {}".format(os.path.abspath(self.path("ghost.pdf"))),
        )
        self.assertEqual(result.payload["pages"], 0)
        self.assertEqual(result.payload["png"], [])

    def test_unsupported_extension(self):
        path = self.write(self.path("notes.txt"), "plain text")
        result = self.render(path)
        self.assertEqual(
            result.payload["error"], "Unsupported file type. Expected .docx, .xlsx, .pptx, or .pdf."
        )
        self.assertFalse(result.payload["success"])

    def test_corrupt_pdf_reports_an_error(self):
        path = self.write(self.path("broken.pdf"), b"%PDF-1.4 not really a pdf")
        result = self.render(path, "--output-dir", self.path("out"))
        self.assertEqual(result.returncode, 1)
        self.assertFalse(result.payload["success"])
        self.assertIn("cannot open PDF", result.payload["error"])

    def test_dpi_and_output_dir_require_render(self):
        result = run_runner(["--inspect", self.path("x.docx"), "--dpi", "72"])
        self.assertEqual(result.returncode, 2)
        self.assertIn("--output-dir and --dpi require --render", result.stderr)


class RenderOfficeTests(ShimTestCase):
    """The Office export half, driven by the fake osascript."""

    def setUp(self):
        super().setUp()
        self.pdf = support.write_pdf(self.path("export.pdf"), [["Exported line"]])

    def render(self, name="deck.docx", *extra, exit_code=0, stderr=""):
        source = self.path(name)
        if not os.path.exists(source):
            support.write_docx(source, [("p", name, None)])
        result = self.invoke(
            "--render",
            source,
            "--output-dir",
            self.path("out"),
            *extra,
            shim_env={
                "MAC_OFFICE_FAKE_OUTPUT_SOURCE": self.pdf,
                "MAC_OFFICE_FAKE_EXIT": str(exit_code),
                "MAC_OFFICE_FAKE_STDERR": stderr,
            },
        )
        self.assertIsNotNone(result.payload, result.stderr)
        return result

    def test_docx_is_exported_then_rasterized(self):
        result = self.render()
        self.assertEqual(result.returncode, 0, result.stderr)
        payload = result.payload
        self.assertTrue(payload["success"])
        self.assertEqual(payload["mode"], "render")
        self.assertEqual(payload["pdf"], os.path.join(os.path.abspath(self.path("out")), "deck.pdf"))
        self.assertTrue(os.path.isfile(payload["pdf"]))
        self.assertEqual(payload["pages"], 1)
        self.assertEqual([line["text"] for line in payload["lines"]], ["Exported line"])

    def test_app_receives_staged_paths_only(self):
        self.render()
        env = self.shim.env_dump()
        stage = env["MAC_OFFICE_STAGE_DIR"]
        self.assertEqual(os.path.dirname(stage), staging_dir())
        self.assertRegex(env["MAC_OFFICE_TARGET_FILE"], r"/deck-[0-9a-f]{8}\.docx$")
        self.assertEqual(os.path.dirname(env["MAC_OFFICE_TARGET_FILE"]), stage)
        self.assertEqual(env["MAC_OFFICE_OUTPUT_FILE"], os.path.join(stage, "export.pdf"))
        script = self.shim.last_script()
        self.assertIn('tell application "Microsoft Word"', script)
        self.assertIn("format PDF", script)
        self.assertNotIn("activate", script)

    def test_export_scripts_never_target_the_active_document(self):
        """A swallowed open must not export or close a document the user has open."""
        script = self.shim_script_for("deck.docx")
        staged_name = os.path.basename(self.shim.env_dump()["MAC_OFFICE_TARGET_FILE"])
        self.assertNotIn("active document", script)
        self.assertIn('if not (exists document "{}") then return'.format(staged_name), script)
        self.assertIn("read only true add to recent files false", script)

    def shim_script_for(self, name):
        self.render(name)
        return self.shim.last_script()

    def test_existing_pdf_in_the_output_dir_is_never_overwritten(self):
        existing = self.write(self.path("out", "deck.pdf"), "user pdf")
        payload = self.render().payload
        self.assertTrue(payload["success"], payload["error"])
        self.assertEqual(payload["pdf"], os.path.join(os.path.abspath(self.path("out")), "deck-2.pdf"))
        self.assertEqual([os.path.basename(path) for path in payload["png"]], ["deck-2-page-1.png"])
        with open(existing) as handle:
            self.assertEqual(handle.read(), "user pdf")

    def test_renders_of_different_files_into_one_dir_keep_their_pngs(self):
        first = self.render("first.docx").payload
        second = self.render("second.docx").payload
        self.assertEqual([os.path.basename(path) for path in first["png"]], ["first-page-1.png"])
        self.assertEqual([os.path.basename(path) for path in second["png"]], ["second-page-1.png"])
        self.assertTrue(all(os.path.isfile(path) for path in first["png"] + second["png"]))

    def test_unchanged_file_is_rendered_from_the_cache_without_office(self):
        first = self.render().payload
        self.assertFalse(first["cached"])
        exports = len(self.shim.scripts())
        second = self.render().payload
        self.assertTrue(second["success"], second["error"])
        self.assertTrue(second["cached"])
        self.assertEqual(len(self.shim.scripts()), exports, "the app was asked to export again")
        self.assertEqual(second["pdf"], os.path.join(os.path.abspath(self.path("out")), "deck-2.pdf"))
        self.assertEqual([line["text"] for line in second["lines"]], ["Exported line"])

    def test_changed_file_is_exported_again(self):
        self.render()
        exports = len(self.shim.scripts())
        source = self.path("deck.docx")
        support.write_docx(source, [("p", "deck changed", None)])
        result = self.invoke(
            "--render", source, "--output-dir", self.path("out"),
            shim_env={"MAC_OFFICE_FAKE_OUTPUT_SOURCE": self.pdf},
        )
        self.assertFalse(result.payload["cached"])
        self.assertEqual(len(self.shim.scripts()), exports + 1)

    def test_failed_export_is_not_cached(self):
        self.render(exit_code=1, stderr="execution error: boom")
        payload = self.render().payload
        self.assertTrue(payload["success"], payload["error"])
        self.assertFalse(payload["cached"])

    def test_no_revised_copy_is_created(self):
        self.render()
        self.assertFalse(os.path.exists(self.path("deck_revised.docx")))

    def test_export_failure_is_reported(self):
        result = self.render(exit_code=1, stderr="execution error: boom (-1728)")
        self.assertEqual(result.returncode, 1)
        payload = result.payload
        self.assertFalse(payload["success"])
        self.assertEqual(payload["error"], "execution error: boom (-1728)")
        self.assertEqual(payload["pages"], 0)
        self.assertEqual(payload["png"], [])

    def test_missing_pdf_output_is_reported(self):
        source = self.path("deck.pptx")
        support.write_pptx(source, [[{"kind": "sp", "name": "T", "id": 2, "text": "hi"}]])
        result = self.invoke(
            "--render",
            source,
            "--output-dir",
            self.path("out"),
            shim_env={"MAC_OFFICE_FAKE_STDOUT": "missing"},
        )
        self.assertEqual(result.returncode, 1)
        self.assertEqual(
            result.payload["error"],
            "Microsoft PowerPoint did not produce a PDF for {}".format(os.path.abspath(source)),
        )

    def test_protected_presentation_is_named_as_the_cause(self):
        source = self.path("labelled.pptx")
        support.write_pptx(source, [[{"kind": "sp", "name": "T", "id": 2, "text": "hi"}]])
        result = self.invoke(
            "--render",
            source,
            "--output-dir",
            self.path("out"),
            shim_env={"MAC_OFFICE_FAKE_STDOUT": "protected"},
        )
        self.assertEqual(result.returncode, 1)
        self.assertIn("MIP/IRM-protected", result.payload["error"])
        self.assertIn("File > Export", result.payload["error_hint"])

    def test_staging_directory_is_cleaned_up(self):
        self.render()
        self.assertEqual(support.staging_leftovers([]), [])


if __name__ == "__main__":
    unittest.main()
