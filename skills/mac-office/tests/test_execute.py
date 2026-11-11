"""Unit tests: ``--app`` AppleScript execution (staging, assets, lazy write-back, errors).

Every osascript call goes to the fake shim on PATH; no Microsoft Office app is launched.
"""

import os
import re
import unittest

import support
from support import ShimTestCase, constants, staging_dir

SIMPLE_SCRIPT = (
    'tell application "Microsoft Word"\n'
    "\treturn \"ran\"\n"
    "end tell\n"
)


class ExecuteTestCase(ShimTestCase):
    """Adds an on-disk AppleScript plus an invoke wrapper: app first, then CLI options."""

    def setUp(self):
        super().setUp()
        self.script_path = self.write_script("simple.applescript", SIMPLE_SCRIPT)

    def invoke_script(self, app, *args, **kwargs):
        return self.invoke("--app", app, "--file", self.script_path, *args, **kwargs)


class PreambleTests(ExecuteTestCase):
    def test_preamble_and_postamble_are_wrapped_per_app(self):
        handlers = constants().ALERT_HANDLERS
        for app in ("word", "excel", "powerpoint"):
            with self.subTest(app=app):
                preamble, postamble = handlers[app]
                result = self.invoke_script(app)
                self.assertEqual(result.returncode, 0, result.stderr)
                self.assertTrue(result.payload["success"])
                received = self.shim.last_script()
                self.assertEqual(received, preamble + SIMPLE_SCRIPT.rstrip() + "\n" + postamble)
                # The runner starts a stopped app hidden itself; the script never launches it.
                self.assertNotIn("launch", received)
                self.assertNotIn("activate", received)
                # The script is written to a throwaway file, never to the caller's file.
                self.assertNotEqual(self.shim.last_script_path(), self.script_path)

    def test_display_alert_handlers_match_app_conventions(self):
        handlers = constants().ALERT_HANDLERS
        self.assertIn("set display alerts to alerts none", handlers["word"][0])
        self.assertIn("set display alerts to false", handlers["excel"][0])
        self.assertEqual(handlers["powerpoint"][1], "")
        self.assertIn("get Version", handlers["powerpoint"][0])

    def test_script_can_come_from_stdin(self):
        result = self.invoke("--app", "word", stdin_text=SIMPLE_SCRIPT)
        self.assertTrue(result.payload["success"])
        preamble, postamble = constants().ALERT_HANDLERS["word"]
        self.assertEqual(
            self.shim.last_script(),
            preamble + SIMPLE_SCRIPT.rstrip() + "\n" + postamble,
        )

    def test_unreadable_script_file(self):
        result = self.invoke("--app", "word", "--file", self.path("nope.applescript"))
        self.assertEqual(result.returncode, 1)
        self.assertFalse(result.payload["success"])
        self.assertTrue(result.payload["error"].startswith("Cannot read AppleScript:"))
        self.assertEqual(self.shim.calls(), [])

    def test_stdout_is_reported_as_output(self):
        result = self.invoke_script("word", shim_env={"MAC_OFFICE_FAKE_STDOUT": "ran ok\n"})
        self.assertEqual(result.payload["app"], "word")
        self.assertEqual(result.payload["mode"], "execute")
        self.assertEqual(result.payload["output"], "ran ok")
        self.assertIsInstance(result.payload["execution_time_sec"], float)
        self.assertIsNone(result.payload["error"])


class EnvironmentTests(ExecuteTestCase):
    def test_paths_exposed_to_the_script(self):
        target = self.write(self.path("report.docx"), "original body")
        output = self.path("out", "report.pdf")
        result = self.invoke_script(
            "word", "--target-file", target, "--output-file", output
        )
        self.assertTrue(result.payload["success"])
        env = self.shim.env_dump()
        self.assertEqual(
            sorted(env),
            [
                "MAC_OFFICE_ASSET_COUNT",
                "MAC_OFFICE_ORIGINAL_FILE",
                "MAC_OFFICE_OUTPUT_FILE",
                "MAC_OFFICE_STAGE_DIR",
                "MAC_OFFICE_TARGET_FILE",
            ],
        )
        stage = env["MAC_OFFICE_STAGE_DIR"]
        # Each run gets its own subdirectory of the staging directory.
        self.assertEqual(os.path.dirname(stage), staging_dir())
        self.assertEqual(env["MAC_OFFICE_TARGET_FILE"], os.path.join(stage, "report.docx"))
        self.assertEqual(env["MAC_OFFICE_OUTPUT_FILE"], os.path.join(stage, "report.pdf"))
        self.assertEqual(env["MAC_OFFICE_ASSET_COUNT"], "0")
        self.assertEqual(env["MAC_OFFICE_ORIGINAL_FILE"], os.path.abspath(target))
        self.assertEqual(result.payload["output_file"], os.path.abspath(output))

    def test_staged_names_are_sanitized(self):
        target = self.write(self.path("My Report (final) #2.docx"), "body")
        self.invoke_script("word", "--target-file", target)
        staged = self.shim.env_dump()["MAC_OFFICE_TARGET_FILE"]
        self.assertEqual(os.path.basename(staged), "My_Report_final_2.docx")
        self.assertNotIn(" ", os.path.basename(staged))

    def test_paths_are_empty_without_files(self):
        self.invoke_script("excel")
        env = self.shim.env_dump()
        self.assertEqual(env["MAC_OFFICE_TARGET_FILE"], "")
        self.assertEqual(env["MAC_OFFICE_OUTPUT_FILE"], "")
        self.assertEqual(env["MAC_OFFICE_ORIGINAL_FILE"], "")
        self.assertEqual(env["MAC_OFFICE_ASSET_COUNT"], "0")
        self.assertEqual(os.path.dirname(env["MAC_OFFICE_STAGE_DIR"]), staging_dir())

    def test_stage_directory_is_a_unique_subdirectory_per_run(self):
        self.invoke_script("word")
        first = self.shim.env_dump()["MAC_OFFICE_STAGE_DIR"]
        self.invoke_script("word")
        second = self.shim.env_dump()["MAC_OFFICE_STAGE_DIR"]
        self.assertNotEqual(first, second)
        self.assertEqual(os.path.dirname(first), staging_dir())
        self.assertEqual(os.path.dirname(second), staging_dir())
        # Both stage dirs are gone once their run returned.
        self.assertFalse(os.path.exists(first))
        self.assertFalse(os.path.exists(second))


class AssetTests(ExecuteTestCase):
    def test_assets_are_staged_and_exposed(self):
        logo = self.write(self.path("assets", "logo.png"), b"\x89PNG\r\n\x1a\n")
        other = self.write(self.path("assets", "diagram.png"), b"\x89PNG\r\n\x1a\nsecond")
        result = self.invoke_script("word", "--asset", logo, "--asset", other)
        self.assertTrue(result.payload["success"])
        env = self.shim.env_dump()
        stage = env["MAC_OFFICE_STAGE_DIR"]
        self.assertEqual(
            sorted(key for key in env if key.startswith("MAC_OFFICE_ASSET_")),
            ["MAC_OFFICE_ASSET_1", "MAC_OFFICE_ASSET_2", "MAC_OFFICE_ASSET_COUNT"],
        )
        self.assertEqual(env["MAC_OFFICE_ASSET_COUNT"], "2")
        self.assertEqual(env["MAC_OFFICE_ASSET_1"], os.path.join(stage, "logo.png"))
        self.assertEqual(env["MAC_OFFICE_ASSET_2"], os.path.join(stage, "diagram.png"))

    def test_assets_with_the_same_basename_get_distinct_paths(self):
        first = self.write(self.path("one", "logo.png"), b"first")
        second = self.write(self.path("two", "logo.png"), b"second")
        self.invoke_script("word", "--asset", first, "--asset", second)
        env = self.shim.env_dump()
        self.assertNotEqual(env["MAC_OFFICE_ASSET_1"], env["MAC_OFFICE_ASSET_2"])
        self.assertTrue(env["MAC_OFFICE_ASSET_2"].endswith("logo-2.png"), env["MAC_OFFICE_ASSET_2"])

    def test_assets_are_deleted_after_the_run(self):
        logo = self.write(self.path("logo.png"), b"png")
        copied = self.path("copied.png")
        result = self.invoke_script(
            "word",
            "--asset",
            logo,
            shim_env={"MAC_OFFICE_FAKE_ASSET_COPY_DEST": copied},
        )
        self.assertTrue(result.payload["success"])
        self.assertTrue(os.path.isfile(copied), "the script never saw a staged asset")
        self.assertEqual(support.staging_leftovers([]), [])

    def test_assets_are_deleted_after_a_failing_run(self):
        logo = self.write(self.path("logo.png"), b"png")
        result = self.invoke_script(
            "word",
            "--asset",
            logo,
            shim_env={"MAC_OFFICE_FAKE_EXIT": "1", "MAC_OFFICE_FAKE_STDERR": "boom"},
        )
        self.assertFalse(result.payload["success"])
        self.assertEqual(support.staging_leftovers([]), [])

    def test_missing_asset_fails_before_running_osascript(self):
        result = self.invoke_script("word", "--asset", self.path("ghost.png"))
        self.assertEqual(result.returncode, 1)
        self.assertFalse(result.payload["success"])
        self.assertEqual(
            result.payload["error"],
            "Asset file not found: {}".format(os.path.abspath(self.path("ghost.png"))),
        )
        self.assertEqual(self.shim.calls(), [])

    def test_asset_without_app_is_an_argparse_error(self):
        result = self.invoke("--asset", self.path("logo.png"))
        self.assertEqual(result.returncode, 2)
        self.assertIn("--asset requires --app", result.stderr)
        self.assertEqual(self.shim.calls(), [])


class TargetFileTests(ExecuteTestCase):
    def test_modified_target_is_written_to_a_revised_copy(self):
        target = self.write(self.path("budget.docx"), "original body")
        payload = self.invoke_script(
            "word",
            "--target-file",
            target,
            shim_env={"MAC_OFFICE_FAKE_APPEND_TARGET": " + edited"},
        ).payload
        revised = self.path("budget_revised.docx")
        self.assertTrue(payload["success"])
        self.assertEqual(payload["target_file"], os.path.abspath(target))
        self.assertEqual(payload["working_file"], revised)
        self.assertTrue(payload["target_modified"])
        self.assertTrue(payload["output_exists"])
        self.assertTrue(payload["output_modified"])
        self.assertFalse(os.path.exists(os.path.abspath(target) + ".bak"))
        with open(revised) as handle:
            self.assertEqual(handle.read(), "original body + edited")
        with open(target) as handle:
            self.assertEqual(handle.read(), "original body")

    def test_no_revised_copy_is_created_before_the_script_runs(self):
        target = self.write(self.path("budget.docx"), "original body")
        self.invoke_script(
            "word",
            "--target-file",
            target,
            shim_env={"MAC_OFFICE_FAKE_APPEND_TARGET": " + edited"},
        )
        staged = self.shim.env_dump()["MAC_OFFICE_TARGET_FILE"]
        self.assertFalse("_revised" in staged, staged)
        self.assertEqual(os.path.basename(staged), "budget.docx")

    def test_untouched_target_writes_nothing_back(self):
        target = self.write(self.path("budget.docx"), "original body")
        payload = self.invoke_script("word", "--target-file", target).payload
        self.assertTrue(payload["success"])
        self.assertEqual(payload["target_file"], os.path.abspath(target))
        self.assertIsNone(payload["working_file"])
        self.assertFalse(payload["target_modified"])
        # Nothing was written, so the deliverable is reported absent.
        self.assertFalse(payload["output_exists"])
        self.assertFalse(payload["output_modified"])
        self.assertFalse(os.path.exists(self.path("budget_revised.docx")))

    def test_in_place_updates_original_and_drops_backup(self):
        target = self.write(self.path("budget.docx"), "original body")
        payload = self.invoke_script(
            "word",
            "--target-file",
            target,
            "--in-place",
            shim_env={"MAC_OFFICE_FAKE_APPEND_TARGET": " + edited"},
        ).payload
        self.assertTrue(payload["success"])
        self.assertEqual(payload["target_file"], os.path.abspath(target))
        self.assertEqual(payload["working_file"], os.path.abspath(target))
        self.assertTrue(payload["target_modified"])
        self.assertTrue(payload["output_modified"])
        self.assertFalse(os.path.exists(target + ".bak"))
        with open(target) as handle:
            self.assertEqual(handle.read(), "original body + edited")

    def test_in_place_untouched_target_keeps_the_original(self):
        target = self.write(self.path("budget.docx"), "original body")
        payload = self.invoke_script("word", "--target-file", target, "--in-place").payload
        self.assertTrue(payload["success"])
        self.assertIsNone(payload["working_file"])
        self.assertFalse(payload["target_modified"])
        self.assertFalse(payload["output_exists"])
        self.assertFalse(os.path.exists(target + ".bak"))
        with open(target) as handle:
            self.assertEqual(handle.read(), "original body")

    def test_in_place_failure_leaves_the_original_untouched(self):
        """The original is only written after a successful run, so no backup is needed."""
        target = self.write(self.path("budget.docx"), "original body")
        payload = self.invoke_script(
            "word",
            "--target-file",
            target,
            "--in-place",
            shim_env={
                "MAC_OFFICE_FAKE_APPEND_TARGET": " corrupted",
                "MAC_OFFICE_FAKE_EXIT": "1",
                "MAC_OFFICE_FAKE_STDERR": "execution error: save failed",
            },
        ).payload
        self.assertFalse(payload["success"])
        self.assertIsNone(payload["error_hint"])
        self.assertIsNone(payload["working_file"])
        self.assertFalse(payload["output_exists"])
        self.assertFalse(os.path.exists(target + ".bak"))
        with open(target) as handle:
            self.assertEqual(handle.read(), "original body")

    def test_failure_does_not_write_back_modified_target(self):
        target = self.write(self.path("budget.docx"), "original body")
        payload = self.invoke_script(
            "word",
            "--target-file",
            target,
            shim_env={
                "MAC_OFFICE_FAKE_REPLACE_TARGET": "broken body",
                "MAC_OFFICE_FAKE_EXIT": "1",
                "MAC_OFFICE_FAKE_STDERR": "execution error: save failed",
            },
        ).payload
        self.assertFalse(payload["success"])
        # The staged copy was modified (and reported), but a failed run never lands on disk.
        self.assertTrue(payload["target_modified"])
        self.assertIsNone(payload["working_file"])
        self.assertFalse(payload["output_exists"])
        self.assertFalse(os.path.exists(self.path("budget_revised.docx")))
        with open(target) as handle:
            self.assertEqual(handle.read(), "original body")

    def test_missing_target_file(self):
        payload = self.invoke_script(
            "word", "--target-file", self.path("ghost.docx")
        ).payload
        self.assertFalse(payload["success"])
        self.assertEqual(
            payload["error"],
            "Target file not found: {}".format(os.path.abspath(self.path("ghost.docx"))),
        )
        self.assertEqual(self.shim.calls(), [])

    def test_legacy_payload_named_docx_is_staged_as_doc_and_written_back_as_docx(self):
        """Word refuses a Word 97 payload named .docx, so the staged copy becomes .doc."""
        target = support.write_cfb(
            self.path("legacy.docx"), {"WordDocument": b"\x01" * 700, "1Table": b"\x02" * 300}
        )
        payload = self.invoke_script(
            "word",
            "--target-file",
            target,
            shim_env={"MAC_OFFICE_FAKE_APPEND_TARGET": "edited"},
        ).payload
        self.assertTrue(payload["success"], payload["error"])
        self.assertEqual(os.path.basename(self.shim.env_dump()["MAC_OFFICE_TARGET_FILE"]), "legacy.doc")
        self.assertEqual(payload["working_file"], self.path("legacy_revised.docx"))
        self.assertTrue(os.path.isfile(self.path("legacy_revised.docx")))

    def test_staged_copies_never_survive(self):
        target = self.write(self.path("budget.docx"), "original body")
        self.invoke_script("word", "--target-file", target)
        self.assertEqual(support.staging_leftovers([]), [])


class OutputFileTests(ExecuteTestCase):
    def test_output_is_copied_on_success(self):
        output = self.path("exports", "created.docx")
        payload = self.invoke_script(
            "word",
            "--output-file",
            output,
            shim_env={"MAC_OFFICE_FAKE_WRITE_OUTPUT": "generated document"},
        ).payload
        self.assertTrue(payload["success"])
        self.assertEqual(payload["output_file"], os.path.abspath(output))
        self.assertIsNone(payload["working_file"])
        self.assertTrue(payload["output_exists"])
        self.assertTrue(payload["output_modified"])
        with open(output) as handle:
            self.assertEqual(handle.read(), "generated document")

    def test_output_is_not_copied_on_failure(self):
        output = self.path("exports", "created.docx")
        payload = self.invoke_script(
            "word",
            "--output-file",
            output,
            shim_env={"MAC_OFFICE_FAKE_WRITE_OUTPUT": "half written", "MAC_OFFICE_FAKE_EXIT": "1"},
        ).payload
        self.assertFalse(payload["success"])
        self.assertFalse(os.path.exists(output))
        self.assertFalse(payload["output_exists"])
        self.assertFalse(payload["output_modified"])

    def test_empty_output_is_not_copied(self):
        output = self.path("exports", "empty.pdf")
        payload = self.invoke_script(
            "word",
            "--output-file",
            output,
            shim_env={"MAC_OFFICE_FAKE_WRITE_OUTPUT": ""},
        ).payload
        self.assertTrue(payload["success"])
        self.assertFalse(os.path.exists(output))
        self.assertFalse(payload["output_exists"])
        self.assertFalse(payload["output_modified"])

    def test_missing_output_is_reported_absent(self):
        output = self.path("exports", "missing.pdf")
        payload = self.invoke_script("word", "--output-file", output).payload
        self.assertTrue(payload["success"])
        self.assertFalse(payload["output_exists"])

    def test_stale_output_file_is_left_alone_on_failure(self):
        output = self.write(self.path("exports", "stale.pdf"), "old export")
        payload = self.invoke_script(
            "word",
            "--output-file",
            output,
            shim_env={"MAC_OFFICE_FAKE_WRITE_OUTPUT": "new export", "MAC_OFFICE_FAKE_EXIT": "1"},
        ).payload
        self.assertFalse(payload["success"])
        self.assertTrue(payload["output_exists"])
        self.assertFalse(payload["output_modified"])
        with open(output) as handle:
            self.assertEqual(handle.read(), "old export")

    def test_modified_target_is_written_back_when_output_is_also_given(self):
        target = self.write(self.path("source.docx"), "source body")
        output = self.path("exports", "source.pdf")
        payload = self.invoke_script(
            "word",
            "--target-file",
            target,
            "--output-file",
            output,
            shim_env={
                "MAC_OFFICE_FAKE_APPEND_TARGET": "!",
                "MAC_OFFICE_FAKE_WRITE_OUTPUT": "%PDF-1.7 fake",
            },
        ).payload
        self.assertTrue(payload["success"])
        self.assertTrue(payload["target_modified"])
        revised = self.path("source_revised.docx")
        self.assertEqual(payload["working_file"], revised)
        with open(revised) as handle:
            self.assertEqual(handle.read(), "source body!")
        with open(target) as handle:
            self.assertEqual(handle.read(), "source body")
        with open(output) as handle:
            self.assertEqual(handle.read(), "%PDF-1.7 fake")
        self.assertTrue(payload["output_exists"])
        self.assertTrue(payload["output_modified"])

    def test_staged_output_never_survives(self):
        output = self.path("exports", "created.docx")
        self.invoke_script(
            "word",
            "--output-file",
            output,
            shim_env={"MAC_OFFICE_FAKE_WRITE_OUTPUT": "generated document"},
        )
        self.assertEqual(support.staging_leftovers([]), [])


class ErrorPathTests(ExecuteTestCase):
    def test_nonzero_exit_is_reported(self):
        result = self.invoke_script(
            "excel",
            shim_env={"MAC_OFFICE_FAKE_EXIT": "3", "MAC_OFFICE_FAKE_STDERR": "execution error: boom"},
        )
        self.assertEqual(result.returncode, 1)
        payload = result.payload
        self.assertFalse(payload["success"])
        self.assertEqual(payload["error"], "execution error: boom")
        self.assertIsNone(payload["error_hint"])

    def test_error_falls_back_to_stdout(self):
        payload = self.invoke_script(
            "excel",
            shim_env={"MAC_OFFICE_FAKE_EXIT": "1", "MAC_OFFICE_FAKE_STDOUT": "partial output"},
        ).payload
        self.assertFalse(payload["success"])
        self.assertEqual(payload["output"], "partial output")
        self.assertEqual(payload["error"], "partial output")

    def test_tcc_error_adds_automation_hint(self):
        payload = self.invoke_script(
            "word",
            shim_env={
                "MAC_OFFICE_FAKE_EXIT": "1",
                "MAC_OFFICE_FAKE_STDERR": "execution error: Not authorized to send Apple events "
                "to Microsoft Word. (-1743)",
            },
        ).payload
        self.assertFalse(payload["success"])
        self.assertEqual(payload["error_hint"], constants().TCC_HINT)

    def test_unknown_app_is_an_argparse_error(self):
        result = self.invoke("--app", "scooby", stdin_text="return 1")
        self.assertEqual(result.returncode, 2)
        self.assertIsNone(result.payload)
        self.assertIn("invalid choice", result.stderr)

    def test_no_mode_is_an_argparse_error(self):
        result = self.invoke()
        self.assertEqual(result.returncode, 2)
        self.assertIsNone(result.payload)
        self.assertIn("Specify either --inspect FILE, --render FILE or --app", result.stderr)

    def test_inspect_without_value_is_an_argparse_error(self):
        result = self.invoke("--inspect")
        self.assertEqual(result.returncode, 2)
        self.assertIn("expected one argument", result.stderr)

    def test_in_place_without_target_is_an_argparse_error(self):
        result = self.invoke_script("word", "--in-place")
        self.assertEqual(result.returncode, 2)
        self.assertIn("--in-place requires --target-file", result.stderr)
        self.assertEqual(self.shim.calls(), [])

    def test_inspect_wins_over_app(self):
        path = self.path("both.docx")
        support.write_docx(path, [("p", "hello", None)])
        result = self.invoke("--inspect", path, "--app", "word")
        self.assertEqual(result.payload["mode"], "inspect")
        self.assertEqual(self.shim.calls(), [])


class TimeoutTests(ExecuteTestCase):
    def test_timeout_without_dialog_reports_timeout(self):
        payload = self.invoke_script(
            "word", "--timeout", "5", shim_env={"MAC_OFFICE_FAKE_SLEEP": "9"}
        ).payload
        self.assertFalse(payload["success"])
        self.assertEqual(payload["error"], "osascript timed out after 5s")
        self.assertTrue(payload["error_hint"].startswith("No dialog found."), payload["error_hint"])

    def test_timeout_with_modal_dialog_reports_hint(self):
        payload = self.invoke_script(
            "word",
            "--timeout",
            "5",
            shim_env={"MAC_OFFICE_FAKE_SLEEP": "9", "MAC_OFFICE_FAKE_MODAL": "true"},
        ).payload
        self.assertFalse(payload["success"])
        self.assertEqual(payload["error_hint"], constants().MODAL_ALERT_HINT)

    def test_timeout_with_unreadable_dialog_state(self):
        payload = self.invoke_script(
            "word",
            "--timeout",
            "5",
            shim_env={"MAC_OFFICE_FAKE_SLEEP": "9", "MAC_OFFICE_FAKE_MODAL": "unreadable"},
        ).payload
        self.assertFalse(payload["success"])
        self.assertIn("could not be read", payload["error_hint"])

    def test_timeout_is_clamped_to_at_least_five_seconds(self):
        result = self.invoke_script("word", "--timeout", "1", shim_env={"MAC_OFFICE_FAKE_SLEEP": "2"})
        self.assertTrue(result.payload["success"], result.payload.get("error"))


class StagedDocumentCleanupTests(ExecuteTestCase):
    def test_only_documents_of_staged_files_are_closed(self):
        target = self.write(self.path("open.docx"), "original body")
        logo = self.write(self.path("logo.png"), b"png")
        payload = self.invoke_script("word", "--target-file", target, "--asset", logo).payload
        self.assertTrue(payload["success"])
        env = self.shim.env_dump()
        close_scripts = [
            script for script in self.shim.e_scripts() if "close document i saving no" in script
        ]
        self.assertEqual(len(close_scripts), 1)
        wanted = re.search(r"set wanted to \{(.*)\}", close_scripts[0]).group(1)
        self.assertEqual(
            sorted(wanted.split(", ")),
            sorted('"{}"'.format(path) for path in (env["MAC_OFFICE_TARGET_FILE"], env["MAC_OFFICE_ASSET_1"])),
        )
        # Each document's name is compared in both forms right before it is closed.
        self.assertIn("if wanted contains fn or wantedHFS contains fn then close document i", close_scripts[0])

    def test_no_close_call_without_staged_files(self):
        self.invoke_script("word")
        self.assertEqual(
            [script for script in self.shim.e_scripts() if "close document i" in script], []
        )


if __name__ == "__main__":
    unittest.main()
