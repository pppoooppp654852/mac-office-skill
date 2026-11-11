"""Unit tests: hidden launch, hiding, focus guard, per-app lock and idle quit.

The fake ``osascript`` answers the NSWorkspace (JXA) desktop-state probe from
``MAC_OFFICE_FAKE_APP_STATE`` / ``MAC_OFFICE_FAKE_FRONT``, and a fake ``open`` stands in
for ``open -g -j -a``, so no real application is started, hidden, quit or focused.
"""

import glob
import json
import os
import shutil
import subprocess
import sys
import time
import unittest

import support
from support import ShimTestCase, constants

constants()  # puts scripts/ on sys.path
from macoffice import app_session  # noqa: E402

SCRIPT = 'tell application "Microsoft Word"\n\treturn "ran"\nend tell\n'


class RunnerStateTestCase(ShimTestCase):
    """Keeps the real per-app lifecycle files out of reach of the fake runs."""

    def setUp(self):
        super().setUp()
        self.state_dir = app_session.state_dir()
        self.saved_state = self.path("saved-state")
        os.makedirs(self.saved_state)
        for path in glob.glob(os.path.join(self.state_dir, "*.json")):
            shutil.move(path, self.saved_state)
        self.addCleanup(self._restore_state)

    def _restore_state(self):
        for path in glob.glob(os.path.join(self.state_dir, "*.json")):
            os.unlink(path)
        for path in glob.glob(os.path.join(self.saved_state, "*.json")):
            shutil.move(path, self.state_dir)

    def run_word(self, **shim_env):
        script = self.write_script("s.applescript", SCRIPT)
        return self.invoke("--app", "word", "--file", script, shim_env=shim_env).payload

    def hide_calls(self):
        return [script for script in self.shim.e_scripts() if "set visible of process" in script]

    def write_ownership(self, app, pid, idle_seconds):
        with open(os.path.join(self.state_dir, f"{app}.json"), "w") as handle:
            json.dump({"pid": pid, "last_used": time.time() - idle_seconds}, handle)

    def ownership(self, app):
        return app_session.read_ownership(app)


class LaunchAndHideTests(RunnerStateTestCase):
    def test_stopped_app_is_launched_hidden_and_owned(self):
        payload = self.run_word(MAC_OFFICE_FAKE_APP_STATE="stopped")
        self.assertTrue(payload["success"], payload["error"])
        self.assertEqual(self.shim.open_calls(), [["-g", "-j", "-a", "Microsoft Word"]])
        self.assertEqual(len(self.hide_calls()), 2, "hidden before and after the script")
        self.assertEqual(self.ownership("word")["pid"], 4243)

    def test_visible_app_is_neither_launched_nor_hidden(self):
        """The user's own visible app is never hidden as a whole."""
        payload = self.run_word(MAC_OFFICE_FAKE_APP_STATE="visible")
        self.assertTrue(payload["success"])
        self.assertEqual(self.shim.open_calls(), [])
        self.assertEqual(self.hide_calls(), [])
        self.assertIsNone(self.ownership("word"))

    def test_hidden_app_stays_hidden(self):
        self.run_word(MAC_OFFICE_FAKE_APP_STATE="hidden")
        self.assertEqual(self.shim.open_calls(), [])
        self.assertEqual(len(self.hide_calls()), 2)

    def test_app_hidden_token_tells_runner_scripts_how_the_app_runs(self):
        source = support.write_ole2_file(self.path("doc.docx"))
        for state, expected in (("hidden", "if not true then"), ("visible", "if not false then")):
            with self.subTest(state=state):
                self.invoke("--inspect", source, shim_env={"MAC_OFFICE_FAKE_APP_STATE": state})
                self.assertIn(expected, self.shim.last_script())

    def test_unreadable_desktop_state_still_runs_the_script(self):
        payload = self.run_word(MAC_OFFICE_FAKE_DESKTOP="unreadable")
        self.assertTrue(payload["success"], payload["error"])
        self.assertEqual(self.shim.open_calls(), [])
        self.assertFalse(payload["focus_restored"])


class FocusGuardTests(RunnerStateTestCase):
    def activations(self):
        return [script for script in self.shim.jxa_scripts() if "activateWithOptions" in script]

    def test_stolen_focus_is_given_back(self):
        payload = self.run_word(MAC_OFFICE_FAKE_FRONT="Terminal,Microsoft Word")
        self.assertTrue(payload["focus_restored"])
        self.assertEqual(len(self.activations()), 1)
        # Terminal is the first name of the fake sequence, so its fake pid is 100.
        self.assertIn("runningApplicationWithProcessIdentifier(100)", self.activations()[0])

    def test_untouched_focus_is_left_alone(self):
        payload = self.run_word(MAC_OFFICE_FAKE_FRONT="Terminal")
        self.assertFalse(payload["focus_restored"])
        self.assertEqual(self.activations(), [])

    def test_user_switching_to_another_app_is_not_a_steal(self):
        payload = self.run_word(MAC_OFFICE_FAKE_FRONT="Terminal,Safari")
        self.assertFalse(payload["focus_restored"])
        self.assertEqual(self.activations(), [])

    def test_focus_was_stolen_decision(self):
        def state(name, pid):
            return {"front": {"name": name, "pid": pid}, "app": None}

        steal = app_session.focus_was_stolen
        self.assertTrue(steal(state("Code", 1), state("Microsoft Excel", 2)))
        self.assertFalse(steal(state("Microsoft Word", 2), state("Microsoft Word", 2)))
        self.assertFalse(steal(state("Microsoft Word", 2), state("Microsoft Excel", 3)))
        self.assertFalse(steal(state("Code", 1), state("Safari", 3)))
        self.assertFalse(steal(None, state("Microsoft Word", 2)))
        self.assertFalse(steal({"front": None, "app": None}, state("Microsoft Word", 2)))


class LifecycleTests(RunnerStateTestCase):
    def quit_calls(self):
        return [script for script in self.shim.e_scripts() if "quit saving no" in script]

    def test_idle_runner_launched_app_is_quit_on_a_later_call(self):
        self.write_ownership("excel", 4242, constants().IDLE_QUIT_SECONDS + 60)
        self.run_word(MAC_OFFICE_FAKE_APP_STATE="hidden")
        quits = self.quit_calls()
        self.assertEqual(len(quits), 1)
        self.assertIn('tell application "Microsoft Excel"', quits[0])
        self.assertIn("if (count of workbooks) is 0 then", quits[0])
        self.assertIsNone(self.ownership("excel"))

    def test_recently_used_app_is_kept(self):
        self.write_ownership("excel", 4242, 5)
        self.run_word(MAC_OFFICE_FAKE_APP_STATE="hidden")
        self.assertEqual(self.quit_calls(), [])
        self.assertIsNotNone(self.ownership("excel"))

    def test_app_the_user_took_over_is_never_quit(self):
        """Visible again means the user showed it: forget it instead of quitting."""
        self.write_ownership("excel", 4242, constants().IDLE_QUIT_SECONDS + 60)
        self.run_word(MAC_OFFICE_FAKE_APP_STATE="visible")
        self.assertEqual(self.quit_calls(), [])
        self.assertIsNone(self.ownership("excel"))

    def test_other_instance_than_the_launched_one_is_never_quit(self):
        self.write_ownership("excel", 999, constants().IDLE_QUIT_SECONDS + 60)
        self.run_word(MAC_OFFICE_FAKE_APP_STATE="hidden")
        self.assertEqual(self.quit_calls(), [])
        self.assertIsNone(self.ownership("excel"))

    def test_app_with_open_documents_is_kept_for_a_later_try(self):
        self.write_ownership("excel", 4242, constants().IDLE_QUIT_SECONDS + 60)
        self.run_word(MAC_OFFICE_FAKE_APP_STATE="hidden", MAC_OFFICE_FAKE_QUIT_STDOUT="kept")
        self.assertEqual(len(self.quit_calls()), 1)
        self.assertIsNotNone(self.ownership("excel"))

    def test_the_app_in_use_is_never_quit(self):
        self.write_ownership("word", 4242, constants().IDLE_QUIT_SECONDS + 60)
        self.run_word(MAC_OFFICE_FAKE_APP_STATE="hidden")
        self.assertEqual(self.quit_calls(), [])
        self.assertGreater(self.ownership("word")["last_used"], time.time() - 60)

    def test_ownership_decisions(self):
        owned = {"pid": 7, "last_used": 0}
        self.assertTrue(app_session.runner_owns(owned, {"pid": 7, "hidden": True}))
        self.assertFalse(app_session.runner_owns(owned, {"pid": 7, "hidden": False}))
        self.assertFalse(app_session.runner_owns(owned, {"pid": 8, "hidden": True}))
        self.assertFalse(app_session.runner_owns(None, {"pid": 7, "hidden": True}))
        self.assertFalse(app_session.runner_owns(owned, None))
        limit = constants().IDLE_QUIT_SECONDS
        self.assertTrue(app_session.is_idle({"pid": 7, "last_used": 100}, 100 + limit + 1))
        self.assertFalse(app_session.is_idle({"pid": 7, "last_used": 100}, 100 + limit))
        self.assertFalse(app_session.is_idle(None, 10 ** 10))


class DesktopSamplerSummaryTests(unittest.TestCase):
    def test_a_steal_is_a_switch_from_a_user_app_to_office(self):
        def sample(front, hidden=False, windows=0):
            return {"front": front, "apps": {"Microsoft Word": {"hidden": hidden, "windows": windows}}}

        summary = support.summarize_desktop_samples(
            [
                sample("Code", hidden=True),
                sample("Microsoft Word", windows=1),
                sample("Microsoft Word", windows=1),
                sample("Code"),
                sample("Microsoft Excel"),
            ]
        )
        self.assertEqual(summary["focus_steals"], 2)
        self.assertEqual(summary["visible_window_samples"], {"Microsoft Word": 2})
        self.assertEqual(summary["hidden_at_start"], ["Microsoft Word"])


class LockTests(RunnerStateTestCase):
    def test_calls_for_one_app_are_serialized(self):
        script = self.write_script("s.applescript", SCRIPT)
        with app_session.app_lock("word"):
            process = subprocess.Popen(
                [sys.executable, support.RUNNER, "--app", "word", "--file", script],
                env=self.shim.environment(),
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )
            time.sleep(1.5)
            self.assertIsNone(process.poll(), "the runner did not wait for the lock")
            self.assertEqual(self.shim.scripts(), [])
        stdout, _ = process.communicate(timeout=60)
        self.assertTrue(json.loads(stdout)["success"])
        self.assertEqual(len(self.shim.scripts()), 1)

    def test_lock_is_per_app(self):
        with app_session.app_lock("word"):
            with app_session.app_lock("excel", blocking=False) as acquired:
                self.assertTrue(acquired)

    def test_non_blocking_lock_reports_a_busy_app(self):
        with app_session.app_lock("word"):
            process = subprocess.run(
                [
                    sys.executable,
                    "-c",
                    "import sys; sys.path.insert(0, sys.argv[1]);"
                    "from macoffice import app_session;"
                    "\nwith app_session.app_lock('word', blocking=False) as acquired: print(acquired)",
                    os.path.join(support.SKILL_DIR, "scripts"),
                ],
                capture_output=True,
                text=True,
                timeout=30,
            )
        self.assertEqual(process.stdout.strip(), "False", process.stderr)


if __name__ == "__main__":
    unittest.main()
