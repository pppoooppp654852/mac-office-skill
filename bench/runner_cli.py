"""Call the skill's office_runner.py from the bench and read its JSON answer."""

import json
import os
import subprocess
import sys

BENCH_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_DIR = os.path.dirname(BENCH_DIR)
SKILL_DIR = os.path.join(REPO_DIR, "skills", "mac-office")
RUNNER = os.path.join(SKILL_DIR, "scripts", "office_runner.py")
TASKS_DIR = os.path.join(BENCH_DIR, "tasks")
RUNS_DIR = os.path.join(BENCH_DIR, "runs")
FIXTURES_DIR = os.path.join(RUNS_DIR, "fixtures")
STAGING_ROOT = os.path.expanduser("~/Library/Group Containers/UBF8T346G9.Office/mac-office-staging")


def run_runner(args, stdin_text=None, timeout=300):
    """Run office_runner.py with ``args``; returns its parsed JSON (``success`` false on failure)."""
    try:
        completed = subprocess.run(
            [sys.executable, RUNNER, *args],
            input=stdin_text,
            capture_output=True,
            text=True,
            timeout=timeout,
        )
    except subprocess.TimeoutExpired:
        return {"success": False, "error": f"runner exceeded {timeout}s"}
    try:
        return json.loads(completed.stdout)
    except json.JSONDecodeError:
        return {"success": False, "error": (completed.stderr or completed.stdout).strip()[-2000:]}


def staging_leftovers():
    """Entries in the Office staging root other than the runner's own state directory."""
    if not os.path.isdir(STAGING_ROOT):
        return []
    return sorted(name for name in os.listdir(STAGING_ROOT) if name != "runner-state" and not name.startswith("."))
