#!/usr/bin/env python3
"""Grade each task's hand-written reference solution: every task must score hard = 1.

A task whose reference fails means the task, its fixture or the judge is wrong, so no agent
score on it is meaningful. Results go to bench/runs/reference_selfcheck.json.
"""

import argparse
import json
import os
import shutil
import sys
import tempfile

from judge import judge
from office_health import app_responsive, wait_until_responsive
from runner_cli import RUNS_DIR, run_runner, staging_leftovers
from task_catalog import load_tasks, materialize_docs


def run_reference(task, docs_dir):
    """Run the reference AppleScript; returns the transcript the judge needs."""
    reference = task["reference"]
    transcript = {"final_text": reference.get("answer", ""), "focus_restored_count": 0}
    if not reference.get("runner_args"):
        return transcript, None
    args = ["--app", task["app"], "--file", os.path.join(task["dir"], "reference_solution.applescript")]
    args += [arg.replace("{docs}", docs_dir) for arg in reference["runner_args"]]
    result = run_runner(args)
    transcript["focus_restored_count"] = int(bool(result.get("focus_restored")))
    return transcript, result


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--task", action="append", help="Only check these task ids")
    args = parser.parse_args()
    summary = {}
    for task in load_tasks(ids=args.task):
        while True:
            wait_until_responsive(task["app"], log=lambda message: print(message, flush=True))
            work = tempfile.mkdtemp(prefix="mac-office-reference-")
            try:
                docs = os.path.join(work, "docs")
                hashes = materialize_docs(task, docs)
                transcript, result = run_reference(task, docs)
                leftovers = staging_leftovers()
                verdict = judge(task, docs, hashes, leftovers, transcript)
            finally:
                shutil.rmtree(work, ignore_errors=True)
            if app_responsive(task["app"]):
                break
            print(f"{task['id']}: {task['app']} wedged during this task; waiting and re-running it", flush=True)
        failed = [check for check in verdict["checks"] if not check["passed"]]
        summary[task["id"]] = {"hard": verdict["hard"], "soft": verdict["soft"], "failed": failed}
        if result and not result.get("success"):
            summary[task["id"]]["runner_error"] = result.get("error")
        print(f"{task['id']}: hard={verdict['hard']} soft={verdict['soft']}", flush=True)
        for check in failed:
            print(f"    FAIL {check['name']}: {check['feedback']}", flush=True)
    os.makedirs(RUNS_DIR, exist_ok=True)
    with open(os.path.join(RUNS_DIR, "reference_selfcheck.json"), "w", encoding="utf-8") as handle:
        json.dump(summary, handle, ensure_ascii=False, indent=2)
    return 0 if all(entry["hard"] == 1 for entry in summary.values()) else 1


if __name__ == "__main__":
    sys.exit(main())
