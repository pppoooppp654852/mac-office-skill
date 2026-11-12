#!/usr/bin/env python3
"""Run benchmark tasks with `agy -p` (Antigravity CLI) against one candidate copy of the skill.

Isolation: every rollout gets a fresh project directory outside the repository, with the
candidate mounted as the project skill `.agents/skills/mac-office`; while a batch runs, the
user-level agy skill of the same name (a symlink into this working tree) is hidden by
`agy_cli.user_skills_without_mac_office`. Rollouts run strictly one after another.

Records land in <out>/<task id>/seed<k>/{record.json, trajectory.jsonl}; an existing
record is reused, so an interrupted batch resumes where it stopped.
"""

import argparse
import json
import os
import shutil
import signal
import subprocess
import sys
import tempfile
import time

from agy_cli import agy_command, is_quota_error, user_skills_without_mac_office
from judge import judge
from office_health import OfficeUnavailable, app_responsive, wait_until_responsive
from runner_cli import REPO_DIR, STAGING_ROOT, staging_leftovers
from task_catalog import load_tasks, materialize_docs, render_prompt

DEFAULT_MODEL = "gemini-4-argon-medium"
SKILL_COPY_IGNORE = shutil.ignore_patterns("tests", "__pycache__", "*.pyc", ".DS_Store")
# agy has no per-command deny list, so the safety boundary the user set is stated in the prompt.
SAFETY_NOTE = (
    "\n\nNever kill or force-quit Microsoft Office apps (no kill, pkill, killall or Force Quit), "
    "and never change app preferences with `defaults write`."
)
STEP_TEXT_LIMIT = 1500
CONSECUTIVE_TIMEOUT_LIMIT = 3
WEDGE_RETRIES = 2


class ProviderFailure(RuntimeError):
    """The model API refused the rollout; nothing about the skill was measured."""


def is_provider_failure(summary):
    """A run cut short by a quota error, or one that errored before the agent did any work."""
    if not summary.get("is_error"):
        return False
    return is_quota_error(summary.get("error")) or summary.get("tool_calls", 0) == 0


class DialogSuspected(RuntimeError):
    """Several rollouts in a row hit timeouts: Office is probably waiting on a dialog."""


def mount_skill(skill_dir, project_dir):
    target = os.path.join(project_dir, ".agents", "skills", "mac-office")
    shutil.copytree(skill_dir, target, ignore=SKILL_COPY_IGNORE)
    return target


def _input_text(name, parameters):
    if name == "run_command":
        return parameters.get("CommandLine", "")
    if name in ("write_file", "edit_file", "read_file"):
        body = parameters.get("Content") or parameters.get("ReplacementContent") or ""
        return f"{parameters.get('Path')}\n{body}".strip()
    return json.dumps(parameters, ensure_ascii=False)


def summarize_trajectory(path):
    """Tool steps, final answer and usage from an agy stream-json transcript."""
    steps_by_index, final = {}, {}
    with open(path, encoding="utf-8", errors="replace") as handle:
        for line in handle:
            try:
                event = json.loads(line)
            except json.JSONDecodeError:
                continue
            if event.get("event") == "result":
                final = event.get("result") or {}
                continue
            update = event.get("step_update") or {}
            if update.get("step_type") != "tool" or update.get("state") == "ACTIVE":
                continue
            info = update.get("tool_info") or {}
            name = update.get("tool_name")
            steps_by_index[update.get("step_index")] = {
                "tool": name,
                "input": _input_text(name, info.get("parameters") or {}),
                "result": str(info.get("output") or update.get("error") or ""),
                "is_error": update.get("state") != "DONE",
            }
    steps = [steps_by_index[index] for index in sorted(steps_by_index, key=lambda value: value or 0)]
    results = [step["result"] for step in steps]
    usage = final.get("usage") or {}
    return {
        "final_text": final.get("response") or "",
        "is_error": final.get("status") != "SUCCESS",
        "error": final.get("error") or "",
        "num_turns": final.get("num_turns"),
        "duration_ms": round((final.get("duration_seconds") or 0) * 1000),
        "cost_usd": None,
        "tokens": {key: usage.get(key, 0) for key in ("input_tokens", "output_tokens", "thinking_tokens", "cache_read_tokens")},
        "tool_calls": len(steps),
        "runner_calls": sum(1 for step in steps if step["tool"] == "run_command" and "office_runner.py" in step["input"]),
        "focus_restored_count": sum(text.count('"focus_restored": true') for text in results),
        "runner_timeouts": sum(1 for text in results if "timed out after" in text),
        "skill_leak": any(REPO_DIR in step["input"] for step in steps),
        "steps": [
            {
                "tool": step["tool"],
                "input": step["input"][:STEP_TEXT_LIMIT],
                "result": step["result"][:STEP_TEXT_LIMIT],
                "is_error": step["is_error"],
            }
            for step in steps
        ],
    }


def clear_staging_leftovers():
    for name in staging_leftovers():
        path = os.path.join(STAGING_ROOT, name)
        if os.path.isdir(path):
            shutil.rmtree(path, ignore_errors=True)
        else:
            os.unlink(path)


def run_agent(model, prompt, project_dir, trajectory_path, timeout_sec):
    """Run the target agent; on timeout terminate its process group (never an Office app)."""
    command = agy_command(model, prompt + SAFETY_NOTE, "stream-json")
    with open(trajectory_path, "w", encoding="utf-8") as stdout, open(trajectory_path + ".stderr", "w") as stderr:
        process = subprocess.Popen(
            command, cwd=project_dir, stdin=subprocess.DEVNULL, stdout=stdout, stderr=stderr,
            text=True, start_new_session=True,
        )
        try:
            process.wait(timeout=timeout_sec)
            return False
        except subprocess.TimeoutExpired:
            os.killpg(process.pid, signal.SIGTERM)
            try:
                process.wait(timeout=15)
            except subprocess.TimeoutExpired:
                os.killpg(process.pid, signal.SIGKILL)
                process.wait()
            return True


def run_rollout(task, skill_dir, model, seed, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    clear_staging_leftovers()
    project = tempfile.mkdtemp(prefix="mac-office-rollout-")
    started = time.time()
    try:
        mount_skill(skill_dir, project)
        docs = os.path.join(project, "docs")
        hashes = materialize_docs(task, docs)
        trajectory_path = os.path.join(out_dir, "trajectory.jsonl")
        timed_out = run_agent(model, render_prompt(task, docs), project, trajectory_path, task["timeout_minutes"] * 60)
        wall_sec = round(time.time() - started, 1)
        leftovers = staging_leftovers()
        summary = summarize_trajectory(trajectory_path)
        verdict = judge(task, docs, hashes, leftovers, summary)
        if is_provider_failure(summary):
            raise ProviderFailure(summary["error"] or summary["final_text"])
        if timed_out:
            verdict["hard"] = 0
        record = {
            "task_id": task["id"],
            "split": task["split"],
            "seed": seed,
            "model": model,
            "skill_dir": skill_dir,
            "timed_out": timed_out,
            "wall_sec": wall_sec,
            "staging_leftovers": leftovers,
            **verdict,
            **summary,
        }
    finally:
        shutil.rmtree(project, ignore_errors=True)
        clear_staging_leftovers()
    with open(os.path.join(out_dir, "record.json"), "w", encoding="utf-8") as handle:
        json.dump(record, handle, ensure_ascii=False, indent=1)
    return record


def load_record(out_dir):
    path = os.path.join(out_dir, "record.json")
    if not os.path.isfile(path):
        return None
    with open(path, encoding="utf-8") as handle:
        return json.load(handle)


def last_runner_call(record):
    calls = [step["input"] for step in record.get("steps", []) if "office_runner.py" in step["input"]]
    return calls[-1][-400:] if calls else "(none)"


def quarantine(out_dir):
    """Keep a rollout graded while its app was wedged as evidence, out of every score."""
    target = f"{out_dir}.wedged-{time.strftime('%Y%m%d-%H%M%S')}"
    os.rename(out_dir, target)
    os.rename(os.path.join(target, "record.json"), os.path.join(target, "record.wedged.json"))
    return target


def run_checked_rollout(task, skill_dir, model, seed, out_dir, log):
    """Run one rollout with the app healthy before and after; rerun it after a wedge."""
    for _ in range(WEDGE_RETRIES):
        wait_until_responsive(task["app"], log=log)
        record = run_rollout(task, skill_dir, model, seed, out_dir)
        if app_responsive(task["app"]):
            return record
        log(f"  {task['app']} wedged after {task['id']} seed{seed}; last runner call: {last_runner_call(record)}")
        log(f"  quarantined as {quarantine(out_dir)}")
    raise OfficeUnavailable(f"{task['app']} wedged {WEDGE_RETRIES} times on {task['id']} seed{seed}")


def run_batch(tasks, skill_dir, model, out_root, seeds, log=print):
    """Run every task × seed sequentially; returns the records (reusing finished ones)."""
    with user_skills_without_mac_office():
        return _run_batch(tasks, skill_dir, model, out_root, seeds, log)


def _run_batch(tasks, skill_dir, model, out_root, seeds, log):
    records, consecutive_timeouts = [], 0
    for seed in range(1, seeds + 1):
        for task in tasks:
            out_dir = os.path.join(out_root, task["id"], f"seed{seed}")
            record = load_record(out_dir)
            if record is None:
                record = run_checked_rollout(task, skill_dir, model, seed, out_dir, log)
                troubled = record["timed_out"] or record["runner_timeouts"] > 0
                consecutive_timeouts = consecutive_timeouts + 1 if troubled else 0
                log(
                    f"  {task['id']} seed{seed}: hard={record['hard']} soft={record['soft']} "
                    f"{record['wall_sec']}s tools={record['tool_calls']}"
                    + (" TIMEOUT" if record["timed_out"] else "")
                )
                if consecutive_timeouts >= CONSECUTIVE_TIMEOUT_LIMIT:
                    raise DialogSuspected(
                        f"{consecutive_timeouts} rollouts in a row hit timeouts (last: {task['id']} seed{seed}); "
                        "an Office app may be waiting on a dialog"
                    )
            records.append(record)
    return records


def scores_by_task(records):
    """{task id: [hard per seed]} ordered by seed."""
    table = {}
    for record in sorted(records, key=lambda r: (r["task_id"], r["seed"])):
        table.setdefault(record["task_id"], []).append(record["hard"])
    return table


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--skill", required=True, help="Candidate skill directory (a copy of skills/mac-office)")
    parser.add_argument("--out", required=True, help="Output directory for this batch")
    parser.add_argument("--split", choices=["train", "val", "test"], help="Only this split")
    parser.add_argument("--task", action="append", help="Only these task ids")
    parser.add_argument("--seeds", type=int, default=2)
    parser.add_argument("--model", default=DEFAULT_MODEL, help="agy model id (see `agy models`)")
    args = parser.parse_args()
    tasks = load_tasks(split=args.split, ids=args.task)
    records = run_batch(tasks, os.path.abspath(args.skill), args.model, os.path.abspath(args.out), args.seeds)
    with open(os.path.join(args.out, "scores.json"), "w", encoding="utf-8") as handle:
        json.dump(scores_by_task(records), handle, indent=1)
    return 0


if __name__ == "__main__":
    sys.exit(main())
