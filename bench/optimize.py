#!/usr/bin/env python3
"""SkillOpt-style optimization of one skill document, driven step by step by an orchestrator.

Heavy work runs on the Antigravity CLI: rollouts (`rollout.py`) and the reflection that reads
failed trajectories and proposes edits. The orchestrator (a human or Claude Code) reads each
command's compact JSON report and decides the next command:

  baseline                        step0 on every task
  propose --batch N [--lr K]      current skill on train batch N, then proposed edits
  try --edits 0,2 | --document F  candidate from chosen edits (or a hand-written document),
                                  gated on val (strict improvement, no per-task regression)
  decide --decision accept|reject --reason TEXT
  final                           the single test-split evaluation, A/B report, write-back

SkillOpt's disciplines still hold: bounded edits per step, the val gate as the default verdict,
a rejected-edit buffer every later reflection sees, and the test split evaluated once at the end.
Content between <!-- PROTECTED_START --> and <!-- PROTECTED_END --> cannot be edited.
Everything lives in --run-dir (default bench/runs/opt); finished rollouts are reused.
"""

import argparse
import json
import os
import random
import shutil
import subprocess
import sys
import tempfile
import time

from ab_report import write_report
from agy_cli import agy_command, is_quota_error, last_json_line
from model_quota import ModelUnavailable, wait_until_model_available
from office_health import OfficeUnavailable
from rollout import SKILL_COPY_IGNORE, DialogSuspected, run_batch
from runner_cli import RUNS_DIR, SKILL_DIR
from task_catalog import load_tasks

OPTIMIZER_MODEL = "gemini-4-argon-high"
BATCH_SIZE = 6
SEEDS = 2
LR_MAX = 4
PROPOSALS_PER_STEP = 6
PROTECTED_START = "<!-- PROTECTED_START"
PROTECTED_END = "<!-- PROTECTED_END -->"
TRAJECTORY_STEP_LIMIT = 40
STEP_INPUT_CHARS = 1200
STEP_RESULT_CHARS = 500

EDIT_SCHEMA = {
    "type": "object",
    "properties": {
        "diagnoses": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "task_id": {"type": "string"},
                    "failure_class": {"type": "string", "enum": ["SKILL_DEFECT", "EXECUTION_LAPSE"]},
                    "root_cause": {"type": "string"},
                },
                "required": ["task_id", "failure_class", "root_cause"],
            },
        },
        "protected_conflicts": {
            "type": "array",
            "description": "Protected statements a trajectory shows to be wrong; for human review",
            "items": {"type": "string"},
        },
        "edits": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "op": {"type": "string", "enum": ["add", "replace", "delete"]},
                    "anchor": {"type": "string"},
                    "old": {"type": "string"},
                    "new": {"type": "string"},
                    "failure_class": {"type": "string", "enum": ["SKILL_DEFECT", "EXECUTION_LAPSE"]},
                    "addresses": {"type": "array", "items": {"type": "string"}},
                    "rationale": {"type": "string"},
                },
                "required": ["op", "failure_class", "addresses", "rationale"],
            },
        },
    },
    "required": ["diagnoses", "protected_conflicts", "edits"],
}

REFLECTION_PROMPT = """You maintain a skill document that teaches a coding agent (Claude Code) to edit Microsoft Office files on macOS through AppleScript run by a helper script. The agent read the document below while doing real user requests, and some requests failed.

Your job: diagnose the failures and propose small, general edits to the document `{target}` so that future agents succeed on this *kind* of request.

Classify every failure:
- SKILL_DEFECT: the document is wrong, missing a rule, or its recipe leads to the failure. Fix the body text where the rule belongs.
- EXECUTION_LAPSE: the document already states the right rule but the agent did not follow it. Make the rule harder to miss (move or restate it where the agent looks first, e.g. in §0 Ground Rules or right next to the recipe it used), do not duplicate long text.

Rules for edits:
- Propose at most {proposals} edits, best first; only the first {lr} will be applied this step.
- Each edit is one of: `replace` (exact `old` substring -> `new`), `delete` (exact `old` substring), `add` (insert `new` on the lines right after the line containing the exact `anchor` substring). `old`/`anchor` must be copied verbatim from the document and must occur exactly once.
- Text between `<!-- PROTECTED_START` and `<!-- PROTECTED_END -->` is frozen: never quote it as `old`/`anchor`. If a trajectory clearly shows a protected statement is wrong in some situation, do not edit it: describe the conflict in `protected_conflicts` (a human will review it) and you may add a narrowly scoped clarification in the body.
- Write general guidance (a rule, a verified recipe, a check to run). Never mention these specific documents, names, values, tests, graders or benchmarks.
- Only state AppleScript syntax that already appears in the document or that you can see succeed in a trajectory below. Do not invent dictionary terms.
- Keep the document concise; prefer sharpening an existing sentence over adding a section.
- Do not re-propose an edit from the rejected list unless you change it substantially and say why.

## Current document `{target}`

```markdown
{document}
```

## Entry document SKILL.md (context only, read-only)

```markdown
{skill_md}
```

## Failed rollouts

{failures}

## Requests that succeeded with this document (do not break them)

{successes}

## Previously rejected edits (did not improve the validation set)

{rejected}
"""


def copy_skill(source, destination):
    if not os.path.isdir(destination):
        shutil.copytree(source, destination, ignore=SKILL_COPY_IGNORE)
    return destination


def protected_span(text):
    start = text.find(PROTECTED_START)
    end = text.find(PROTECTED_END)
    if start < 0 or end < 0:
        return None
    return start, end + len(PROTECTED_END)


def apply_edit(text, edit):
    """Return the edited text, or raise ValueError explaining why the edit cannot apply."""
    key = "anchor" if edit["op"] == "add" else "old"
    needle = edit.get(key) or ""
    if not needle:
        raise ValueError(f"missing `{key}`")
    if text.count(needle) != 1:
        raise ValueError(f"`{key}` occurs {text.count(needle)} times (must be exactly once)")
    if PROTECTED_START in (edit.get("new") or "") or PROTECTED_END in (edit.get("new") or ""):
        raise ValueError("`new` contains protected markers")
    position = text.index(needle)
    span = protected_span(text)
    if span and position < span[1] and position + len(needle) > span[0]:
        raise ValueError("touches the protected block")
    if edit["op"] == "add":
        line_end = text.find("\n", position + len(needle))
        line_end = len(text) if line_end < 0 else line_end
        if span and span[0] <= line_end < span[1]:
            raise ValueError("would insert inside the protected block")
        return text[:line_end] + "\n" + edit["new"].rstrip("\n") + text[line_end:]
    replacement = "" if edit["op"] == "delete" else edit.get("new", "")
    return text[:position] + replacement + text[position + len(needle):]


def format_failure(record, task):
    lines = [f"### Task `{record['task_id']}` (seed {record['seed']})", "", f"User request: {task['prompt']}", ""]
    lines.append("What was wrong with the result:")
    lines += [f"- {check['feedback']}" for check in record["checks"] if not check["passed"]]
    if record.get("timed_out"):
        lines.append("- The agent ran out of time before finishing.")
    lines += ["", "Agent trajectory (condensed):"]
    for number, step in enumerate(record["steps"][:TRAJECTORY_STEP_LIMIT], start=1):
        lines.append(f"{number}. [{step['tool']}] {step['input'][:STEP_INPUT_CHARS]}")
        if step.get("result"):
            lines.append(f"   -> {step['result'][:STEP_RESULT_CHARS]}")
    lines += ["", f"Agent's final reply: {(record.get('final_text') or '')[:800]}", ""]
    return "\n".join(lines)


def reflect(target, document, skill_md, failures, successes, rejected, lr, log_path):
    prompt = REFLECTION_PROMPT.format(
        target=target, document=document, skill_md=skill_md, lr=lr, proposals=PROPOSALS_PER_STEP,
        failures="\n".join(failures) or "(none)",
        successes="\n".join(f"- {text}" for text in successes) or "(none)",
        rejected="\n".join(
            f"- [{item['op']}] old/anchor: {json.dumps(item.get('old') or item.get('anchor'), ensure_ascii=False)[:300]} "
            f"new: {json.dumps(item.get('new', ''), ensure_ascii=False)[:400]} — rejected because {item['reason']}"
            for item in rejected
        ) or "(none)",
    )
    while True:
        with tempfile.TemporaryDirectory(prefix="mac-office-reflect-") as cwd:
            completed = subprocess.run(
                agy_command(OPTIMIZER_MODEL, prompt, "json", "--json-schema", json.dumps(EDIT_SCHEMA)),
                capture_output=True, text=True, cwd=cwd, timeout=1800,
            )
        answer = last_json_line(completed.stdout) or {}
        if not (answer.get("status") != "SUCCESS" and is_quota_error(answer.get("error"))):
            break
        wait_until_model_available([OPTIMIZER_MODEL])
    with open(log_path, "w", encoding="utf-8") as handle:
        handle.write(prompt + "\n\n=== RESPONSE ===\n" + completed.stdout + completed.stderr)
    if answer.get("status") != "SUCCESS" or not isinstance(answer.get("structured_output"), dict):
        raise RuntimeError(f"reflection failed: {completed.stdout[-500:]} {completed.stderr[-500:]}")
    return answer["structured_output"], 0.0


def mean(values):
    return sum(values) / len(values) if values else 0.0


def per_task_means(records):
    table = {}
    for record in records:
        table.setdefault(record["task_id"], []).append(record["hard"])
    return {task_id: mean(scores) for task_id, scores in sorted(table.items())}


def gate(current, candidate):
    """(accepted, reason) for candidate vs current per-task val means."""
    current_mean, candidate_mean = mean(list(current.values())), mean(list(candidate.values()))
    regressions = [task_id for task_id in current if candidate[task_id] < current[task_id]]
    if regressions:
        return False, f"val regressed on {', '.join(regressions)} (mean {current_mean:.3f} -> {candidate_mean:.3f})"
    if candidate_mean <= current_mean:
        return False, f"val mean did not strictly improve ({current_mean:.3f} -> {candidate_mean:.3f})"
    return True, f"val mean {current_mean:.3f} -> {candidate_mean:.3f}, no per-task regression"


class Run:
    def __init__(self, run_dir, target):
        self.dir = run_dir
        self.target = target
        self.state_path = os.path.join(run_dir, "state.json")
        os.makedirs(run_dir, exist_ok=True)
        if os.path.isfile(self.state_path):
            with open(self.state_path, encoding="utf-8") as handle:
                self.state = json.load(handle)
        else:
            self.state = {"target": target, "current": "step0", "best": "step0", "steps": [], "rejected": [],
                          "stopped": None, "optimizer_cost_usd": 0.0}

    def save(self):
        with open(self.state_path, "w", encoding="utf-8") as handle:
            json.dump(self.state, handle, ensure_ascii=False, indent=1)

    def skill_dir(self, label):
        return os.path.join(self.dir, "skills", label)

    def rollouts(self, label):
        return os.path.join(self.dir, "rollouts", label)

    def log(self, message):
        line = f"[{time.strftime('%H:%M:%S')}] {message}"
        print(line, flush=True)
        with open(os.path.join(self.dir, "optimize.log"), "a", encoding="utf-8") as handle:
            handle.write(line + "\n")

    def rollout_seconds(self):
        """(total wall seconds, rollout count) over every rollout of this run."""
        total, count = 0.0, 0
        for root, _, files in os.walk(os.path.join(self.dir, "rollouts")):
            if "record.json" in files:
                with open(os.path.join(root, "record.json"), encoding="utf-8") as handle:
                    total += json.load(handle).get("wall_sec") or 0.0
                count += 1
        return total, count

    def batch(self, label, tasks):
        return run_batch(tasks, self.skill_dir(label), self.rollouts(label), SEEDS, log=self.log)


def train_batches(tasks):
    """The fixed, shuffled train batches of Word tasks (BATCH_SIZE each)."""
    train = sorted(task_id for task_id, task in tasks.items() if task["split"] == "train" and task["app"] == "word")
    random.Random(0).shuffle(train)
    return [train[i:i + BATCH_SIZE] for i in range(0, len(train), BATCH_SIZE)]


def current_val_means(run, tasks):
    label = run.state["current"]
    val = [task for task in tasks.values() if task["split"] == "val"]
    return per_task_means(run.batch(label, val))


def split_summary(records):
    table = {}
    for record in records:
        hits, total = table.get(record["split"], (0, 0))
        table[record["split"]] = (hits + record["hard"], total + 1)
    return {split: f"{hits}/{total}" for split, (hits, total) in sorted(table.items())}


def command_baseline(run, tasks, args):
    copy_skill(SKILL_DIR, run.skill_dir("step0"))
    records = run.batch("step0", list(tasks.values()))
    print(json.dumps({"hard_by_split": split_summary(records),
                      "per_task": per_task_means(records)}, ensure_ascii=False, indent=1))


def command_propose(run, tasks, args):
    """Roll out the current skill on one train batch and ask the reflection model for edits."""
    step_number = len(run.state["steps"]) + 1
    batch_ids = train_batches(tasks)[args.batch - 1]
    label = run.state["current"]
    records = run.batch(label, [tasks[task_id] for task_id in batch_ids])
    failed = [record for record in records if record["hard"] == 0]
    step_dir = os.path.join(run.dir, "steps", f"step{step_number}")
    os.makedirs(step_dir, exist_ok=True)
    report = {"step": step_number, "current": label, "batch": batch_ids, "train": per_task_means(records)}
    if failed:
        with open(os.path.join(run.skill_dir(label), run.target), encoding="utf-8") as handle:
            document = handle.read()
        with open(os.path.join(run.skill_dir(label), "SKILL.md"), encoding="utf-8") as handle:
            skill_md = handle.read()
        reflection, _ = reflect(
            run.target, document, skill_md,
            [format_failure(record, tasks[record["task_id"]]) for record in failed],
            sorted({tasks[record["task_id"]]["prompt"] for record in records if record["hard"] == 1}),
            run.state["rejected"], args.lr, os.path.join(step_dir, "reflection_prompt.txt"),
        )
        with open(os.path.join(step_dir, "reflection.json"), "w", encoding="utf-8") as handle:
            json.dump(reflection, handle, ensure_ascii=False, indent=1)
        report["diagnoses"] = reflection["diagnoses"]
        report["protected_conflicts"] = reflection.get("protected_conflicts", [])
        report["edits"] = [dict(edit, index=index) for index, edit in enumerate(reflection["edits"])]
    print(json.dumps(report, ensure_ascii=False, indent=1))


def command_try(run, tasks, args):
    """Build stepN_candidate from chosen proposed edits (or a hand-written document) and gate it on val."""
    step_number = len(run.state["steps"]) + 1
    label = run.state["current"]
    candidate = f"step{step_number}_candidate"
    with open(os.path.join(run.skill_dir(label), run.target), encoding="utf-8") as handle:
        text = handle.read()
    applied = []
    if args.document:
        with open(args.document, encoding="utf-8") as handle:
            text = handle.read()
    else:
        with open(os.path.join(run.dir, "steps", f"step{step_number}", "reflection.json"), encoding="utf-8") as handle:
            edits = json.load(handle)["edits"]
        for index in args.edits:
            text = apply_edit(text, edits[index])
            applied.append(edits[index])
    if os.path.isdir(run.skill_dir(candidate)):
        shutil.rmtree(run.skill_dir(candidate))
        shutil.rmtree(run.rollouts(candidate), ignore_errors=True)
    copy_skill(run.skill_dir(label), run.skill_dir(candidate))
    with open(os.path.join(run.skill_dir(candidate), run.target), "w", encoding="utf-8") as handle:
        handle.write(text)
    current = current_val_means(run, tasks)
    val = [task for task in tasks.values() if task["split"] == "val"]
    proposed = per_task_means(run.batch(candidate, val))
    accepted, reason = gate(current, proposed)
    pending = {"step": step_number, "current": label, "candidate": candidate, "applied": applied,
               "val_current": current, "val_candidate": proposed, "gate_passed": accepted, "gate_reason": reason}
    with open(os.path.join(run.dir, "steps", f"step{step_number}", "pending.json"), "w", encoding="utf-8") as handle:
        json.dump(pending, handle, ensure_ascii=False, indent=1)
    print(json.dumps({key: pending[key] for key in ("val_current", "val_candidate", "gate_passed", "gate_reason")},
                     ensure_ascii=False, indent=1))


def command_decide(run, tasks, args):
    """Record the orchestrator's accept/reject decision for the pending candidate."""
    step_number = len(run.state["steps"]) + 1
    with open(os.path.join(run.dir, "steps", f"step{step_number}", "pending.json"), encoding="utf-8") as handle:
        pending = json.load(handle)
    accepted = args.decision == "accept"
    entry = dict(pending, accepted=accepted, reason=args.reason)
    if accepted:
        run.state["current"] = run.state["best"] = pending["candidate"]
    else:
        run.state["rejected"] += [dict(edit, reason=args.reason) for edit in pending["applied"]]
    run.state["steps"].append(entry)
    run.save()
    print(json.dumps({"step": step_number, "accepted": accepted, "current": run.state["current"]}))


def final_evaluation(run):
    """The single test-split evaluation of the best skill, then the A/B report."""
    tasks = load_tasks()
    best = run.state["best"]
    if best != "step0":
        run.log(f"test split: {best}")
        run.batch(best, [task for task in tasks if task["split"] == "test"])
    report = write_report(run.dir, run.state)
    best_doc = os.path.join(run.skill_dir(best), run.target)
    if best != "step0":
        shutil.copy2(best_doc, os.path.join(SKILL_DIR, run.target))
        run.log(f"copied {best}/{run.target} into the working tree")
    run.log(f"report: {report}")


COMMANDS = {
    "baseline": command_baseline,
    "propose": command_propose,
    "try": command_try,
    "decide": command_decide,
    "final": lambda run, tasks, args: final_evaluation(run),
}


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("command", choices=sorted(COMMANDS))
    parser.add_argument("--run-dir", default=os.path.join(RUNS_DIR, "opt"))
    parser.add_argument("--target", default="references/word.md", help="Skill document to train (relative to the skill)")
    parser.add_argument("--batch", type=int, default=1, help="propose: train batch number (1-based)")
    parser.add_argument("--lr", type=int, default=LR_MAX, help="propose: edits the reflection may count on")
    parser.add_argument("--edits", type=lambda text: [int(item) for item in text.split(",") if item], default=[],
                        help="try: comma-separated indices of proposed edits to apply")
    parser.add_argument("--document", help="try: a hand-written candidate document instead of proposed edits")
    parser.add_argument("--decision", choices=["accept", "reject"], help="decide: the orchestrator's verdict")
    parser.add_argument("--reason", default="", help="decide: why")
    args = parser.parse_args()
    run = Run(os.path.abspath(args.run_dir), args.target)
    try:
        COMMANDS[args.command](run, {task["id"]: task for task in load_tasks()}, args)
    except (DialogSuspected, OfficeUnavailable, ModelUnavailable) as error:
        run.log(f"PAUSED: {error}")
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
