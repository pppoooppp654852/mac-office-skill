#!/usr/bin/env python3
"""Write the optimization report: baseline per task/seed, step history, and the test A/B.

The paired statistics come from SkillOpt's evalkit (McNemar + bootstrap CI), run as
`python -m skillopt_sleep.evalkit` with PYTHONPATH pointing at a SkillOpt checkout
($SKILLOPT_DIR, default ~/dev/SkillOpt); the checkout is only read.

A/A noise floor: the baseline's two seeds are two independent runs of the same skill, so
seed 1 vs seed 2 is a real A/A comparison (on the test split and on all tasks).
"""

import argparse
import difflib
import json
import os
import subprocess
import sys

SKILLOPT_DIR = os.path.expanduser(os.environ.get("SKILLOPT_DIR", "~/dev/SkillOpt"))


def load_records(rollout_dir):
    records = []
    for root, _, files in os.walk(rollout_dir):
        if "record.json" in files:
            with open(os.path.join(root, "record.json"), encoding="utf-8") as handle:
                records.append(json.load(handle))
    return sorted(records, key=lambda r: (r["task_id"], r["seed"]))


def outcomes(records, seed=None):
    table = {}
    for record in records:
        if seed is None or record["seed"] == seed:
            table.setdefault(record["task_id"], []).append(record["hard"])
    return {task_id: scores if seed is None else scores[0] for task_id, scores in table.items()}


def evalkit(report_dir, name, manifest, a, b=None):
    paths = {}
    for label, payload in (("manifest", manifest), ("a", a), ("b", b)):
        if payload is not None:
            paths[label] = os.path.join(report_dir, f"{name}_{label}.json")
            with open(paths[label], "w", encoding="utf-8") as handle:
                json.dump(payload, handle, indent=1)
    command = [sys.executable, "-m", "skillopt_sleep.evalkit", "--manifest", paths["manifest"], "--a", paths["a"], "--json"]
    command += ["--b", paths["b"]] if b is not None else ["--aa"]
    completed = subprocess.run(command, capture_output=True, text=True, env=dict(os.environ, PYTHONPATH=SKILLOPT_DIR))
    if completed.returncode != 0:
        return {"error": (completed.stderr or completed.stdout).strip()}
    return json.loads(completed.stdout)


def efficiency(records):
    if not records:
        return {}
    count = len(records)

    def average(key):
        return round(sum(r.get(key) or 0 for r in records) / count, 2)

    return {
        "rollouts": count,
        "hard": average("hard"),
        "soft": round(sum(r["soft"] for r in records) / count, 4),
        "tool_calls": average("tool_calls"),
        "runner_calls": average("runner_calls"),
        "wall_sec": average("wall_sec"),
        "cost_usd": round(sum(r.get("cost_usd") or 0 for r in records) / count, 4),
        "output_tokens": round(sum(r["tokens"]["output_tokens"] for r in records) / count),
        "input_tokens_incl_cache": round(
            sum(r["tokens"]["input_tokens"] + r["tokens"]["cache_read_input_tokens"] + r["tokens"]["cache_creation_input_tokens"] for r in records) / count
        ),
        "focus_restored_total": sum(r["focus_restored_count"] for r in records),
        "timeouts": sum(1 for r in records if r["timed_out"]),
        "runner_timeouts": sum(r["runner_timeouts"] for r in records),
        "skill_leaks": sum(1 for r in records if r.get("skill_leak")),
    }


def describe(result):
    if "error" in result:
        return f"evalkit error: {result['error']}"
    boot = result["bootstrap"]
    text = (
        f"rate A={result['rate_a']:.3f} B={result['rate_b']:.3f} Δ={result['delta']:+.3f}, "
        f"95% bootstrap CI [{boot['low']:+.3f}, {boot['high']:+.3f}]"
    )
    if result.get("mcnemar"):
        m = result["mcnemar"]
        text += f", McNemar a_only={m['a_only']} b_only={m['b_only']} p_exact={m['p_exact']:.3g}"
    return text


def write_report(run_dir, state):
    report_dir = os.path.join(run_dir, "report")
    os.makedirs(report_dir, exist_ok=True)
    best = state["best"]
    baseline = load_records(os.path.join(run_dir, "rollouts", "step0"))
    best_records = load_records(os.path.join(run_dir, "rollouts", best))
    test_ids = sorted({r["task_id"] for r in baseline if r["split"] == "test"})
    all_ids = sorted({r["task_id"] for r in baseline})
    baseline_test = [r for r in baseline if r["split"] == "test"]
    best_test = [r for r in best_records if r["split"] == "test"]

    stats = {
        "aa_test_seed1_vs_seed2": evalkit(report_dir, "aa_test", test_ids,
                                          {k: v for k, v in outcomes(baseline_test, 1).items()},
                                          {k: v for k, v in outcomes(baseline_test, 2).items()}),
        "aa_all_seed1_vs_seed2": evalkit(report_dir, "aa_all", all_ids, outcomes(baseline, 1), outcomes(baseline, 2)),
        "aa_identity_smoke": evalkit(report_dir, "aa_identity", test_ids, outcomes(baseline_test)),
    }
    if best != "step0" and best_test:
        stats["ab_test"] = evalkit(report_dir, "ab_test", test_ids, outcomes(baseline_test), outcomes(best_test))
        stats["ab_test_seed1"] = evalkit(report_dir, "ab_test_seed1", test_ids, outcomes(baseline_test, 1), outcomes(best_test, 1))
        stats["ab_test_seed2"] = evalkit(report_dir, "ab_test_seed2", test_ids, outcomes(baseline_test, 2), outcomes(best_test, 2))

    target = state["target"]
    with open(os.path.join(run_dir, "skills", "step0", target), encoding="utf-8") as handle:
        before = handle.readlines()
    with open(os.path.join(run_dir, "skills", best, target), encoding="utf-8") as handle:
        after = handle.readlines()
    diff = "".join(difflib.unified_diff(before, after, f"step0/{target}", f"{best}/{target}"))
    with open(os.path.join(report_dir, "best.diff"), "w", encoding="utf-8") as handle:
        handle.write(diff)

    summary = {
        "best": best,
        "stopped": state.get("stopped"),
        "optimizer_cost_usd": round(state.get("optimizer_cost_usd", 0.0), 2),
        "efficiency": {
            "baseline_all": efficiency(baseline),
            **{f"baseline_{split}": efficiency([r for r in baseline if r["split"] == split]) for split in ("train", "val", "test")},
            "best_test": efficiency(best_test) if best != "step0" else None,
        },
        "baseline_per_task": {
            task_id: [{"seed": r["seed"], "hard": r["hard"], "soft": r["soft"],
                       "failed": [c["name"] for c in r["checks"] if not c["passed"]]}
                      for r in baseline if r["task_id"] == task_id]
            for task_id in all_ids
        },
        "steps": state["steps"],
        "stats": stats,
    }
    with open(os.path.join(report_dir, "report.json"), "w", encoding="utf-8") as handle:
        json.dump(summary, handle, ensure_ascii=False, indent=1)

    lines = ["# mac-office skill optimization report", "", f"- best skill: `{best}`", f"- stopped: {state.get('stopped')}",
             f"- optimizer cost: ${summary['optimizer_cost_usd']}", "", "## Baseline (step0) per task and seed", "",
             "| task | split | seed1 hard/soft | seed2 hard/soft | failed checks |", "|---|---|---|---|---|"]
    split_of = {r["task_id"]: r["split"] for r in baseline}
    for task_id, rows in summary["baseline_per_task"].items():
        cells = {row["seed"]: f"{row['hard']} / {row['soft']:.2f}" for row in rows}
        failed = sorted({name for row in rows for name in row["failed"]})
        lines.append(f"| {task_id} | {split_of[task_id]} | {cells.get(1, '-')} | {cells.get(2, '-')} | {', '.join(failed)} |")
    lines += ["", "## Efficiency", "", "| set | rollouts | hard | soft | tools | runner calls | wall s | $ | focus_restored | timeouts |",
              "|---|---|---|---|---|---|---|---|---|---|"]
    for name, eff in summary["efficiency"].items():
        if eff:
            lines.append(f"| {name} | {eff['rollouts']} | {eff['hard']} | {eff['soft']} | {eff['tool_calls']} | {eff['runner_calls']} | "
                         f"{eff['wall_sec']} | {eff['cost_usd']} | {eff['focus_restored_total']} | {eff['timeouts']} |")
    lines += ["", "## Steps", ""]
    for step in state["steps"]:
        lines.append(f"- step {step['step']} (epoch {step['epoch']}, lr {step['lr']}, current {step['current']}): "
                     f"train hard {step['train_hard']:.2f}, failures {step['train_failures']}; "
                     f"{'ACCEPTED' if step['accepted'] else 'rejected'} — {step.get('reason')}")
    lines += ["", "## Statistics (evalkit)", ""]
    for name, result in stats.items():
        lines.append(f"- {name}: {describe(result)}")
    lines += ["", "## Diff of the best skill document", "", "```diff", diff or "(no change)", "```", ""]
    path = os.path.join(report_dir, "report.md")
    with open(path, "w", encoding="utf-8") as handle:
        handle.write("\n".join(lines))
    return path


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--run-dir", required=True)
    args = parser.parse_args()
    with open(os.path.join(args.run_dir, "state.json"), encoding="utf-8") as handle:
        state = json.load(handle)
    print(write_report(os.path.abspath(args.run_dir), state))
    return 0


if __name__ == "__main__":
    sys.exit(main())
