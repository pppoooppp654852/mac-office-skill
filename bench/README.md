# mac-office skill benchmark

Rule-graded tasks for the `mac-office` skill, used as a regression check after editing the skill
documents or the runner. Not part of the shipped skill.

Requires macOS with Microsoft Word, Excel and PowerPoint, and for agent rollouts the Antigravity
CLI `agy`. Everything generated goes to `bench/runs/` (git-ignored: on a MIP-managed Mac every
fixture Office saves is a protected container for internal use only).

## Layout

| Path | Role |
|---|---|
| `tasks/<id>/task.json` | user prompt, split, fixture name, expected files, graded file, reference runner args |
| `tasks/<id>/build_fixture.applescript` | builds the synthetic fixture inside the real app |
| `tasks/<id>/checks.json` | rule checks for the judge (each failure carries a semantic message) |
| `tasks/<id>/reference_solution.applescript` | hand-written correct solution, used to validate the task |
| `build_fixtures.py` | builds every fixture with `office_runner.py --app … --output-file` |
| `verify_reference_solutions.py` | grades every reference solution; every task must score hard = 1 |
| `judge.py` | rule-based grader (inspect / render / read-only Excel probe; hygiene; focus) |
| `agy_cli.py` | `agy` command line and hiding the user-level `mac-office` skill from agy |
| `rollout.py` | runs tasks with `agy -p` against one skill copy, sequentially |

Tasks: Word w01–w18, Excel x01–x03, PowerPoint p01–p03 (each `task.json` also carries a fixed
train/val/test split).

## Commands

Run from `bench/`, one at a time (Office has a single instance; never run two in parallel).

```bash
python3 build_fixtures.py                       # once; --missing-only to add new tasks
python3 verify_reference_solutions.py           # must print hard=1 for every task
python3 rollout.py --skill ../skills/mac-office --out runs/check --split val --seeds 1
python3 rollout.py --skill ../skills/mac-office --out runs/check --task w06_signature_in_table_fixed_row --model gemini-4-argon-medium
```

Records land in `<out>/<task>/seed<k>/{record.json, trajectory.jsonl}` and `<out>/scores.json`;
finished rollouts are reused, so re-running a command resumes it. A rollout the model provider
refused (quota) stops the batch without a record. Several consecutive timeouts also stop it: an
Office dialog is probably waiting for the user, so answer it in the app and run again.

## Rollout isolation

Each rollout runs in a fresh temporary project with the skill copied to
`.agents/skills/mac-office`. agy also loads user skills from `~/.gemini/antigravity-cli/skills`,
a symlink to `~/.claude/skills` whose `mac-office` entry points at this working tree, and a
project skill cannot override a user skill of the same name. So while a batch runs,
`agy_cli.user_skills_without_mac_office` replaces that symlink with a directory linking every
other user skill and restores it afterwards (a leftover from a crash is restored on the next
run). Claude Code reads `~/.claude/skills` directly and is unaffected. `record.json` flags
`skill_leak` when the agent touched this repository. agy has no per-command deny list, so the
prompt forbids killing Office apps and `defaults write`.

## Grading

hard = every check passes, soft = fraction passed. Besides the task checks, every task checks
that the expected files exist and no other file appeared in the user's folder, originals are
byte-identical (unless an in-place edit was requested), the Office staging folder is empty, and
no runner call reported `focus_restored: true`. Efficiency (tool calls, time, tokens) is only
recorded.
