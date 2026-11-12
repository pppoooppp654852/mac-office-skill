# mac-office skill benchmark and optimizer

Rule-graded tasks for the `mac-office` skill and a SkillOpt-style loop that trains one skill
document against them. Not part of the shipped skill.

Requires macOS with Microsoft Word, Excel and PowerPoint, the Antigravity CLI `agy` (target
`gemini-4-argon-medium`, optimizer `gemini-4-argon-high`), and (for the
final statistics) a read-only SkillOpt checkout (`$SKILLOPT_DIR`, default `~/dev/SkillOpt`).
Everything generated goes to `bench/runs/` (git-ignored: on a MIP-managed Mac every fixture
Office saves is a protected container for internal use only).

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
| `agy_cli.py` | `agy` command line, quota probe, and hiding the user-level `mac-office` skill from agy |
| `rollout.py` | runs tasks with `agy -p` against one candidate skill copy, sequentially |
| `optimize.py` | orchestrator commands: baseline, propose (reflection), try (val gate), decide, final |
| `ab_report.py` | report with evalkit paired A/B, A/A and efficiency |

## Splits (fixed)

| split | tasks |
|---|---|
| train (12) | w01–w10, x01, p01 |
| val (6) | w11–w14, x02, p02 |
| test (6) | w15–w18, x03, p03 |

Only Word train tasks feed the reflection; Excel and PowerPoint tasks are regression checks.

## Commands

Run from `bench/`, one at a time (Office has a single instance; never run two in parallel).

```bash
python3 build_fixtures.py                       # once; --missing-only to add new tasks
python3 verify_reference_solutions.py           # must print hard=1 for every task
python3 rollout.py --skill ../skills/mac-office --out runs/adhoc --task w01_form_fill_paragraph_fields --seeds 1
python3 optimize.py baseline                    # step0 on all tasks
python3 optimize.py propose --batch 1           # current skill on train batch 1, then proposed edits
python3 optimize.py try --edits 0,2             # candidate from chosen edits, gated on val
python3 optimize.py decide --decision accept --reason "..."
python3 optimize.py final                       # test split once, A/B report, write-back
python3 ab_report.py --run-dir runs/opt         # rebuild the report
```

`optimize.py` writes `runs/opt/{state.json, optimize.log, skills/, rollouts/, steps/, report/}`.
Each command prints a compact JSON report for the orchestrator, who decides the next command.
Finished rollouts are reused, so re-running a command resumes it; after a pause (several
consecutive timeouts, which usually means an Office dialog is waiting for the user) answer the
dialog in the app and run it again. `final` copies the best document into `skills/mac-office/`
as an uncommitted change for review.

## Rollout isolation

Each rollout runs in a fresh temporary project with the candidate copied to
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
