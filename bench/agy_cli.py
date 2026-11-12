"""Run the Antigravity CLI (`agy`) headlessly for rollouts.

Skill isolation: agy loads user skills from ``~/.gemini/antigravity-cli/skills``, a symlink to
``~/.claude/skills`` whose ``mac-office`` entry points at this working tree, and a project
skill cannot override a user skill of the same name. While rollouts run,
:func:`user_skills_without_mac_office` swaps that symlink for a directory linking every other
user skill, and restores the symlink afterwards, also after a crash on the next run. Claude
Code reads ``~/.claude/skills`` directly and is never affected.
"""

import os
import shutil
import signal
from contextlib import contextmanager

AGY_SKILLS_LINK = os.path.expanduser("~/.gemini/antigravity-cli/skills")
ISOLATION_MARKER = ".mac-office-bench-original-target"
HIDDEN_SKILL = "mac-office"
QUOTA_MARKERS = ("quota", "rate limit", "429", "resource_exhausted", "usage limit")


def agy_command(model, prompt, output_format, *extra):
    return [
        "agy",
        "--model", model,
        "--output-format", output_format,
        "--dangerously-skip-permissions",
        "--print-timeout", "0s",
        *extra,
        f"-p={prompt}",
    ]


def is_quota_error(message):
    lowered = (message or "").lower()
    return any(marker in lowered for marker in QUOTA_MARKERS)


def _restore(original_target):
    if os.path.isdir(AGY_SKILLS_LINK) and not os.path.islink(AGY_SKILLS_LINK):
        shutil.rmtree(AGY_SKILLS_LINK)
    if not os.path.lexists(AGY_SKILLS_LINK):
        os.symlink(original_target, AGY_SKILLS_LINK)


def _exit_on_sigterm(signum, frame):
    raise SystemExit(128 + signum)


def _restore_leftover_isolation():
    marker = os.path.join(AGY_SKILLS_LINK, ISOLATION_MARKER)
    if not os.path.islink(AGY_SKILLS_LINK) and os.path.isfile(marker):
        with open(marker, encoding="utf-8") as handle:
            _restore(handle.read().strip())


@contextmanager
def user_skills_without_mac_office():
    """Hide the user-level mac-office skill from agy for the duration of the block."""
    _restore_leftover_isolation()
    if not os.path.islink(AGY_SKILLS_LINK):
        raise RuntimeError(f"{AGY_SKILLS_LINK} is not a symlink; refusing to modify it")
    original_target = os.readlink(AGY_SKILLS_LINK)
    previous_sigterm = signal.signal(signal.SIGTERM, _exit_on_sigterm)
    os.unlink(AGY_SKILLS_LINK)
    try:
        os.mkdir(AGY_SKILLS_LINK)
        with open(os.path.join(AGY_SKILLS_LINK, ISOLATION_MARKER), "w", encoding="utf-8") as handle:
            handle.write(original_target)
        for name in os.listdir(original_target):
            if name != HIDDEN_SKILL and not name.startswith("."):
                os.symlink(os.path.join(original_target, name), os.path.join(AGY_SKILLS_LINK, name))
        yield
    finally:
        _restore(original_target)
        signal.signal(signal.SIGTERM, previous_sigterm)
