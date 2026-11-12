#!/usr/bin/env python3
"""Build every task's synthetic fixture with real Office through office_runner.py --app.

Output per task in bench/runs/fixtures/<id>/:
  docs/                 the files handed to the agent (document + generated image assets)
  before_render.json    --render of a pristine Word document (layout reference for the judge;
                        PowerPoint writes no PDF for a MIP-protected deck, so decks have none)
  before_inspect.json   --inspect of the pristine document (structure reference)

Fixtures are never hand-assembled OOXML: each document is written by the owning app from
bench/tasks/<id>/build_fixture.applescript, so on a MIP-managed Mac it is the same kind of
container a user's own files are.
"""

import argparse
import json
import os
import shutil
import struct
import sys
import zlib

from runner_cli import run_runner
from task_catalog import fixture_dir, fixture_docs_dir, load_tasks


def write_png(path, width, height, rgb):
    """A striped RGB PNG, so a picture is visibly a picture when rendered."""
    rows = []
    for y in range(height):
        row = bytearray([0])
        for x in range(width):
            stripe = ((x + y) // 12) % 2
            row.extend(rgb if stripe else (255, 255, 255))
        rows.append(bytes(row))

    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)

    header = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    with open(path, "wb") as handle:
        handle.write(b"\x89PNG\r\n\x1a\n")
        handle.write(chunk(b"IHDR", header))
        handle.write(chunk(b"IDAT", zlib.compress(b"".join(rows), 9)))
        handle.write(chunk(b"IEND", b""))


def build(task):
    out_dir = fixture_dir(task)
    docs = fixture_docs_dir(task)
    if os.path.isdir(out_dir):
        shutil.rmtree(out_dir)
    os.makedirs(docs)
    for asset in task.get("assets", []):
        write_png(os.path.join(docs, asset["name"]), asset["width"], asset["height"], asset["rgb"])

    built_name = task.get("build_as", task["fixture"])
    built_path = os.path.join(docs, built_name)
    result = run_runner(
        ["--app", task["app"], "--output-file", built_path,
         "--file", os.path.join(task["dir"], "build_fixture.applescript"), "--timeout", "120"]
    )
    if not result.get("success") or not os.path.isfile(built_path):
        raise RuntimeError(f"{task['id']}: fixture build failed: {result.get('error')} {result.get('error_hint')}")
    fixture_path = os.path.join(docs, task["fixture"])
    if built_name != task["fixture"]:
        os.rename(built_path, fixture_path)

    if fixture_path.endswith(".docx"):
        render = run_runner(["--render", fixture_path, "--output-dir", os.path.join(out_dir, "before_render"), "--dpi", "60"])
        if not render.get("success"):
            raise RuntimeError(f"{task['id']}: before render failed: {render.get('error')}")
        with open(os.path.join(out_dir, "before_render.json"), "w", encoding="utf-8") as handle:
            json.dump(render, handle, ensure_ascii=False, indent=1)
    inspect = run_runner(["--inspect", fixture_path])
    with open(os.path.join(out_dir, "before_inspect.json"), "w", encoding="utf-8") as handle:
        json.dump(inspect, handle, ensure_ascii=False, indent=1)
    return inspect.get("container")


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--task", action="append", help="Only build these task ids")
    parser.add_argument("--missing-only", action="store_true", help="Skip tasks whose fixture exists")
    args = parser.parse_args()
    failures = 0
    for task in load_tasks(ids=args.task):
        if args.missing_only and os.path.isdir(fixture_docs_dir(task)):
            continue
        try:
            container = build(task)
            print(f"built {task['id']} ({container})", flush=True)
        except RuntimeError as error:
            failures += 1
            print(f"FAILED {error}", file=sys.stderr, flush=True)
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
