"""Load benchmark tasks (bench/tasks/<id>/) and copy their fixtures into a working folder."""

import hashlib
import json
import os
import shutil

from runner_cli import FIXTURES_DIR, TASKS_DIR

SPLITS = ("train", "val", "test")


def load_tasks(split=None, ids=None):
    tasks = []
    for name in sorted(os.listdir(TASKS_DIR)):
        task_dir = os.path.join(TASKS_DIR, name)
        task_file = os.path.join(task_dir, "task.json")
        if not os.path.isfile(task_file):
            continue
        with open(task_file, encoding="utf-8") as handle:
            task = json.load(handle)
        with open(os.path.join(task_dir, "checks.json"), encoding="utf-8") as handle:
            task["checks"] = json.load(handle)["checks"]
        task["dir"] = task_dir
        if split and task["split"] != split:
            continue
        if ids and task["id"] not in ids:
            continue
        tasks.append(task)
    return tasks


def fixture_dir(task):
    return os.path.join(FIXTURES_DIR, task["id"])


def fixture_docs_dir(task):
    return os.path.join(fixture_dir(task), "docs")


def _fixture_json(task, name):
    path = os.path.join(fixture_dir(task), name)
    if not os.path.isfile(path):
        return None
    with open(path, encoding="utf-8") as handle:
        return json.load(handle)


def before_render(task):
    return _fixture_json(task, "before_render.json")


def before_inspect(task):
    return _fixture_json(task, "before_inspect.json")


def md5(path):
    digest = hashlib.md5()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def materialize_docs(task, docs_dir):
    """Copy the task's pristine fixture files into ``docs_dir``; returns {name: md5}."""
    source = fixture_docs_dir(task)
    if not os.path.isdir(source):
        raise RuntimeError(f"fixture for {task['id']} is missing; run bench/build_fixtures.py first")
    os.makedirs(docs_dir, exist_ok=True)
    hashes = {}
    for name in sorted(os.listdir(source)):
        shutil.copy2(os.path.join(source, name), os.path.join(docs_dir, name))
        hashes[name] = md5(os.path.join(docs_dir, name))
    return hashes


def render_prompt(task, docs_dir):
    return task["prompt"].replace("{docs}", docs_dir)
