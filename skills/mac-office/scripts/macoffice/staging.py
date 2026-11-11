"""Per-run staging inside the Office Group Container.

macOS asks the user to grant file access whenever an Office app touches a path outside its
own sandbox container, so every file handed to an AppleScript is first copied into a unique
per-run subdirectory of the shared staging directory and removed again afterwards.
"""

import os
import re
import secrets
import shutil
import tempfile
import time

from .constants import GROUP_CONTAINER_DIR


def staging_root():
    """The shared staging directory (created on demand, falls back to the temp directory)."""
    base = os.path.dirname(GROUP_CONTAINER_DIR)
    if os.path.isdir(base):
        os.makedirs(GROUP_CONTAINER_DIR, exist_ok=True)
        return GROUP_CONTAINER_DIR
    fallback = os.path.join(tempfile.gettempdir(), "mac-office-staging")
    os.makedirs(fallback, exist_ok=True)
    return fallback


def safe_name(name, limit=60):
    """Reduce ``name`` to a filesystem-safe single path component, keeping its extension."""
    stem, extension = os.path.splitext(name)
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", stem).strip("._")[:limit] or "file"
    return cleaned + re.sub(r"[^A-Za-z0-9._-]+", "", extension)


class StageDir:
    """A unique staging subdirectory for one runner invocation, removed on exit.

    ``path`` is what ``MAC_OFFICE_STAGE_DIR`` reports, so scripts can drop intermediate
    files there without ever touching the user's directories.
    """

    def __init__(self, label="run"):
        self.label = safe_name(label or "run", limit=40)
        self.path = None
        self._used_names = set()

    def __enter__(self):
        root = staging_root()
        self.path = tempfile.mkdtemp(prefix=self.label + "-", dir=root)
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        if self.path:
            _remove_tree(self.path)
        self.path = None
        return False

    def unique_name(self, name):
        """A single path component inside the stage dir, disambiguated on collision."""
        candidate = safe_name(name)
        stem, extension = os.path.splitext(candidate)
        counter = 1
        while candidate in self._used_names:
            counter += 1
            candidate = f"{stem}-{counter}{extension}"
        self._used_names.add(candidate)
        return candidate

    def path_for(self, name):
        if self.path is None:
            raise RuntimeError("StageDir is not active; use it as a context manager")
        return os.path.join(self.path, self.unique_name(name))

    def stage(self, source_path, name=None):
        """Copy ``source_path`` into the stage directory and return the staged path."""
        target = self.path_for(name or os.path.basename(source_path))
        shutil.copy2(source_path, target)
        return target

    def stage_uniquely(self, source_path, extension):
        """Stage a copy under a name no other open document can have.

        Runner-owned scripts look up the copy they opened by document name, so the name gets
        a random token and can never match a document the user already has open.
        """
        stem = safe_name(os.path.splitext(os.path.basename(source_path))[0], limit=40)
        return self.stage(source_path, f"{stem}-{secrets.token_hex(4)}{extension}")


def _remove_tree(path, attempts=10):
    """Delete the stage directory, tolerating a short-lived lock from the Office app.

    The apps can still hold a file handle for a moment after ``close``, which makes
    ``shutil.rmtree`` fail; retry briefly instead of leaving the run's files behind.
    """
    for attempt in range(attempts):
        shutil.rmtree(path, ignore_errors=True)
        if not os.path.exists(path):
            return
        time.sleep(0.05 * (attempt + 1))


def hfs_to_posix(path):
    """Convert an HFS path ("Disk:Users:me:a.docx") to POSIX ("/Users/me/a.docx")."""
    if not path or path in ("missing value", "/missing value"):
        return None
    if path.startswith("/"):
        return path
    return "/" + path.replace(":", "/").split("/", 1)[-1]
