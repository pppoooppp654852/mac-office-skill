"""OLE2 container detection.

Detection must happen on the file header *before* any zip parsing: a Word 97-2003 file and a
MIP (Microsoft Purview / RMS) container both start with the OLE2 magic, and a MIP container
also embeds a readable theme zip. ``zipfile`` therefore opens such a file and the required
part check would wrongly report "missing word/document.xml".

The directory is read straight from the container so the reported classification rests on
stream names, not on byte patterns (``EncryptedPackage``, ``WordDocument``, ...).
"""

import os
import struct

from .constants import CONTAINER_HINTS, OLE2_MAGIC

OPENXML = "openxml"
OLE2 = "ole2"
OLE2_MIP = "ole2_mip"

MIP_STREAMS = ("EncryptedPackage", "DRMEncryptedTransform")

# Stream names that identify the legacy binary format behind an ``ole2`` container. The
# owning app refuses to open a legacy payload whose file extension does not match, so these
# also drive the extension used for the staged copy.
LEGACY_STREAMS = {
    "docx": ("WordDocument",),
    "xlsx": ("Workbook", "Book"),
    "pptx": ("PowerPoint Document",),
}
LEGACY_EXTENSIONS = {"docx": ".doc", "xlsx": ".xls", "pptx": ".ppt"}

MIP_MARKERS = (b"EncryptedPackage", b"DRMEncryptedTransform")

ENDOFCHAIN = 0xFFFFFFFE
FREESECT = 0xFFFFFFFF
MAX_SECTORS = 1 << 20


def _read_sectors(handle, sector_size, indices):
    chunks = []
    for index in indices:
        handle.seek((index + 1) * sector_size)
        chunks.append(handle.read(sector_size))
    return b"".join(chunks)


def _chain(fat, start, limit=MAX_SECTORS):
    indices = []
    current = start
    while current not in (ENDOFCHAIN, FREESECT) and current < len(fat) and len(indices) < limit:
        indices.append(current)
        current = fat[current]
    return indices


def ole2_stream_names(path):
    """Return the stream names in an OLE2 (Compound File) container.

    Returns an empty list when the file is not a container this reader understands.
    """
    try:
        with open(path, "rb") as handle:
            header = handle.read(512)
            if len(header) < 512 or header[:8] != OLE2_MAGIC:
                return []
            sector_shift = struct.unpack_from("<H", header, 30)[0]
            fat_sector_count = struct.unpack_from("<I", header, 44)[0]
            first_directory_sector = struct.unpack_from("<I", header, 48)[0]
            if not 7 <= sector_shift <= 20 or not 0 < fat_sector_count < MAX_SECTORS:
                return []
            sector_size = 1 << sector_shift
            difat = [
                value
                for value in struct.unpack_from("<109I", header, 76)
                if value not in (FREESECT, ENDOFCHAIN)
            ][:fat_sector_count]
            fat = []
            for sector in difat:
                raw = _read_sectors(handle, sector_size, [sector])
                fat.extend(struct.unpack_from("<%dI" % (len(raw) // 4), raw))
            directory = _read_sectors(handle, sector_size, _chain(fat, first_directory_sector))
            names = []
            for offset in range(0, len(directory) - 127, 128):
                entry = directory[offset : offset + 128]
                name_length = struct.unpack_from("<H", entry, 64)[0]
                object_type = entry[66]
                if object_type != 2 or not 2 <= name_length <= 64:
                    continue
                name = entry[: name_length - 2].decode("utf-16-le", "replace")
                names.append(name)
            return names
    except (OSError, struct.error):
        return []


def detect_container(path):
    """Return ``openxml``, ``ole2`` or ``ole2_mip`` for ``path``."""
    try:
        with open(path, "rb") as handle:
            if handle.read(8) != OLE2_MAGIC:
                return OPENXML
    except OSError:
        return OPENXML
    names = set(ole2_stream_names(path))
    if any(name in names for name in MIP_STREAMS):
        return OLE2_MIP
    if names:
        return OLE2
    # Unreadable directory: fall back to a byte scan so a MIP wrapper still routes correctly.
    try:
        with open(path, "rb") as handle:
            content = handle.read()
    except OSError:
        return OLE2
    return OLE2_MIP if any(marker in content for marker in MIP_MARKERS) else OLE2


def staged_extension(path, kind):
    """Extension to give the staged copy so the app does not see a format mismatch.

    A legacy payload with a modern extension makes the app show a format-mismatch dialog
    (and swallow the open), so the staged copy gets the extension its streams imply.
    """
    if detect_container(path) == OLE2:
        names = set(ole2_stream_names(path))
        if any(name in names for name in LEGACY_STREAMS.get(kind, ())):
            return LEGACY_EXTENSIONS[kind]
    return os.path.splitext(path)[1]


def container_hint(container):
    return CONTAINER_HINTS.get(container)
