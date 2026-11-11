"""Text and spreadsheet-column helpers."""

import re


def trim_text(value, limit=200):
    """Collapse whitespace and truncate with an ellipsis."""
    collapsed = " ".join((value or "").split())
    return collapsed if len(collapsed) <= limit else collapsed[: limit - 1] + "…"


def column_index(cell_reference):
    """``"AB12"`` -> 28."""
    letters = re.match(r"[A-Z]+", cell_reference or "")
    if not letters:
        return None
    index = 0
    for character in letters.group(0):
        index = index * 26 + (ord(character) - ord("A") + 1)
    return index


def column_name(index):
    """28 -> ``"AB"``."""
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(ord("A") + remainder) + name
    return name
