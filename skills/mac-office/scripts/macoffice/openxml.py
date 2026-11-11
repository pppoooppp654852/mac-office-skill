"""Headless zipfile + ElementTree inspectors for .docx / .xlsx / .pptx.

Every returned collection carries ``<collection>_count`` (the total the document holds) and
``<collection>_truncated`` (True when the returned list was capped).
"""

import re
import xml.etree.ElementTree as ET

from .constants import A_NS, EMU_PER_POINT, P_NS, PKG_REL_NS, R_NS, S_NS, W_NS
from .text import column_index, column_name, trim_text

PREVIEW_ROWS = 5
PREVIEW_CELL_CHARS = 80
SHEET_PREVIEW_ROWS = 20
SHAPE_TEXT_CHARS = 120
SLIDE_PART = r"ppt/slides/slide\d+\.xml"


def inspect_docx(archive, max_items):
    body = ET.fromstring(archive.read("word/document.xml")).find(W_NS + "body")
    paragraphs, tables = [], []
    for child in body:
        if child.tag == W_NS + "p":
            style_element = child.find(f"{W_NS}pPr/{W_NS}pStyle")
            paragraphs.append(
                {
                    "index": len(paragraphs) + 1,
                    "style": style_element.get(W_NS + "val") if style_element is not None else None,
                    "text": trim_text("".join(node.text or "" for node in child.iter(W_NS + "t"))),
                }
            )
        elif child.tag == W_NS + "tbl":
            rows = child.findall(W_NS + "tr")
            preview = [
                [
                    trim_text(
                        "".join(node.text or "" for node in cell.iter(W_NS + "t")),
                        PREVIEW_CELL_CHARS,
                    )
                    for cell in row.findall(W_NS + "tc")
                ]
                for row in rows[:PREVIEW_ROWS]
            ]
            tables.append(
                {
                    "index": len(tables) + 1,
                    "rows": len(rows),
                    "columns": max((len(row.findall(W_NS + "tc")) for row in rows), default=0),
                    "preview": preview,
                }
            )
    return {
        "paragraph_count": len(paragraphs),
        "paragraphs_truncated": len(paragraphs) > max_items,
        "paragraphs": paragraphs[:max_items],
        "table_count": len(tables),
        "tables_truncated": len(tables) > max_items,
        "tables": tables[:max_items],
    }


def shared_strings(archive):
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    return [
        "".join(node.text or "" for node in item.iter(S_NS + "t"))
        for item in root.findall(S_NS + "si")
    ]


def workbook_sheet_targets(archive):
    relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {}
    for relationship in relationships.findall(PKG_REL_NS + "Relationship"):
        target = relationship.get("Target").lstrip("/")
        targets[relationship.get("Id")] = target if target.startswith("xl/") else "xl/" + target
    sheets_root = ET.fromstring(archive.read("xl/workbook.xml"))
    return [
        (sheet.get("name"), targets.get(sheet.get(R_NS + "id")))
        for sheet in sheets_root.iter(S_NS + "sheet")
    ]


def sheet_preview(archive, sheet_path, strings, max_items):
    if not sheet_path or sheet_path not in archive.namelist():
        return {"dimension": None, "rows": [], "truncated": False}
    root = ET.fromstring(archive.read(sheet_path))
    rows, max_row, max_column, truncated = [], 0, 0, False
    row_limit = min(max_items, SHEET_PREVIEW_ROWS)
    for row in root.iter(S_NS + "row"):
        max_row = max(max_row, int(row.get("r") or 0))
        if len(rows) >= row_limit:
            truncated = True
            break
        values = {}
        for cell in row.findall(S_NS + "c"):
            index = column_index(cell.get("r")) or (len(values) + 1)
            max_column = max(max_column, index)
            value_node = cell.find(S_NS + "v")
            if cell.get("t") == "s" and value_node is not None:
                try:
                    value = strings[int(value_node.text)]
                except (ValueError, IndexError):
                    value = value_node.text
            elif cell.get("t") == "inlineStr":
                value = "".join(node.text or "" for node in cell.iter(S_NS + "t"))
            else:
                value = value_node.text if value_node is not None else None
            values[index] = value
        row_values = [values.get(index) for index in range(1, max(values, default=0) + 1)]
        while row_values and row_values[-1] is None:
            row_values.pop()
        rows.append(row_values)
    dimension = root.find(S_NS + "dimension")
    used_range = dimension.get("ref") if dimension is not None else None
    if used_range is None and max_row and max_column:
        used_range = f"A1:{column_name(max_column)}{max_row}"
    return {"dimension": used_range, "rows": rows, "truncated": truncated}


def inspect_xlsx(archive, max_items):
    strings = shared_strings(archive)
    worksheets = []
    for name, target in workbook_sheet_targets(archive):
        preview = sheet_preview(archive, target, strings, max_items)
        worksheets.append(
            {
                "name": name,
                "used_range": preview["dimension"],
                "preview": preview["rows"],
                "truncated": preview["truncated"],
            }
        )
    return {
        "sheet_count": len(worksheets),
        "sheets_truncated": len(worksheets) > max_items,
        "worksheets": worksheets[:max_items],
    }


def slide_targets(archive):
    relationships = ET.fromstring(archive.read("ppt/_rels/presentation.xml.rels"))
    target_by_id = {
        relationship.get("Id"): "ppt/" + relationship.get("Target").lstrip("/").removeprefix("ppt/")
        for relationship in relationships.findall(PKG_REL_NS + "Relationship")
    }
    presentation = ET.fromstring(archive.read("ppt/presentation.xml"))
    targets = []
    slide_list = presentation.find(P_NS + "sldIdLst")
    for slide_id in (slide_list if slide_list is not None else []):
        target = target_by_id.get(slide_id.get(R_NS + "id"))
        if target:
            targets.append(target)
    if not targets:
        targets = sorted(
            (name for name in archive.namelist() if re.fullmatch(SLIDE_PART, name)),
            key=lambda name: int(re.search(r"(\d+)", name).group(1)),
        )
    return targets


def shape_records(slide_root):
    records = []
    for shape in slide_root.iter():
        if shape.tag not in (P_NS + "sp", P_NS + "pic", P_NS + "graphicFrame"):
            continue
        name_node = shape.find(f".//{P_NS}cNvPr")
        offset = shape.find(f"{P_NS}spPr/{A_NS}xfrm/{A_NS}off")
        extent = shape.find(f"{P_NS}spPr/{A_NS}xfrm/{A_NS}ext")
        text = " ".join(node.text or "" for node in shape.iter(A_NS + "t"))
        records.append(
            {
                "index": len(records) + 1,
                "id": int(name_node.get("id")) if name_node is not None and name_node.get("id") else None,
                "name": name_node.get("name") if name_node is not None else None,
                "kind": shape.tag.split("}")[-1],
                "left_pt": round(int(offset.get("x")) / EMU_PER_POINT, 2) if offset is not None else None,
                "top_pt": round(int(offset.get("y")) / EMU_PER_POINT, 2) if offset is not None else None,
                "width_pt": round(int(extent.get("cx")) / EMU_PER_POINT, 2) if extent is not None else None,
                "height_pt": round(int(extent.get("cy")) / EMU_PER_POINT, 2) if extent is not None else None,
                "text": trim_text(text, SHAPE_TEXT_CHARS),
            }
        )
    return records


def inspect_pptx(archive, max_items):
    slides = []
    for index, target in enumerate(slide_targets(archive), start=1):
        records = shape_records(ET.fromstring(archive.read(target)))
        slides.append(
            {
                "index": index,
                "shape_count": len(records),
                "shapes_truncated": len(records) > max_items,
                "shapes": records[:max_items],
            }
        )
    return {
        "slide_count": len(slides),
        "slides_truncated": len(slides) > max_items,
        "slides": slides[:max_items],
    }
