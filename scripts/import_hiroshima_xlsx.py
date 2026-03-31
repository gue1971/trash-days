#!/usr/bin/env python3

import argparse
import json
import re
import zipfile
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

GARBAGE_TYPES = {
    "BURNABLE": {"fullName": "可燃ごみ", "label": "可燃\nごみ"},
    "RECYCLE_PLASTIC": {"fullName": "リサイクルプラ", "label": "リサイ\nプラ"},
    "OTHER_PLASTIC": {"fullName": "その他プラ", "label": "その他\nプラ"},
    "RECYCLABLE_HAZARDOUS": {"fullName": "資源ごみ・有害ごみ", "label": "資源\nごみ"},
    "NON_BURNABLE": {"fullName": "不燃ごみ", "label": "不燃\nごみ"},
    "LARGE_WASTE": {"fullName": "大型ごみ", "label": "大型\nごみ"},
    "NONE": {"fullName": "収集なし", "label": ""},
}

FILL_TO_TYPE = {
    "FF00B400": "LARGE_WASTE",
    "FFFDFD63": "BURNABLE",
    "FFFF6699": "RECYCLE_PLASTIC",
    "FF66CCFF": "OTHER_PLASTIC",
    "FFFF9933": "NON_BURNABLE",
    "FFAFFFAF": "RECYCLABLE_HAZARDOUS",
}

MONTH_TRANSLATION = str.maketrans("０１２３４５６７８９", "0123456789")


def col_to_index(col_letters):
    value = 0
    for char in col_letters:
        value = value * 26 + (ord(char) - 64)
    return value


def index_to_col(index):
    chars = []
    while index:
        index, rem = divmod(index - 1, 26)
        chars.append(chr(65 + rem))
    return "".join(reversed(chars))


def parse_ref(ref):
    match = re.match(r"([A-Z]+)(\d+)$", ref)
    if not match:
        raise ValueError(f"Unsupported cell reference: {ref}")
    return col_to_index(match.group(1)), int(match.group(2))


def fullwidth_int(value):
    return int(str(value).translate(MONTH_TRANSLATION))


def normalized_text(value):
    return re.sub(r"\s+", "", str(value or ""))


def load_shared_strings(archive):
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    strings = []
    for si in root.findall(f"{NS}si"):
        parts = []
        for child in si:
            if child.tag == f"{NS}t":
                parts.append(child.text or "")
            elif child.tag == f"{NS}r":
                parts.extend(node.text or "" for node in child.findall(f"{NS}t"))
        strings.append("".join(parts))
    return strings


def load_style_fill_map(archive):
    root = ET.fromstring(archive.read("xl/styles.xml"))

    fills = []
    for fill in root.find(f"{NS}fills"):
        pattern = fill.find(f"{NS}patternFill")
        fg = pattern.find(f"{NS}fgColor") if pattern is not None else None
        fills.append(fg.attrib if fg is not None else {})

    style_fill_map = []
    for xf in root.find(f"{NS}cellXfs"):
        style_fill_map.append(int(xf.attrib.get("fillId", "0")))

    def resolve_fill(style_id):
        fill = fills[style_fill_map[style_id]]
        return fill.get("rgb") or fill.get("indexed") or fill.get("theme")

    return resolve_fill


def load_cells(archive, shared_strings):
    root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    cells = {}
    for cell in root.iter(f"{NS}c"):
        ref = cell.attrib["r"]
        style_id = int(cell.attrib.get("s", "0"))
        value_type = cell.attrib.get("t")
        value_node = cell.find(f"{NS}v")
        inline_node = cell.find(f"{NS}is")
        value = None

        if value_type == "s" and value_node is not None:
            value = shared_strings[int(value_node.text)]
        elif value_type == "inlineStr" and inline_node is not None:
            value = "".join(node.text or "" for node in inline_node.iter(f"{NS}t"))
        elif value_node is not None:
            value = value_node.text

        cells[ref] = {"value": value, "style_id": style_id}

    return cells


def find_cell_by_value(cells, predicate):
    for ref, cell in cells.items():
        if predicate(cell["value"]):
            return ref, cell["value"]
    return None, None


def detect_block_starts(cells):
    starts = []
    for ref, cell in cells.items():
        value = cell["value"]
        if not isinstance(value, str):
            continue
        if not re.fullmatch(r"[０-９]+月", value):
            continue
        col_index, row_index = parse_ref(ref)
        starts.append({
            "month": fullwidth_int(value[:-1]),
            "start_row": row_index,
            "start_col": col_index - 3,
        })

    starts.sort(key=lambda item: (item["start_row"], item["start_col"]))
    if len(starts) != 12:
        raise ValueError(f"Expected 12 month blocks, found {len(starts)}")
    return starts


def classify_entry(text, fill):
    compact = normalized_text(text)

    if "資源" in compact or "有害" in compact:
        return "RECYCLABLE_HAZARDOUS"
    if "大型" in compact:
        return "LARGE_WASTE"
    if "可燃" in compact:
        return "BURNABLE"
    if "ﾘｻｲｸﾙ" in compact or "リサイクル" in compact:
        return "RECYCLE_PLASTIC"
    if "その他" in compact:
        return "OTHER_PLASTIC"
    if "不燃" in compact:
        return "NON_BURNABLE"
    if compact == "ごみ" and fill == "FFFDFD63":
        return "BURNABLE"
    if compact == "ごみ" and fill == "FFFF9933":
        return "NON_BURNABLE"
    if compact == "プラ" and fill == "FFFF6699":
        return "RECYCLE_PLASTIC"
    if compact == "プラ" and fill == "FF66CCFF":
        return "OTHER_PLASTIC"
    if compact == "(予約制)" and fill == "FF00B400":
        return "LARGE_WASTE"

    return FILL_TO_TYPE.get(fill)


def build_schedule(cells, resolve_fill, fiscal_start_year):
    schedule = {}
    for block in detect_block_starts(cells):
        month = block["month"]
        year = fiscal_start_year if month >= 4 else fiscal_start_year + 1

        for week_index in range(6):
            date_row = block["start_row"] + 2 + week_index * 3

            for day_offset in range(1, 8):
                col_index = block["start_col"] + day_offset
                date_ref = f"{index_to_col(col_index)}{date_row}"
                date_cell = cells.get(date_ref)

                if not date_cell or date_cell["value"] is None:
                    continue

                day = int(float(date_cell["value"]))
                date_key = f"{year}-{month:02d}-{day:02d}"
                schedule[date_key] = []

                for label_row in (date_row + 1, date_row + 2):
                    label_ref = f"{index_to_col(col_index)}{label_row}"
                    label_cell = cells.get(label_ref)
                    if not label_cell:
                        continue

                    garbage_type = classify_entry(
                        label_cell["value"],
                        resolve_fill(label_cell["style_id"]),
                    )
                    if garbage_type and garbage_type not in schedule[date_key]:
                        schedule[date_key].append(garbage_type)

    return schedule


def build_payload(xlsx_path, cells, schedule):
    title_ref, title = find_cell_by_value(cells, lambda value: isinstance(value, str) and value.startswith("令和") and "年度版" in value)
    if not title:
        raise ValueError("Could not locate fiscal year title in workbook")

    area_ref, area_name = find_cell_by_value(cells, lambda value: isinstance(value, str) and "丁目" in value and "東山町" in value)
    district_ref, district = find_cell_by_value(cells, lambda value: isinstance(value, str) and re.fullmatch(r".+区\d+", value))
    contact_ref, contact = find_cell_by_value(cells, lambda value: isinstance(value, str) and "TEL" in value)

    match = re.search(r"令和(\d+)年度版", title)
    if not match:
        raise ValueError(f"Could not parse fiscal year from title: {title}")

    reiwa_year = int(match.group(1))
    fiscal_start_year = 2018 + reiwa_year

    ordered_days = dict(sorted(schedule.items()))
    start_date = next(iter(ordered_days))
    end_date = next(reversed(ordered_days))

    return {
        "version": 1,
        "generatedAt": datetime.now().isoformat(timespec="seconds"),
        "source": {
            "xlsxFile": Path(xlsx_path).name,
            "sheetName": "#2",
            "titleCell": title_ref,
            "areaCell": area_ref,
            "districtCell": district_ref,
            "contactCell": contact_ref,
        },
        "fiscalYear": {
            "japaneseLabel": title,
            "reiwaYear": reiwa_year,
            "startYear": fiscal_start_year,
        },
        "area": {
            "name": area_name,
            "district": district,
            "contact": contact,
        },
        "dateRange": {
            "start": start_date,
            "end": end_date,
        },
        "garbageTypes": GARBAGE_TYPES,
        "days": ordered_days,
    }


def write_outputs(payload, output_json, output_js):
    output_json.parent.mkdir(parents=True, exist_ok=True)
    output_js.parent.mkdir(parents=True, exist_ok=True)

    json_text = json.dumps(payload, ensure_ascii=False, indent=2) + "\n"
    output_json.write_text(json_text, encoding="utf-8")
    output_js.write_text(f"window.TRASH_SCHEDULE_DATA = {json_text};", encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(description="Import Hiroshima trash schedule workbook into app data files.")
    parser.add_argument("xlsx_path", help="Path to Hiroshima schedule workbook (.xlsx)")
    parser.add_argument("--output-json", default="data/schedule.json", help="Path to generated JSON output")
    parser.add_argument("--output-js", default="data/schedule-data.js", help="Path to generated JS output")
    args = parser.parse_args()

    with zipfile.ZipFile(args.xlsx_path) as archive:
        shared_strings = load_shared_strings(archive)
        resolve_fill = load_style_fill_map(archive)
        cells = load_cells(archive, shared_strings)

        title_ref, title = find_cell_by_value(cells, lambda value: isinstance(value, str) and value.startswith("令和") and "年度版" in value)
        match = re.search(r"令和(\d+)年度版", title or "")
        if not match:
            raise ValueError("Could not determine fiscal year from workbook title")
        fiscal_start_year = 2018 + int(match.group(1))

        schedule = build_schedule(cells, resolve_fill, fiscal_start_year)
        payload = build_payload(args.xlsx_path, cells, schedule)

    write_outputs(payload, Path(args.output_json), Path(args.output_js))
    print(f"Wrote {len(payload['days'])} scheduled dates to {args.output_json}")
    print(f"Wrote browser data wrapper to {args.output_js}")


if __name__ == "__main__":
    main()
