"""
Check the user/lookup ids and recurrence fields of extracted generic lists.

For each list of the given sites: whether it was extracted with the current
format (.state.json "format"), and for each visible User/UserMulti/Lookup/
LookupMulti column, how many items carry <Name>Id and with which value types.
For calendars, items with fRecurrence versus items with RecurrenceData.
Also lists the dropped_columns found in run_summary.json for these sites.
Only names, counts and types are printed.

Usage (in the container, from the project root):
    python src/tools/check_list_id_fields.py hr team projects
"""

import json
import os
import sys
from collections import Counter

SRC_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, SRC_DIR)
sys.path.insert(0, os.path.join(SRC_DIR, "extraction"))
from extract_metadata import (  # noqa: E402
    GENERIC_LIST_FORMAT,
    ID_FIELD_EXCLUDED,
    ID_FIELD_TYPES,
)


def value_type(value):
    if isinstance(value, list):
        inner = sorted({type(v).__name__ for v in value}) or ["empty"]
        return f"list[{'|'.join(inner)}]"
    if isinstance(value, dict):
        return f"dict{{{','.join(sorted(value))}}}"
    return type(value).__name__


def load_json(path):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return None


def find_dropped(node, path=""):
    if isinstance(node, dict):
        for key, value in node.items():
            if key in ("dropped_columns", "dropped_fields") and value:
                yield path, value
            else:
                yield from find_dropped(value, f"{path}/{key}")
    elif isinstance(node, list):
        for index, value in enumerate(node):
            label = value.get("title", index) if isinstance(value, dict) else index
            yield from find_dropped(value, f"{path}/{label}")


def check_list(list_dir):
    schema = load_json(os.path.join(list_dir, "_schema_columns.json"))
    items = load_json(os.path.join(list_dir, "items.json"))
    state = load_json(os.path.join(list_dir, ".state.json")) or {}
    if not isinstance(schema, list) or not isinstance(items, list):
        return None
    lines = []
    fmt = state.get("format")
    status = "current" if fmt == GENERIC_LIST_FORMAT else f"OLD (format={fmt})"
    lines.append(f"{len(items)} items, {status}")
    names = {c.get("internal_name") for c in schema}
    for column in schema:
        name = column.get("internal_name")
        if (
            column.get("hidden")
            or column.get("type") not in ID_FIELD_TYPES
            or name in ID_FIELD_EXCLUDED
        ):
            continue
        key = f"{name}Id"
        present = sum(1 for i in items if key in i)
        filled = [i[key] for i in items if i.get(key) not in (None, "", [], {})]
        types = ", ".join(sorted(Counter(value_type(v) for v in filled))) or "-"
        line = (
            f"{key} ({column.get('type')}): key on {present}, "
            f"filled {len(filled)} [{types}]"
        )
        if column.get("type") in ("User", "UserMulti"):
            named = sum(1 for i in items if i.get(name) not in (None, "", [], {}))
            line += f", names {named}"
        lines.append(line)
    for author_key in ("Author", "Editor"):
        with_id = sum(1 for i in items if i.get(f"{author_key}Id"))
        named = sum(1 for i in items if i.get(author_key))
        lines.append(f"{author_key}: id {with_id}, names {named}")
    if "fRecurrence" in names:
        recurrent = sum(1 for i in items if i.get("fRecurrence") is True)
        with_data = sum(1 for i in items if i.get("RecurrenceData"))
        event_types = dict(
            sorted(Counter(i.get("EventType") for i in items).items(), key=str)
        )
        lines.append(
            f"recurrence: fRecurrence={recurrent}, RecurrenceData={with_data}, "
            f"EventType={event_types}"
        )
    return lines


def main():
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    root, slugs = sys.argv[1], sys.argv[2:]
    output_dir = os.environ.get("SPX_OUTPUT_DIR", "./discovery_output")
    root_dir = os.path.join(output_dir, "extraction", root)
    summary = load_json(os.path.join(root_dir, "run_summary.json")) or {}

    for slug in slugs:
        print(f"\n=== {slug} ===")
        lists_dir = os.path.join(root_dir, slug, "lists")
        if not os.path.isdir(lists_dir):
            print("  no lists folder")
            continue
        old = 0
        for title in sorted(os.listdir(lists_dir)):
            lines = check_list(os.path.join(lists_dir, title))
            if not lines:
                continue
            if "OLD" in lines[0]:
                old += 1
            print(f"  {title}: {lines[0]}")
            for line in lines[1:]:
                print(f"      {line}")
        print(f"  lists not yet re-extracted: {old}")
        dropped = list(find_dropped(summary.get(slug, {})))
        print(f"  dropped columns in run_summary: {len(dropped)}")
        for path, columns in dropped:
            print(f"      {path}: {columns}")


if __name__ == "__main__":
    main()
