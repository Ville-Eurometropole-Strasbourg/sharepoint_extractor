"""
Run extract_generic_list on one list into a temporary folder and report which
keys came back, how many items fill each, and the Python type of the values.
Nothing is written under discovery_output/ and no value is printed.

Usage (in the container, from the project root):
    python src/tools/probe_list_fields.py /sites/hr/team "Tasks"
"""

import json
import logging
import os
import shutil
import sys
import tempfile
from collections import Counter

SRC_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, SRC_DIR)
sys.path.insert(0, os.path.join(SRC_DIR, "extraction"))
from sharepoint_client import get_config, build_context  # noqa: E402
from persistent_session import enable_persistent_session  # noqa: E402
from extract_metadata import extra_list_fields, extract_generic_list  # noqa: E402


def value_type(value):
    if isinstance(value, list):
        inner = sorted({type(v).__name__ for v in value}) or ["empty"]
        return f"list[{'|'.join(inner)}]"
    if isinstance(value, dict):
        return f"dict{{{','.join(sorted(value))}}}"
    return type(value).__name__


def main():
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    url_suffix, list_title = sys.argv[1], sys.argv[2]

    enable_persistent_session()
    site_base_url, username, password, ca_bundle_path = get_config()
    host = "/".join(site_base_url.split("/")[:3])
    ctx = build_context(host + url_suffix, username, password, ca_bundle_path)

    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    logger = logging.getLogger("probe_list_fields")

    tmp_dir = tempfile.mkdtemp(prefix="probe_list_")
    try:
        sp_list = ctx.web.lists.get_by_title(list_title)
        result = extract_generic_list(sp_list, list_title, tmp_dir, logger)
        list_dir = os.path.join(tmp_dir, "lists", list_title.replace("/", "_"))
        with open(
            os.path.join(list_dir, "_schema_columns.json"), encoding="utf-8"
        ) as f:
            schema = json.load(f)
        with open(os.path.join(list_dir, "items.json"), encoding="utf-8") as f:
            items = json.load(f)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    requested = extra_list_fields(schema)
    print(f"\nResult: {result}")
    print(f"Extra fields requested: {requested}")
    print(f"Items: {len(items)}")

    keys = Counter(k for item in items for k in item)
    filled = Counter(
        k for item in items for k, v in item.items() if v not in (None, "", [], {})
    )
    types = {}
    for item in items:
        for k, v in item.items():
            if v not in (None, "", [], {}):
                types.setdefault(k, set()).add(value_type(v))
    for key in sorted(keys):
        mark = "  <- extra" if key in requested else ""
        kinds = ", ".join(sorted(types.get(key, {"-"})))
        print(f"  {key}: {filled[key]}/{len(items)} [{kinds}]{mark}")
    missing = [name for name in requested if name not in keys]
    if missing:
        print(f"Requested but absent: {missing}")


if __name__ == "__main__":
    main()
