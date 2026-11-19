"""
Add creation / modification dates and authors to the .meta.json of every
file of an extracted site.

For each library folder under <site>/documents/, one paged query on the
SharePoint list (title = folder name) reads, for every item:
    Id, Created, Modified, Author (created by), Editor (modified by)
and each .meta.json whose "ID" matches gets four keys:
    "Created", "Modified": ISO dates as returned by SharePoint (UTC)
    "Author", "Editor":    {"login_name": ..., "title": ...}
Other keys are kept. Each file is written atomically (tmp + os.replace).

Only counts are printed, never names or values.

Usage (in the extractor container, like the other extraction scripts):
    python3 src/extraction/enrich_file_meta.py hr team
"""

import glob
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from sharepoint_client import build_context, get_config  # noqa: E402

FIELDS = [
    "Id",
    "Created",
    "Modified",
    "Author/Title",
    "Author/Name",
    "Editor/Title",
    "Editor/Name",
]


def person(value):
    value = value or {}
    return {"login_name": value.get("Name"), "title": value.get("Title")}


def write_json_atomic(path, data):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def library_items(ctx, title):
    """{item Id: {Created, Modified, Author, Editor}} for one library."""
    items = (
        ctx.web.lists.get_by_title(title)
        .items.select(FIELDS)
        .expand(["Author", "Editor"])
        .get_all(5000)
        .execute_query()
    )
    result = {}
    for item in items:
        p = item.properties
        result[p.get("Id")] = {
            "Created": p.get("Created"),
            "Modified": p.get("Modified"),
            "Author": person(p.get("Author")),
            "Editor": person(p.get("Editor")),
        }
    return result


def main():
    if len(sys.argv) != 3:
        sys.exit("usage: enrich_file_meta.py <root_name> <site_slug>")
    root, slug = sys.argv[1], sys.argv[2]
    output_dir = os.environ.get("SPX_OUTPUT_DIR", "./discovery_output")
    site_dir = os.path.join(output_dir, "extraction", root, slug)
    documents_dir = os.path.join(site_dir, "documents")
    with open(os.path.join(site_dir, "site_metadata.json"), encoding="utf-8") as f:
        site_url = json.load(f)["Url"]

    _, username, password, ca_bundle_path = get_config()
    ctx = build_context(site_url, username, password, ca_bundle_path)

    totals = {"updated": 0, "not_found": 0, "libraries_failed": 0}
    for library in sorted(os.listdir(documents_dir)):
        lib_dir = os.path.join(documents_dir, library)
        if not os.path.isdir(lib_dir):
            continue
        metas = glob.glob(os.path.join(lib_dir, "**", "*.meta.json"), recursive=True)
        if not metas:
            continue
        try:
            by_id = library_items(ctx, library)
        except Exception as exc:  # report and continue with the next library
            totals["libraries_failed"] += 1
            print(f"{library}: query FAILED ({type(exc).__name__}: {exc})")
            continue
        updated = not_found = 0
        for meta_path in metas:
            with open(meta_path, encoding="utf-8") as f:
                meta = json.load(f)
            info = by_id.get(meta.get("ID") or meta.get("Id"))
            if info is None:
                not_found += 1
                continue
            meta.update(info)
            write_json_atomic(meta_path, meta)
            updated += 1
        print(f"{library}: {updated} updated, {not_found} without matching item")
        totals["updated"] += updated
        totals["not_found"] += not_found
    print(f"\nDone: {totals}")


if __name__ == "__main__":
    main()
