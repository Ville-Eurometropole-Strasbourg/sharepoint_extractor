"""
Full version history of one site page: body, author and date per version.

Usage:
    python src/extraction/extract_page_history.py <root_name> <site_slug> [<page_name>]

page_name (a page of SitePages) defaults to the site's welcome page
(WelcomePage in site_metadata.json).

Writes under <SPX_OUTPUT_DIR>/extraction/<root>/<slug>/pages/history/<page>/:
    <label>.html     page body of that version (never printed)
    history.json     per version: label, length, sha256, author, date, comment
"""

import hashlib
import json
import os
import sys
from urllib.parse import urlparse

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from sharepoint_client import build_context, get_config
from extract_versions import fetch_versions

BODY_FIELD = "WikiField"


def safe_name(name):
    return "".join(c if c.isalnum() or c in "-_." else "_" for c in name)


def main():
    if len(sys.argv) not in (3, 4):
        print(__doc__)
        sys.exit(1)
    root_name, site_slug = sys.argv[1], sys.argv[2]

    output_dir = os.environ.get("SPX_OUTPUT_DIR", "./discovery_output")
    site_dir = os.path.join(output_dir, "extraction", root_name, site_slug)
    with open(os.path.join(site_dir, "site_metadata.json"), encoding="utf-8") as f:
        site_metadata = json.load(f)
    site_path = urlparse(site_metadata["Url"]).path

    if len(sys.argv) == 4:
        page_name = sys.argv[3]
    else:
        welcome_page = site_metadata.get("WelcomePage") or ""
        if not welcome_page.startswith("SitePages/"):
            print(
                f"[ERROR] The welcome page '{welcome_page}' is not in SitePages: "
                "pass a page name."
            )
            sys.exit(1)
        page_name = welcome_page[len("SitePages/") :]

    base_url, username, password, ca_bundle = get_config()
    parsed = urlparse(base_url)
    ctx = build_context(
        f"{parsed.scheme}://{parsed.netloc}{site_path}", username, password, ca_bundle
    )
    file_ref = f"{site_path}/SitePages/{page_name}"

    # Bodies from item versions (current one included).
    item_versions = ctx.web.get_file_by_server_relative_url(
        file_ref
    ).listItemAllFields.versions.select(
        ["VersionLabel", BODY_FIELD, "IsCurrentVersion"]
    )
    ctx.load(item_versions)
    ctx.execute_query()

    # Authors and dates from file versions.
    info = fetch_versions(ctx, file_ref)
    meta_by_label = {
        h["label"]: {
            "author": h["author"],
            "date": h["created"],
            "comment": h.get("comment"),
        }
        for h in info["history"]
    }
    current = info["current"]
    meta_by_label[current["label"]] = {
        "author": current["author"],
        "date": current["modified"],
        "comment": None,
    }

    out_dir = os.path.join(site_dir, "pages", "history", safe_name(page_name))
    os.makedirs(out_dir, exist_ok=True)

    versions = []
    for v in item_versions:
        label = v.properties.get("VersionLabel")
        body = v.properties.get(BODY_FIELD)
        entry = {
            "label": label,
            "is_current": bool(v.properties.get("IsCurrentVersion")),
            "chars": None,
            "sha256": None,
            **meta_by_label.get(label, {"author": None, "date": None, "comment": None}),
            "has_file_metadata": label in meta_by_label,
        }
        if body is not None:
            with open(
                os.path.join(out_dir, f"{safe_name(label)}.html"), "w", encoding="utf-8"
            ) as f:
                f.write(body)
            entry["chars"] = len(body)
            entry["sha256"] = hashlib.sha256(body.encode("utf-8")).hexdigest()
        versions.append(entry)

    versions.sort(key=lambda e: float(e["label"]))
    with open(os.path.join(out_dir, "history.json"), "w", encoding="utf-8") as f:
        json.dump(
            {"file_ref": file_ref, "versions": versions},
            f,
            ensure_ascii=False,
            indent=2,
            default=str,
        )

    print(f"[INFO] versions: {len(versions)}, labels: {[e['label'] for e in versions]}")
    print(
        f"[INFO] without body: {sum(1 for e in versions if e['sha256'] is None)}, "
        f"without file metadata: {sum(1 for e in versions if not e['has_file_metadata'])}"
    )

    # The current version must match the body already on disk.
    disk_path = os.path.join(
        site_dir, "pages", "content", f"{safe_name(page_name)}.{BODY_FIELD}.html"
    )
    cur = [e for e in versions if e["is_current"]]
    if os.path.exists(disk_path) and cur:
        with open(disk_path, "rb") as f:
            disk_hash = hashlib.sha256(f.read()).hexdigest()
        print(
            f"[INFO] current version matches inventory file: {disk_hash == cur[0]['sha256']}"
        )
    print(f"[INFO] written to {out_dir}")


if __name__ == "__main__":
    main()
