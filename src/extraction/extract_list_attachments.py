"""
Write a .meta.json for each list item attachment (metadata only):
    lists/<list>/attachments/<item id>/<file name>.meta.json
Uses extract_metadata.py's list skip rules. Binaries come later (download).

Resume: lists unchanged since the last error-free pass are skipped
(.attachments_state.json). Attachments deleted on SharePoint are not cleaned up.

Usage (root = SPX_SITE_URL):
    python src/extraction/extract_list_attachments.py <site_folder>
    python src/extraction/extract_list_attachments.py --all
"""

import json
import logging
import os
import sys

from extract_metadata import LIST_FILTER_FIELDS, list_skip_reason
from persistent_session import enable_persistent_session
from sharepoint_client import build_context, get_config
from site_urls import get_root_name

STATE_FILE = ".attachments_state.json"


def safe_name(name):
    return (name or "").replace("/", "_")


def _read_state(list_dir):
    try:
        with open(os.path.join(list_dir, STATE_FILE), encoding="utf-8") as f:
            return json.load(f).get("last_item_modified_date")
    except (OSError, json.JSONDecodeError):
        return None


def _write_state(list_dir, value):
    os.makedirs(list_dir, exist_ok=True)
    with open(os.path.join(list_dir, STATE_FILE), "w", encoding="utf-8") as f:
        json.dump({"last_item_modified_date": value}, f)


def extract_site_attachments(site_dir, username, password, ca_bundle_path, logger):
    with open(os.path.join(site_dir, "site_metadata.json"), encoding="utf-8") as f:
        site_url = json.load(f)["Url"]
    ctx = build_context(site_url, username, password, ca_bundle_path)
    lists = (
        ctx.web.lists.select(
            [
                "Id",
                "BaseType",
                "EnableAttachments",
                "LastItemModifiedDate",
            ]
            + LIST_FILTER_FIELDS
        )
        .get()
        .execute_query()
    )
    counts = {
        "lists_checked": 0,
        "lists_unchanged": 0,
        "items_with_attachments": 0,
        "attachments": 0,
        "errors": 0,
    }
    for lst in lists:
        p = lst.properties
        title = p.get("Title")
        if p.get("BaseType") == 1 or not p.get("EnableAttachments"):
            continue
        if list_skip_reason(p):
            continue
        list_dir = os.path.join(site_dir, "lists", safe_name(title))
        server_date = str(p.get("LastItemModifiedDate"))
        if _read_state(list_dir) == server_date:
            counts["lists_unchanged"] += 1
            continue
        counts["lists_checked"] += 1
        keyed = ctx.web.lists.get_by_id(p.get("Id"))
        try:
            items = (
                keyed.items.select(["Id", "Attachments"])
                .get_all(page_size=100)
                .execute_query()
            )
        except Exception as e:
            counts["errors"] += 1
            logger.warning(f"  Failed to read the items of a list: {e}")
            continue

        list_errors = 0
        for item in items:
            if not item.properties.get("Attachments"):
                continue
            item_id = item.properties.get("Id")
            counts["items_with_attachments"] += 1
            try:
                files = (
                    keyed.items.get_by_id(item_id)
                    .attachment_files.get()
                    .execute_query()
                )
            except Exception as e:
                list_errors += 1
                logger.warning(
                    f"  Failed to list the attachments of item {item_id}: {e}"
                )
                continue
            for att in files:
                ap = att.properties
                name = safe_name(ap.get("FileName"))
                meta_path = os.path.join(
                    list_dir, "attachments", str(item_id), name + ".meta.json"
                )
                os.makedirs(os.path.dirname(meta_path), exist_ok=True)
                with open(meta_path, "w", encoding="utf-8") as f:
                    json.dump(
                        {
                            "FileName": ap.get("FileName"),
                            "ServerRelativeUrl": ap.get("ServerRelativeUrl"),
                            "ItemId": item_id,
                            "ListId": p.get("Id"),
                            "ListTitle": title,
                            "_relative_path": f"attachments/{item_id}/{name}",
                        },
                        f,
                        ensure_ascii=False,
                        indent=2,
                        default=str,
                    )
                counts["attachments"] += 1

        counts["errors"] += list_errors
        if not list_errors:
            _write_state(list_dir, server_date)
    return counts


def write_summary(root_dir, summaries, logger):
    """Merge per-site counts into attachments_run_summary.json."""
    path = os.path.join(root_dir, "attachments_run_summary.json")
    merged = {}
    if os.path.exists(path):
        try:
            with open(path, encoding="utf-8") as f:
                merged = json.load(f)
        except (json.JSONDecodeError, OSError) as e:
            logger.error(f"Could not read {path} ({e}): not overwriting it.")
            return
    merged.update(summaries)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(merged, f, ensure_ascii=False, indent=2)


def main():
    enable_persistent_session()
    site_base_url, username, password, ca_bundle_path = get_config()
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s"
    )
    logger = logging.getLogger("extract_list_attachments")

    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(1)

    root_name = get_root_name()
    root_dir = os.path.join(
        os.environ.get("SPX_OUTPUT_DIR", "./discovery_output"),
        "extraction",
        root_name,
    )
    if sys.argv[1] == "--all":
        slugs = sorted(
            d
            for d in os.listdir(root_dir)
            if os.path.exists(os.path.join(root_dir, d, "site_metadata.json"))
        )
    else:
        slugs = [sys.argv[1]]
        if not os.path.exists(os.path.join(root_dir, slugs[0], "site_metadata.json")):
            print(f"[ERROR] {root_dir}/{slugs[0]}/site_metadata.json not found.")
            sys.exit(1)

    totals, failed = {}, []
    for i, slug in enumerate(slugs, start=1):
        try:
            counts = extract_site_attachments(
                os.path.join(root_dir, slug), username, password, ca_bundle_path, logger
            )
        except Exception as e:
            logger.error(f"  Site-level failure for '{slug}': {e}")
            counts = {"error": str(e)}
            failed.append(slug)
        if counts.get("attachments") or counts.get("errors") or "error" in counts:
            logger.info(f"[{i}/{len(slugs)}] {slug}: {counts}")
        write_summary(root_dir, {slug: counts}, logger)
        for k, v in counts.items():
            if isinstance(v, int):
                totals[k] = totals.get(k, 0) + v

    logger.info(f"Done: {len(slugs)} site(s), {len(failed)} failed, totals: {totals}")
    if failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
