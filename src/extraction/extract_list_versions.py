"""
Write lists/<list>/items_versions.json for each generic list whose items
have a history: every version of each such item, the current one included,
with its field values as returned by SharePoint.

- Lists: extract_metadata.py's skip rules. Libraries are extract_versions.py's.
- One paged read per non-empty list for the current version labels, versioned
  or not: a list whose versioning was turned off keeps its past versions.
  Then one /versions call per item that can have a history (same rule as
  files, see extract_versions.needs_history). Items without a label are
  counted, not queried.
- Resume: a list unchanged since its last error-free pass is skipped
  (.versions_state.json). In a list, an item whose label is the one already
  saved is not read again, and the file is saved every SAVE_EVERY items.
  Items deleted on SharePoint are dropped from the file.
- --dry-run reads the labels only and counts the items that would be queried.

Usage (root = SPX_SITE_URL):
    python src/extraction/extract_list_versions.py <site_folder> [--dry-run]
    python src/extraction/extract_list_versions.py --all [--dry-run]
"""

import json
import logging
import os
import sys
from datetime import datetime, timezone

from extract_metadata import LIST_FILTER_FIELDS, list_skip_reason
from extract_versions import needs_history
from persistent_session import enable_persistent_session
from sharepoint_client import build_context, get_config
from site_urls import get_root_name

VERSIONS_FILE = "items_versions.json"
STATE_FILE = ".versions_state.json"
# items_versions.json and .versions_state.json format.
LIST_VERSIONS_FORMAT = 1
LABEL_FIELD = "OData__UIVersionString"
ITEMS_PAGE_SIZE = 1000
SAVE_EVERY = 100


def list_dir_of(site_dir, title):
    """Same folder as extract_metadata.extract_generic_list."""
    return os.path.join(site_dir, "lists", (title or "").replace("/", "_"))


def _load_json(path):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return None


def _write_json_atomic(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str)
    os.replace(tmp, path)


def _read_state(list_dir):
    state = _load_json(os.path.join(list_dir, STATE_FILE)) or {}
    if state.get("format") != LIST_VERSIONS_FORMAT:
        return None
    return state.get("last_item_modified_date")


def _write_state(list_dir, value):
    _write_json_atomic(
        os.path.join(list_dir, STATE_FILE),
        {"last_item_modified_date": value, "format": LIST_VERSIONS_FORMAT},
    )


def _saved_items(list_dir):
    """{item id (str): entry} already on disk, {} when absent or older."""
    data = _load_json(os.path.join(list_dir, VERSIONS_FILE)) or {}
    if data.get("format") != LIST_VERSIONS_FORMAT:
        return {}
    return data.get("items") or {}


def fetch_item_versions(sp_list, item_id):
    """Every version of one item, newest first, as raw field values."""
    versions = sp_list.items.get_by_id(item_id).versions.get().execute_query()
    return [dict(v.properties) for v in versions]


def extract_list(sp_list, props, list_dir, counts, logger, dry_run=False):
    """
    Versions of one list's items. Returns the number of item errors, or None
    when the labels could not be read.
    """
    try:
        items = (
            sp_list.items.select(["Id", LABEL_FIELD])
            .get_all(page_size=ITEMS_PAGE_SIZE)
            .execute_query()
        )
    except Exception as e:
        logger.warning(f"  Failed to read the version labels of a list: {e}")
        return None

    minor_versions = props.get("EnableMinorVersions")
    with_history = []
    for item in items:
        p = item.properties
        counts["items_read"] += 1
        label = p.get(LABEL_FIELD)
        if label is None:
            counts["items_without_label"] += 1
        elif needs_history(label, minor_versions):
            with_history.append((str(p.get("Id")), label))
    counts["items_with_history"] += len(with_history)

    saved = _saved_items(list_dir)
    to_query = [
        (i, label)
        for i, label in with_history
        if (saved.get(i) or {}).get("current_label") != label
    ]
    counts["items_already_saved"] += len(with_history) - len(to_query)
    if dry_run:
        counts["items_to_query"] += len(to_query)
        return 0

    path = os.path.join(list_dir, VERSIONS_FILE)
    if not with_history:
        if os.path.exists(path):
            os.remove(path)
        return 0

    # Only items that still exist and can have a history are kept.
    keep = {i for i, _ in with_history}
    entries = {i: e for i, e in saved.items() if i in keep}
    errors = []

    def save():
        _write_json_atomic(
            path,
            {
                "format": LIST_VERSIONS_FORMAT,
                "list_id": props.get("Id"),
                "list_title": props.get("Title"),
                "enable_versioning": props.get("EnableVersioning"),
                "items": entries,
                "errors": errors,
                "extracted_at": datetime.now(timezone.utc).isoformat(),
            },
        )

    for n, (item_id, label) in enumerate(to_query, start=1):
        try:
            versions = fetch_item_versions(sp_list, int(item_id))
        except Exception as e:
            # An older entry is kept: its label differs, so it is retried.
            errors.append({"item_id": int(item_id), "error": str(e)})
            logger.warning(f"  Failed to read the versions of item {item_id}: {e}")
        else:
            entries[item_id] = {"current_label": label, "versions": versions}
            counts["items_queried"] += 1
            counts["versions"] += len(versions)
        if n % SAVE_EVERY == 0:
            save()
            logger.info(f"    {n}/{len(to_query)} items read")
    save()
    return len(errors)


def extract_site_list_versions(
    site_dir, username, password, ca_bundle_path, logger, dry_run=False
):
    with open(os.path.join(site_dir, "site_metadata.json"), encoding="utf-8") as f:
        site_url = json.load(f)["Url"]
    ctx = build_context(site_url, username, password, ca_bundle_path)
    lists = (
        ctx.web.lists.select(
            [
                "Id",
                "BaseType",
                "ItemCount",
                "EnableVersioning",
                "EnableMinorVersions",
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
        "lists_failed": 0,
        "items_read": 0,
        "items_without_label": 0,
        "items_with_history": 0,
        "items_already_saved": 0,
        "items_queried": 0,
        "versions": 0,
        "errors": 0,
    }
    if dry_run:
        counts["items_to_query"] = 0
    for lst in lists:
        p = lst.properties
        if p.get("BaseType") == 1 or list_skip_reason(p):
            continue
        list_dir = list_dir_of(site_dir, p.get("Title"))
        server_date = str(p.get("LastItemModifiedDate"))
        if _read_state(list_dir) == server_date:
            counts["lists_unchanged"] += 1
            continue
        if not p.get("ItemCount"):
            stale = os.path.join(list_dir, VERSIONS_FILE)
            if not dry_run and os.path.exists(stale):
                os.remove(stale)
            continue
        counts["lists_checked"] += 1
        keyed = ctx.web.lists.get_by_id(p.get("Id"))
        item_errors = extract_list(keyed, p, list_dir, counts, logger, dry_run)
        if item_errors is None:
            counts["lists_failed"] += 1
            continue
        counts["errors"] += item_errors
        if not item_errors and not dry_run:
            _write_state(list_dir, server_date)
    return counts


def write_summary(root_dir, summaries, logger):
    """Merge per-site counts into list_versions_run_summary.json."""
    path = os.path.join(root_dir, "list_versions_run_summary.json")
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
    _, username, password, ca_bundle_path = get_config()
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s"
    )
    logger = logging.getLogger("extract_list_versions")

    args = sys.argv[1:]
    dry_run = "--dry-run" in args
    args = [a for a in args if a != "--dry-run"]
    if len(args) != 1:
        print(__doc__)
        sys.exit(1)

    root_dir = os.path.join(
        os.environ.get("SPX_OUTPUT_DIR", "./discovery_output"),
        "extraction",
        get_root_name(),
    )
    if args[0] == "--all":
        slugs = (
            sorted(
                d
                for d in os.listdir(root_dir)
                if os.path.exists(os.path.join(root_dir, d, "site_metadata.json"))
            )
            if os.path.isdir(root_dir)
            else []
        )
        if not slugs:
            print(f"[ERROR] No extracted site found under {root_dir}.")
            sys.exit(1)
    else:
        slugs = [args[0]]
        if not os.path.exists(os.path.join(root_dir, slugs[0], "site_metadata.json")):
            print(f"[ERROR] {root_dir}/{slugs[0]}/site_metadata.json not found.")
            sys.exit(1)

    totals, failed = {}, []
    for i, slug in enumerate(slugs, start=1):
        try:
            counts = extract_site_list_versions(
                os.path.join(root_dir, slug),
                username,
                password,
                ca_bundle_path,
                logger,
                dry_run,
            )
        except Exception as e:
            logger.error(f"  Site-level failure for '{slug}': {e}")
            counts = {"error": str(e)}
            failed.append(slug)
        logger.info(f"[{i}/{len(slugs)}] {slug}: {counts}")
        if not dry_run:
            write_summary(root_dir, {slug: counts}, logger)
        for k, v in counts.items():
            if isinstance(v, int):
                totals[k] = totals.get(k, 0) + v

    logger.info(
        f"Done{' (dry run)' if dry_run else ''}: {len(slugs)} site(s), "
        f"{len(failed)} failed, totals: {totals}"
    )
    if failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
