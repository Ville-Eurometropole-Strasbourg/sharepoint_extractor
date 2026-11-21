#!/usr/bin/env python3
"""
Write a <file>.versions.json next to each extracted .meta.json: current
version plus every past version (label, size, date, author, comment, URL).

- One call per file; files with a .versions.json are skipped (resume).
- /versions excludes the current version, read from the File instead.
- 0.1 drafts have no history: no call made. Neither has 1.0 in a library
  without minor versions; the setting is read once per library.

Usage (root = SPX_SITE_URL):
    python extract_versions.py <site_slug>
    python extract_versions.py --all
"""

import glob
import json
import logging
import os
import sys
from datetime import datetime, timezone

from sharepoint_client import get_config, build_context
from site_urls import get_root_name
from persistent_session import enable_persistent_session

# First unpublished draft: never has a history.
NO_HISTORY_LABELS = ("0.1",)
# First published version: no history in a library without minor versions,
# but with minor versions it is usually published from drafts 0.1, 0.2...
FIRST_MAJOR_LABEL = "1.0"


def needs_history(label, minor_versions):
    """
    Whether a file whose current version is `label` can have a history.
    minor_versions is the library's EnableMinorVersions, None when unknown:
    an unknown setting is treated as enabled, so no history is missed.
    """
    if label in NO_HISTORY_LABELS:
        return False
    if label == FIRST_MAJOR_LABEL:
        return minor_versions is not False
    return True


def _user_info(user):
    """Author fields kept for XWiki."""
    props = user.properties
    return {
        "login_name": props.get("LoginName"),
        "title": props.get("Title"),
        "email": props.get("Email"),
    }


def _authors_from_item_versions(ctx, file_ref):
    """
    Authors from item versions (Editor's name and email, no LoginName), for
    deleted accounts.

    Returns ({version_label: author}, current_version_author).
    """
    item_versions = ctx.web.get_file_by_server_relative_url(
        file_ref
    ).listItemAllFields.versions.select(["VersionLabel", "Editor", "IsCurrentVersion"])
    ctx.load(item_versions)
    ctx.execute_query()
    by_label, current = {}, {"login_name": None, "title": None, "email": None}
    for v in item_versions:
        editor = v.properties.get("Editor") or {}
        author = {
            "login_name": None,
            "title": editor.get("LookupValue"),
            "email": editor.get("Email"),
        }
        by_label[v.properties.get("VersionLabel")] = author
        if v.properties.get("IsCurrentVersion"):
            current = author
    return by_label, current


def fetch_versions(ctx, file_ref):
    """Current version and /versions history of one file."""
    file_obj = ctx.web.get_file_by_server_relative_url(file_ref).expand(["ModifiedBy"])
    versions = file_obj.versions.expand(["CreatedBy"])
    ctx.load(file_obj)
    ctx.load(versions)
    current_author = None
    try:
        ctx.execute_query()
        current_author = _user_info(file_obj.modified_by)
    except Exception as e:
        if "Impossible de trouver l'utilisateur" not in str(e):
            raise
        file_obj = ctx.web.get_file_by_server_relative_url(file_ref)
        versions = file_obj.versions.expand(["CreatedBy"])
        ctx.load(file_obj)
        ctx.load(versions)
        ctx.execute_query()

    fp = file_obj.properties
    history = []
    for v in versions:
        vp = v.properties
        history.append(
            {
                "label": vp.get("VersionLabel"),
                "size": vp.get("Size"),
                "created": str(vp.get("Created")),
                "author": _user_info(v.created_by),
                "author_from_item_fallback": False,
                "comment": vp.get("CheckInComment"),
                "url": vp.get("Url"),
            }
        )

    author_fallback = current_author is None
    missing = [h for h in history if not h["author"]["login_name"]]
    if author_fallback or missing:
        by_label, fallback_current = _authors_from_item_versions(ctx, file_ref)
        if author_fallback:
            current_author = fallback_current
        for h in missing:
            if h["label"] in by_label:
                h["author"] = by_label[h["label"]]
                h["author_from_item_fallback"] = True

    return {
        "file_ref": file_ref,
        "current": {
            "label": fp.get("UIVersionLabel"),
            "size": int(fp.get("Length") or 0),
            "modified": str(fp.get("TimeLastModified")),
            "author": current_author,
            "author_from_item_fallback": author_fallback,
        },
        "history": history,
        "history_count": len(history),
        "history_size_bytes": sum(h["size"] or 0 for h in history),
        "extracted_at": datetime.now(timezone.utc).isoformat(),
    }


def library_minor_versions(ctx, file_ref):
    """EnableMinorVersions of the library holding file_ref (one call)."""
    library = (
        ctx.web.get_file_by_server_relative_url(file_ref)
        .listItemAllFields.parent_list.get()
        .select(["EnableMinorVersions"])
        .execute_query()
    )
    return library.properties.get("EnableMinorVersions")


def extract_site_versions(site_dir, username, password, ca_bundle_path, logger):
    with open(os.path.join(site_dir, "site_metadata.json"), "r", encoding="utf-8") as f:
        site_url = json.load(f)["Url"]
    ctx = build_context(site_url, username, password, ca_bundle_path)
    documents_dir = os.path.join(site_dir, "documents")
    # Library folder under documents/ -> EnableMinorVersions, read once per
    # library holding a 1.0 file; None when it could not be read.
    minor_versions_by_library = {}

    counts = {
        "queried": 0,
        "skipped_no_history": 0,
        "skipped_already_done": 0,
        "errors": 0,
    }
    errors = []
    meta_paths = sorted(
        glob.glob(
            os.path.join(site_dir, "documents", "**", "*.meta.json"), recursive=True
        )
    )
    logger.info(f"{len(meta_paths)} .meta.json found under {site_dir}")

    for meta_path in meta_paths:
        versions_path = meta_path[: -len(".meta.json")] + ".versions.json"
        if os.path.exists(versions_path):
            counts["skipped_already_done"] += 1
            continue

        with open(meta_path, "r", encoding="utf-8") as f:
            meta = json.load(f)
        label = meta.get("OData__UIVersionString")
        file_ref = meta.get("FileRef")
        minor_versions = None
        if label == FIRST_MAJOR_LABEL:
            library = os.path.relpath(meta_path, documents_dir).split(os.sep)[0]
            if library not in minor_versions_by_library:
                try:
                    minor_versions_by_library[library] = library_minor_versions(
                        ctx, file_ref
                    )
                except Exception as e:
                    minor_versions_by_library[library] = None
                    logger.warning(
                        f"  Could not read whether '{library}' keeps minor versions, "
                        f"querying its 1.0 files: {e}"
                    )
            minor_versions = minor_versions_by_library[library]
        if not needs_history(label, minor_versions):
            counts["skipped_no_history"] += 1
            continue

        try:
            result = fetch_versions(ctx, file_ref)
        except Exception as e:
            counts["errors"] += 1
            errors.append({"file_ref": file_ref, "error": str(e)})
            logger.warning(f"  Failed on '{file_ref}': {e}")
            continue

        with open(versions_path, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2, default=str)
        counts["queried"] += 1

    # Latest pass only: failed files are retried next run.
    errors_path = os.path.join(site_dir, "versions_errors.json")
    if errors:
        with open(errors_path, "w", encoding="utf-8") as f:
            json.dump(errors, f, ensure_ascii=False, indent=2)
    elif os.path.exists(errors_path):
        os.remove(errors_path)
    return counts


def write_versions_summary(root_dir, summaries, logger):
    """Merge per-site counts into versions_run_summary.json (left as is if unreadable)."""
    summary_path = os.path.join(root_dir, "versions_run_summary.json")
    merged = {}
    if os.path.exists(summary_path):
        try:
            with open(summary_path, "r", encoding="utf-8") as f:
                merged = json.load(f)
        except (json.JSONDecodeError, OSError) as e:
            logger.error(
                f"Could not read existing {summary_path} ({e}): not overwriting it. "
                f"This run's summary: {json.dumps(summaries)[:500]}"
            )
            return
    merged.update(summaries)
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(merged, f, ensure_ascii=False, indent=2)
    logger.info(f"Summary written to {summary_path}")


def main():
    enable_persistent_session()

    site_base_url, username, password, ca_bundle_path = get_config()

    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s"
    )
    logger = logging.getLogger("extract_versions")

    if len(sys.argv) < 2:
        print("Usage: python extract_versions.py <site_slug> | --all")
        sys.exit(1)

    root_name = get_root_name()
    root_dir = os.path.join(
        os.environ.get("SPX_OUTPUT_DIR", "./discovery_output"),
        "extraction",
        root_name,
    )

    if sys.argv[1] == "--all":
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
            print(
                f"[ERROR] No extracted site found under {root_dir}: run extract_metadata.py first."
            )
            sys.exit(1)
        logger.info(f"{len(slugs)} extracted site(s) found under {root_dir}")
    else:
        slugs = [sys.argv[1]]
        if not os.path.exists(os.path.join(root_dir, slugs[0], "site_metadata.json")):
            print(
                f"[ERROR] {root_dir}/{slugs[0]}/site_metadata.json not found: run extract_metadata.py on this site first."
            )
            sys.exit(1)

    summaries = {}
    for i, slug in enumerate(slugs, start=1):
        logger.info(f"=== [{i}/{len(slugs)}] {slug} ===")
        try:
            counts = extract_site_versions(
                os.path.join(root_dir, slug), username, password, ca_bundle_path, logger
            )
        except Exception as e:
            logger.error(f"  Site-level failure for '{slug}': {e}")
            counts = {"error": str(e)}
        logger.info(f"Done for '{slug}': {counts}")
        summaries[slug] = counts
        write_versions_summary(root_dir, {slug: counts}, logger)

    # Retry sites with per-file errors; done files are skipped.
    retry = [s for s, c in summaries.items() if c.get("errors")]
    if retry:
        logger.info(f"=== Retrying failed files on {len(retry)} site(s): {retry} ===")
        for slug in retry:
            try:
                counts = extract_site_versions(
                    os.path.join(root_dir, slug),
                    username,
                    password,
                    ca_bundle_path,
                    logger,
                )
            except Exception as e:
                logger.error(f"  Site-level failure for '{slug}': {e}")
                counts = {"error": str(e)}
            logger.info(f"Retry done for '{slug}': {counts}")
            summaries[slug] = counts
            write_versions_summary(root_dir, {slug: counts}, logger)

    # Per-file errors don't fail the run.
    with_file_errors = [s for s, c in summaries.items() if c.get("errors")]
    if with_file_errors:
        logger.warning(
            f"{len(with_file_errors)} site(s) with per-file errors left: {with_file_errors}"
        )
    failed = [s for s, c in summaries.items() if "error" in c]
    if failed:
        logger.warning(f"{len(failed)} site(s) failed entirely: {failed}")
        sys.exit(1)


if __name__ == "__main__":
    main()
