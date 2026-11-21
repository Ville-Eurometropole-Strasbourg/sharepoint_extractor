#!/usr/bin/env python3
"""
Download past version binaries listed in .versions.json files:
    Documents/Rapport.docx.history/1.0.docx
    Documents/Rapport.docx.history/2.0.docx
The current version comes from download_files.py.

Each version is read from its _vti_history URL. SharePoint 2019 answers 404
there for pages (.aspx), so on a 404 the version is read through the REST API.

Written files must match the version Size exactly, else they are deleted
and counted as errors. Files already at the right size are skipped.

Usage (root = SPX_SITE_URL):
    python download_versions.py <site_slug>
"""

import glob
import json
import logging
import os
import sys
from urllib.parse import quote

import requests
from office365.runtime.http.request_options import RequestOptions

from sharepoint_client import get_config, build_context
from site_urls import get_root_name
from persistent_session import enable_persistent_session


def _history_path(versions_path, label):
    """<file>.history/<label><ext>, the extension kept from the file name."""
    base = versions_path[: -len(".versions.json")]
    ext = os.path.splitext(base)[1]
    return os.path.join(base + ".history", f"{label}{ext}")


def _open_stream(ctx, url):
    """Stream a GET of url. An HTTP error status raises requests.HTTPError."""
    opts = RequestOptions(url)
    opts.stream = True
    return ctx.pending_request().execute_request_direct(opts)


def _download_version(ctx, site_url, file_ref, rel_url, local_path, expected_size):
    """Stream one past version to local_path. Returns the written size.

    rel_url is the version's _vti_history/<id>/<path> URL.
    """
    try:
        response = _open_stream(ctx, site_url.rstrip("/") + "/" + quote(rel_url))
    except requests.exceptions.HTTPError as e:
        if getattr(e.response, "status_code", None) != 404:
            raise
        # Pages (.aspx): the REST API serves the version that _vti_history refuses.
        versions = ctx.web.get_file_by_server_relative_url(file_ref).versions
        version = versions.get_by_id(int(rel_url.split("/")[1]))
        response = _open_stream(ctx, version.resource_url + "/$value")
    os.makedirs(os.path.dirname(local_path), exist_ok=True)
    try:
        with open(local_path, "wb") as f:
            for chunk in response.iter_content(chunk_size=1024 * 1024):
                f.write(chunk)
    finally:
        response.close()
    size = os.path.getsize(local_path)
    if expected_size is not None and size != expected_size:
        os.remove(local_path)
        raise ValueError(
            f"size mismatch: got {size} bytes, /versions said {expected_size}"
        )
    return size


def download_site_versions(site_dir, username, password, ca_bundle_path, logger):
    with open(os.path.join(site_dir, "site_metadata.json"), "r", encoding="utf-8") as f:
        site_url = json.load(f)["Url"]
    ctx = build_context(site_url, username, password, ca_bundle_path)

    counts = {"downloaded": 0, "skipped_already_done": 0, "errors": 0, "bytes": 0}
    errors = []
    versions_paths = sorted(
        glob.glob(
            os.path.join(site_dir, "documents", "**", "*.versions.json"), recursive=True
        )
    )
    logger.info(f"{len(versions_paths)} .versions.json found under {site_dir}")

    for versions_path in versions_paths:
        with open(versions_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        for h in data["history"]:
            local_path = _history_path(versions_path, h["label"])
            if os.path.exists(local_path) and os.path.getsize(local_path) == h["size"]:
                counts["skipped_already_done"] += 1
                continue
            try:
                size = _download_version(
                    ctx, site_url, data["file_ref"], h["url"], local_path, h["size"]
                )
            except Exception as e:
                counts["errors"] += 1
                errors.append(
                    {"file_ref": data["file_ref"], "label": h["label"], "error": str(e)}
                )
                logger.warning(f"  Failed on '{data['file_ref']}' v{h['label']}: {e}")
                continue
            counts["downloaded"] += 1
            counts["bytes"] += size

    summary = {"counts": counts, "errors": errors}
    with open(
        os.path.join(site_dir, "versions_download_summary.json"), "w", encoding="utf-8"
    ) as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    return counts


def main():
    enable_persistent_session()

    site_base_url, username, password, ca_bundle_path = get_config()

    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s"
    )
    logger = logging.getLogger("download_versions")

    if len(sys.argv) != 2:
        print("Usage: python download_versions.py <site_slug>")
        sys.exit(1)

    root_name = get_root_name()
    site_dir = os.path.join(
        os.environ.get("SPX_OUTPUT_DIR", "./discovery_output"),
        "extraction",
        root_name,
        sys.argv[1],
    )
    if not os.path.exists(os.path.join(site_dir, "site_metadata.json")):
        print(
            f"[ERROR] {site_dir}/site_metadata.json not found: run extract_metadata.py on this site first."
        )
        sys.exit(1)

    counts = download_site_versions(
        site_dir, username, password, ca_bundle_path, logger
    )
    logger.info(f"Done for '{sys.argv[1]}': {counts}")
    if counts["errors"]:
        sys.exit(1)


if __name__ == "__main__":
    main()
