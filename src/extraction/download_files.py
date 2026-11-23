"""
Download file binaries.

Usage:
    python download_files.py <site_url_suffix> <library_title> [--workers N] [--limit N]
    python download_files.py --site <site_slug> [--workers N] [--limit N]
    python download_files.py --all [total_limit] [--workers N]

--site and --all use the .meta.json already under
SPX_OUTPUT_DIR/extraction/<SPX_ROOT_NAME>, for the active root only.
Each file is written to <file>.spx-part, then renamed once complete.
"""

import os
import sys
import json
import logging
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urlparse

from sharepoint_client import get_config, build_context
from site_urls import absolute_url, get_root_name
from persistent_session import enable_persistent_session
from extract_metadata import relative_path_in_library, site_slug

ITEM_SELECT_FIELDS = [
    "FileLeafRef",
    "FileRef",
    "FileSystemObjectType",
    "Id",
    "GUID",
    "FileSizeDisplay",
]
DEFAULT_WORKER_COUNT = 6
MAX_WORKER_COUNT = 8
PART_SUFFIX = ".spx-part"
SIZE_TOLERANCE_BYTES = 1
_thread_local = threading.local()


def _parse_size(value):
    """Size as int (REST often returns a string), or None."""
    if value is None:
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _get_thread_context(site_url, username, password, ca_bundle_path):
    """ClientContext cached per (thread, site)."""
    if not hasattr(_thread_local, "ctx_by_site"):
        _thread_local.ctx_by_site = {}
    if site_url not in _thread_local.ctx_by_site:
        _thread_local.ctx_by_site[site_url] = build_context(
            site_url, username, password, ca_bundle_path
        )
    return _thread_local.ctx_by_site[site_url]


def _download_one_file(
    site_url,
    username,
    password,
    ca_bundle_path,
    library_title,
    lib_dir,
    name,
    file_ref,
    item_id,
    logger,
    root_name=None,
    site_relative_url=None,
    expected_size=None,
    relative_path=None,
    by_url=False,
):
    """
    Download one file to lib_dir/relative_path with this thread's context.
    Skipped if the file already exists. A size mismatch with expected_size
    is only logged. root_name/site_relative_url are copied into the result
    for the summary. by_url fetches by FileRef instead of list item Id.
    """
    relative_path = relative_path or name
    local_path = os.path.join(lib_dir, relative_path)
    os.makedirs(os.path.dirname(local_path) or lib_dir, exist_ok=True)

    base_result = {
        "name": name,
        "path": relative_path,
        "local_path": os.path.abspath(local_path),
        "site_url": site_url,
        "library_title": library_title,
    }
    if root_name is not None:
        base_result["root_name"] = root_name
    if site_relative_url is not None:
        base_result["site_relative_url"] = site_relative_url

    if os.path.exists(local_path):
        logger.info(f"  SKIP '{relative_path}': already downloaded.")
        return {
            **base_result,
            "status": "skipped",
            "size_bytes": os.path.getsize(local_path),
        }

    ctx = _get_thread_context(site_url, username, password, ca_bundle_path)
    part_path = local_path + PART_SUFFIX

    try:
        if by_url:
            target = ctx.web.get_file_by_server_relative_url(file_ref)
        else:
            target = (
                ctx.web.lists.get_by_title(library_title).items.get_by_id(item_id).file
            )
        with open(part_path, "wb") as f:
            target.download(f).execute_query()
        os.replace(part_path, local_path)
        size_bytes = os.path.getsize(local_path)

        if expected_size is not None and size_bytes != expected_size:
            logger.warning(
                f"  SIZE MISMATCH '{relative_path}': downloaded {size_bytes} bytes, "
                f"metadata said {expected_size} (kept -- metadata field is unreliable here)."
            )

        logger.info(f"  OK '{relative_path}': {size_bytes} bytes.")
        return {**base_result, "status": "ok", "size_bytes": size_bytes}
    except Exception as e:
        logger.error(f"  FAILED '{relative_path}': {e}")
        if os.path.exists(part_path):
            os.remove(part_path)
        return {**base_result, "status": "error", "error": str(e)}


def download_library(
    site_url,
    username,
    password,
    ca_bundle_path,
    library_title,
    lib_dir,
    logger,
    worker_count,
    limit=None,
):
    os.makedirs(lib_dir, exist_ok=True)
    try:
        ctx = build_context(site_url, username, password, ca_bundle_path)
        sp_list = ctx.web.lists.get_by_title(library_title)
        if limit is not None:
            items = (
                sp_list.items.select(ITEM_SELECT_FIELDS)
                .order_by("Id")
                .top(limit)
                .get()
                .execute_query()
            )
        else:
            items = (
                sp_list.items.select(ITEM_SELECT_FIELDS)
                .get_all(page_size=100)
                .execute_query()
            )
    except Exception as e:
        logger.error(f"Failed to read items from '{library_title}': {e}")
        sys.exit(1)

    # Files are placed from the root folder URL, not the title: a library
    # titled "Documents" may live in "Shared Documents".
    library_root = None
    try:
        library_root = (
            sp_list.root_folder.get()
            .execute_query()
            .properties.get("ServerRelativeUrl")
        )
    except Exception as e:
        logger.warning(
            f"Could not read the root folder of '{library_title}', "
            f"writing files flat: {e}"
        )

    files_to_download = []
    for item in items:
        props = item.properties
        if props.get("FileSystemObjectType") == 1:
            continue
        name = props.get("FileLeafRef") or f"item_{props.get('Id')}"
        # Same layout as the .meta.json written by extract_metadata.py.
        relative = relative_path_in_library(props.get("FileRef"), library_root)
        files_to_download.append(
            (
                name,
                props.get("FileRef"),
                props.get("Id"),
                _parse_size(props.get("FileSizeDisplay")),
                (
                    os.path.join(*relative.split("/"))
                    if relative
                    else name.replace("/", "_")
                ),
            )
        )
    if limit is not None and len(files_to_download) > limit:
        files_to_download = files_to_download[:limit]

    logger.info(
        f"Downloading {len(files_to_download)} file(s) with {worker_count} worker(s)."
    )

    results = []
    with ThreadPoolExecutor(max_workers=worker_count) as executor:
        futures = {
            executor.submit(
                _download_one_file,
                site_url,
                username,
                password,
                ca_bundle_path,
                library_title,
                lib_dir,
                name,
                file_ref,
                item_id,
                logger,
                expected_size=expected_size,
                relative_path=relative_path,
            ): name
            for name, file_ref, item_id, expected_size, relative_path in files_to_download
        }
        for future in as_completed(futures):
            results.append(future.result())

    return results


def _site_relative_url(site_dir):
    """
    Site's server-relative URL from its site_metadata.json, or None.
    A host root's Url has no path (https://host): its URL is "/".
    """
    try:
        with open(
            os.path.join(site_dir, "site_metadata.json"), "r", encoding="utf-8"
        ) as f:
            url = json.load(f)["Url"]
    except (FileNotFoundError, KeyError, json.JSONDecodeError):
        return None
    if not url:
        return None
    return urlparse(url).path or "/"


def _find_site_dir(root_dir, url_suffix):
    """Folder of the extracted site whose URL is url_suffix (case-insensitive), or None."""
    wanted = url_suffix.rstrip("/").lower()
    if os.path.isdir(root_dir):
        for site_slug in sorted(os.listdir(root_dir)):
            site_dir = os.path.join(root_dir, site_slug)
            if not os.path.isdir(site_dir):
                continue
            url = _site_relative_url(site_dir)
            if url and url.rstrip("/").lower() == wanted:
                return site_dir
    return None


def discover_all_downloadable_files(
    extraction_dir, logger, limit=None, root_filter=None, site_filter=None
):
    """
    Download tasks for every .meta.json on disk (library files and list
    attachments); each binary lands next to its .meta.json. `limit` caps
    the task count, already downloaded files included.

    Layout:
        <extraction_dir>/<root>/<site>/documents/<library>/...
        <extraction_dir>/<root>/<site>/lists/<list>/attachments/<item id>/...

    Sites without a readable site_metadata.json are skipped.
    """
    tasks = []

    for root_name in sorted(os.listdir(extraction_dir)):
        if root_filter is not None and root_name != root_filter:
            continue
        root_path = os.path.join(extraction_dir, root_name)
        if not os.path.isdir(root_path):
            continue

        for site_slug in sorted(os.listdir(root_path)):
            if site_filter is not None and site_slug != site_filter:
                continue
            site_dir = os.path.join(root_path, site_slug)
            if not os.path.isdir(site_dir):
                continue
            server_relative_url = _site_relative_url(site_dir)
            if not server_relative_url:
                logger.warning(
                    f"Skipping site '{site_slug}' under root '{root_name}': "
                    f"no readable site_metadata.json (re-run the metadata extraction)."
                )
                continue

            # (container, title, .meta.json folder, kind)
            sources = []
            documents_dir = os.path.join(site_dir, "documents")
            if os.path.isdir(documents_dir):
                for title in sorted(os.listdir(documents_dir)):
                    d = os.path.join(documents_dir, title)
                    if os.path.isdir(d):
                        sources.append(("documents", title, d, "library"))
            lists_dir = os.path.join(site_dir, "lists")
            if os.path.isdir(lists_dir):
                for title in sorted(os.listdir(lists_dir)):
                    d = os.path.join(lists_dir, title, "attachments")
                    if os.path.isdir(d):
                        sources.append(("lists", title, d, "attachment"))

            for container, title, walk_dir, kind in sources:
                container_dir = os.path.join(site_dir, container, title)
                for dirpath, _, filenames in os.walk(walk_dir):
                    for fname in sorted(filenames):
                        if not fname.endswith(".meta.json"):
                            continue
                        with open(
                            os.path.join(dirpath, fname), "r", encoding="utf-8"
                        ) as f:
                            meta = json.load(f)
                        tasks.append(
                            {
                                "kind": kind,
                                "container": container,
                                "root_name": root_name,
                                "site_slug": site_slug,
                                "server_relative_url": server_relative_url,
                                "library_title": title,
                                "name": meta.get("FileLeafRef")
                                or meta.get("FileName")
                                or fname,
                                "file_ref": meta.get("FileRef")
                                or meta.get("ServerRelativeUrl"),
                                "item_id": meta.get("Id") or meta.get("ItemId"),
                                "expected_size": _parse_size(
                                    meta.get("FileSizeDisplay")
                                ),
                                "relative_path": os.path.relpath(
                                    os.path.join(dirpath, fname), container_dir
                                )[: -len(".meta.json")],
                            }
                        )
                        if limit is not None and len(tasks) >= limit:
                            return tasks
    return tasks


def _run_scoped_mode(
    global_limit,
    worker_count,
    site_base_url,
    username,
    password,
    ca_bundle_path,
    logger,
    root_filter,
    site_filter=None,
):
    """
    --all / --site: download every extracted file of the active root in
    scope with one thread pool. global_limit=None means everything (caller
    must confirm).
    """
    output_dir = os.environ.get("SPX_OUTPUT_DIR", "./discovery_output")
    extraction_dir = os.path.join(output_dir, "extraction")

    scope_desc = (
        f"site '{site_filter}' under root '{root_filter}'"
        if site_filter is not None
        else f"all sites/libraries under root '{root_filter}'"
    )
    limit_desc = (
        "no limit (all files)" if global_limit is None else f"limit={global_limit}"
    )
    logger.info(
        f"[scoped mode] Scanning {extraction_dir} for downloadable files ({scope_desc}, {limit_desc})."
    )
    tasks = discover_all_downloadable_files(
        extraction_dir,
        logger,
        global_limit,
        root_filter=root_filter,
        site_filter=site_filter,
    )
    logger.info(
        f"Found {len(tasks)} file(s) to download, using {worker_count} worker(s)."
    )

    if not tasks:
        logger.warning(
            "No downloadable files found for this scope -- check the root/site "
            "name and that metadata extraction has run for this site."
        )
        return

    results = []
    with ThreadPoolExecutor(max_workers=worker_count) as executor:
        futures = {
            executor.submit(
                _download_one_file,
                absolute_url(site_base_url, task["server_relative_url"]),
                username,
                password,
                ca_bundle_path,
                task["library_title"],
                os.path.join(
                    extraction_dir,
                    task["root_name"],
                    task["site_slug"],
                    task["container"],
                    task["library_title"],
                ),
                task["name"],
                task["file_ref"],
                task["item_id"],
                logger,
                root_name=task["root_name"],
                site_relative_url=task["server_relative_url"],
                expected_size=task["expected_size"],
                relative_path=task["relative_path"],
                by_url=task["kind"] == "attachment",
            ): task["name"]
            for task in tasks
        }
        for future in as_completed(futures):
            results.append(future.result())

    summary_name = (
        f"download_all_summary_{root_filter}.json"
        if site_filter is None
        else f"download_site_summary_{root_filter}_{site_filter}.json"
    )
    summary_path = os.path.join(output_dir, summary_name)
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    ok_count = sum(1 for r in results if r["status"] == "ok")
    skipped_count = sum(1 for r in results if r["status"] == "skipped")
    error_count = len(results) - ok_count - skipped_count
    logger.info(
        f"Done: {ok_count} downloaded, {skipped_count} already present (skipped), "
        f"{error_count} failed, out of {len(results)}."
    )
    logger.info(f"Summary written to {summary_path}")


def main():
    enable_persistent_session()

    site_base_url, username, password, ca_bundle_path = get_config()
    root_name = get_root_name()

    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s"
    )
    logger = logging.getLogger("download_files")

    args = sys.argv[1:]

    if "--site" in args:
        idx = args.index("--site")
        if idx + 1 >= len(args) or args[idx + 1].startswith("--"):
            print(
                "Usage: python download_files.py --site <site_slug> [--workers N] [--limit N]"
            )
            print("  e.g.: python download_files.py --site team-a")
            sys.exit(1)
        site_filter = args[idx + 1]

        global_limit = None
        if "--limit" in args:
            l_idx = args.index("--limit")
            try:
                global_limit = int(args[l_idx + 1])
            except (IndexError, ValueError):
                print("[ERROR] --limit requires an integer argument.")
                sys.exit(1)
            if global_limit < 1:
                print("[ERROR] --limit must be at least 1.")
                sys.exit(1)

        worker_count = DEFAULT_WORKER_COUNT
        if "--workers" in args:
            w_idx = args.index("--workers")
            try:
                worker_count = int(args[w_idx + 1])
            except (IndexError, ValueError):
                print("[ERROR] --workers requires an integer argument.")
                sys.exit(1)
            if not (1 <= worker_count <= MAX_WORKER_COUNT):
                print(f"[ERROR] --workers must be between 1 and {MAX_WORKER_COUNT}.")
                sys.exit(1)

        _run_scoped_mode(
            global_limit,
            worker_count,
            site_base_url,
            username,
            password,
            ca_bundle_path,
            logger,
            root_filter=root_name,
            site_filter=site_filter,
        )
        return

    if "--all" in args:
        idx = args.index("--all")
        global_limit = None
        if idx + 1 < len(args) and not args[idx + 1].startswith("--"):
            try:
                global_limit = int(args[idx + 1])
            except ValueError:
                print(
                    "[ERROR] --all's optional argument must be an integer, e.g. --all 50"
                )
                sys.exit(1)
            if global_limit < 1:
                print("[ERROR] --all's limit must be at least 1.")
                sys.exit(1)

        worker_count = DEFAULT_WORKER_COUNT
        if "--workers" in args:
            w_idx = args.index("--workers")
            try:
                worker_count = int(args[w_idx + 1])
            except (IndexError, ValueError):
                print("[ERROR] --workers requires an integer argument.")
                sys.exit(1)
            if not (1 <= worker_count <= MAX_WORKER_COUNT):
                print(f"[ERROR] --workers must be between 1 and {MAX_WORKER_COUNT}.")
                sys.exit(1)

        _run_scoped_mode(
            global_limit,
            worker_count,
            site_base_url,
            username,
            password,
            ca_bundle_path,
            logger,
            root_filter=root_name,
        )
        return

    if len(args) < 2:
        print(
            "Usage: python download_files.py <site_url_suffix> <library_title> [--workers N] [--limit N]"
        )
        print("  e.g.: python download_files.py /sites/hr/team Documents")
        print(
            "Usage: python download_files.py --site <site_slug> [--workers N] [--limit N]"
        )
        print("  e.g.: python download_files.py --site team-a")
        print("Usage: python download_files.py --all <total_limit> [--workers N]")
        print("  e.g.: python download_files.py --all 50")
        sys.exit(1)

    url_suffix, library_title = args[0], args[1]
    if not url_suffix.startswith("/"):
        url_suffix = "/" + url_suffix

    worker_count = DEFAULT_WORKER_COUNT
    if "--workers" in args:
        idx = args.index("--workers")
        try:
            worker_count = int(args[idx + 1])
        except (IndexError, ValueError):
            print("[ERROR] --workers requires an integer argument.")
            sys.exit(1)
        if not (1 <= worker_count <= MAX_WORKER_COUNT):
            print(f"[ERROR] --workers must be between 1 and {MAX_WORKER_COUNT}.")
            sys.exit(1)

    limit = None
    if "--limit" in args:
        idx = args.index("--limit")
        try:
            limit = int(args[idx + 1])
        except (IndexError, ValueError):
            print("[ERROR] --limit requires an integer argument.")
            sys.exit(1)
        if limit < 1:
            print("[ERROR] --limit must be at least 1.")
            sys.exit(1)

    site_url = absolute_url(site_base_url, url_suffix)

    output_dir = os.path.join(
        os.environ.get("SPX_OUTPUT_DIR", "./discovery_output"),
        "extraction",
        root_name,
    )
    # Reuse extract_metadata.py's folder; if not extracted, the name it would get.
    site_dir = _find_site_dir(output_dir, url_suffix) or os.path.join(
        output_dir, site_slug(url_suffix, site_base_url)
    )
    lib_dir = os.path.join(site_dir, "documents", library_title)

    logger.info(f"Connecting to {site_url} to download from '{library_title}'")

    results = download_library(
        site_url,
        username,
        password,
        ca_bundle_path,
        library_title,
        lib_dir,
        logger,
        worker_count,
        limit,
    )

    summary_path = os.path.join(lib_dir, "download_summary.json")
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    ok_count = sum(1 for r in results if r["status"] == "ok")
    skipped_count = sum(1 for r in results if r["status"] == "skipped")
    error_count = len(results) - ok_count - skipped_count
    logger.info(
        f"Done: {ok_count} downloaded, {skipped_count} already present (skipped), "
        f"{error_count} failed, out of {len(results)}."
    )
    logger.info(f"Summary written to {summary_path}")


if __name__ == "__main__":
    main()
