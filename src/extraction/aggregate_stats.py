"""
Stats from the extraction on disk (no SharePoint call), per site, per root
and overall: sites with content, dead sites (404), file count by extension
group, total size and version history.

Migration volume = total_size_bytes + history_size_bytes (history excludes
the current version). Unknown sizes/histories are counted apart, never as 0.

Usage:
    python aggregate_stats.py [--output stats_summary.json]
"""

import os
import csv
import json
import argparse
from collections import defaultdict

EXTENSION_GROUPS = {
    "images": {
        ".jpg",
        ".jpeg",
        ".png",
        ".gif",
        ".bmp",
        ".tiff",
        ".tif",
        ".svg",
        ".webp",
        ".heic",
    },
    "pdf": {".pdf"},
    "office_word": {".doc", ".docx", ".dot", ".dotx", ".rtf"},
    "office_excel": {".xls", ".xlsx", ".xlsm", ".csv"},
    "office_powerpoint": {".ppt", ".pptx", ".pps", ".ppsx"},
    "onenote": {".one", ".onetoc2", ".onetemp"},
    "video": {".mp4", ".mov", ".avi", ".wmv", ".mkv"},
    "audio": {".mp3", ".wav", ".wma"},
    "web_pages": {".aspx", ".html", ".htm"},
    "archives": {".zip", ".rar", ".7z"},
}

# Size field names, by preference.
SIZE_FIELD_CANDIDATES = ("FileSizeDisplay", "File_x0020_Size", "FileSize", "Length")


def classify_extension(filename):
    """Extension group of a filename, or 'other'."""
    ext = os.path.splitext(filename)[1].lower()
    for group, extensions in EXTENSION_GROUPS.items():
        if ext in extensions:
            return group
    return "other"


def read_meta(meta_path):
    """A .meta.json's content, None when unreadable."""
    try:
        with open(meta_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return None


def read_file_size(meta):
    """(size_in_bytes, version_label) from a .meta.json's content, None when unknown."""
    if meta is None:
        return None, None

    label = meta.get("OData__UIVersionString")
    for field in SIZE_FIELD_CANDIDATES:
        value = meta.get(field)
        if value is not None:
            try:
                return int(value), label
            except (TypeError, ValueError):
                continue
    return None, label


def read_history(meta_path, version_label):
    """
    (history_versions, history_size_bytes) from the .versions.json, (0, 0)
    for "1.0"/"0.1" labels, else None (unknown).
    """
    versions_path = meta_path[: -len(".meta.json")] + ".versions.json"
    if os.path.exists(versions_path):
        try:
            with open(versions_path, "r", encoding="utf-8") as f:
                versions = json.load(f)
            return versions["history_count"], versions["history_size_bytes"]
        except (OSError, json.JSONDecodeError, KeyError):
            return None
    if version_label in ("1.0", "0.1"):
        return 0, 0
    return None


def load_dead_sites(errors_csv_path):
    """{root_name: set(slugs)} of sites in 404_not_found."""
    dead_sites = defaultdict(set)
    if not os.path.exists(errors_csv_path):
        return dead_sites
    with open(errors_csv_path, "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        for row in reader:
            if len(row) < 3:
                continue
            root, slug, category = row[0], row[1], row[2]
            if category == "404_not_found":
                dead_sites[root].add(slug)
    return dead_sites


def new_counters():
    """Empty counters, same shape for site/root/total."""
    return {
        "total_files": 0,
        "by_extension_group": defaultdict(int),
        "total_size_bytes": 0,
        "files_with_unknown_size": 0,
        "versioned_files": 0,
        "history_versions": 0,
        "history_size_bytes": 0,
        "files_with_unknown_history": 0,
    }


# All counters except by_extension_group.
SUMMED_COUNTERS = (
    "total_files",
    "total_size_bytes",
    "files_with_unknown_size",
    "versioned_files",
    "history_versions",
    "history_size_bytes",
    "files_with_unknown_history",
)


def accumulate(target, group, size_bytes, history):
    target["total_files"] += 1
    target["by_extension_group"][group] += 1
    if size_bytes is None:
        target["files_with_unknown_size"] += 1
    else:
        target["total_size_bytes"] += size_bytes
    if history is None:
        target["files_with_unknown_history"] += 1
    else:
        history_versions, history_size_bytes = history
        if history_versions > 0:
            target["versioned_files"] += 1
        target["history_versions"] += history_versions
        target["history_size_bytes"] += history_size_bytes


def merge_into(target, source):
    """Add source counters into target."""
    for key in SUMMED_COUNTERS:
        target[key] += source[key]
    for group, count in source["by_extension_group"].items():
        target["by_extension_group"][group] += count


def finalize(counters):
    """Counters as plain dicts for JSON."""
    return {
        **{key: counters[key] for key in SUMMED_COUNTERS},
        "by_extension_group": dict(counters["by_extension_group"]),
    }


def scan_extraction_dir(extraction_dir):
    """
    Count per site and per root. Sites without any documents/*.meta.json are ignored.

    Returns {root_name: {"sites": {site_slug: counters}, "root_totals": counters}}.
    """
    stats = {}

    for root_name in sorted(os.listdir(extraction_dir)):
        root_path = os.path.join(extraction_dir, root_name)
        if not os.path.isdir(root_path):
            continue

        root_totals = new_counters()
        sites = {}

        for site_slug in sorted(os.listdir(root_path)):
            documents_dir = os.path.join(root_path, site_slug, "documents")
            if not os.path.isdir(documents_dir):
                continue

            site_counters = new_counters()
            for dirpath, _, filenames in os.walk(documents_dir):
                for fname in filenames:
                    if not fname.endswith(".meta.json"):
                        continue
                    meta_path = os.path.join(dirpath, fname)
                    meta = read_meta(meta_path)
                    if meta is not None and meta.get("FileSystemObjectType") == 1:
                        continue  # folder or document set
                    real_name = fname[: -len(".meta.json")]
                    group = classify_extension(real_name)
                    size_bytes, version_label = read_file_size(meta)
                    history = read_history(meta_path, version_label)
                    accumulate(site_counters, group, size_bytes, history)

            if site_counters["total_files"] > 0:
                sites[site_slug] = site_counters
                merge_into(root_totals, site_counters)

        stats[root_name] = {"sites": sites, "root_totals": root_totals}

    return stats


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output",
        default=None,
        help="Path to write the aggregated stats as JSON. Prints to stdout if omitted.",
    )
    args = parser.parse_args()

    output_dir = os.environ.get("SPX_OUTPUT_DIR", "./discovery_output")
    extraction_dir = os.path.join(output_dir, "extraction")
    errors_csv_path = "./errors_consolidated.csv"

    print(f"[INFO] Scanning {extraction_dir} ...")
    scanned = scan_extraction_dir(extraction_dir)

    print(f"[INFO] Reading dead-site list from {errors_csv_path} ...")
    dead_sites = load_dead_sites(errors_csv_path)

    summary = {}
    grand_totals = new_counters()
    total_sites_with_content = 0
    total_dead = 0

    for root_name, root_data in scanned.items():
        root_dead_slugs = dead_sites.get(root_name, set())

        sites_summary = {}
        for site_slug, site_counters in root_data["sites"].items():
            sites_summary[site_slug] = finalize(site_counters)
            sites_summary[site_slug]["is_dead"] = site_slug in root_dead_slugs

        summary[root_name] = {
            "sites_with_content": len(root_data["sites"]),
            "dead_sites": len(root_dead_slugs),
            **finalize(root_data["root_totals"]),
            "sites": sites_summary,
        }

        total_sites_with_content += len(root_data["sites"])
        total_dead += len(root_dead_slugs)
        merge_into(grand_totals, root_data["root_totals"])

    summary["_totals"] = {
        "sites_with_content": total_sites_with_content,
        "dead_sites": total_dead,
        **finalize(grand_totals),
    }

    output_json = json.dumps(summary, ensure_ascii=False, indent=2)
    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(output_json)
        print(f"[INFO] Written to {args.output}")
    else:
        print(output_json)


if __name__ == "__main__":
    main()
