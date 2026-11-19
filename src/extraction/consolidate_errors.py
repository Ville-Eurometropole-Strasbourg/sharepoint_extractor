"""
Aggregate errors and skip reasons from every run_summary.json on disk, all
roots included.

Usage:
    python3 consolidate_errors.py [base_dir]   # default <SPX_OUTPUT_DIR>/extraction

Writes errors_consolidated.csv (one row per root, slug, error) and prints
counts by root and category.
"""

import os
import sys
import csv
import json
import glob
from collections import Counter


def find_run_summaries(base_dir):
    """(root, path) of every <root>/run_summary.json or <root>/<site>/run_summary.json."""
    patterns = [
        os.path.join(base_dir, "*", "run_summary.json"),
        os.path.join(base_dir, "*", "*", "run_summary.json"),
    ]
    found = []
    for pattern in patterns:
        for path in glob.glob(pattern):
            rel = os.path.relpath(path, base_dir)
            root = rel.split(os.sep)[0]
            found.append((root, path))
    return found


def classify_error_message(msg):
    """Coarse error category."""
    if "404" in msg:
        return "404_not_found"
    if (
        "ConnectionError" in msg
        or "Connection refused" in msg
        or "NewConnectionError" in msg
    ):
        return "connection_error"
    if "Timeout" in msg or "timed out" in msg.lower():
        return "timeout"
    if "SSL" in msg or "certificate" in msg.lower():
        return "ssl_error"
    if "401" in msg or "403" in msg:
        return "auth_error"
    return "other"


def main():
    base_dir = (
        sys.argv[1]
        if len(sys.argv) > 1
        else os.path.join(
            os.environ.get("SPX_OUTPUT_DIR", "./discovery_output"), "extraction"
        )
    )

    if not os.path.isdir(base_dir):
        print(
            f"[ERROR] '{base_dir}' is not a directory. Pass the correct path as an argument."
        )
        sys.exit(1)

    summaries = find_run_summaries(base_dir)
    print(f"[INFO] Found {len(summaries)} run_summary.json file(s) under {base_dir}")

    error_rows = []
    skip_rows = []

    for root, path in summaries:
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except (json.JSONDecodeError, OSError) as e:
            print(f"[WARN] Could not read '{path}': {e}")
            continue

        for slug, content in data.items():
            if not isinstance(content, dict):
                continue

            if "error" in content:
                err = content["error"]
                msg = err[2] if isinstance(err, list) and len(err) > 2 else str(err)
                error_rows.append(
                    {
                        "root": root,
                        "slug": slug,
                        "source_file": path,
                        "error_category": classify_error_message(msg),
                        "error_message": msg,
                    }
                )

            for skipped_item in content.get("skipped", []):
                skip_rows.append(
                    {
                        "root": root,
                        "slug": slug,
                        "list_title": skipped_item.get("title"),
                        "item_count": skipped_item.get("item_count"),
                        "reason": skipped_item.get("reason"),
                    }
                )

            for section in ("document_libraries", "generic_lists"):
                for list_title, list_result in content.get(section, {}).items():
                    if isinstance(list_result, dict) and "error" in list_result:
                        category = (
                            list_result.get("error_category")
                            or f"partial_{section}_error"
                        )
                        error_rows.append(
                            {
                                "root": root,
                                "slug": slug,
                                "source_file": path,
                                "error_category": category,
                                "error_message": f"[{list_title}] {list_result['error']}",
                            }
                        )

    out_path = "errors_consolidated.csv"
    fieldnames = ["root", "slug", "error_category", "error_message", "source_file"]
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in error_rows:
            writer.writerow(row)

    print(f"\n[INFO] {len(error_rows)} site/list-level error(s) written to {out_path}")
    scanned_roots = sorted(set(root for root, _ in summaries))
    by_root = Counter(r["root"] for r in error_rows)
    print("\n[SUMMARY] Errors by root (all scanned roots shown, including clean ones):")
    for root in scanned_roots:
        print(f"  {root}: {by_root.get(root, 0)}")

    by_category = Counter(r["error_category"] for r in error_rows)
    print("\n[SUMMARY] Errors by category:")
    for cat, count in by_category.most_common():
        print(f"  {cat}: {count}")

    if skip_rows:
        by_reason = Counter(r["reason"] for r in skip_rows)
        print(
            f"\n[INFO] {len(skip_rows)} skipped list/library entrie(s) (not errors, informational):"
        )
        for reason, count in by_reason.most_common():
            print(f"  {reason}: {count}")


if __name__ == "__main__":
    main()
