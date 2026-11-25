"""Read-only recount of broken permission inheritance from permissions.json.

Expected layout:
    {
      "format": 2,
      "site_has_unique_role_assignments": bool,
      "site_permissions": [principal, ...],
      "lists_with_broken_inheritance": {"<list title>": [principal, ...]},
      "items_with_broken_inheritance": {
        "<list title>": [{id, kind, url, permissions: [principal, ...]}]
      },
      "items_scan_errors": [{list, item_id, error}]
    }
    principal = {login_name, title, type, roles, members_resolvable, members?}
    kind = folder, file or item

A list appears in lists_with_broken_inheritance only if its inheritance is
broken, so the total number of lists is not known from this file. Files
without "format" predate item permissions: extract_metadata.py
--permissions-all redoes them.
Site-level counts are split into top-level sites and sub-sites (parent found
through site_metadata.json). Only names and counts are printed.

Usage:
    python3 src/tools/check_broken_inheritance.py <root> [--base DIR] [--csv FILE]
"""

import argparse
import csv
import json
import os
import sys
from collections import Counter

DEFAULT_BASE = os.path.join(
    os.environ.get("SPX_OUTPUT_DIR", "./discovery_output"), "extraction"
)


def load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def site_urls(root_dir):
    """{slug: lower-case ServerRelativeUrl without trailing slash}."""
    urls = {}
    for slug in os.listdir(root_dir):
        path = os.path.join(root_dir, slug, "site_metadata.json")
        if not os.path.isfile(path):
            continue
        try:
            url = load_json(path).get("ServerRelativeUrl")
        except (OSError, ValueError):
            continue
        if url:
            urls[slug] = url.rstrip("/").lower()
    return urls


def has_parent(url, all_urls):
    parts = url.split("/")
    return any("/".join(parts[:i]) in all_urls for i in range(len(parts) - 1, 0, -1))


def main():
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("root", help="extraction root, e.g. hr")
    parser.add_argument("--base", default=DEFAULT_BASE)
    parser.add_argument("--csv", help="write one row per site to this file")
    args = parser.parse_args()

    root_dir = os.path.join(args.base, args.root)
    if not os.path.isdir(root_dir):
        sys.exit(f"not a directory: {root_dir}")

    urls = site_urls(root_dir)
    url_set = set(urls.values())

    rows = []
    unreadable = []
    broken_titles = Counter()
    principal_types = Counter()
    principals_total = principals_unresolvable = 0
    item_kinds = Counter()
    item_principals = 0
    without_items = []
    sites_with_scan_errors = 0
    scan_errors_total = 0

    for slug in sorted(os.listdir(root_dir)):
        path = os.path.join(root_dir, slug, "permissions.json")
        if not os.path.isfile(path):
            continue
        try:
            data = load_json(path)
        except (OSError, ValueError) as exc:
            unreadable.append(f"{slug} ({type(exc).__name__})")
            continue
        broken = data.get("lists_with_broken_inheritance") or {}
        for title in broken:
            broken_titles[title] += 1
        for principals in [data.get("site_permissions") or []] + list(broken.values()):
            for p in principals:
                principals_total += 1
                principal_types[str(p.get("type"))] += 1
                if p.get("members_resolvable") is False:
                    principals_unresolvable += 1
        items_by_list = data.get("items_with_broken_inheritance")
        if items_by_list is None:
            without_items.append(slug)
            items_by_list = {}
        site_items = [i for items in items_by_list.values() for i in items]
        for item in site_items:
            item_kinds[str(item.get("kind"))] += 1
            item_principals += len(item.get("permissions") or [])
        scan_errors = len(data.get("items_scan_errors") or [])
        scan_errors_total += scan_errors
        sites_with_scan_errors += scan_errors > 0
        url = urls.get(slug)
        kind = (
            "unknown" if url is None else ("sub" if has_parent(url, url_set) else "top")
        )
        rows.append(
            (
                slug,
                kind,
                bool(data.get("site_has_unique_role_assignments")),
                len(broken),
                "" if slug in without_items else len(site_items),
                "" if slug in without_items else scan_errors,
            )
        )

    n = len(rows)
    print(f"permissions.json read: {n}")
    print(f"unreadable: {len(unreadable)}")
    for name in unreadable:
        print(f"   {name}")

    for kind in ("top", "sub", "unknown"):
        subset = [r for r in rows if r[1] == kind]
        if subset:
            unique = sum(1 for r in subset if r[2])
            print(
                f"{kind}-level sites: {len(subset)}, unique permissions at site level: {unique}"
            )

    with_broken = sum(1 for r in rows if r[3] > 0)
    print(f"sites with at least one list with broken inheritance: {with_broken}")
    if n:
        print(f"  = {100 * with_broken / n:.1f} % of sites read")
    print(f"lists with broken inheritance (total): {sum(r[3] for r in rows)}")
    print(f"distinct list titles: {len(broken_titles)}")
    print("most frequent titles among broken lists:")
    for title, c in broken_titles.most_common(20):
        print(f"  {c:>5} x {title}")

    print(f"principal entries (site + lists): {principals_total}")
    for t, c in principal_types.most_common():
        print(f"  type={t}: {c}")
    print(f"entries with members_resolvable = false: {principals_unresolvable}")

    print(
        f"permissions.json without item permissions (older format): "
        f"{len(without_items)}"
    )
    if without_items:
        print("  run extract_metadata.py --permissions-all to complete them")
    sites_with_items = sum(1 for r in rows if r[4] not in ("", 0))
    print(f"sites with at least one item with broken inheritance: {sites_with_items}")
    print(
        f"items with broken inheritance (total): {sum(item_kinds.values())}"
        + (
            " (" + ", ".join(f"{k}={c}" for k, c in sorted(item_kinds.items())) + ")"
            if item_kinds
            else ""
        )
    )
    print(f"principal entries on items: {item_principals}")
    print(
        f"item scan errors: {scan_errors_total} "
        f"(sites with at least one: {sites_with_scan_errors})"
    )

    if args.csv:
        with open(args.csv, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(
                [
                    "slug",
                    "level",
                    "site_unique",
                    "lists_broken",
                    "items_broken",
                    "items_scan_errors",
                ]
            )
            writer.writerows(rows)
        print(f"csv written: {args.csv}")


if __name__ == "__main__":
    main()
