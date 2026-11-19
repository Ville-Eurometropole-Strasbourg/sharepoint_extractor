"""Read-only recount of broken permission inheritance from permissions.json.

Expected layout:
    {
      "site_has_unique_role_assignments": bool,
      "site_permissions": [principal, ...],
      "lists_with_broken_inheritance": {"<list title>": [principal, ...]}
    }
    principal = {login_name, title, type, roles, members_resolvable, members?}

A list appears in lists_with_broken_inheritance only if its inheritance is
broken, so the total number of lists is not known from this file.
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

    if args.csv:
        with open(args.csv, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["slug", "level", "site_unique", "lists_broken"])
            writer.writerows(rows)
        print(f"csv written: {args.csv}")


if __name__ == "__main__":
    main()
