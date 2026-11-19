"""Read-only check: sub-site names that collide with another name.

Compares, for one extraction root, the sites found on disk
(site_metadata.json -> ServerRelativeUrl) and reports:

A. sites whose ServerRelativeUrl is identical ignoring case;
B. slugs (folder names) that are identical ignoring case;
C. sub-sites whose last URL segment equals, ignoring case, an entry of
   their parent site on disk (a list/library folder, a list under lists/,
   a page under pages/content/).

Only names and counts are printed, never file contents.

Usage:
    python3 src/tools/check_subweb_name_conflicts.py <root> [--base DIR]
    e.g. python3 src/tools/check_subweb_name_conflicts.py hr
"""

import argparse
import json
import os
import sys
from collections import defaultdict

DEFAULT_BASE = os.path.join(
    os.environ.get("SPX_OUTPUT_DIR", "./discovery_output"), "extraction"
)


def load_sites(root_dir):
    """Return {slug: server_relative_url} for every site folder with metadata."""
    sites = {}
    for slug in sorted(os.listdir(root_dir)):
        meta_path = os.path.join(root_dir, slug, "site_metadata.json")
        if not os.path.isfile(meta_path):
            continue
        try:
            with open(meta_path, encoding="utf-8") as f:
                url = json.load(f).get("ServerRelativeUrl")
        except (OSError, ValueError):
            print(f"unreadable site_metadata.json: {slug}", file=sys.stderr)
            continue
        if url:
            sites[slug] = url.rstrip("/")
    return sites


def find_parent(url, urls_by_lower):
    """Longest other site URL that is a path prefix of `url`."""
    parts = url.lower().split("/")
    for i in range(len(parts) - 1, 0, -1):
        candidate = "/".join(parts[:i])
        if candidate in urls_by_lower:
            return urls_by_lower[candidate]
    return None


def names_on_disk(site_dir, subweb_slugs):
    """{name: kind} of things stored in a site folder, excluding sub-sites.

    kind is "folder" (top-level folder), "list" (under lists/) or
    "page:<file name>" (under pages/content/).
    """
    names = {}
    if not os.path.isdir(site_dir):
        return names
    for entry in os.listdir(site_dir):
        if entry in subweb_slugs or entry in ("lists", "pages"):
            continue
        if os.path.isdir(os.path.join(site_dir, entry)):
            names[entry] = "folder"
    lists_dir = os.path.join(site_dir, "lists")
    if os.path.isdir(lists_dir):
        for entry in os.listdir(lists_dir):
            names[entry] = "list"
    content_dir = os.path.join(site_dir, "pages", "content")
    if os.path.isdir(content_dir):
        for page in os.listdir(content_dir):
            names[page.split(".")[0]] = f"page:{page}"
    return names


def main():
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("root", help="extraction root, e.g. hr")
    parser.add_argument("--base", default=DEFAULT_BASE)
    args = parser.parse_args()

    root_dir = os.path.join(args.base, args.root)
    if not os.path.isdir(root_dir):
        sys.exit(f"not a directory: {root_dir}")

    sites = load_sites(root_dir)
    print(f"sites with metadata: {len(sites)}")

    by_url = defaultdict(list)
    by_slug = defaultdict(list)
    for slug, url in sites.items():
        by_url[url.lower()].append(slug)
        by_slug[slug.lower()].append(slug)

    dup_urls = {u: s for u, s in by_url.items() if len(s) > 1}
    dup_slugs = {k: s for k, s in by_slug.items() if len(s) > 1}
    print(f"A. same URL ignoring case: {len(dup_urls)}")
    for url, slugs in dup_urls.items():
        print(f"   {url}: {slugs}")
    print(f"B. same slug ignoring case: {len(dup_slugs)}")
    for slugs in dup_slugs.values():
        print(f"   {slugs}")

    urls_by_lower = {url.lower(): slug for slug, url in sites.items()}
    children = defaultdict(list)
    for slug, url in sites.items():
        parent = find_parent(url, urls_by_lower)
        if parent:
            children[parent].append(slug)
    n_subwebs = sum(len(v) for v in children.values())
    print(f"sub-sites with a parent found on disk: {n_subwebs}")

    conflicts = []
    for parent, subs in children.items():
        taken = names_on_disk(os.path.join(root_dir, parent), set(subs))
        taken_lower = {n.lower(): (n, kind) for n, kind in taken.items()}
        for sub in subs:
            segment = sites[sub].rsplit("/", 1)[-1]
            if segment.lower() in taken_lower:
                entry, kind = taken_lower[segment.lower()]
                conflicts.append((parent, sub, segment, entry, kind))
    print(f"C. sub-site name equal to an entry of its parent: {len(conflicts)}")
    for parent, sub, segment, entry, kind in conflicts:
        print(
            f"   parent={parent} sub-site={sub} segment={segment} entry={entry} kind={kind}"
        )


if __name__ == "__main__":
    main()
