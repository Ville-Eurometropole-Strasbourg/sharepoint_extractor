"""Read-only inventory of one extracted site (reference counts).

Prints only names of libraries/lists/pages and counts, never contents or
people. Layout assumed:
    <site>/documents/<library>/...   files
    <site>/lists/<list>/items.json   + other files (schema, views, ...)
    <site>/pages/...                 files
    <site>/permissions.json

Usage:
    python3 src/tools/inventory_site.py <root> <site_slug> [--base DIR]
"""

import argparse
import json
import os
import sys
from collections import Counter

DEFAULT_BASE = os.path.join(
    os.environ.get("SPX_OUTPUT_DIR", "./discovery_output"), "extraction"
)
LIMITED_ACCESS_ROLE_TYPE = 1  # SP.RoleType.Guest ("Limited Access")


def walk_files(path):
    for dirpath, _, filenames in os.walk(path):
        for name in filenames:
            yield os.path.join(dirpath, name)


def human(size):
    for unit in ("o", "Ko", "Mo", "Go"):
        if size < 1024:
            return f"{size:.0f} {unit}"
        size /= 1024
    return f"{size:.1f} To"


def documents_section(site_dir):
    docs = os.path.join(site_dir, "documents")
    print("== documents ==")
    if not os.path.isdir(docs):
        print("  (none)")
        return
    for library in sorted(os.listdir(docs)):
        lib_dir = os.path.join(docs, library)
        if not os.path.isdir(lib_dir):
            continue
        files = list(walk_files(lib_dir))
        size = sum(os.path.getsize(f) for f in files)
        exts = Counter(os.path.splitext(f)[1].lower() or "<none>" for f in files)
        top = ", ".join(f"{e}:{c}" for e, c in exts.most_common(5))
        print(f"  {library}: {len(files)} files, {human(size)} [{top}]")


def lists_section(site_dir):
    lists_dir = os.path.join(site_dir, "lists")
    print("== lists ==")
    if not os.path.isdir(lists_dir):
        print("  (none)")
        return set()
    names = set()
    file_kinds = Counter()
    total_items = 0
    for name in sorted(os.listdir(lists_dir)):
        list_dir = os.path.join(lists_dir, name)
        if not os.path.isdir(list_dir):
            continue
        names.add(name)
        files = list(walk_files(list_dir))
        for f in files:
            file_kinds[os.path.relpath(f, list_dir).split(os.sep)[0]] += 1
        items = "?"
        items_path = os.path.join(list_dir, "items.json")
        if os.path.isfile(items_path):
            try:
                with open(items_path, encoding="utf-8") as fh:
                    items = len(json.load(fh))
                total_items += items
            except (OSError, ValueError):
                items = "unreadable"
        print(f"  {name}: {items} items, {len(files)} files")
    print(f"  -> {len(names)} lists, {total_items} items")
    print("  entries found in list folders (name: count):")
    for kind, c in file_kinds.most_common(15):
        print(f"    {kind}: {c}")
    return names


def pages_section(site_dir):
    pages = os.path.join(site_dir, "pages")
    print("== pages ==")
    if not os.path.isdir(pages):
        print("  (none)")
        return
    for f in sorted(walk_files(pages)):
        print(f"  {os.path.relpath(f, pages)} ({human(os.path.getsize(f))})")


def permissions_section(site_dir, list_names):
    path = os.path.join(site_dir, "permissions.json")
    print("== permissions ==")
    if not os.path.isfile(path):
        print("  (none)")
        return
    with open(path, encoding="utf-8") as fh:
        data = json.load(fh)
    print(
        f"  site_has_unique_role_assignments: {data.get('site_has_unique_role_assignments')}"
    )
    site = data.get("site_permissions") or []
    broken = data.get("lists_with_broken_inheritance") or {}
    print(f"  site-level principals: {len(site)}")
    print(f"  lists/libraries with broken inheritance: {len(broken)}")
    types = Counter()
    roles = Counter()
    for principals in [site] + list(broken.values()):
        for p in principals:
            types[str(p.get("type"))] += 1
            for r in p.get("roles") or []:
                roles[r] += 1
    print(
        "  principal entries by type: "
        + ", ".join(f"{t}={c}" for t, c in types.most_common())
    )
    print("  roles used: " + ", ".join(f"{r}={c}" for r, c in roles.most_common()))
    print("  broken lists/libraries (principals, where):")
    for title, principals in sorted(broken.items(), key=lambda kv: -len(kv[1])):
        where = "list" if title in list_names else "library or other"
        print(f"    {title}: {len(principals)} ({where})")

    # Group broken lists by identical assignments, ignoring Limited Access
    # (needs role_types, absent from older permissions.json files).
    def signature(principals):
        result = set()
        for p in principals:
            roles = p.get("roles") or []
            role_types = p.get("role_types") or [None] * len(roles)
            for role, role_type in zip(roles, role_types):
                if role_type != LIMITED_ACCESS_ROLE_TYPE:
                    result.add((p.get("login_name"), role))
        return frozenset(result)

    site_sig = signature(site)
    groups = {}
    for title, principals in broken.items():
        groups.setdefault(signature(principals), []).append(title)
    print(f"  distinct permission sets among broken lists: {len(groups)}")
    for i, (sig, titles) in enumerate(
        sorted(groups.items(), key=lambda kv: -len(kv[1])), 1
    ):
        same = " = same as site" if sig == site_sig else ""
        print(f"    set {i}: {len(sig)} assignments, {len(titles)} lists{same}")
        print(f"      {', '.join(sorted(titles))}")


def main():
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("root")
    parser.add_argument("site_slug")
    parser.add_argument("--base", default=DEFAULT_BASE)
    args = parser.parse_args()

    site_dir = os.path.join(args.base, args.root, args.site_slug)
    if not os.path.isdir(site_dir):
        sys.exit(f"not a directory: {site_dir}")
    documents_section(site_dir)
    names = lists_section(site_dir)
    pages_section(site_dir)
    permissions_section(site_dir, names)


if __name__ == "__main__":
    main()
