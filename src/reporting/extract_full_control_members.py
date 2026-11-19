"""
CSV with one row per (site, person) holding full control at site level
(site_permissions in permissions.json; list permissions are ignored).

Full control is the built-in role (RoleTypeKind 5, any site language), plus
the custom roles named with --role. permissions.json files written before
role_types existed only match --role names.

Both SharePoint group members and direct grants (users, AD groups,
technical accounts) are kept; direct grants get group_title="(direct)".

Usage:
    python3 extract_full_control_members.py [--base-dir DIR] [--output CSV]
        [--exclude ROOT ...] [--role NAME ...]

Every root under --base-dir is scanned except the --exclude ones.
"""

import os
import csv
import json
import glob
import argparse

FULL_CONTROL_ROLE_TYPE = 5  # SP.RoleType.Administrator


def is_full_control(entry, extra_roles=()):
    if FULL_CONTROL_ROLE_TYPE in (entry.get("role_types") or []):
        return True
    return bool(set(entry.get("roles") or []) & set(extra_roles))


FIELDNAMES = [
    "site_id",
    "site_name",
    "site_url",
    "group_title",
    "person_name",
    "person_login_name",
]


def find_sites(base_dir, excluded_roots=()):
    """(root, slug, metadata_path, permissions_path) for each site with both
    files, either base_dir/<root>/<slug>/ or base_dir/<root>/ (single-site root)."""
    found = []
    for root_dir in sorted(glob.glob(os.path.join(base_dir, "*"))):
        if not os.path.isdir(root_dir):
            continue
        root = os.path.basename(root_dir)
        if root in excluded_roots:
            continue

        metadata_path = os.path.join(root_dir, "site_metadata.json")
        permissions_path = os.path.join(root_dir, "permissions.json")
        if os.path.isfile(metadata_path) and os.path.isfile(permissions_path):
            found.append((root, root, metadata_path, permissions_path))

        for site_dir in sorted(glob.glob(os.path.join(root_dir, "*"))):
            if not os.path.isdir(site_dir):
                continue
            slug = os.path.basename(site_dir)
            metadata_path = os.path.join(site_dir, "site_metadata.json")
            permissions_path = os.path.join(site_dir, "permissions.json")
            if os.path.isfile(metadata_path) and os.path.isfile(permissions_path):
                found.append((root, slug, metadata_path, permissions_path))
    return found


def extract_rows_for_site(site_id, site_name, site_url, permissions, extra_roles=()):
    """Full-control rows from site_permissions."""
    rows = []
    for entry in permissions.get("site_permissions") or []:
        if not is_full_control(entry, extra_roles):
            continue

        if entry.get("type") == "SharePointGroup":
            for member in entry.get("members") or []:
                rows.append(
                    {
                        "site_id": site_id,
                        "site_name": site_name,
                        "site_url": site_url,
                        "group_title": entry.get("title"),
                        "person_name": member.get("title"),
                        "person_login_name": member.get("login_name"),
                    }
                )
        else:
            rows.append(
                {
                    "site_id": site_id,
                    "site_name": site_name,
                    "site_url": site_url,
                    "group_title": "(direct)",
                    "person_name": entry.get("title"),
                    "person_login_name": entry.get("login_name"),
                }
            )
    return rows


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--base-dir",
        default=os.path.join(
            os.environ.get("SPX_OUTPUT_DIR", "./discovery_output"), "extraction"
        ),
        help="Root directory holding one subdirectory per site root (default: %(default)s)",
    )
    parser.add_argument(
        "--output",
        default="full_control_members.csv",
        help="Output CSV path (default: %(default)s)",
    )
    parser.add_argument(
        "--exclude",
        nargs="*",
        default=[],
        metavar="ROOT",
        help="Roots to skip",
    )
    parser.add_argument(
        "--role",
        nargs="*",
        default=[],
        metavar="NAME",
        help="Custom role names that also count as full control",
    )
    args = parser.parse_args()

    if not os.path.isdir(args.base_dir):
        print(f"[ERROR] '{args.base_dir}' is not a directory.")
        raise SystemExit(1)

    sites = find_sites(args.base_dir, set(args.exclude))
    print(
        f"[INFO] Found {len(sites)} site(s) with both site_metadata.json and permissions.json under {args.base_dir}"
    )

    all_rows = []
    skipped = 0
    without_role_types = 0
    for root, slug, metadata_path, permissions_path in sites:
        try:
            with open(metadata_path, "r", encoding="utf-8") as f:
                metadata = json.load(f)
        except (json.JSONDecodeError, OSError) as e:
            print(f"[WARN] Could not read '{metadata_path}': {e}")
            skipped += 1
            continue

        try:
            with open(permissions_path, "r", encoding="utf-8") as f:
                permissions = json.load(f)
        except (json.JSONDecodeError, OSError) as e:
            print(f"[WARN] Could not read '{permissions_path}': {e}")
            skipped += 1
            continue

        site_id = metadata.get("Id")
        site_name = metadata.get("Title")
        site_url = metadata.get("Url")

        entries = permissions.get("site_permissions") or []
        if any("role_types" not in entry for entry in entries):
            without_role_types += 1
        all_rows.extend(
            extract_rows_for_site(site_id, site_name, site_url, permissions, args.role)
        )

    with open(args.output, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writeheader()
        for row in all_rows:
            writer.writerow(row)

    print(f"\n[INFO] {len(all_rows)} (site, person) row(s) written to {args.output}")
    if skipped:
        print(
            f"[INFO] {skipped} site(s) skipped (unreadable metadata or permissions file)"
        )
    if without_role_types:
        print(
            f"[WARN] {without_role_types} site(s) have a permissions.json without "
            "role_types: only --role names were matched there. Re-extract their "
            "permissions (extract_metadata.py --permissions-all)."
        )


if __name__ == "__main__":
    main()
