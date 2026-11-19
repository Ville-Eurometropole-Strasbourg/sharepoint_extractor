"""
Extract the user table of a site's collection to site_users.json, from the
hidden User Information List (former users and SharePoint groups
included). Each entry: Id, Title, LoginName, UserName, kind (person, group or
the raw ContentTypeId), deleted, hidden, and active (still returned by
_api/web/siteusers). These Ids are the values of the <Column>Id fields of list
items (AuthorId, AssignedToId...); they are scoped to the site collection, so
the file is written for each site. No contact details are kept.

Prints counts only, and compares the total with the item count of the hidden
list recorded in run_summary.json, when present.

Usage (in the container, from the project root):
    python src/extraction/extract_site_users.py /sites/hr/team team
"""

import json
import os
import sys
from collections import Counter

SRC_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, SRC_DIR)
from sharepoint_client import get_config, build_context  # noqa: E402
from site_urls import absolute_url, get_root_name
from persistent_session import enable_persistent_session  # noqa: E402

INFO_FIELDS = [
    "Id",
    "Title",
    "Name",
    "UserName",
    "ContentTypeId",
    "Deleted",
    "UserInfoHidden",
]
CONTENT_TYPE_KINDS = {"0x010A": "person", "0x010B": "group"}
USER_INFO_TEMPLATE = 112  # BaseTemplate of the User Information List


def kind_of(content_type_id):
    prefix = (content_type_id or "")[:6].upper().replace("0X", "0x")
    return CONTENT_TYPE_KINDS.get(prefix, content_type_id or "unknown")


def fetch_user_table(ctx):
    items = (
        ctx.web.site_user_info_list.items.select(INFO_FIELDS)
        .get_all(page_size=100)
        .execute_query()
    )
    active_users = ctx.web.site_users.select(["Id"]).get().execute_query()
    active_ids = {user.properties.get("Id") for user in active_users}
    table = []
    for item in items:
        props = item.properties
        table.append(
            {
                "Id": props.get("Id"),
                "Title": props.get("Title"),
                "LoginName": props.get("Name"),
                "UserName": props.get("UserName"),
                "kind": kind_of(props.get("ContentTypeId")),
                "deleted": bool(props.get("Deleted")),
                "hidden": bool(props.get("UserInfoHidden")),
                "active": props.get("Id") in active_ids,
            }
        )
    return table, len(active_ids)


def write_atomic(path, data):
    tmp_path = path + ".part"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp_path, path)


def user_info_list_count(root_dir, slug):
    try:
        with open(os.path.join(root_dir, "run_summary.json"), encoding="utf-8") as f:
            entry = json.load(f).get(slug, {})
    except (FileNotFoundError, json.JSONDecodeError):
        return None
    for skipped in entry.get("skipped", []):
        if skipped.get("base_template") == USER_INFO_TEMPLATE:
            return skipped.get("item_count")
    return None


def main():
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    url_suffix, slug = sys.argv[1], sys.argv[2]

    enable_persistent_session()
    site_base_url, username, password, ca_bundle_path = get_config()
    root_name = get_root_name()
    root_dir = os.path.join(
        os.environ.get("SPX_OUTPUT_DIR", "./discovery_output"),
        "extraction",
        root_name,
    )
    site_dir = os.path.join(root_dir, slug)
    if not os.path.isdir(site_dir):
        sys.exit(f"Site folder not found: {site_dir}")

    ctx = build_context(
        absolute_url(site_base_url, url_suffix), username, password, ca_bundle_path
    )
    table, active_count = fetch_user_table(ctx)
    write_atomic(os.path.join(site_dir, "site_users.json"), table)

    ids = [entry["Id"] for entry in table]
    print(f"Site: {slug}")
    print(f"Entries: {len(table)} (distinct ids: {len(set(ids))})")
    print(f"User Information List item count: {user_info_list_count(root_dir, slug)}")
    print(f"By kind: {dict(Counter(entry['kind'] for entry in table))}")
    print(
        f"Active: {sum(entry['active'] for entry in table)} "
        f"(siteusers returned {active_count})"
    )
    print(f"Deleted: {sum(entry['deleted'] for entry in table)}")
    print(f"Hidden: {sum(entry['hidden'] for entry in table)}")
    print(f"Without title: {sum(1 for entry in table if not entry['Title'])}")


if __name__ == "__main__":
    main()
