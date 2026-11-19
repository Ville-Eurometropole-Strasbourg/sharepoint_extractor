"""Read-only check of the lists, libraries and views of one SharePoint site.

Compares each list's ItemCount with what the extraction wrote on disk, and
summarises its views (displayed fields, filter, sort, grouping).
Prints names and counts only: no item values, no filter values.
"""

import argparse
import json
import os
import re
import sys

import requests
from requests_ntlm import HttpNtlmAuth

TIMEOUT = (10, 60)
HEADERS = {
    "Accept": "application/json;odata=nometadata",
    "X-FORMS_BASED_AUTH_ACCEPTED": "f",
}
USER_ENV = "SPX_USERNAME"
PASSWORD_ENV = "SPX_PASSWORD"
CA_BUNDLE_ENV = "SPX_CA_BUNDLE"
LISTS_TOP = 500
DOC_LIBRARY_BASE_TYPE = 1


def get_json(session, url):
    response = session.get(url, headers=HEADERS, timeout=TIMEOUT)
    response.raise_for_status()
    return response.json()


def fetch_lists(session, site_url):
    url = (
        f"{site_url}/_api/web/lists"
        "?$select=Id,Title,ItemCount,BaseType,BaseTemplate,Hidden,RootFolder/Name"
        f"&$expand=RootFolder&$top={LISTS_TOP}"
    )
    lists = get_json(session, url)["value"]
    if len(lists) >= LISTS_TOP:
        print(f"WARNING: {LISTS_TOP} lists returned, result may be truncated")
    return lists


def fetch_views(session, site_url, list_id):
    url = f"{site_url}/_api/web/lists(guid'{list_id}')/views?$expand=ViewFields"
    return get_json(session, url)["value"]


def summarize_query(caml):
    """Return the field names used by Where, OrderBy and GroupBy, never values."""
    caml = caml or ""
    summary = {}
    for tag in ("Where", "OrderBy", "GroupBy"):
        match = re.search(rf"<{tag}\b.*?</{tag}>", caml, re.S)
        summary[tag] = (
            re.findall(r'<FieldRef[^>]*\bName="([^"]+)"', match.group(0))
            if match
            else []
        )
    return summary


def view_fields(view):
    fields = view.get("ViewFields") or {}
    return fields.get("Items") or []


def count_items(path):
    try:
        with open(path, encoding="utf-8") as handle:
            data = json.load(handle)
    except (OSError, ValueError):
        return None
    if isinstance(data, list):
        return len(data)
    if isinstance(data, dict):
        for key in ("value", "items", "Items"):
            if isinstance(data.get(key), list):
                return len(data[key])
    return None


def count_disk(site_dir):
    """Count items.json entries and .meta.json files per top-level container."""
    items = {}
    metas = {}
    attachments = 0
    for dirpath, dirnames, filenames in os.walk(site_dir):
        dirnames[:] = [d for d in dirnames if d != ".history"]
        parts = os.path.relpath(dirpath, site_dir).split(os.sep)
        if "items.json" in filenames:
            key = os.sep.join(parts[:2])
            items[key] = count_items(os.path.join(dirpath, "items.json"))
        n_meta = sum(1 for f in filenames if f.endswith(".meta.json"))
        if not n_meta:
            continue
        if "attachments" in parts:
            attachments += n_meta
        else:
            key = os.sep.join(parts[:2])
            metas[key] = metas.get(key, 0) + n_meta
    return items, metas, attachments


def match_disk(sp_list, disk_counts):
    names = {sp_list["Title"].lower(), sp_list["RootFolder"]["Name"].lower()}
    for key, count in disk_counts.items():
        if key.split(os.sep)[-1].lower() in names:
            return key, count
    return None, None


def print_views(views):
    for view in views:
        query = summarize_query(view.get("ViewQuery"))
        flag = "*" if view.get("DefaultView") else ("H" if view.get("Hidden") else " ")
        print(
            f"   {flag} view '{view['Title']}': {len(view_fields(view))} fields,"
            f" row limit {view.get('RowLimit')}, paged {view.get('Paged')}"
        )
        print(f"       fields: {view_fields(view)}")
        print(
            f"       where on {query['Where']} | order by {query['OrderBy']}"
            f" | group by {query['GroupBy']}"
        )


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("site_url", help="full URL of the site")
    parser.add_argument("site_dir", help="extraction folder of the site on disk")
    parser.add_argument("--include-hidden", action="store_true")
    parser.add_argument(
        "--views", metavar="TITLE", help="detail every view of this list only"
    )
    args = parser.parse_args()

    user = os.environ.get(USER_ENV)
    password = os.environ.get(PASSWORD_ENV)
    ca_bundle = os.environ.get(CA_BUNDLE_ENV, "").strip()
    if not user or not password:
        sys.exit(f"Missing {USER_ENV} or {PASSWORD_ENV} in the environment")
    if not os.path.isdir(args.site_dir):
        sys.exit(f"Not a directory: {args.site_dir}")

    site_url = args.site_url.rstrip("/")
    session = requests.Session()
    session.auth = HttpNtlmAuth(user, password)
    session.verify = ca_bundle or True
    disk_items, disk_metas, disk_attachments = count_disk(args.site_dir)
    matched = set()
    sp_lists = fetch_lists(session, site_url)

    for sp_list in sorted(sp_lists, key=lambda x: x["Title"].lower()):
        if args.views and sp_list["Title"] != args.views:
            continue
        if sp_list["Hidden"] and not args.include_hidden and not args.views:
            continue
        is_library = sp_list["BaseType"] == DOC_LIBRARY_BASE_TYPE
        key, on_disk = match_disk(sp_list, disk_metas if is_library else disk_items)
        if key:
            matched.add(key)
        kind = "library" if is_library else "list"
        print(
            f"[{kind}] '{sp_list['Title']}' template {sp_list['BaseTemplate']}"
            f" | ItemCount {sp_list['ItemCount']} | on disk"
            f" {'-' if on_disk is None else on_disk} ({key or 'no match'})"
        )
        views = fetch_views(session, site_url, sp_list["Id"])
        if args.views:
            print_views(views)
        else:
            default = next((v for v in views if v.get("DefaultView")), None)
            visible = sum(1 for v in views if not v.get("Hidden"))
            print(
                f"   {visible} visible views, default"
                f" '{default['Title'] if default else '-'}'"
                f" with {len(view_fields(default)) if default else 0} fields"
            )

    if not args.views:
        unmatched = [k for k in list(disk_items) + list(disk_metas) if k not in matched]
        print(f"\nDisk entries without a matching list: {len(unmatched)}")
        for key in sorted(unmatched):
            print(f"   {key}")
        print(f"List attachment .meta.json on disk: {disk_attachments}")


if __name__ == "__main__":
    main()
