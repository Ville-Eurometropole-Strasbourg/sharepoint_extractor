"""
Read-only inventory of one site's wiki/publishing pages.

Usage:
    python src/extraction/inventory_site_pages.py <root_name> <site_slug>

Only the domain comes from SPX_SITE_URL; the site path comes from
site_metadata.json.

Writes under <SPX_OUTPUT_DIR>/extraction/<root>/<slug>/pages/:
    pages_inventory.json   fields, web parts, navigation, version labels
    content/<name>.html    raw page HTML (never printed)
"""

import json
import os
import sys
from urllib.parse import urlparse

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from sharepoint_client import build_context, get_config
from extract_versions import fetch_versions

# Fields holding page body HTML, depending on the page type.
CONTENT_FIELDS = ("WikiField", "PublishingPageContent", "CanvasContent1")
# Page libraries: 119 = Site Pages (wiki), 850 = Pages (publishing).
PAGE_LIBRARY_TEMPLATES = (119, 850)


def safe_name(name):
    return "".join(c if c.isalnum() or c in "-_." else "_" for c in name)


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(1)
    root_name, site_slug = sys.argv[1], sys.argv[2]

    output_dir = os.environ.get("SPX_OUTPUT_DIR", "./discovery_output")
    site_dir = os.path.join(output_dir, "extraction", root_name, site_slug)
    with open(os.path.join(site_dir, "site_metadata.json"), encoding="utf-8") as f:
        site_url = json.load(f)["Url"]

    base_url, username, password, ca_bundle = get_config()
    parsed = urlparse(base_url)
    site_url = f"{parsed.scheme}://{parsed.netloc}{urlparse(site_url).path}"
    ctx = build_context(site_url, username, password, ca_bundle)

    out_dir = os.path.join(site_dir, "pages")
    content_dir = os.path.join(out_dir, "content")
    os.makedirs(content_dir, exist_ok=True)

    inventory = {"site_url": site_url, "errors": []}

    # Match by template.
    lists = ctx.web.lists.get().execute_query()
    page_libs = [
        lst
        for lst in lists
        if lst.properties.get("BaseTemplate") in PAGE_LIBRARY_TEMPLATES
    ]
    inventory["page_libraries"] = [
        {
            "title": lst.properties.get("Title"),
            "root_folder": lst.properties.get("EntityTypeName"),
            "item_count": lst.properties.get("ItemCount"),
            "versioning": lst.properties.get("EnableVersioning"),
            "base_template": lst.properties.get("BaseTemplate"),
        }
        for lst in page_libs
    ]
    print(f"[INFO] page libraries: {len(page_libs)}")

    inventory["pages"] = []
    for lib in page_libs:
        try:
            # Every page, through server-driven paging: a single .top(N) read
            # silently stopped at N pages.
            items = (
                lib.items.select(["*", "FileLeafRef", "FileRef"])
                .get_all(page_size=100)
                .execute_query()
            )
        except Exception as e:
            inventory["errors"].append({"step": "list_items", "error": str(e)[:300]})
            continue
        for item in items:
            props = item.properties
            name = props.get("FileLeafRef") or f"item_{props.get('Id')}"
            page = {
                "library": lib.properties.get("Title"),
                "name": name,
                "id": props.get("Id"),
                "content_type": props.get("ContentTypeId"),
                "version": props.get("OData__UIVersionString"),
                "created": props.get("Created"),
                "modified": props.get("Modified"),
                "author_id": props.get("AuthorId"),
                "editor_id": props.get("EditorId"),
                "moderation": props.get("OData__ModerationStatus"),
                "fs_object_type": props.get("FileSystemObjectType"),
                "layout": props.get("PublishingPageLayout"),
                "field_names": sorted(props.keys()),
                "content_fields": {},
                "web_parts": [],
            }
            for field in CONTENT_FIELDS:
                value = props.get(field)
                if value:
                    path = os.path.join(content_dir, f"{safe_name(name)}.{field}.html")
                    with open(path, "w", encoding="utf-8") as f:
                        f.write(value)
                    page["content_fields"][field] = len(value.encode("utf-8"))

            file_url = props.get("FileRef")

            if file_url:
                try:
                    res = fetch_versions(ctx, file_url)
                    page["history_labels"] = [h["label"] for h in res["history"]]
                except Exception as e:
                    page["history_error"] = str(e)[:300]

            if not file_url:
                page["web_parts_error"] = "FileRef missing, web parts not queried"
            elif name.lower().endswith(".aspx"):
                try:
                    wpm = ctx.web.get_file_by_server_relative_url(
                        file_url
                    ).get_limited_webpart_manager(
                        1
                    )  # 1 = shared view
                    parts = (
                        wpm.web_parts.select(
                            ["Id", "WebPart/Title", "WebPart/TitleUrl"]
                        )
                        .expand(["WebPart"])
                        .get()
                        .execute_query()
                    )
                    for part in parts:
                        wp = part.properties.get("WebPart") or {}
                        page["web_parts"].append(
                            {
                                "id": part.properties.get("Id"),
                                "title": (
                                    wp.get("Title")
                                    if isinstance(wp, dict)
                                    else getattr(wp, "properties", {}).get("Title")
                                ),
                                "webpart_type": type(wp).__name__,
                                "webpart_keys": (
                                    sorted(wp.keys())
                                    if isinstance(wp, dict)
                                    else sorted(getattr(wp, "properties", {}).keys())
                                ),
                                "raw_keys": sorted(part.properties.keys()),
                            }
                        )
                except Exception as e:
                    page["web_parts_error"] = str(e)[:300]
            inventory["pages"].append(page)

    # Quick Launch (left navigation).
    try:
        nodes = ctx.web.navigation.quick_launch.get().execute_query()
        inventory["quick_launch"] = [
            {
                "title": n.properties.get("Title"),
                "url": n.properties.get("Url"),
                "id": n.properties.get("Id"),
                "is_external": n.properties.get("IsExternal"),
            }
            for n in nodes
        ]
    except Exception as e:
        inventory["errors"].append({"step": "quick_launch", "error": str(e)[:300]})

    with open(
        os.path.join(out_dir, "pages_inventory.json"), "w", encoding="utf-8"
    ) as f:
        json.dump(inventory, f, ensure_ascii=False, indent=2, default=str)

    with_content = sum(1 for p in inventory["pages"] if p["content_fields"])
    with_parts = sum(1 for p in inventory["pages"] if p["web_parts"])
    with_parts_errors = sum(1 for p in inventory["pages"] if p.get("web_parts_error"))
    with_history = sum(1 for p in inventory["pages"] if "history_labels" in p)
    with_history_errors = sum(1 for p in inventory["pages"] if p.get("history_error"))
    print(
        f"[INFO] pages: {len(inventory['pages'])}, with body: {with_content}, "
        f"with web parts: {with_parts}, web parts errors: {with_parts_errors}, "
        f"history queried: {with_history}, history errors: {with_history_errors}, "
        f"errors: {len(inventory['errors'])}"
    )
    print(f"[INFO] written to {out_dir}")


if __name__ == "__main__":
    main()
