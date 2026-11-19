"""Diagnostic: list the lists/libraries and subwebs under a subsite path."""

import os

from sharepoint_client import get_config, build_context


def main():
    site_url, username, password, ca_bundle_path = get_config()

    # Relative to the root, e.g. "" (root itself) or "/team".
    subpath = os.environ.get("SPX_INSPECT_PATH", "")
    target_url = site_url.rstrip("/") + subpath

    print(f"[INFO] Inspecting lists/libraries under: {target_url}\n")

    ctx = build_context(target_url, username, password, ca_bundle_path)

    try:
        lists = (
            ctx.web.lists.select(["Title", "BaseTemplate", "ItemCount", "Hidden", "Id"])
            .get()
            .execute_query()
        )
    except Exception as e:
        print(f"[ERROR] Failed to read lists of {target_url}: {e}")
        lists = []

    if lists:
        visible_lists = [l for l in lists if not l.properties.get("Hidden")]
        print(
            f"[OK] Found {len(lists)} list(s)/librarie(s) total, "
            f"{len(visible_lists)} visible (non-hidden):\n"
        )
        for lst in visible_lists:
            props = lst.properties
            print(
                f"  - {props.get('Title')!r:40s} "
                f"template={props.get('BaseTemplate')!s:5s} "
                f"items={props.get('ItemCount')}"
            )

    try:
        subwebs = (
            ctx.web.webs.select(["Title", "Url", "ServerRelativeUrl", "WebTemplate"])
            .get()
            .execute_query()
        )
        print(f"\n[INFO] {subpath or '(root)'} has {len(subwebs)} direct subweb(s):")
        for w in list(subwebs)[:20]:
            props = w.properties
            print(
                f"  - {props.get('Title')!s:40s} "
                f"template={props.get('WebTemplate')!s:15s} "
                f"-> {props.get('Url')}"
            )
        if len(subwebs) > 20:
            print(f"  ... and {len(subwebs) - 20} more")
    except Exception as e:
        print(f"\n[WARNING] Could not check subwebs of {subpath or '(root)'}: {e}")

    list_title = os.environ.get("SPX_INSPECT_LIST")
    if list_title:
        print(f"\n[INFO] Inspecting list '{list_title}' schema and sample items...")
        try:
            fields = (
                ctx.web.lists.get_by_title(list_title)
                .fields.select(["Title", "InternalName", "TypeAsString", "Hidden"])
                .get()
                .execute_query()
            )
            visible_fields = [f for f in fields if not f.properties.get("Hidden")]
            print(f"\n[OK] {len(visible_fields)} visible field(s):")
            for f in visible_fields:
                props = f.properties
                print(
                    f"  - {props.get('InternalName')!s:30s} "
                    f"type={props.get('TypeAsString')!s:15s} "
                    f"label={props.get('Title')}"
                )
        except Exception as e:
            print(f"[ERROR] Failed to read fields of '{list_title}': {e}")

        try:
            items = (
                ctx.web.lists.get_by_title(list_title)
                .items.get_all(page_size=5)
                .execute_query()
            )
            print(f"\n[OK] First {min(5, len(items))} raw item(s) (all properties):")
            for item in list(items)[:5]:
                print(f"  {item.properties}\n")
        except Exception as e:
            print(f"[ERROR] Failed to read sample items of '{list_title}': {e}")


if __name__ == "__main__":
    main()
