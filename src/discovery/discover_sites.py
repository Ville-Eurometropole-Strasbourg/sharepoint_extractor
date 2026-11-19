"""
Discover every site under a SharePoint root.

1. Site collections: SiteData.asmx (VirtualServer, then each ContentDatabase)
   lists the collections of the web application. Those on the root's host and
   under the root's path are kept. If the enumeration fails (account without
   read access to the whole web application), only the root collection is
   crawled.
2. Subsites: breadth-first over _api/web/webs inside each collection.
"""

import os
import json
import logging
from collections import deque
import sys
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

from requests_ntlm import HttpNtlmAuth

from persistent_session import enable_persistent_session, get_shared_session
from sharepoint_client import get_config, build_context
from site_urls import get_root_name, root_path, server_relative_path, site_origin

WEB_SELECT_FIELDS = [
    "Id",
    "Title",
    "Url",
    "ServerRelativeUrl",
    "WebTemplate",
    "Created",
]

SOAP_NS = "http://schemas.microsoft.com/sharepoint/soap/"
SITEDATA_PATH = "/_vti_bin/SiteData.asmx"


def web_to_node(web_properties, parent_guid, depth):
    """Web properties -> index node."""
    return {
        "guid": web_properties.get("Id"),
        "title": web_properties.get("Title") or "(no title)",
        "url": web_properties.get("Url"),
        "server_relative_url": web_properties.get("ServerRelativeUrl"),
        "web_template": web_properties.get("WebTemplate"),
        "created": web_properties.get("Created"),
        "parent_guid": parent_guid,
        "depth": depth,
        "status": "ok",
    }


# --- Site collections (SiteData.asmx) ---------------------------------------


def get_content_envelope(object_type, object_id=None, last_item_id=None):
    """SOAP body of SiteData.GetContent."""
    parts = [f"<objectType>{object_type}</objectType>"]
    if object_id:
        parts.append(f"<objectId>{escape(object_id)}</objectId>")
    parts.append(
        "<retrieveChildItems>true</retrieveChildItems>"
        "<securityOnly>false</securityOnly>"
    )
    if last_item_id:
        parts.append(f"<lastItemIdOnPage>{escape(last_item_id)}</lastItemIdOnPage>")
    return (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">'
        f'<soap:Body><GetContent xmlns="{SOAP_NS}">{"".join(parts)}</GetContent>'
        "</soap:Body></soap:Envelope>"
    )


def parse_get_content(response_bytes):
    """
    SOAP response -> (result element, lastItemIdOnPage or None). The result
    is an XML document carried as text inside GetContentResult.
    """
    envelope = ET.fromstring(response_bytes)
    result = envelope.find(f".//{{{SOAP_NS}}}GetContentResult")
    if result is None or not (result.text or "").strip():
        raise ValueError("GetContentResult missing or empty")
    last = envelope.find(f".//{{{SOAP_NS}}}lastItemIdOnPage")
    last_id = (last.text or "").strip() if last is not None else ""
    return ET.fromstring(result.text), (last_id or None)


def content_database_ids(virtual_server):
    return [e.get("ID") for e in virtual_server.iter("ContentDatabase") if e.get("ID")]


def site_collection_urls(content_database):
    return [e.get("URL") for e in content_database.iter("Site") if e.get("URL")]


def collections_under_root(urls, site_url):
    """
    Server-relative paths of the collections on site_url's host and under its
    path (all of the host for a root at /), root excluded, sorted, deduplicated.
    """
    base = root_path(site_url)
    base_key = base.lower()
    kept = {}
    for url in urls:
        path = server_relative_path(url, site_url)
        if not path:
            continue
        path = path.rstrip("/") or "/"
        key = path.lower()
        if key == base_key:
            continue
        if base_key != "/" and not key.startswith(base_key + "/"):
            continue
        kept.setdefault(key, path)
    return sorted(kept.values(), key=str.lower)


def sitedata_call(site_url, auth, verify, object_type, object_id=None, last=None):
    response = get_shared_session().post(
        site_origin(site_url) + root_path(site_url).rstrip("/") + SITEDATA_PATH,
        data=get_content_envelope(object_type, object_id, last).encode("utf-8"),
        headers={
            "Content-Type": "text/xml; charset=utf-8",
            "SOAPAction": SOAP_NS + "GetContent",
            "X-FORMS_BASED_AUTH_ACCEPTED": "f",
        },
        auth=auth,
        verify=verify,
        timeout=300,
    )
    response.raise_for_status()
    return parse_get_content(response.content)


def enumerate_site_collections(site_url, username, password, verify, logger):
    """Absolute URLs of every collection of the web application (SiteData.asmx)."""
    auth = HttpNtlmAuth(username, password)
    virtual_server, _ = sitedata_call(site_url, auth, verify, "VirtualServer")
    db_ids = content_database_ids(virtual_server)
    logger.info(f"SiteData: {len(db_ids)} content database(s)")
    urls = []
    for db_id in db_ids:
        last = None
        while True:
            database, next_last = sitedata_call(
                site_url, auth, verify, "ContentDatabase", db_id, last
            )
            page = site_collection_urls(database)
            urls.extend(page)
            if not page or not next_last or next_last == last:
                break
            last = next_last
    return urls


# --- Subsites (_api/web/webs) -----------------------------------------------


def crawl_collection(
    root_name, collection_url, username, password, ca_bundle_path, logger
):
    """
    Breadth-first crawl of one collection. Returns (nodes, tree, errors);
    depth is relative to the collection (its root web is depth 0).
    """
    nodes = []
    errors = []

    try:
        ctx = build_context(collection_url, username, password, ca_bundle_path)
        top_props = ctx.web.get().execute_query().properties
    except Exception as e:
        logger.error(f"[{root_name}] Failed to resolve {collection_url}: {e}")
        errors.append({"url": collection_url, "error": str(e)})
        return nodes, None, errors

    collection_path = top_props.get("ServerRelativeUrl")
    top_node = web_to_node(top_props, parent_guid=None, depth=0)
    top_node["root_name"] = root_name
    top_node["site_collection"] = collection_path
    nodes.append(top_node)
    tree = {**top_node, "children": []}

    # (url, tree_node, depth)
    queue = deque([(collection_url, tree, 0)])

    while queue:
        current_url, current_tree_node, depth = queue.popleft()

        try:
            ctx = build_context(current_url, username, password, ca_bundle_path)
            subwebs = ctx.web.webs.select(WEB_SELECT_FIELDS).get().execute_query()
        except Exception as e:
            logger.error(f"[{root_name}] Failed to list subwebs of {current_url}: {e}")
            errors.append({"url": current_url, "error": str(e)})
            continue

        for subweb in subwebs:
            props = subweb.properties
            node = web_to_node(
                props, parent_guid=current_tree_node["guid"], depth=depth + 1
            )
            node["root_name"] = root_name
            node["site_collection"] = collection_path
            nodes.append(node)

            child_tree_node = {**node, "children": []}
            current_tree_node["children"].append(child_tree_node)

            child_url = props.get("Url")
            if child_url:
                queue.append((child_url, child_tree_node, depth + 1))
            else:
                logger.warning(
                    f"[{root_name}] Subweb {props.get('Title')} under {current_url} "
                    f"has no Url property, cannot descend further into it."
                )

    return nodes, tree, errors


def discover_root(root_name, root_url, username, password, ca_bundle_path, logger):
    """
    Returns (flat_index, tree, errors): one node per site, the root's tree
    with the other collections' trees under "site_collections", and
    {"url", "error"} per failure. flat_index is empty if the root failed.
    """
    nodes, tree, errors = crawl_collection(
        root_name, root_url, username, password, ca_bundle_path, logger
    )
    if tree is None:
        return [], {}, errors

    try:
        all_urls = enumerate_site_collections(
            root_url, username, password, ca_bundle_path, logger
        )
        collection_paths = collections_under_root(all_urls, root_url)
        logger.info(
            f"[{root_name}] SiteData: {len(all_urls)} collection(s) in the web "
            f"application, {len(collection_paths)} more under the root."
        )
    except Exception as e:
        logger.warning(
            f"[{root_name}] Site collection enumeration (SiteData.asmx) failed, "
            f"only the root collection is crawled: {e}"
        )
        errors.append({"url": root_url.rstrip("/") + SITEDATA_PATH, "error": str(e)})
        collection_paths = []

    seen = {(n["server_relative_url"] or "").lower() for n in nodes}
    flat_index = list(nodes)
    tree["site_collections"] = []
    origin = site_origin(root_url)

    for i, path in enumerate(collection_paths, start=1):
        if path.lower() in seen:
            continue
        logger.info(f"[{root_name}] Collection {i}/{len(collection_paths)}: {path}")
        c_nodes, c_tree, c_errors = crawl_collection(
            root_name, origin + path, username, password, ca_bundle_path, logger
        )
        errors.extend(c_errors)
        for node in c_nodes:
            key = (node["server_relative_url"] or "").lower()
            if key not in seen:
                seen.add(key)
                flat_index.append(node)
        if c_tree is not None:
            tree["site_collections"].append(c_tree)

    return flat_index, tree, errors


def main():
    site_url, username, password, ca_bundle_path = get_config()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
    )
    logger = logging.getLogger("discover_sites")

    root_name = get_root_name()
    enable_persistent_session()

    logger.info(f"Starting discovery for root '{root_name}' at {site_url}")
    flat_index, tree, errors = discover_root(
        root_name, site_url, username, password, ca_bundle_path, logger
    )
    if not flat_index:
        logger.error(f"[{root_name}] Root not reached, existing index left untouched.")
        sys.exit(1)

    output_dir = os.environ.get("SPX_OUTPUT_DIR", "./discovery_output")
    os.makedirs(output_dir, exist_ok=True)

    index_path = os.path.join(output_dir, f"sites_index_{root_name}.json")
    tree_path = os.path.join(output_dir, f"sites_tree_{root_name}.json")
    errors_path = os.path.join(output_dir, f"discovery_errors_{root_name}.json")

    with open(index_path, "w", encoding="utf-8") as f:
        json.dump(flat_index, f, ensure_ascii=False, indent=2, default=str)

    with open(tree_path, "w", encoding="utf-8") as f:
        json.dump(tree, f, ensure_ascii=False, indent=2, default=str)

    with open(errors_path, "w", encoding="utf-8") as f:
        json.dump(errors, f, ensure_ascii=False, indent=2, default=str)

    collections = 1 + len(tree.get("site_collections", []))
    logger.info(
        f"Discovery finished for '{root_name}': {len(flat_index)} site(s) in "
        f"{collections} collection(s), {len(errors)} error(s)."
    )
    logger.info(f"Output written to {output_dir}/")


if __name__ == "__main__":
    main()
