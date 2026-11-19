"""
List people profiles and their personal sites through the Search REST API,
with raw requests + NTLM. SPX_SITE_URL is the personal sites (My Site) host.

Usage:
    python discover_personal_sites.py [--rowlimit N]
"""

import os
import sys
import argparse
import xml.etree.ElementTree as ET

import requests
from requests_ntlm import HttpNtlmAuth

from sharepoint_client import get_config

# OData/AtomPub namespaces: "d" (data) and "m" (metadata).
NS = {
    "d": "http://schemas.microsoft.com/ado/2007/08/dataservices",
    "m": "http://schemas.microsoft.com/ado/2007/08/dataservices/metadata",
}


def build_search_url(base_url, querytext="*", rowlimit=10):
    """_api/search/query URL ('*' matches everything)."""
    return (
        f"{base_url.rstrip('/')}/_api/search/query"
        f"?querytext='{querytext}'&rowlimit={rowlimit}"
    )


def fetch_search_results(base_url, username, password, ca_bundle_path, rowlimit):
    """Raw NTLM GET against the search API, same header and CA bundle as build_context."""
    url = build_search_url(base_url, rowlimit=rowlimit)
    headers = {
        "Accept": "application/atom+xml",
        "X-FORMS_BASED_AUTH_ACCEPTED": "f",
    }
    response = requests.get(
        url,
        auth=HttpNtlmAuth(username, password),
        headers=headers,
        verify=ca_bundle_path,
        timeout=30,
    )
    response.raise_for_status()
    return response.text


def parse_people_results(xml_text):
    """(WorkId, PreferredName, OriginalPath) per search row, None if absent."""
    root = ET.fromstring(xml_text)
    results = []
    current = {}
    for elem in root.iter("{%s}element" % NS["d"]):
        key_elem = elem.find("d:Key", NS)
        value_elem = elem.find("d:Value", NS)
        if key_elem is None or value_elem is None:
            continue
        key = key_elem.text
        value = (
            None if value_elem.get("{%s}null" % NS["m"]) == "true" else value_elem.text
        )

        if key == "WorkId":
            if current:
                results.append(current)
            current = {"WorkId": value}
        elif key in ("PreferredName", "OriginalPath"):
            current[key] = value

    if current:
        results.append(current)
    return results


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--rowlimit",
        type=int,
        default=10,
        help="Number of results to request (default: 10, kept small for this first exploratory pass).",
    )
    args = parser.parse_args()

    site_url, username, password, ca_bundle_path = get_config()

    print(f"[INFO] Querying search API at {site_url} (rowlimit={args.rowlimit}).")
    try:
        xml_text = fetch_search_results(
            site_url, username, password, ca_bundle_path, args.rowlimit
        )
    except requests.exceptions.RequestException as e:
        print(f"[ERROR] Request failed: {e}")
        sys.exit(1)

    results = parse_people_results(xml_text)
    print(f"[INFO] Parsed {len(results)} result(s).\n")
    for r in results:
        print(f"  WorkId: {r.get('WorkId')}")
        print(f"  PreferredName: {r.get('PreferredName')}")
        print(f"  OriginalPath: {r.get('OriginalPath')}")
        print()

    if not results:
        print(
            "[WARNING] No results parsed -- either the XML shape differs from "
            "what's expected here, or the query genuinely returned nothing "
            "(e.g. no indexed profiles on this farm)."
        )


if __name__ == "__main__":
    main()
