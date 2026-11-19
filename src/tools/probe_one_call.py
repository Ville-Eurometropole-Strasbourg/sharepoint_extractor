"""One read-only call (web title and URL) to check authenticated access to a host."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from sharepoint_client import build_context, get_config


def main():
    site_url, username, password, ca_bundle = get_config()
    bundle = "default (certifi)" if ca_bundle is True else ca_bundle
    print(f"target: {site_url} | bundle: {bundle}")
    ctx = build_context(site_url, username, password, ca_bundle)
    try:
        web = ctx.web.get().select(["Title", "Url"]).execute_query()
    except Exception as exc:
        status = getattr(getattr(exc, "response", None), "status_code", None)
        print(f"FAIL {type(exc).__name__} http_status={status} {str(exc)[:120]}")
        sys.exit(1)
    print(f"OK {web.url} | {web.title}")


if __name__ == "__main__":
    main()
