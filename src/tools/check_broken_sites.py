"""
Unauthenticated HTTP check of the errored sites in run_summary.json.
Results are compared with the optional known-good sites (server-relative
paths on the SPX_SITE_URL host).

Usage:
    python3 check_broken_sites.py /path/to/run_summary.json [good_path ...]

Writes check_broken_sites_report.csv in the current directory.
"""

import sys
import csv
import json
import os
import time
import requests
from urllib.parse import urlparse

TLS_VERIFY = os.environ.get("SPX_CA_BUNDLE", "").strip() or True
REQUEST_DELAY_SECONDS = 0.3
TIMEOUT_SECONDS = 15


def load_errored_sites(run_summary_path):
    """{slug: site_root_url} of errored sites, URL parsed from the error message."""
    with open(run_summary_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    errored = {}
    for slug, content in data.items():
        if not isinstance(content, dict) or "error" not in content:
            continue
        err = content["error"]
        msg = err[2] if isinstance(err, list) and len(err) > 2 else str(err)
        marker = "for url: "
        idx = msg.find(marker)
        if idx == -1:
            print(
                f"[WARN] Could not extract URL from error for '{slug}', skipping: {msg[:120]}"
            )
            continue
        full_url = msg[idx + len(marker) :].strip()
        site_root = full_url.rsplit("/_api/Web", 1)[0]
        errored[slug] = site_root

    return errored


def probe(url):
    """Unauthenticated GET without following redirects."""
    try:
        resp = requests.get(
            url, timeout=TIMEOUT_SECONDS, allow_redirects=False, verify=TLS_VERIFY
        )
        return {
            "status_code": resp.status_code,
            "body_length": len(resp.content),
            "content_type": resp.headers.get("Content-Type", ""),
            "location": resp.headers.get("Location", ""),
            "www_authenticate": resp.headers.get("WWW-Authenticate", ""),
            "request_error": "",
        }
    except requests.exceptions.SSLError as e:
        return {"request_error": f"SSLError (certificate rejected): {e}"}
    except requests.exceptions.RequestException as e:
        return {"request_error": f"{type(e).__name__}: {e}"}


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    run_summary_path = sys.argv[1]
    good_paths = sys.argv[2:]
    site_url = os.environ.get("SPX_SITE_URL", "")
    if good_paths and not site_url:
        print("[ERROR] SPX_SITE_URL is required to probe known-good sites.")
        sys.exit(1)
    parsed = urlparse(site_url)
    origin = f"{parsed.scheme}://{parsed.netloc}"
    errored = load_errored_sites(run_summary_path)
    print(f"[INFO] {len(errored)} errored site(s) extracted from {run_summary_path}")

    baseline = {}
    if good_paths:
        print("[INFO] Probing known-good reference sites first...")
    for path in good_paths:
        url = origin + "/" + path.lstrip("/")
        baseline[path] = probe(url)
        print(f"  {path}: {baseline[path]}")
        time.sleep(REQUEST_DELAY_SECONDS)

    baseline_all_empty = bool(baseline) and all(
        b.get("body_length", 0) == 0 and not b.get("request_error")
        for b in baseline.values()
    )
    if baseline_all_empty:
        print(
            "\n[WARNING] The known-good reference sites ALSO returned an empty "
            "body without authentication. This means anonymous requests are "
            "uniformly blocked/redirected here, and body length alone will "
            "NOT distinguish a broken site from a healthy one in this run. "
            "The report below is still worth reading for status codes, "
            "redirect targets (Location header), and WWW-Authenticate "
            "differences, but do not trust 'empty body' as a verdict by "
            "itself -- cross-check a few of the most surprising cases by "
            "hand (browser or the real authenticated script) before "
            "concluding anything.\n"
        )

    rows = []
    print(f"[INFO] Probing {len(errored)} errored site(s)...")
    for i, (slug, url) in enumerate(errored.items(), 1):
        result = probe(url)
        result["slug"] = slug
        result["url"] = url
        rows.append(result)
        if i % 10 == 0 or i == len(errored):
            print(f"  ...{i}/{len(errored)} done")
        time.sleep(REQUEST_DELAY_SECONDS)

    out_path = "check_broken_sites_report.csv"
    fieldnames = [
        "slug",
        "url",
        "status_code",
        "body_length",
        "content_type",
        "location",
        "www_authenticate",
        "request_error",
    ]
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow({k: row.get(k, "") for k in fieldnames})
    shapes = {}
    for row in rows:
        key = (
            row.get("status_code"),
            row.get("body_length"),
            bool(row.get("request_error")),
        )
        shapes.setdefault(key, []).append(row["slug"])

    print(f"\n[INFO] Report written to {out_path}")
    print("\n[SUMMARY] Grouped by (status_code, body_length, had_request_error):")
    for (status, length, had_err), slugs in sorted(
        shapes.items(), key=lambda x: -len(x[1])
    ):
        print(
            f"  status={status} body_length={length} request_error={had_err}: {len(slugs)} site(s)"
        )
        print(f"    e.g. {slugs[:5]}")


if __name__ == "__main__":
    main()
