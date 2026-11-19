"""
Offline analysis of lists skipped as hidden, from the "skipped" entries of
one or more run_summary.json files (pooled before sampling).

Usage:
    python check_hidden_list.py path/to/run_summary.json [path2 ...] [--sample N]
"""

import argparse
import json
import random
from collections import Counter, defaultdict
from pathlib import Path


def load_sites_from_summary(summary_path: Path):
    """{"<root>/<site>": site_data} from one run_summary.json, {} on failure."""
    try:
        data = json.loads(summary_path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as e:
        print(f"[WARN] Could not read {summary_path}: {e}")
        return {}

    if not isinstance(data, dict):
        print(f"[WARN] {summary_path} is not a dict at the top level as expected.")
        return {}

    root_label = summary_path.parent.name or summary_path.stem
    tagged = {}
    for site_name, site_data in data.items():
        if isinstance(site_data, dict) and "skipped" in site_data:
            tagged[f"{root_label}/{site_name}"] = site_data
    return tagged


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "summary_paths",
        nargs="+",
        help="One or more run_summary.json paths (e.g. one per root)",
    )
    parser.add_argument(
        "--sample",
        type=int,
        default=40,
        help="Number of sites to sample across all provided files (default: 40)",
    )
    parser.add_argument("--seed", type=int, default=42, help="Random seed for sampling")
    args = parser.parse_args()

    all_sites = {}
    for p in args.summary_paths:
        path = Path(p)
        if not path.exists():
            print(f"[ERROR] {path} does not exist, skipping.")
            continue
        sites = load_sites_from_summary(path)
        print(f"[INFO] {path}: {len(sites)} site(s) with a 'skipped' key loaded.")
        all_sites.update(sites)

    print(f"\n[INFO] {len(all_sites)} site(s) total pooled across all provided files.")
    if not all_sites:
        print("[INFO] Nothing to analyze.")
        return

    random.seed(args.seed)
    sample_size = min(args.sample, len(all_sites))
    sampled_keys = random.sample(list(all_sites.keys()), sample_size)
    print(f"[INFO] Sampling {sample_size} site(s) out of {len(all_sites)}.\n")

    title_counter = Counter()
    title_to_sites = defaultdict(set)
    title_to_itemcounts = defaultdict(list)
    sites_with_no_hidden = []

    for site_key in sampled_keys:
        site_data = all_sites[site_key]
        skipped = site_data.get("skipped", [])
        hidden_entries = [e for e in skipped if e.get("reason") == "hidden"]
        if not hidden_entries:
            sites_with_no_hidden.append(site_key)
            continue
        for entry in hidden_entries:
            title = entry.get("title") or "(no title)"
            title_counter[title] += 1
            title_to_sites[title].add(site_key)
            ic = entry.get("item_count")
            if ic is not None:
                title_to_itemcounts[title].append(ic)

    print("=== Aggregated hidden-list titles across the sample ===\n")
    print(f"{'Title':<45} {'# sites':<10} {'Occurrences':<14} {'Item counts (sample)'}")
    print("-" * 110)
    for title, count in title_counter.most_common():
        n_sites = len(title_to_sites[title])
        counts = title_to_itemcounts[title]
        nonzero = [c for c in counts if c and c > 0]
        counts_preview = counts[:6]
        flag = "  <-- non-zero item_count found, worth a manual look" if nonzero else ""
        print(f"{title[:44]:<45} {n_sites:<10} {count:<14} {counts_preview}{flag}")

    print(
        f"\n[INFO] {len(sites_with_no_hidden)} sampled site(s) had zero hidden-skip "
        f"entries (i.e. nothing was hidden there, or everything hidden was already "
        f"in KEEP_DESPITE_HIDDEN_TEMPLATES): {sites_with_no_hidden[:10]}"
        + (" (truncated)" if len(sites_with_no_hidden) > 10 else "")
    )

    print(
        "\n[NOTE] A non-zero item_count is a signal worth checking manually "
        "(e.g. a quick live query on one such list), not proof of real "
        "content by itself — could still be plumbing that happens to hold a "
        "few rows (e.g. a per-user settings row). Titles with item_count "
        "always at/near 0 across the sample support the 'plumbing' "
        "assumption; titles with meaningfully non-zero counts across "
        "multiple sites are the ones that actually deserve a look."
    )


if __name__ == "__main__":
    main()
