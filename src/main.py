"""
Interactive menu (Docker default entry point): pick a step, then its scope.
Each step runs its standalone script as a subprocess.
"""

import os
import subprocess
import sys

from site_urls import ROOT_NAME_ENV

SRC_DIR = os.path.dirname(os.path.abspath(__file__))
DISCOVERY_DIR = os.path.join(SRC_DIR, "discovery")
EXTRACTION_DIR = os.path.join(SRC_DIR, "extraction")


def ask_choice(prompt, options):
    """Numbered menu over (key, label) tuples; returns the chosen key."""
    print(f"\n{prompt}")
    for i, (_, label) in enumerate(options, start=1):
        print(f"  {i}. {label}")
    while True:
        raw = input("> ").strip()
        if raw.isdigit() and 1 <= int(raw) <= len(options):
            return options[int(raw) - 1][0]
        print(f"Please enter a number between 1 and {len(options)}.")


def ask_site_path():
    """Prompt for a server-relative site path, e.g. /sites/hr/team."""
    raw = input("Site path (e.g. /sites/hr/team): ").strip()
    if not raw.startswith("/"):
        raw = "/" + raw
    return raw


def run_script(path, args=None, env_overrides=None):
    """Run a script, streaming its output. env_overrides is merged over os.environ."""
    cmd = [sys.executable, "-u", path] + (args or [])
    env = {**os.environ, **(env_overrides or {})}
    print(f"\n--- Running {os.path.basename(path)} {' '.join(args or [])} ---\n")
    result = subprocess.run(cmd, env=env)
    if result.returncode != 0:
        print(
            f"\n[WARNING] {os.path.basename(path)} exited with an error (code {result.returncode})."
        )
    return result.returncode == 0


def sites_index_path():
    return os.path.join(
        os.environ.get("SPX_OUTPUT_DIR", "./discovery_output"),
        f"sites_index_{os.environ.get(ROOT_NAME_ENV, '')}.json",
    )


def handle_discovery():
    script = ask_choice(
        "Discovery: which script?",
        [
            ("bfs", "Recursive crawl of the root's sites"),
            ("personal", "Personal sites, estimate (root = My Site host)"),
            ("inspect", "Inspect a path (SPX_INSPECT_PATH)"),
            ("cancel", "Cancel"),
        ],
    )
    if script == "cancel":
        return
    script_map = {
        "bfs": os.path.join(DISCOVERY_DIR, "discover_sites.py"),
        "personal": os.path.join(DISCOVERY_DIR, "discover_personal_sites.py"),
        "inspect": os.path.join(DISCOVERY_DIR, "inspect_path.py"),
    }
    run_script(script_map[script])


def handle_metadata_extraction():
    scope = ask_choice(
        "Metadata extraction: which scope?",
        [
            ("site", "One site"),
            (
                "all",
                "All sites of the root (after Discovery > Recursive crawl)",
            ),
            ("cancel", "Cancel"),
        ],
    )
    if scope == "cancel":
        return

    script = os.path.join(EXTRACTION_DIR, "extract_metadata.py")
    if scope == "site":
        site_path = ask_site_path()
        run_script(script, [site_path])
        return

    index_path = sites_index_path()
    if not os.path.exists(index_path):
        print(f"[ERROR] '{index_path}' not found. Run the site discovery first.")
        return
    root_url = os.environ.get("SPX_SITE_URL", "")
    confirm = (
        input(
            f"\nThis will process every site listed in '{index_path}' "
            f"on {root_url}. It can take a long time. Continue? [y/N] "
        )
        .strip()
        .lower()
    )
    if confirm == "y":
        run_script(script, ["--all-from-discovery", index_path])
    else:
        print("Cancelled.")


def handle_versions_extraction():
    root_url = os.environ.get("SPX_SITE_URL", "")
    root_name = os.environ.get(ROOT_NAME_ENV, "")
    scope = ask_choice(
        f"Version history (active root: {root_url}): which scope?",
        [
            ("site", "One site (already extracted)"),
            ("all", "All sites already extracted from this root"),
            ("cancel", "Cancel"),
        ],
    )
    if scope == "cancel":
        return

    script = os.path.join(EXTRACTION_DIR, "extract_versions.py")
    if scope == "site":
        site_slug = input(
            f"Site id (folder under discovery_output/extraction/{root_name}/): "
        ).strip()
        run_script(script, [site_slug])
    elif scope == "all":
        confirm = (
            input(
                f"\nThis will query the history of every versioned file of "
                f"every site already extracted from {root_name}. It can take a "
                f"long time (resumable: files already processed are skipped). "
                f"Continue? [y/N] "
            )
            .strip()
            .lower()
        )
        if confirm == "y":
            run_script(script, ["--all"])
        else:
            print("Cancelled.")


def handle_list_attachments():
    root_url = os.environ.get("SPX_SITE_URL", "")
    root_name = os.environ.get(ROOT_NAME_ENV, "")
    scope = ask_choice(
        f"List attachments (active root: {root_url}): which scope?",
        [
            ("site", "One site (already extracted)"),
            ("all", "All sites already extracted from this root"),
            ("cancel", "Cancel"),
        ],
    )
    if scope == "cancel":
        return
    script = os.path.join(EXTRACTION_DIR, "extract_list_attachments.py")
    if scope == "site":
        site_slug = input(
            f"Site id (folder under discovery_output/extraction/{root_name}/): "
        ).strip()
        run_script(script, [site_slug])
    else:
        confirm = (
            input(
                f"\nThis will go through the lists of every site already extracted "
                f"from {root_name} (resumable). Continue? [y/N] "
            )
            .strip()
            .lower()
        )
        if confirm == "y":
            run_script(script, ["--all"])
        else:
            print("Cancelled.")


def handle_list_versions():
    root_url = os.environ.get("SPX_SITE_URL", "")
    root_name = os.environ.get(ROOT_NAME_ENV, "")
    scope = ask_choice(
        f"List item history (active root: {root_url}): which scope?",
        [
            ("site", "One site (already extracted)"),
            ("all", "All sites already extracted from this root"),
            ("count", "All sites, count the items to query only (--dry-run)"),
            ("cancel", "Cancel"),
        ],
    )
    if scope == "cancel":
        return
    script = os.path.join(EXTRACTION_DIR, "extract_list_versions.py")
    if scope == "site":
        site_slug = input(
            f"Site id (folder under discovery_output/extraction/{root_name}/): "
        ).strip()
        run_script(script, [site_slug])
    elif scope == "count":
        run_script(script, ["--all", "--dry-run"])
    else:
        confirm = (
            input(
                f"\nThis will query the history of every list item with past "
                f"versions on every site already extracted from {root_name}: "
                f"one call per item (resumable). Continue? [y/N] "
            )
            .strip()
            .lower()
        )
        if confirm == "y":
            run_script(script, ["--all"])
        else:
            print("Cancelled.")


def handle_download():
    scope = ask_choice(
        "File download: which scope?",
        [
            ("library", "One library of one site"),
            ("site", "Every document of one site (all its libraries)"),
            ("all", "All sites already extracted from this root"),
            ("cancel", "Cancel"),
        ],
    )
    if scope == "cancel":
        return

    script = os.path.join(EXTRACTION_DIR, "download_files.py")
    if scope == "library":
        site_path = ask_site_path()
        library_title = input("Library title (e.g. Documents): ").strip()
        run_script(script, [site_path, library_title])
    elif scope == "site":
        root_name = os.environ.get(ROOT_NAME_ENV, "")
        site_slug = input(
            f"Site id (folder under discovery_output/extraction/{root_name}/): "
        ).strip()
        raw_limit = input(
            "Maximum number of files to download for this site "
            "(empty for no limit): "
        ).strip()
        cmd_args = ["--site", site_slug]
        if raw_limit:
            if not raw_limit.isdigit() or int(raw_limit) < 1:
                print("[ERROR] The limit must be an integer of at least 1.")
                return
            cmd_args += ["--limit", raw_limit]
        run_script(script, cmd_args)
    elif scope == "all":
        raw_limit = input(
            "Maximum total number of files to download for this root "
            "(empty for no limit): "
        ).strip()

        if raw_limit == "":
            print(
                "\n[WARNING] No limit: this will try to download EVERY file "
                "already discovered on the active root, with no resume logic "
                "at that scale if interrupted."
            )
            confirm = input("Start without a limit? [y/N] ").strip().lower()
            if confirm != "y":
                print("Cancelled.")
                return
            run_script(script, ["--all"])
            return

        if not raw_limit.isdigit() or int(raw_limit) < 1:
            print("[ERROR] The limit must be an integer of at least 1.")
            return
        run_script(script, ["--all", raw_limit])


def handle_stats():
    repo_dir = os.path.dirname(SRC_DIR)
    stats_path = os.path.join(repo_dir, "stats_summary.json")
    print(
        "\n[INFO] Read-only pass over the .meta.json / .versions.json already "
        "on disk, no SharePoint call."
    )
    run_script(
        os.path.join(EXTRACTION_DIR, "aggregate_stats.py"), ["--output", stats_path]
    )


def handle_permissions():
    print(
        "\n[INFO] This mode processes every site already extracted without "
        "error (from the latest run_summary.json), without re-extracting "
        "metadata."
    )
    script = os.path.join(EXTRACTION_DIR, "extract_metadata.py")
    run_script(script, ["--permissions-all"])


def handle_consolidate_errors():
    print(
        "\n[INFO] This mode aggregates the errors of every run_summary.json "
        "already on disk (read-only, no network access)."
    )
    script = os.path.join(EXTRACTION_DIR, "consolidate_errors.py")
    run_script(script, [])


def main():
    print("=== SharePoint Extraction ===")
    print(
        f"Active root: {os.environ.get(ROOT_NAME_ENV, '')} "
        f"({os.environ.get('SPX_SITE_URL', '')})"
    )
    while True:
        step = ask_choice(
            "What do you want to do?",
            [
                ("discovery", "Site discovery"),
                ("metadata", "Metadata and permissions extraction"),
                (
                    "attachments",
                    "List attachments extraction (sites already extracted)",
                ),
                (
                    "versions",
                    "Version history extraction (sites already extracted)",
                ),
                (
                    "list_versions",
                    "List item history extraction (sites already extracted)",
                ),
                ("errors", "Consolidate errors (every run_summary.json)"),
                ("download", "File download"),
                ("stats", "Statistics (stats_summary.json)"),
                (
                    "permissions",
                    "Permissions catch-up (sites already extracted)",
                ),
                ("quit", "Quit"),
            ],
        )

        if step == "quit":
            print("Bye.")
            break
        elif step == "discovery":
            handle_discovery()
        elif step == "metadata":
            handle_metadata_extraction()
        elif step == "versions":
            handle_versions_extraction()
        elif step == "list_versions":
            handle_list_versions()
        elif step == "attachments":
            handle_list_attachments()
        elif step == "permissions":
            handle_permissions()
        elif step == "errors":
            handle_consolidate_errors()
        elif step == "stats":
            handle_stats()
        elif step == "download":
            handle_download()


if __name__ == "__main__":
    main()
