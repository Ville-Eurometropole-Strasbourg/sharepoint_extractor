"""Root name (SPX_ROOT_NAME) and absolute site URLs."""

import os
import re
import sys
from urllib.parse import urlparse

ROOT_NAME_ENV = "SPX_ROOT_NAME"
ROOT_NAME_PATTERN = re.compile(r"^[A-Za-z0-9_-]+$")


def get_root_name():
    """SPX_ROOT_NAME, used as a folder and file name. Exits if missing or invalid."""
    name = os.environ.get(ROOT_NAME_ENV, "").strip()
    if not name:
        print(f"[ERROR] Missing environment variable: {ROOT_NAME_ENV}")
        sys.exit(1)
    if not ROOT_NAME_PATTERN.match(name):
        print(
            f"[ERROR] {ROOT_NAME_ENV}='{name}': only letters, digits, '-' and '_' "
            "are allowed."
        )
        sys.exit(1)
    return name


def site_origin(site_url):
    """https://host/sites/hr -> https://host."""
    parsed = urlparse(site_url)
    return f"{parsed.scheme}://{parsed.netloc}"


def root_path(site_url):
    """https://host/sites/hr/ -> /sites/hr; https://host -> /."""
    return urlparse(site_url).path.rstrip("/") or "/"


def absolute_url(site_url, server_relative_url):
    """(https://host/sites/hr, /sites/hr/team) -> https://host/sites/hr/team."""
    path = server_relative_url or "/"
    if not path.startswith("/"):
        path = "/" + path
    return site_origin(site_url) + path


def server_relative_path(url, site_url):
    """Path of an absolute url on the same host as site_url, else None."""
    parsed = urlparse(url)
    if parsed.netloc.lower() != urlparse(site_url).netloc.lower():
        return None
    return parsed.path or None
