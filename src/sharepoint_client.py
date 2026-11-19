"""Shared SharePoint connection helpers: environment config and authenticated ClientContext."""

import os
import sys

from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential


def tls_verify():
    """
    requests' `verify`: SPX_CA_BUNDLE if set, else True (certifi's CA
    list). Exits if the bundle is set but missing.
    """
    ca_bundle_path = os.environ.get("SPX_CA_BUNDLE", "").strip()
    if not ca_bundle_path:
        return True
    if not os.path.isfile(ca_bundle_path):
        print(f"[ERROR] SPX_CA_BUNDLE not found: {ca_bundle_path}")
        sys.exit(1)
    return ca_bundle_path


def get_config():
    """
    Read SharePoint connection config from environment variables. The last
    value is requests' `verify` (see tls_verify()).
    """
    site_url = os.environ.get("SPX_SITE_URL")
    username = os.environ.get("SPX_USERNAME")
    password = os.environ.get("SPX_PASSWORD")

    missing = [
        name
        for name, value in [
            ("SPX_SITE_URL", site_url),
            ("SPX_USERNAME", username),
            ("SPX_PASSWORD", password),
        ]
        if not value
    ]
    if missing:
        print(f"[ERROR] Missing environment variables: {', '.join(missing)}")
        sys.exit(1)

    return site_url, username, password, tls_verify()


def build_context(site_url, username, password, ca_bundle_path):
    """
    Authenticated ClientContext for site_url. The header forces NTLM on
    mixed FBA/NTLM zones. ca_bundle_path is requests' `verify` (a CA bundle
    path, or True).
    """
    ctx = ClientContext(site_url, allow_ntlm=True).with_credentials(
        UserCredential(username, password)
    )

    def _inject_fba_header(request):
        request.headers["X-FORMS_BASED_AUTH_ACCEPTED"] = "f"
        request.verify = ca_bundle_path

    ctx.pending_request().beforeExecute += _inject_fba_header
    return ctx
