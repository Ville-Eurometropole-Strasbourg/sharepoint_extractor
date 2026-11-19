"""Pre-flight TLS check, to run once after installing the repository.

Verifies that the certificate of each SharePoint host is accepted, with
SPX_CA_BUNDLE if set, else certifi's CA list, and says why when it is
not: missing CA in the bundle, or a network device presenting its own
certificate (TLS inspection). Only certificate names are printed, never
secrets or proxy values.

Usage: python3 src/tools/check_tls.py [host ...]   (default: SPX_SITE_URL's host)
"""

import os
import socket
import ssl
import sys
import warnings

from urllib.parse import urlparse

import certifi
import requests

BUNDLE_ENV = "SPX_CA_BUNDLE"
PROXY_VARS = ("https_proxy", "http_proxy", "no_proxy")
TIMEOUT_S = 10


def bundle_subjects(path):
    """Return the subjects of the certificates in the bundle, or None."""
    try:
        from cryptography import x509
    except ImportError:
        return None
    with open(path, "rb") as handle, warnings.catch_warnings():
        warnings.simplefilter("ignore")
        certs = x509.load_pem_x509_certificates(handle.read())
    return {cert.subject.rfc4514_string() for cert in certs}


def presented_issuer(host):
    """Return the issuer of the certificate actually received on this network."""
    context = ssl.create_default_context()
    context.check_hostname = False
    context.verify_mode = ssl.CERT_NONE
    with socket.create_connection((host, 443), timeout=TIMEOUT_S) as sock:
        with context.wrap_socket(sock, server_hostname=host) as tls:
            der = tls.getpeercert(binary_form=True)
    try:
        from cryptography import x509
    except ImportError:
        return None
    return x509.load_der_x509_certificate(der).issuer.rfc4514_string()


def check_direct(host, bundle):
    """Raise if a direct TLS handshake fails with the bundle."""
    context = ssl.create_default_context(cafile=bundle)
    with socket.create_connection((host, 443), timeout=TIMEOUT_S) as sock:
        with context.wrap_socket(sock, server_hostname=host):
            pass


def explain_failure(host, subjects):
    """Print what was presented and whether the bundle knows its issuer."""
    try:
        issuer = presented_issuer(host)
    except OSError as exc:
        print(f"     issuer presented: unavailable ({exc.__class__.__name__})")
        return
    if issuer is None:
        print("     issuer presented: unknown (cryptography not installed)")
        return
    print(f"     issuer presented: {issuer}")
    if subjects is None:
        return
    if issuer in subjects:
        print("     issuer is in the bundle: chain problem on the server side")
    else:
        print(
            "     issuer is NOT in the bundle: TLS inspection on this network, "
            "or a CA missing from the bundle (compare with a working machine)"
        )


def check_host(host, bundle, subjects):
    """Check one host directly and through requests. Return True if usable."""
    direct = "ok"
    try:
        check_direct(host, bundle)
        print(f"INFO {host}: direct TLS accepted")
    except ssl.SSLCertVerificationError as exc:
        direct = "rejected"
        print(f"INFO {host}: direct TLS rejected ({exc.verify_message})")
        explain_failure(host, subjects)
    except OSError as exc:
        direct = "unreachable"
        print(f"INFO {host}: no direct route ({exc.__class__.__name__})")

    try:
        requests.head(f"https://{host}/", verify=bundle, timeout=TIMEOUT_S)
    except requests.exceptions.SSLError:
        if direct == "ok":
            print(f"FAIL {host}: rejected through the proxy only (TLS inspection?)")
        else:
            print(f"FAIL {host}: certificate rejected")
        return False
    except requests.RequestException as exc:
        print(f"FAIL {host}: unreachable ({exc.__class__.__name__})")
        return False
    print(f"OK   {host}: certificate accepted")
    return True


def main():
    default_host = urlparse(os.environ.get("SPX_SITE_URL", "")).hostname
    hosts = sys.argv[1:] or ([default_host] if default_host else [])
    if not hosts:
        print("FAIL no host given and SPX_SITE_URL is unset")
        return 1
    bundle = os.environ.get(BUNDLE_ENV, "").strip()
    if not bundle:
        bundle = certifi.where()
        print(f"INFO {BUNDLE_ENV} unset: using certifi's CA list")
    elif not os.path.isfile(bundle):
        print(f"FAIL bundle not found: {bundle}")
        return 1

    subjects = bundle_subjects(bundle)
    count = "?" if subjects is None else len(subjects)
    print(f"OK   bundle {bundle} ({count} certificates)")

    proxies = sorted(k for k in os.environ if k.lower() in PROXY_VARS)
    print(f"INFO proxy variables set: {', '.join(proxies) or 'none'}")

    results = [check_host(host, bundle, subjects) for host in hosts]
    return 0 if all(results) else 1


if __name__ == "__main__":
    sys.exit(main())
