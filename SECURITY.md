# Security Policy

## Supported versions

Security fixes are applied to the latest commit on `main`.

| Version | Supported |
|---|---|
| latest `main` | yes |
| older commits / forks | no |

## Reporting a vulnerability

**Please do not open a public issue for security problems.**

Use one of these channels:

- GitHub's private vulnerability reporting: on the repository page, go to **Security > Advisories > Report a vulnerability** (preferred, if the project is hosted on GitHub);
- or email: `SECURITY-CONTACT@example.org` — **replace this address with the organization's security contact before publishing**.

A good report includes:

- a description of the issue and its security impact;
- the affected file(s) and, if relevant, the commit;
- steps to reproduce, or a proof of concept;
- whether you believe data (credentials, tokens, site content) can be exposed, leaked or logged.

**Never include real credentials, tokens, cookies, private keys, internal hostnames or extracted site content in a report.** Redact them first.

You can expect:

- an acknowledgement within **5 business days**;
- a status update (confirmed / not confirmed, plan) within **30 days**;
- coordinated disclosure: we will agree on a publication date with you, typically after a fix is available, within **90 days** of the report.

## Scope

In scope:

- the code in this repository and the Docker image built from it;
- handling of credentials and the CA bundle (`SPX_USERNAME`, `SPX_PASSWORD`, `SPX_CA_BUNDLE`, `.env`, `certs/`);
- TLS verification in the extraction path;
- accidental logging or committing of secrets and of extracted data (`.gitignore` / `.dockerignore` coverage);
- dependencies listed in `requirements.txt` / `requirements-dev.txt` with known CVEs.

Out of scope:

- the SharePoint farm being extracted and its own configuration;
- the network infrastructure the tool runs on (Docker subnets, proxies, Grafana instance);
- denial of service against the project's infrastructure;
- social engineering.

## Safe harbor

We support good-faith security research. If you follow this policy — test only on systems and data you own or are authorized to test, do not exfiltrate or expose data, do not disrupt the service, and report privately — we will not pursue legal action against you, and we will credit you in the fix announcement if you wish.

## Specific note for this project

This tool authenticates with NTLM credentials and extracts full site contents (files, versions, permissions, user tables). Reports about:

- secrets ending up in logs, console output, `run_summary.json` or any output file;
- TLS verification being bypassable in the extraction path;
- output or certificate files being committable or baked into the Docker image;

are especially welcome.
