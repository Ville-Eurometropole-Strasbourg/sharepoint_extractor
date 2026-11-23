# SharePoint On-Premises Extractor

Raw, exhaustive extraction of an on-premises SharePoint farm (built and used on SharePoint Server 2019), to prepare a migration to another platform.

It extracts, site by site:

- document libraries: one `.meta.json` per file, then the binaries;
- the version history of each file, with the binaries of past versions;
- generic lists: items, schema, and item attachments;
- permissions of each site, and of lists with broken inheritance;
- wiki and publishing pages (inventory, body, page history);
- the user table of each site collection.

Nothing is transformed: files (OneNote notebooks included) are copied as they are, and every library or list schema is read from the server rather than assumed.

## How it works

- Python, run in Docker.
- NTLM authentication through `office365-rest-python-client` 2.6.2 and `requests_ntlm`. Version 3.0.0 of the client no longer supports NTLM.
- Classic SharePoint REST API (`_api/...`), not Microsoft Graph.
- All calls go through one persistent HTTP session with adaptive throttling: a base delay of 0.3 s between calls, doubled when the server is under strain (up to 5 s), then eased back.
- Every step can resume. Lists and libraries unchanged since the last success are skipped (`LastItemModifiedDate`), file versions are resumed per file. A single-site run is merged into the root's `run_summary.json`; `--all-from-discovery` rewrites it at the end of the run.
- Downloads use several threads. Each file is written to `<file>.spx-part`, then renamed, so a file under its final name is always complete.
- Every HTTP call is logged to SQLite (`<SPX_OUTPUT_DIR>/monitoring.sqlite3`), which the bundled Grafana can read.

System lists are recognised by their internal name (`EntityTypeName`) and template (`BaseTemplate`), and roles by their type (`RoleTypeKind`), so the extraction does not depend on the site language.

## Requirements

- Docker with Docker Compose.
- Network access to the SharePoint farm.
- An account with read access to the sites to extract (NTLM, `DOMAIN\user`). To find every site collection under the root, it needs read access to the whole web application (e.g. a *Full Read* user policy); without it, only the root collection is crawled and the failure is listed in `discovery_errors_<root>.json`.
- If the farm's certificate comes from an internal CA: the CA bundle in PEM format (see [TLS](#tls)).

## Quick start

```bash
cp .env.example .env
# fill in .env (see Configuration)

docker compose build

# once: check that the server certificate is accepted
docker compose run --rm spx-extractor python src/tools/check_tls.py

# interactive menu
docker compose run --rm spx-extractor
```

The menu runs nothing without confirmation.

## Configuration

All settings are environment variables, read from `.env` by Docker Compose.

| Variable | Required | Description |
|---|---|---|
| `SPX_SITE_URL` | yes | URL of the site collection to extract (the *root*). Host-named (`https://intranet.example.org`) or path-based (`https://sp.example.org/sites/hr`). |
| `SPX_ROOT_NAME` | yes | Name you give that root: letters, digits, `-` and `_`. Used for the output folder and file names. |
| `SPX_USERNAME` | yes | NTLM account, e.g. `DOMAIN\user`. |
| `SPX_PASSWORD` | yes | Its password. |
| `SPX_CA_BUNDLE` | no | Path of a PEM CA bundle, as seen from the container (e.g. `/app/certs/ca_bundle.pem`). Empty: certifi's public CA list. |
| `SPX_OUTPUT_DIR` | no | Output folder, default `./discovery_output`. |
| `SPX_INSPECT_PATH` | no | Path inspected by the diagnostic `inspect_path.py`, default the root itself. |
| `SPX_INSPECT_LIST` | no | List whose fields `inspect_path.py` also prints. |
| `SPX_DOCKER_SUBNET` | no | Subnet of the extractor's Docker network, default `10.200.0.0/24` (see [Docker network](#docker-network)). |
| `SPX_MONITOR_DB` | no | SQLite file for the HTTP call log, default `<SPX_OUTPUT_DIR>/monitoring.sqlite3`. |

One root is one `SPX_SITE_URL` + `SPX_ROOT_NAME` pair, and every script works on that root only. To process another one without editing `.env`, override both:

```bash
docker compose run --rm \
  -e SPX_SITE_URL=https://sp.example.org/sites/finance \
  -e SPX_ROOT_NAME=finance \
  spx-extractor
```

## Workflow

For each root, in this order:

| Step | Menu entry | Command |
|---|---|---|
| 1. Find the sites | Site discovery > Recursive crawl | `python src/discovery/discover_sites.py` |
| 2. Metadata and permissions | Metadata and permissions extraction > All sites | `python src/extraction/extract_metadata.py --all-from-discovery discovery_output/sites_index_<root>.json` |
| 3. List attachments | List attachments extraction | `python src/extraction/extract_list_attachments.py --all` |
| 4. Version history | Version history extraction | `python src/extraction/extract_versions.py --all` |
| 5. Check errors | Consolidate errors | `python src/extraction/consolidate_errors.py` |
| 6. Download files | File download | `python src/extraction/download_files.py --all [limit]` |
| 7. Download past versions | (command line only) | `python src/extraction/download_versions.py <site>` |
| 8. Statistics | Statistics | `python src/extraction/aggregate_stats.py --output stats_summary.json` |

Commands run in the container, from the project root: `docker compose run --rm spx-extractor python ...`.

- Step 1 lists the root's site collections through `SiteData.asmx` (for a root at `/`, every collection on its host; for `/sites/hr`, only those under that path), then crawls the subsites of each.
- Paths use the default `SPX_OUTPUT_DIR` (`discovery_output`).
- `<site>` is the site's folder name under `discovery_output/extraction/<root>/`: the last segment of its URL, or the whole path joined with `-` when several sites of the root share that segment.
- A single site can be extracted with `extract_metadata.py /sites/hr/team`. Its result is merged into the root's `run_summary.json`.
- `extract_metadata.py --permissions-all` (menu: Permissions catch-up) re-extracts the permissions of every site already extracted without error.
- `download_files.py` also works per site (`--site <site>`) or per library (`<site path> <library title>`), with `--workers N` (default 6) and `--limit N`.
- Do not run two scripts on the same root at the same time.

## Output

```
discovery_output/
├── sites_index_<root>.json, sites_tree_<root>.json, discovery_errors_<root>.json
└── extraction/<root>/
    ├── run_summary.json, permissions_run_summary.json,
    │   versions_run_summary.json, attachments_run_summary.json
    └── <site>/
        ├── site_metadata.json
        ├── permissions.json                        site permissions, lists with broken inheritance
        ├── documents/<library>/
        │   ├── _schema_columns.json, .state.json
        │   └── <folders>/<file>                    binary
        │       <folders>/<file>.meta.json          metadata
        │       <folders>/<file>.versions.json      version history
        │       <folders>/<file>.history/<version><ext>
        ├── lists/<list>/
        │   ├── items.json, _schema_columns.json, .state.json
        │   └── attachments/<item id>/<file>        attachment + <file>.meta.json
        └── pages/                                  page inventory, bodies and history
```

`run_summary.json` lists, per site, the libraries and lists extracted, those skipped and why, and the errors.

## Scripts

Entry point: `src/main.py`, the interactive menu started by default in the container. Every script also runs on its own; run it without arguments to see its usage.

| Folder | Script | Purpose |
|---|---|---|
| `discovery/` | `discover_sites.py` | Every site collection under the root (`SiteData.asmx`), then every subsite of each (breadth-first over `_api/web/webs`) |
| | `discover_personal_sites.py` | Estimate of personal sites through the Search API (`SPX_SITE_URL` = My Site host) |
| | `inspect_path.py` | Lists, libraries and subsites under `SPX_INSPECT_PATH` |
| `extraction/` | `extract_metadata.py` | Schemas, file metadata, list items and permissions |
| | `extract_list_attachments.py` | Metadata of list item attachments |
| | `extract_versions.py` | `<file>.versions.json` for each versioned file |
| | `download_files.py` | Binaries of files and attachments |
| | `download_versions.py` | Binaries of past versions, size checked to the byte |
| | `enrich_file_meta.py` | Adds created / modified dates and authors to a site's `.meta.json` |
| | `extract_site_users.py` | User table of a site collection (`site_users.json`), to resolve the `<Column>Id` fields of list items |
| | `inventory_site_pages.py`, `extract_page_history.py` | Page inventory of a site, version history of a page |
| | `aggregate_stats.py` | Counts and volumes from the files on disk, no SharePoint call |
| | `consolidate_errors.py` | All errors of all `run_summary.json` files, to `errors_consolidated.csv` |
| `reporting/` | `extract_full_control_members.py` | CSV of people with full control on each site (`--role` adds custom roles) |
| `tools/` | `check_tls.py` | TLS pre-flight check, and why a certificate is refused |
| | `probe_one_call.py` | One authenticated call, to check access to a host |
| | `check_broken_sites.py` | Unauthenticated HTTP check of the sites in error |
| | `check_hidden_list.py` | Analysis of the hidden lists that were skipped |
| | `check_broken_inheritance.py` | Broken permission inheritance per root |
| | `check_subweb_name_conflicts.py` | Subsite names that collide with another name |
| | `check_list_id_fields.py`, `probe_list_fields.py` | Checks of user / lookup id fields in extracted lists |
| | `inventory_site.py`, `inspect_lists_and_views.py` | Reference counts of an extracted site; lists and views on the server compared with the disk |
| `src/` | `sharepoint_client.py`, `persistent_session.py`, `site_urls.py`, `monitoring.py` | Configuration and connection, HTTP session and throttling, URLs, HTTP call log |

## Version history

- `extract_versions.py` writes `<file>.versions.json` for each file that can have a history: the current version, then each past one (label, size, date, author, comment, URL). `0.1` drafts are skipped, and so are `1.0` files in a library without minor versions. With minor versions, `1.0` is usually published from drafts `0.1`, `0.2`…, so these files are queried; the setting is read once per library.
- `download_versions.py <site>` downloads past versions to `<file>.history/<label><ext>`. A file whose size differs from the server's is deleted and counted as an error.
- Each version is read from its `_vti_history` URL. SharePoint 2019 answers 404 there for pages (`.aspx`), so on a 404 the version is read through the REST API (`versions/getById(<id>)/$value`).

Known limits:

- Unpublished `0.1` drafts have no history and are skipped.
- Versions the library did not keep (versioning disabled, version limit reached) cannot be recovered through the API.
- Deleted AD accounts have no login any more, only their display name.

## TLS

Without `SPX_CA_BUNDLE`, the server certificate is checked against certifi's public CA list. TLS verification is never disabled.

If the farm uses an internal CA:

1. Get the **root** certificate of that CA from your IT team, in PEM format. A root's `subject` and `issuer` are identical:
   ```bash
   openssl x509 -in root.pem -noout -subject -issuer -fingerprint -sha256
   ```
   Have the fingerprint confirmed by your IT team. Do not take it from the output of `openssl s_client -showcerts`, which shows what the server (or a proxy) sends.
2. Put it, with any other required root, in a bundle under `certs/`, e.g. `certs/ca_bundle.pem`. The project folder is mounted at `/app` in the container, and `certs/` is ignored by Git.
3. Set `SPX_CA_BUNDLE=/app/certs/ca_bundle.pem` in `.env`.
4. Check: `docker compose run --rm spx-extractor python src/tools/check_tls.py [host ...]`. It tells a missing CA from a network device that presents its own certificate (TLS inspection).

## Monitoring

`docker compose up -d grafana` starts Grafana on port 3000, with the SQLite plugin, reading `SPX_OUTPUT_DIR` (mounted at `/app/discovery_output`). Add a SQLite data source on `/app/discovery_output/monitoring.sqlite3` to follow calls, errors and delays during a long run. `python src/monitoring.py` prints a refreshing summary in the terminal.

## Docker network

`docker-compose.yml` gives the extractor network a fixed subnet, `10.200.0.0/24` by default. If that range overlaps your SharePoint servers, your LAN or another Docker network ("Pool overlaps with other one on this address space"), set another one with `SPX_DOCKER_SUBNET` in `.env`. A Docker bridge on the same range as a server captures its traffic, and connections fail with "No route to host".

## Security

- `.env` holds credentials. It is ignored by Git (as is any `.env*` copy, except `.env.example`): never commit it, and keep it private (`chmod 600 .env`).
- Outputs (the `SPX_OUTPUT_DIR` folder, CSV reports, `stats_summary.json`) hold site, list and people names, and the downloaded files themselves. `discovery_output*/`, `*.csv` and `stats_summary.json` are ignored by Git: if you change `SPX_OUTPUT_DIR`, keep a `discovery_output` prefix or a folder outside the repository.
- Console output is limited to names and counts, never file contents or field values.
- A full extraction reads everything the account can read. Use an account scoped to what you are allowed to extract.

## Development

See [CONTRIBUTING.md](CONTRIBUTING.md) for how to report bugs and propose changes.

```bash
pip install -r requirements.txt -r requirements-dev.txt
python3 -m unittest discover -s tests
black src tests
```

Tests make no network call.

## License

[MIT](LICENSE)

## Security

To report a vulnerability, see [SECURITY.md](SECURITY.md). Do not open a public issue for security problems.
