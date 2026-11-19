# Contributing

Thanks for your interest in this project. Bug reports, fixes and improvements are welcome.

## Reporting a bug

Open an issue describing:

- what you expected, and what happened instead;
- the exact command you ran (Docker or bare Python);
- the error message, or the relevant entries of `run_summary.json` / `discovery_errors_<root>.json`.

Never paste credentials, tokens, cookie values or real internal hostnames: redact URLs and account names before sharing logs or JSON output.

## Proposing a change

1. Fork the repository and create a branch from `main`: `fix/...`, `feat/...`, `docs/...`.
2. Make one focused change per branch.
3. Add or update tests in `tests/` when the change affects behavior. Tests must not make any network call.
4. Run the test suite and the formatter before opening the pull request:

   ```bash
   pip install -r requirements.txt -r requirements-dev.txt
   python3 -m unittest discover -s tests
   black src tests
   ```

5. Open a pull request describing the problem and the solution, and reference the issue it fixes if there is one.

## Code style

- Python, formatted with `black` (see `requirements-dev.txt`); run `black src tests` before committing.
- Scripts under `src/` read their configuration from environment variables only (see `get_config()` in `src/sharepoint_client.py`). Never hardcode URLs, accounts, passwords or machine-specific paths.
- TLS verification is never disabled in the extraction path. `SPX_CA_BUNDLE` is the only supported way to trust an internal CA.
- Console output stays limited to names and counts — never file contents or field values.
- Anything produced by a run (`discovery_output*/`, `*.csv`, `stats_summary.json`, `certs/`, `.env*`) stays out of Git; extend `.gitignore` and `.dockerignore` if you add a new output.
- Keep the README up to date when you change the workflow, the configuration variables or the scripts table.

## Commit messages

Short imperative subject (e.g. `Skip hidden lists without a business template`), with a body explaining *why* when it is not obvious from the subject.

## Licensing

By contributing, you agree that your contributions are licensed under the [MIT License](LICENSE), like the rest of the project.

## Code of conduct

This project follows the [Contributor Covenant](https://www.contributor-covenant.org), version 2.1. Please be respectful and constructive in issues and pull requests.
