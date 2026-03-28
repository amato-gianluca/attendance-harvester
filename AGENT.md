# AGENT.md

This repository contains a Python CLI that harvests Microsoft Teams attendance data and exports it to CSV and JSON.

## Entry Points

- `main.py`: CLI entry point and top-level workflow dispatch.
- `src/auth.py`: MSAL authentication and token cache handling.
- `src/graph_client.py`: Microsoft Graph API client and retry behavior.
- `src/meeting_resolver.py`: Meeting discovery and attendance extraction.
- `src/exporter.py`: CSV/JSON export logic.
- `src/sharepoint_csv_uploader.py`: Optional SharePoint upload support for CSV exports.
- `src/team_filter.py`: Team filtering by regex.

## Common Commands

Install dependencies:

```bash
pip install -r requirements.txt
```

Run the default harvest flow:

```bash
python main.py
```

Rebuild CSV files from existing JSON exports:

```bash
python main.py --rebuild-csv output
```

Upload existing CSV exports to SharePoint:

```bash
python main.py --upload-csv-to-sharepoint
```

Basic syntax check:

```bash
python -m py_compile main.py src/*.py
```

## Configuration Notes

- Main runtime configuration lives in `config.yaml`.
- Start from `config.yaml.template` if a local config is missing.
- Authentication can use either public device-code flow or confidential client credentials.
- `TEAMS_HARVESTER_CLIENT_SECRET` can provide the client secret when it is not stored in config.

## Editing Guidance

- Keep `main.py` focused on CLI orchestration and mode selection.
- Prefer small helpers for isolated execution paths instead of growing `main()` further.
- Preserve current behavior around logging summaries and explicit exit handling.
- Avoid committing local artifacts such as `cache/`, generated `output/`, notebooks, or ad hoc CSV files unless explicitly requested.

## Validation

- At minimum, run `python -m py_compile main.py src/*.py` after Python changes.
- If behavior changes touch authentication, export, or SharePoint upload flows, validate the affected CLI mode directly when credentials and config are available.

## Commit

- Prepare commit messaged made from a short summary and a list of items with more specific details.