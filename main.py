#!/usr/bin/env python3
"""
Microsoft Teams Attendance Harvester - Main Entry Point

Scans subscribed Teams, filters by name regex, discovers meetings,
downloads attendance logs, and exports to CSV/JSON.
"""
import argparse
import json
import logging
import os
import sys
from pathlib import Path
from typing import Any
from uuid import UUID

import yaml

from src.app_config import AppConfig, AuthConfig, load_app_config
from src.auth import AuthenticationError, Authenticator
from src.exporter import AttendanceExporter
from src.graph_client import GraphAPIError, GraphClient
from src.meeting_resolver import MeetingResolver
from src.sharepoint_csv_uploader import SharePointCSVUploader
from src.team_filter import TeamFilter


def _is_valid_guid(value: Any) -> bool:
    """Return True if value is a valid GUID/UUID string."""
    if not value or not isinstance(value, str):
        return False
    try:
        UUID(value)
        return True
    except (ValueError, TypeError, AttributeError):
        return False


def setup_logging(verbose: bool = False):
    """Configure logging."""
    level = logging.DEBUG if verbose else logging.INFO

    logging.basicConfig(
        level=level,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    # Reduce noise from libraries
    logging.getLogger("msal").setLevel(logging.WARNING)
    logging.getLogger("urllib3").setLevel(logging.WARNING)


def load_attendance_from_json_inputs(json_inputs: list[str]) -> list[dict]:
    """
    Load attendance payloads from JSON files, directories, or glob patterns.

    Args:
        json_inputs: One or more filesystem inputs

    Returns:
        List of attendance data dictionaries
    """
    json_files: list[Path] = []

    for raw_input in json_inputs:
        path = Path(raw_input)

        if path.is_dir():
            json_files.extend(sorted(path.rglob("*.json")))
            continue

        matches = sorted(Path().glob(raw_input))
        if matches:
            json_files.extend(match for match in matches if match.is_file() and match.suffix.lower() == ".json")
            continue

        if path.is_file() and path.suffix.lower() == ".json":
            json_files.append(path)
            continue

        raise FileNotFoundError(f"No JSON files found for input: {raw_input}")

    unique_files = list(dict.fromkeys(file.resolve() for file in json_files))
    attendance_data: list[dict] = []

    for json_file in unique_files:
        with open(json_file, "r", encoding="utf-8") as f:
            payload = json.load(f)

        if not isinstance(payload, dict):
            raise ValueError(f"JSON file must contain a single attendance object: {json_file}")

        attendance_data.append(payload)

    return attendance_data


def build_exporter(config: AppConfig) -> AttendanceExporter:
    """Build exporter with per-format output directories."""
    output_config = config.output

    return AttendanceExporter(
        output_dir=str(output_config.directory),
        filename_pattern=output_config.filename_pattern,
        csv_output_dir=str(output_config.csv_directory),
        json_output_dir=str(output_config.json_directory),
        min_csv_report_duration_seconds=output_config.min_csv_report_duration_seconds,
        team_directories_file=output_config.team_directories_file
    )


def get_csv_output_dir(config: AppConfig) -> Path:
    """Resolve the CSV output directory from config."""
    return config.output.csv_directory


def acquire_access_token_from_config(
    config: AppConfig,
    auth_config: AuthConfig,
    log_details: bool = False
) -> tuple[str, str]:
    """Resolve auth settings from config, build an authenticator, and acquire an access token."""
    logger = logging.getLogger(__name__)

    if log_details:
        well_known_clients = {
            "14d82eec-204b-4c2f-b7e8-296a70dab67e": "Microsoft Graph PowerShell",
            "04b07795-8ddb-461a-bbee-02f9e1bf7b46": "Azure CLI",
            "d3590ed6-52b3-4102-aeff-aad2292ab01c": "Microsoft Office"
        }

        if auth_config.mode == "confidential":
            logger.info("Using confidential client credentials mode")
            logger.info("Using custom Azure AD application with app-only token")
        elif auth_config.client_id in well_known_clients:
            logger.info("Using public device code mode")
            logger.info(f"Using {well_known_clients[auth_config.client_id]} public client (no Azure app needed)")
            logger.info("You'll authenticate with your Microsoft credentials in the browser")
        else:
            logger.info("Using public device code mode")
            logger.info("Using custom Azure AD application")

    cache_dir_path = config.cache.directory
    cache_dir_path.mkdir(parents=True, exist_ok=True)
    client_secret = auth_config.client_secret or os.getenv("TEAMS_HARVESTER_CLIENT_SECRET")

    authenticator = Authenticator(
        client_id=auth_config.client_id,
        authority=auth_config.authority,
        scopes=auth_config.scopes,
        cache_path=cache_dir_path / auth_config.token_cache,
        auth_mode=auth_config.mode,
        client_secret=client_secret
    )

    if auth_config.clear_cache:
        logger.info("Clearing token cache as requested")
        authenticator.clear_cache()

    return authenticator.acquire_token(), auth_config.mode


def build_sharepoint_csv_uploader(config: AppConfig, force_enable: bool = False):
    """Build optional SharePoint uploader for CSV exports."""
    sharepoint_config = config.output.sharepoint_csv
    if not force_enable and not sharepoint_config.auto_upload:
        return None

    access_token, _ = acquire_access_token_from_config(
        config=config,
        auth_config=sharepoint_config.auth
    )
    graph_client = GraphClient(
        access_token=access_token,
        max_retries=config.api.max_retries,
        retry_backoff_factor=config.api.retry_backoff_factor,
        timeout=config.api.timeout
    )

    return SharePointCSVUploader(
        graph_client=graph_client,
        site_id=sharepoint_config.site_id,
        site_hostname=sharepoint_config.site_hostname,
        site_path=sharepoint_config.site_path,
        drive_id=sharepoint_config.drive_id,
        drive_name=sharepoint_config.drive_name,
        folder_path=sharepoint_config.folder_path
    )


def upload_csv_exports_to_sharepoint(uploader, exporter: AttendanceExporter, created_files: list[Path]) -> list[str]:
    """Upload newly created CSV exports to SharePoint if configured."""
    if not uploader:
        return []

    csv_files = [Path(path) for path in created_files if Path(path).suffix.lower() == ".csv"]
    if not csv_files:
        return []

    return uploader.upload_files(csv_files, exporter.csv_output_dir)


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description="Microsoft Teams Attendance Harvester - Download attendance logs from Teams meetings"
    )
    parser.add_argument(
        "-c", "--config",
        default="config.yaml",
        help="Path to configuration file (default: config.yaml)"
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose logging"
    )
    parser.add_argument(
        "--clear-cache",
        action="store_true",
        help="Clear authentication token cache and re-authenticate"
    )
    parser.add_argument(
        "--skip-processed",
        action="store_true",
        default=True,
        help="Skip meetings already processed (default: True)"
    )
    parser.add_argument(
        "--team-regex",
        help="Override team filter regex from config"
    )
    parser.add_argument(
        "--lookback-days",
        type=int,
        help="Override lookback days from config"
    )
    parser.add_argument(
        "--lookahead-days",
        type=int,
        help="Override lookahead days from config"
    )
    parser.add_argument(
        "--rebuild-csv",
        nargs="+",
        metavar="PATH",
        help="Force CSV regeneration from existing attendance JSON file(s), directory, or glob without connecting to Teams"
    )
    parser.add_argument(
        "--min-csv-report-duration-seconds",
        type=int,
        help="Only export CSV for reports whose meeting duration is at least this many seconds; JSON export is unaffected"
    )
    parser.add_argument(
        "--upload-csv-to-sharepoint",
        action="store_true",
        help="Upload existing local CSV exports to SharePoint and exit"
    )

    return parser.parse_args()


def run_rebuild_csv(config: AppConfig, json_inputs: list[str]) -> None:
    """Rebuild CSV exports from existing attendance JSON payloads."""
    logger = logging.getLogger(__name__)

    logger.info("Step 1: Loading attendance data from existing JSON files")
    attendance_data = load_attendance_from_json_inputs(json_inputs)

    if not attendance_data:
        logger.warning("No attendance data found in the provided JSON input(s).")
        return

    logger.info(f"✓ Loaded {len(attendance_data)} attendance payload(s)")

    logger.info("Step 2: Exporting attendance data")
    exporter = build_exporter(config)
    sharepoint_csv_uploader = build_sharepoint_csv_uploader(config)

    export_format = "csv"
    created_files = exporter.export_batch(attendance_data, format=export_format)
    uploaded_files = upload_csv_exports_to_sharepoint(sharepoint_csv_uploader, exporter, created_files)

    logger.info(f"✓ Created {len(created_files)} output files in {exporter.csv_output_dir}")
    if uploaded_files:
        logger.info(f"✓ Uploaded {len(uploaded_files)} CSV files to SharePoint")

    logger.info("=" * 70)
    logger.info("SUMMARY")
    logger.info("=" * 70)
    logger.info(f"JSON files loaded: {len(attendance_data)}")
    logger.info(f"CSV output directory: {exporter.csv_output_dir}")
    logger.info(f"Files created: {len(created_files)}")
    if uploaded_files:
        logger.info(f"SharePoint CSV uploads: {len(uploaded_files)}")
    logger.info("=" * 70)
    logger.info("✓ CSV rebuild completed successfully")


def run_upload_csv_to_sharepoint(config: AppConfig) -> None:
    """Upload existing local CSV exports to SharePoint."""
    logger = logging.getLogger(__name__)

    uploader = build_sharepoint_csv_uploader(config, force_enable=True)
    if not uploader:
        raise ValueError("SharePoint CSV upload could not be initialized from output.sharepoint_csv")

    csv_root = get_csv_output_dir(config)
    csv_files = sorted(csv_root.rglob("*.csv"))
    uploaded_files = uploader.upload_files(csv_files, csv_root)

    logger.info("=" * 70)
    logger.info("SUMMARY")
    logger.info("=" * 70)
    logger.info(f"CSV output directory: {csv_root}")
    logger.info(f"CSV files: {len(csv_files)}")
    logger.info(f"SharePoint CSV uploads: {len(uploaded_files)}")
    logger.info("=" * 70)
    logger.info("✓ CSV SharePoint upload completed successfully")


def run_harvest(config: AppConfig) -> None:
    """Run the default Teams attendance harvesting workflow."""
    logger = logging.getLogger(__name__)

    # Step 1: Authenticate
    logger.info("Step 1: Authenticating with Microsoft Graph API")
    access_token, auth_mode = acquire_access_token_from_config(
        config=config,
        auth_config=config.auth,
        log_details=True
    )
    logger.info("✓ Authentication successful")

    # Step 2: Initialize Graph client
    logger.info("\nStep 2: Initializing Graph API client")
    target_user_id = config.auth.target_user_id
    if auth_mode == "confidential" and not target_user_id:
        raise ValueError(
            "Confidential mode requires auth.target_user_id in config (UPN or object ID), "
            "used for /users/{id}/... Graph calls."
        )

    graph_client = GraphClient(
        access_token=access_token,
        max_retries=config.api.max_retries,
        retry_backoff_factor=config.api.retry_backoff_factor,
        timeout=config.api.timeout,
        user_id=target_user_id if auth_mode == "confidential" else None,
        metadata_cache_file=str(config.cache.metadata_cache_file)
    )
    logger.info("✓ Graph API client initialized")

    if config.auth.clear_cache:
        graph_client.clear_metadata_cache()

    # Step 3: Get and filter teams
    logger.info("\nStep 3: Fetching and filtering teams")
    teams = graph_client.get_joined_teams()

    if config.meetings.include_associated_teams:
        associated = graph_client.get_associated_teams()
        team_ids = {t["id"] for t in teams}
        for assoc_team in associated:
            assoc_team_id = assoc_team.get("teamId")
            if not _is_valid_guid(assoc_team_id):
                logger.debug(
                    "Skipping associated team with non-GUID teamId: %s (displayName=%s)",
                    assoc_team_id,
                    assoc_team.get("displayName", "Unknown")
                )
                continue

            if assoc_team_id not in team_ids:
                teams.append({
                    "id": assoc_team_id,
                    "displayName": assoc_team.get("displayName", "Unknown")
                })
                team_ids.add(assoc_team_id)

    team_filter = TeamFilter(config.team_filter.regex)
    filtered_teams = team_filter.filter_teams(teams)

    if not filtered_teams:
        logger.warning("No teams matched the filter criteria. Exiting.")
        graph_client.sync_filtered_teams_cache([])
        return

    graph_client.sync_filtered_teams_cache(filtered_teams)

    logger.info(f"✓ Found {len(filtered_teams)} matching teams")
    for team in filtered_teams:
        logger.info(f"  - {team['displayName']}")

    # Step 4: Get channels (optionally only the primary/General channel)
    logger.info("\nStep 4: Fetching channels")
    teams_with_channels = []

    for team in filtered_teams:
        try:
            if not _is_valid_guid(team.get("id", "")):
                logger.warning(
                    "  ✗ Skipping %s: invalid team id '%s' (not a GUID)",
                    team.get("displayName", "Unknown"),
                    team.get("id", "")
                )
                continue

            if config.meetings.general_channel_only:
                general_channel = graph_client.get_team_primary_channel(team["id"])
                if general_channel:
                    teams_with_channels.append({
                        "team": team,
                        "channel": general_channel
                    })
                    logger.info(f"  ✓ {team['displayName']}: Found primary channel")
                else:
                    logger.warning(f"  ✗ {team['displayName']}: Primary channel not found")
            else:
                channels = graph_client.get_team_channels(team["id"])
                for channel in channels:
                    teams_with_channels.append({
                        "team": team,
                        "channel": channel
                    })
                logger.info(f"  ✓ {team['displayName']}: {len(channels)} channels")

        except GraphAPIError as e:
            logger.error(f"  ✗ Failed to get channels for {team['displayName']}: {e}")

    logger.info(f"✓ Total team-channel combinations: {len(teams_with_channels)}")

    # Step 5: Discover meetings and extract attendance
    logger.info("\nStep 5: Discovering meetings and extracting attendance")
    meeting_resolver = MeetingResolver(
        graph_client=graph_client,
        json_output_dir=str(config.output.json_directory)
    )

    attendance_data = meeting_resolver.extract_all_attendance(
        teams_with_channels=teams_with_channels,
        lookback_days=config.meetings.lookback_days,
        lookahead_days=config.meetings.lookahead_days
    )

    if not attendance_data:
        logger.warning("No attendance data found. This could mean:")
        logger.warning("  - No meetings in the specified time range")
        logger.warning("  - No attendance reports available (reports are organizer-only)")
        logger.warning("  - All meetings were already processed (use --skip-processed=False to re-process)")
        return

    logger.info(f"✓ Extracted {len(attendance_data)} attendance reports")

    # Step 6: Export attendance data
    logger.info("\nStep 6: Exporting attendance data")
    exporter = build_exporter(config)

    export_format = config.output.format
    sharepoint_csv_uploader = None
    if export_format in ("csv", "both"):
        sharepoint_csv_uploader = build_sharepoint_csv_uploader(config)
    created_files = exporter.export_batch(attendance_data, format=export_format)
    uploaded_files = upload_csv_exports_to_sharepoint(
        sharepoint_csv_uploader,
        exporter,
        created_files
    )

    logger.info(f"✓ Created {len(created_files)} output files")
    if uploaded_files:
        logger.info(f"✓ Uploaded {len(uploaded_files)} CSV files to SharePoint")

    logger.info("=" * 70)
    logger.info("SUMMARY")
    logger.info("=" * 70)
    logger.info(f"Teams scanned: {len(filtered_teams)}")
    logger.info(f"Attendance reports extracted: {len(attendance_data)}")
    logger.info(f"Files created: {len(created_files)}")
    if uploaded_files:
        logger.info(f"SharePoint CSV uploads: {len(uploaded_files)}")
    if export_format in ("csv", "both"):
        logger.info(f"CSV output directory: {config.output.csv_directory}")
    if export_format in ("json", "both"):
        logger.info(f"JSON output directory: {config.output.json_directory}")
    logger.info("=" * 70)
    logger.info("✓ Attendance harvesting completed successfully")


def main():
    """Main execution function."""
    args = parse_args()

    # Setup logging
    setup_logging(args.verbose)
    logger = logging.getLogger(__name__)

    logger.info("=" * 70)
    logger.info("Microsoft Teams Attendance Harvester")
    logger.info("=" * 70)

    try:
        config = load_app_config(args.config, args)
        if args.rebuild_csv:
            run_rebuild_csv(config, args.rebuild_csv)
        elif args.upload_csv_to_sharepoint:
            run_upload_csv_to_sharepoint(config)
        else:
            run_harvest(config)
    except FileNotFoundError:
        print(f"Error: Configuration file not found: {args.config}")
        print("Please copy config.yaml.template to config.yaml and fill in your settings.")
        sys.exit(1)
    except KeyboardInterrupt:
        logger.warning("\nOperation cancelled by user")
        sys.exit(130)
    except AuthenticationError as e:
        logger.error(f"Authentication failed: {e}")
        sys.exit(1)
    except GraphAPIError as e:
        logger.error(f"Graph API error: {e}")
        sys.exit(1)
    except (OSError, ValueError, KeyError, yaml.YAMLError) as e:
        logger.error(f"Failed to initialize or execute: {e}")
        sys.exit(1)
    except Exception as e:
        logger.exception(f"Unexpected error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
