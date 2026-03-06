#!/usr/bin/env python3
"""
Microsoft Teams Attendance Harvester - Main Entry Point

Scans subscribed Teams, filters by name regex, discovers meetings,
downloads attendance logs, and exports to CSV/JSON.
"""
import argparse
import logging
import sys
from pathlib import Path

import yaml

from src.auth import Authenticator, AuthenticationError
from src.graph_client import GraphClient, GraphAPIError
from src.team_filter import TeamFilter
from src.meeting_resolver import MeetingResolver
from src.exporter import AttendanceExporter


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


def load_config(config_path: str) -> dict:
    """Load configuration from YAML file."""
    config_file = Path(config_path)

    if not config_file.exists():
        print(f"Error: Configuration file not found: {config_path}")
        print("Please copy config.yaml.template to config.yaml and fill in your settings.")
        sys.exit(1)

    with open(config_file, "r") as f:
        config = yaml.safe_load(f)

    return config


def main():
    """Main execution function."""
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

    args = parser.parse_args()

    # Setup logging
    setup_logging(args.verbose)
    logger = logging.getLogger(__name__)

    logger.info("=" * 70)
    logger.info("Microsoft Teams Attendance Harvester")
    logger.info("=" * 70)

    # Load configuration
    try:
        config = load_config(args.config)
    except Exception as e:
        logger.error(f"Failed to load configuration: {e}")
        sys.exit(1)

    # Override config with command-line arguments
    if args.team_regex:
        config["team_filter"]["regex"] = args.team_regex
    if args.lookback_days:
        config["meetings"]["lookback_days"] = args.lookback_days

    try:
        # Step 1: Authenticate
        logger.info("Step 1: Authenticating with Microsoft Graph API")

        client_id = config["azure"]["client_id"]

        # Check if using well-known public client
        well_known_clients = {
            "14d82eec-204b-4c2f-b7e8-296a70dab67e": "Microsoft Graph PowerShell",
            "04b07795-8ddb-461a-bbee-02f9e1bf7b46": "Azure CLI",
            "d3590ed6-52b3-4102-aeff-aad2292ab01c": "Microsoft Office"
        }

        if client_id in well_known_clients:
            logger.info(f"Using {well_known_clients[client_id]} public client (no Azure app needed)")
            logger.info("You'll authenticate with your Microsoft credentials in the browser")
        else:
            logger.info("Using custom Azure AD application")

        authenticator = Authenticator(
            client_id=client_id,
            authority=config["azure"]["authority"],
            scopes=config["scopes"],
            cache_dir=config["cache"]["directory"],
            cache_filename=config["cache"]["token_cache"]
        )

        if args.clear_cache:
            logger.info("Clearing token cache as requested")
            authenticator.clear_cache()

        access_token = authenticator.acquire_token()
        logger.info("✓ Authentication successful")

        # Step 2: Initialize Graph client
        logger.info("\nStep 2: Initializing Graph API client")
        graph_client = GraphClient(
            access_token=access_token,
            max_retries=config["api"]["max_retries"],
            retry_backoff_factor=config["api"]["retry_backoff_factor"],
            timeout=config["api"]["timeout"]
        )
        logger.info("✓ Graph API client initialized")

        # Step 3: Get and filter teams
        logger.info("\nStep 3: Fetching and filtering teams")
        teams = graph_client.get_joined_teams()

        if config["meetings"].get("include_associated_teams", False):
            associated = graph_client.get_associated_teams()
            # Merge, avoiding duplicates by ID
            team_ids = {t["id"] for t in teams}
            for assoc_team in associated:
                if assoc_team.get("teamId") not in team_ids:
                    # Convert associated team format to regular team format
                    teams.append({
                        "id": assoc_team.get("teamId"),
                        "displayName": assoc_team.get("displayName", "Unknown")
                    })

        team_filter = TeamFilter(config["team_filter"]["regex"])
        filtered_teams = team_filter.filter_teams(teams)

        if not filtered_teams:
            logger.warning("No teams matched the filter criteria. Exiting.")
            return

        logger.info(f"✓ Found {len(filtered_teams)} matching teams")
        for team in filtered_teams:
            logger.info(f"  - {team['displayName']}")

        # Step 4: Get channels (optionally filter to General only)
        logger.info("\nStep 4: Fetching channels")
        teams_with_channels = []

        for team in filtered_teams:
            try:
                channels = graph_client.get_team_channels(team["id"])

                if config["meetings"].get("general_channel_only", True):
                    general_channel = team_filter.get_general_channel(channels)
                    if general_channel:
                        teams_with_channels.append({
                            "team": team,
                            "channel": general_channel
                        })
                        logger.info(f"  ✓ {team['displayName']}: Found General channel")
                    else:
                        logger.warning(f"  ✗ {team['displayName']}: General channel not found")
                else:
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
        checkpoint_file = Path(config["cache"]["directory"]) / config["cache"]["checkpoint_file"]

        meeting_resolver = MeetingResolver(
            graph_client=graph_client,
            checkpoint_file=str(checkpoint_file)
        )

        lookback_days = config["meetings"]["lookback_days"]
        attendance_data = meeting_resolver.extract_all_attendance(
            teams_with_channels=teams_with_channels,
            lookback_days=lookback_days
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
        exporter = AttendanceExporter(
            output_dir=config["output"]["directory"],
            filename_pattern=config["output"]["filename_pattern"]
        )

        export_format = config["output"]["format"]
        created_files = exporter.export_batch(attendance_data, format=export_format)

        logger.info(f"✓ Created {len(created_files)} output files in {config['output']['directory']}")

        # Summary
        logger.info("\n" + "=" * 70)
        logger.info("SUMMARY")
        logger.info("=" * 70)
        logger.info(f"Teams scanned: {len(filtered_teams)}")
        logger.info(f"Attendance reports extracted: {len(attendance_data)}")
        logger.info(f"Files created: {len(created_files)}")
        logger.info(f"Output directory: {config['output']['directory']}")
        logger.info("=" * 70)
        logger.info("✓ Attendance harvesting completed successfully")

    except AuthenticationError as e:
        logger.error(f"Authentication failed: {e}")
        sys.exit(1)
    except GraphAPIError as e:
        logger.error(f"Graph API error: {e}")
        sys.exit(1)
    except KeyboardInterrupt:
        logger.warning("\nOperation cancelled by user")
        sys.exit(130)
    except Exception as e:
        logger.exception(f"Unexpected error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
