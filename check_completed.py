#!/usr/bin/env python3
"""Report completed courses based on instructor in-meeting duration across all CSV exports."""
from __future__ import annotations

import argparse
import csv
import io
import logging
import os
import re
from pathlib import Path

from src.app_config import load_app_config
from src.auth import Authenticator
from src.graph_client import GraphClient
from src.sharepoint_csv_uploader import SharePointCSVUploader


LOGGER = logging.getLogger(__name__)
EXPECTED_HOURS_RE = re.compile(r"\[(\d+)\]\s*$")
DURATION_RE = re.compile(r"^(?:(\d+)\s+days?,\s+)?(\d+):(\d{2}):(\d{2})$")
HMS_TEXT_RE = re.compile(r"^\s*(?:(\d+)h)?\s*(?:(\d+)m)?\s*(?:(\d+)s)?\s*$")


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description="Check which courses are completed from attendance CSV folders"
    )
    parser.add_argument(
        "-c", "--config",
        default="config.yaml",
        help="Path to configuration file (default: config.yaml)",
    )
    parser.add_argument(
        "--csv-dir",
        help="Optional CSV root directory override (default: output.csv_directory from config)",
    )
    parser.add_argument(
        "--team-dirs-csv",
        default="team_dirs.csv",
        help="Path to team_dirs.csv (default: team_dirs.csv)",
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose logging",
    )
    return parser.parse_args()


def setup_logging(verbose: bool = False) -> None:
    """Configure logging output."""
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def parse_email_list(raw_value: str) -> set[str]:
    """Parse comma-separated email list into a normalized set."""
    if not raw_value:
        return set()
    return {
        email.strip().lower()
        for email in raw_value.split(",")
        if email.strip()
    }


def load_course_teacher_emails(team_dirs_csv: Path) -> dict[str, set[str]]:
    """Load teacher emails per course directory from team_dirs.csv."""
    content = read_text_with_fallbacks(team_dirs_csv)
    reader = csv.DictReader(io.StringIO(content))

    expected_columns = {"directory", "team_owner", "additional_teachers"}
    if not reader.fieldnames or not expected_columns.issubset(set(reader.fieldnames)):
        raise ValueError(
            f"Missing required columns in {team_dirs_csv}: {sorted(expected_columns)}"
        )

    mapping: dict[str, set[str]] = {}
    for row in reader:
        directory_name = (row.get("directory") or "").strip()
        if not directory_name:
            continue

        teacher_emails: set[str] = set()
        team_owner = (row.get("team_owner") or "").strip().lower()
        if team_owner:
            teacher_emails.add(team_owner)

        teacher_emails.update(parse_email_list(row.get("additional_teachers") or ""))

        if teacher_emails:
            mapping[directory_name] = teacher_emails

    return mapping


def parse_expected_hours(directory_name: str) -> int | None:
    """Extract expected hours from the trailing [N] part of a directory name."""
    match = EXPECTED_HOURS_RE.search(directory_name)
    if not match:
        return None
    return int(match.group(1))


def parse_duration_to_seconds(raw_value: str) -> int:
    """Convert a duration formatted as HH:MM:SS or 'N day(s), HH:MM:SS' to seconds."""
    text = raw_value.strip()
    match = DURATION_RE.match(text)
    if match:
        days = int(match.group(1) or 0)
        hours = int(match.group(2))
        minutes = int(match.group(3))
        seconds = int(match.group(4))
        return (((days * 24) + hours) * 60 + minutes) * 60 + seconds

    hms_match = HMS_TEXT_RE.match(text)
    if hms_match and any(part is not None for part in hms_match.groups()):
        hours = int(hms_match.group(1) or 0)
        minutes = int(hms_match.group(2) or 0)
        seconds = int(hms_match.group(3) or 0)
        return ((hours * 60) + minutes) * 60 + seconds

    raise ValueError(f"Unsupported duration format: {raw_value!r}")


def read_text_with_fallbacks(path: Path) -> str:
    """Read text trying common encodings used by exported attendance CSV files."""
    raw_bytes = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-16", "cp1252"):
        try:
            return raw_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue

    raise UnicodeDecodeError("unknown", b"", 0, 1, f"Cannot decode file {path}")


def extract_max_teacher_duration(csv_file: Path, teacher_emails: set[str]) -> int:
    """
    Extract the max 'In-Meeting Duration' among teachers present in a CSV file.

    A CSV is considered valid when at least one participant email matches a teacher email.
    """
    content = read_text_with_fallbacks(csv_file)
    delimiters = ("\t", ",", ";")
    found_participants_section = False

    for delimiter in delimiters:
        reader = csv.reader(io.StringIO(content), delimiter=delimiter)
        rows = list(reader)

        # Find the section starting with "2." in the first column (language-independent)
        participants_idx = None
        for i, row in enumerate(rows):
            if row and row[0].strip().startswith("2."):
                participants_idx = i
                break

        if participants_idx is None:
            continue  # Try next delimiter
        found_participants_section = True

        if participants_idx + 1 >= len(rows):
            raise ValueError("Participants section header found but no rows follow")

        # Columns are at fixed positions in the standard export format:
        # 0: Name, 1: First Join, 2: Last Leave, 3: In-Meeting Duration, 4: Email, 5: Participant ID, 6: Role
        email_idx = 4
        duration_idx = 3
        matched_durations: list[int] = []

        # Search all teacher participants in this CSV and keep the maximum duration.
        for row in rows[participants_idx + 2:]:
            if not any(cell.strip() for cell in row):
                break
            if len(row) <= max(email_idx, duration_idx):
                continue

            email = row[email_idx].strip().lower()
            if email in teacher_emails:
                try:
                    duration_str = row[duration_idx].strip()
                    duration_secs = parse_duration_to_seconds(duration_str)
                    matched_durations.append(duration_secs)
                except Exception as exc:
                    raise ValueError(f"Could not parse duration '{duration_str}': {exc}")

        if matched_durations:
            return max(matched_durations)

    if not found_participants_section:
        raise ValueError("Could not find Participants section in any delimiter variant")

    raise ValueError("No teacher found in Participants section")


def format_seconds(total_seconds: int) -> str:
    """Format seconds as HH:MM:SS with total hours (can be > 24)."""
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:02}:{minutes:02}:{seconds:02}"


def evaluate_courses(
    csv_root_dir: Path,
    course_teacher_emails: dict[str, set[str]],
    tolerance_seconds: int = 0,
) -> tuple[list[dict], list[dict]]:
    """Evaluate all course directories and split them into completed/non-completed.

    For each course, considers as teachers the emails from team_owner/additional_teachers,
    and for each CSV sums the maximum duration among teachers present.

    Args:
        csv_root_dir: Root directory containing course subdirectories
        course_teacher_emails: Mapping from course directory name to teacher emails
        tolerance_seconds: Tolerance in seconds for completion threshold (default: 0)
                          If set to 600, a course requiring 6 hours is considered complete
                          with 5h 50m or more.
    """
    completed: list[dict] = []
    not_completed: list[dict] = []

    for course_dir in sorted(path for path in csv_root_dir.iterdir() if path.is_dir()):
        expected_hours = parse_expected_hours(course_dir.name)
        if expected_hours is None:
            LOGGER.warning("Skipping %s: expected hours [N] not found", course_dir.name)
            continue

        teacher_emails = course_teacher_emails.get(course_dir.name, set())
        if not teacher_emails:
            LOGGER.warning("No teachers configured in team_dirs.csv for %s", course_dir.name)

        total_seconds = 0
        parsed_files = 0
        csv_files = sorted(course_dir.glob("*.csv"))

        if teacher_emails:
            for csv_file in csv_files:
                try:
                    file_duration_seconds = extract_max_teacher_duration(csv_file, teacher_emails)
                    total_seconds += file_duration_seconds
                    parsed_files += 1
                except Exception as exc:
                    LOGGER.warning("Skipping %s: %s", csv_file, exc)

        expected_seconds = expected_hours * 3600
        completion_threshold = expected_seconds - tolerance_seconds

        course_info = {
            "name": course_dir.name,
            "expected_hours": expected_hours,
            "expected_seconds": expected_seconds,
            "total_seconds": total_seconds,
            "parsed_files": parsed_files,
            "total_files": len(csv_files),
        }

        # Consider the course completed when total >= (expected - tolerance).
        if total_seconds >= completion_threshold:
            completed.append(course_info)
        else:
            not_completed.append(course_info)

    return completed, not_completed


def build_sharepoint_uploader(config) -> SharePointCSVUploader | None:
    """Build SharePoint uploader client for folder-status checks."""
    sharepoint_config = config.output.sharepoint_csv

    client_secret = sharepoint_config.auth.client_secret or os.getenv("TEAMS_HARVESTER_CLIENT_SECRET")
    cache_dir_path = config.cache.directory
    cache_dir_path.mkdir(parents=True, exist_ok=True)

    authenticator = Authenticator(
        client_id=sharepoint_config.auth.client_id,
        authority=sharepoint_config.auth.authority,
        scopes=sharepoint_config.auth.scopes,
        cache_path=cache_dir_path / sharepoint_config.auth.token_cache,
        auth_mode=sharepoint_config.auth.mode,
        client_secret=client_secret,
    )
    if sharepoint_config.auth.clear_cache:
        authenticator.clear_cache()

    access_token = authenticator.acquire_token()
    graph_client = GraphClient(
        access_token=access_token,
        max_retries=config.api.max_retries,
        retry_backoff_factor=config.api.retry_backoff_factor,
        timeout=config.api.timeout,
    )
    return SharePointCSVUploader(
        graph_client=graph_client,
        site_id=sharepoint_config.site_id,
        site_hostname=sharepoint_config.site_hostname,
        site_path=sharepoint_config.site_path,
        drive_id=sharepoint_config.drive_id,
        drive_name=sharepoint_config.drive_name,
        folder_path=sharepoint_config.folder_path,
    )


def get_processed_courses_on_sharepoint(uploader: SharePointCSVUploader, course_names: list[str]) -> set[str]:
    """Return course names marked as processed on SharePoint via '[closed]' folder suffix."""
    drive_id, root_folder = uploader._resolve_root_folder()
    root_item = uploader._get_item_by_path(drive_id, root_folder)
    if not root_item:
        return set()

    children = uploader._list_children(drive_id, root_item["id"])
    child_folder_names = {
        str(child.get("name", "")).strip()
        for child in children
        if isinstance(child.get("folder"), dict)
    }

    processed: set[str] = set()
    for course_name in course_names:
        if f"{course_name} [closed]" in child_folder_names:
            processed.add(course_name)

    return processed


def print_section(title: str, courses: list[dict]) -> None:
    """Print one section of the final report."""
    print(title)
    print("-" * len(title))
    if not courses:
        print("(nessuno)")
        print()
        return

    for course in courses:
        print(
            f"- {course['name']}: "
            f"totale={format_seconds(course['total_seconds'])} "
            f"attese={course['expected_hours']}h "
            f"file={course['parsed_files']}/{course['total_files']}"
        )
    print()


def main() -> None:
    """CLI entry point."""
    args = parse_args()
    setup_logging(args.verbose)
    config = load_app_config(args.config, args)

    csv_root_dir = Path(args.csv_dir) if args.csv_dir else config.output.csv_directory
    if not csv_root_dir.exists() or not csv_root_dir.is_dir():
        raise FileNotFoundError(f"CSV directory not found: {csv_root_dir}")

    team_dirs_csv = Path(args.team_dirs_csv)
    if not team_dirs_csv.exists() or not team_dirs_csv.is_file():
        raise FileNotFoundError(f"team_dirs.csv not found: {team_dirs_csv}")

    course_teacher_emails = load_course_teacher_emails(team_dirs_csv)

    # Get tolerance from config
    tolerance_minutes = config.completion.tolerance_minutes
    tolerance_seconds = tolerance_minutes * 60

    if tolerance_minutes > 0:
        LOGGER.info("Using completion tolerance: %d minute(s) (%d second(s))", tolerance_minutes, tolerance_seconds)

    completed, not_completed = evaluate_courses(csv_root_dir, course_teacher_emails, tolerance_seconds)

    processed_course_names: set[str] = set()
    try:
        uploader = build_sharepoint_uploader(config)
        if uploader is not None:
            processed_course_names = get_processed_courses_on_sharepoint(
                uploader,
                [course["name"] for course in completed],
            )
    except Exception as exc:
        LOGGER.warning("SharePoint processed-check unavailable: %s", exc)

    completed_processed = [course for course in completed if course["name"] in processed_course_names]
    completed_not_processed = [course for course in completed if course["name"] not in processed_course_names]

    print()
    print(f"Directory CSV analizzata: {csv_root_dir}")
    print(f"Totale corsi valutati: {len(completed) + len(not_completed)}")
    if tolerance_minutes > 0:
        print(f"Tolleranza di completamento: {tolerance_minutes} minuti")
    print()
    print_section("Corsi terminati e processati", completed_processed)
    print_section("Corsi terminati e NON processati", completed_not_processed)
    print_section("Corsi non terminati", not_completed)


if __name__ == "__main__":
    main()