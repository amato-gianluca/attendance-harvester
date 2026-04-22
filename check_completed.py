#!/usr/bin/env python3
"""Report completed courses based on instructor in-meeting duration across all CSV exports."""
from __future__ import annotations

import argparse
import csv
import io
import logging
import re
from pathlib import Path

from src.app_config import load_app_config


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


def extract_instructor_name(directory_name: str) -> str:
    """Extract instructor name from directory name, removing all bracketed suffixes.

    Example: "CLAUDIO CRIVELLARI [a] [24]" -> "CLAUDIO CRIVELLARI"
    """
    # Remove everything after and including the first [ character
    name = directory_name.split("[")[0].strip()
    return name


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


def extract_instructor_duration(csv_file: Path, instructor_name: str) -> int:
    """
    Extract the 'In-Meeting Duration' for a specific instructor from a CSV file.

    Matches the instructor name in the first column (case-insensitive).
    """
    content = read_text_with_fallbacks(csv_file)
    delimiters = ("\t", ",", ";")

    instructor_name_lower = instructor_name.lower()

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

        if participants_idx + 1 >= len(rows):
            raise ValueError("Participants section header found but no rows follow")

        # Columns are at fixed positions in the standard export format:
        # 0: Name, 1: First Join, 2: Last Leave, 3: In-Meeting Duration, 4: Email, 5: Participant ID, 6: Role
        name_idx = 0
        duration_idx = 3

        # Search for the instructor in this CSV
        for row in rows[participants_idx + 2:]:
            if not any(cell.strip() for cell in row):
                break
            if len(row) <= max(name_idx, duration_idx):
                continue

            name = row[name_idx].strip()
            if name.lower() == instructor_name_lower:
                try:
                    duration_str = row[duration_idx].strip()
                    duration_secs = parse_duration_to_seconds(duration_str)
                    return duration_secs
                except Exception as exc:
                    raise ValueError(f"Could not parse duration '{duration_str}': {exc}")

        # If we found the section but not the instructor, continue to next delimiter
        raise ValueError(f"Instructor '{instructor_name}' not found in Participants section")

    raise ValueError("Could not find Participants section in any delimiter variant")


def format_seconds(total_seconds: int) -> str:
    """Format seconds as HH:MM:SS with total hours (can be > 24)."""
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:02}:{minutes:02}:{seconds:02}"


def evaluate_courses(csv_root_dir: Path, tolerance_seconds: int = 0) -> tuple[list[dict], list[dict]]:
    """Evaluate all course directories and split them into completed/non-completed.

    For each course, finds the In-Meeting Duration of the course instructor (matched by name)
    across all CSV files in that course, taking the maximum.

    Args:
        csv_root_dir: Root directory containing course subdirectories
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

        instructor_name = extract_instructor_name(course_dir.name)

        total_seconds = 0
        parsed_files = 0
        csv_files = sorted(course_dir.glob("*.csv"))

        for csv_file in csv_files:
            try:
                file_duration_seconds = extract_instructor_duration(csv_file, instructor_name)
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

    # Get tolerance from config
    tolerance_minutes = config.completion.tolerance_minutes
    tolerance_seconds = tolerance_minutes * 60

    if tolerance_minutes > 0:
        LOGGER.info("Using completion tolerance: %d minute(s) (%d second(s))", tolerance_minutes, tolerance_seconds)

    completed, not_completed = evaluate_courses(csv_root_dir, tolerance_seconds)

    print()
    print(f"Directory CSV analizzata: {csv_root_dir}")
    print(f"Totale corsi valutati: {len(completed) + len(not_completed)}")
    if tolerance_minutes > 0:
        print(f"Tolleranza di completamento: {tolerance_minutes} minuti")
    print()
    print_section("Corsi terminati", completed)
    print_section("Corsi non terminati", not_completed)


if __name__ == "__main__":
    main()