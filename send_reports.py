#!/usr/bin/env python3
"""
Send registro.xlsx files found in SharePoint folders and mark them as sent.
"""
from __future__ import annotations

import argparse
import csv
import logging
import smtplib
from email.message import EmailMessage
from pathlib import Path, PurePosixPath

from main import build_sharepoint_csv_uploader, setup_logging
from src.app_config import AppConfig, load_app_config


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments for report delivery."""
    parser = argparse.ArgumentParser(
        description="Send registro.xlsx files found in SharePoint team folders"
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
        "--dry-run",
        action="store_true",
        help="Scan and log the actions without sending mail or creating SENT markers"
    )
    parser.add_argument("--team-regex")
    parser.add_argument("--lookback-days", type=int)
    parser.add_argument("--lookahead-days", type=int)
    parser.add_argument("--min-csv-report-duration-seconds", type=int)
    return parser.parse_args()


def load_team_directory_rows(team_dirs_file: str) -> dict[str, dict[str, str]]:
    """Load team directory metadata keyed by SharePoint folder name."""
    with open(team_dirs_file, "r", encoding="utf-8", newline="") as csv_file:
        reader = csv.DictReader(csv_file)
        rows: dict[str, dict[str, str]] = {}
        for row in reader:
            directory = (row.get("directory") or "").strip()
            if directory:
                rows[directory] = {key: (value or "").strip() for key, value in row.items() if key}
        return rows


def parse_email_list(raw_value: str) -> list[str]:
    """Split a comma/semicolon separated list of emails."""
    normalized = raw_value.replace("||", ",").replace("|", ",").replace(";", ",")
    return [part.strip() for part in normalized.split(",") if part.strip()]


def render_message(template_file: Path, team_name: str) -> str:
    """Render the mail template for the given team."""
    template = template_file.read_text(encoding="utf-8")
    return template.replace("{team_name}", team_name)


def get_team_directory_name(base_folder: str, parent_relative_path: PurePosixPath) -> str:
    """Extract the top-level team folder name from a SharePoint relative path."""
    root = PurePosixPath(base_folder)
    try:
        relative_path = parent_relative_path.relative_to(root)
    except ValueError:
        relative_path = parent_relative_path

    if not relative_path.parts:
        return ""

    return relative_path.parts[0]


def build_email_message(
    *,
    sender: str,
    to_recipients: list[str],
    cc_recipients: list[str],
    bcc_recipients: list[str],
    subject: str,
    body_text: str,
    attachment_name: str,
    attachment_content: bytes,
) -> EmailMessage:
    """Build the outbound email with attachment."""
    message = EmailMessage()
    message["From"] = sender
    message["To"] = ", ".join(to_recipients)
    if cc_recipients:
        message["Cc"] = ", ".join(cc_recipients)
    if bcc_recipients:
        message["Bcc"] = ", ".join(bcc_recipients)
    message["Subject"] = subject
    message.set_content(body_text)
    message.add_attachment(
        attachment_content,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=attachment_name,
    )
    return message


def send_email_via_smtp(config: AppConfig, message: EmailMessage, recipients: list[str]) -> None:
    """Send a prepared email via SMTP."""
    email_config = config.reports_email

    if email_config.smtp_ssl:
        smtp: smtplib.SMTP = smtplib.SMTP_SSL(email_config.smtp_hostname, email_config.smtp_port, timeout=config.api.timeout)
    else:
        smtp = smtplib.SMTP(email_config.smtp_hostname, email_config.smtp_port, timeout=config.api.timeout)

    try:
        smtp.ehlo()
        if email_config.smtp_starttls:
            smtp.starttls()
            smtp.ehlo()

        if email_config.smtp_username:
            smtp.login(email_config.smtp_username, email_config.smtp_password or "")

        smtp.send_message(message, to_addrs=recipients)
    finally:
        try:
            smtp.quit()
        except Exception:
            smtp.close()


def run_send_reports(config: AppConfig, *, dry_run: bool = False) -> None:
    """Scan SharePoint folders for registro.xlsx files and send them by email."""
    logger = logging.getLogger(__name__)

    uploader = build_sharepoint_csv_uploader(config, force_enable=True)
    team_directories_file = config.output.team_directories_file
    if not uploader:
        raise ValueError("SharePoint CSV upload could not be initialized from output.sharepoint_csv")
    if not team_directories_file:
        raise ValueError("output.team_directories_file is required")

    team_rows = load_team_directory_rows(team_directories_file)
    template_file = config.reports_email.template_file
    if not template_file.exists():
        raise FileNotFoundError(f"Mail template file not found: {template_file}")

    attachment_name = config.reports_email.attachment_filename
    sent_marker_name = config.reports_email.sent_marker_filename
    report_files = uploader.find_files_by_name(attachment_name)

    logger.info("Found %d SharePoint file(s) named %s", len(report_files), attachment_name)
    if not report_files:
        return

    processed = 0
    skipped = 0

    for report_file in report_files:
        parent_relative_path = report_file["_parent_relative_path"]
        if uploader.folder_contains_name(parent_relative_path, sent_marker_name):
            logger.info("Skipping %s because %s already exists", report_file["_relative_path"], sent_marker_name)
            skipped += 1
            continue

        team_directory = get_team_directory_name(
            uploader.folder_path,
            parent_relative_path
        )
        team_row = team_rows.get(team_directory)
        if not team_row:
            logger.warning("Skipping %s: no team_dirs.csv row found for directory %s", report_file["_relative_path"], team_directory)
            skipped += 1
            continue

        team_name = team_row.get("team_displayname", "").strip()
        team_owner_emails = parse_email_list(team_row.get("team_owner", ""))
        additional_emails = parse_email_list(team_row.get("additional_email", ""))

        if not team_name:
            logger.warning("Skipping %s: team_displayname is empty", report_file["_relative_path"])
            skipped += 1
            continue
        if not team_owner_emails:
            logger.warning("Skipping %s: team_owner email is missing for team %s", report_file["_relative_path"], team_name)
            skipped += 1
            continue

        subject = config.reports_email.subject_template.format(team_name=team_name)
        body_text = render_message(template_file, team_name)

        logger.info(
            "Preparing report email for %s -> to=%s cc=%s",
            team_name,
            ", ".join(team_owner_emails),
            ", ".join(additional_emails) if additional_emails else "-"
        )

        if dry_run:
            processed += 1
            continue

        attachment_content = uploader.download_file_content(report_file["_relative_path"])
        message = build_email_message(
            sender=config.reports_email.sender,
            to_recipients=team_owner_emails,
            cc_recipients=additional_emails,
            bcc_recipients=config.reports_email.bcc_recipients,
            subject=subject,
            body_text=body_text,
            attachment_name=attachment_name,
            attachment_content=attachment_content,
        )
        send_email_via_smtp(
            config,
            message,
            team_owner_emails + additional_emails + config.reports_email.bcc_recipients
        )
        uploader.create_empty_file(parent_relative_path, sent_marker_name)
        logger.info("Sent report for %s and created %s marker", team_name, sent_marker_name)
        processed += 1

    logger.info("=" * 70)
    logger.info("SUMMARY")
    logger.info("=" * 70)
    logger.info("Report files found: %d", len(report_files))
    logger.info("Processed: %d", processed)
    logger.info("Skipped: %d", skipped)
    logger.info("Dry-run: %s", dry_run)
    logger.info("=" * 70)


def main() -> None:
    """CLI entry point."""
    args = parse_args()
    setup_logging(args.verbose)
    config = load_app_config(args.config, args)
    run_send_reports(config, dry_run=args.dry_run)


if __name__ == "__main__":
    main()
