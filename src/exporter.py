"""
Export attendance data to CSV and JSON formats.
"""
import csv
import json
import logging
import re
from datetime import datetime, timedelta
from pathlib import Path

logger = logging.getLogger(__name__)


class AttendanceExporter:
    """Exports attendance data to various formats."""

    def __init__(
        self,
        output_dir: str = "./output",
        filename_pattern: str | None = None,
        csv_output_dir: str | None = None,
        json_output_dir: str | None = None,
        min_csv_report_duration_seconds: int = 0
    ):
        """
        Initialize exporter.

        Args:
            output_dir: Base directory for output files
            filename_pattern: Pattern for output filenames
            csv_output_dir: Directory for CSV exports
            json_output_dir: Directory for JSON exports
            min_csv_report_duration_seconds: Minimum report duration in seconds
                required before exporting a CSV
        """
        base_output_dir = Path(output_dir)
        self.csv_output_dir = Path(csv_output_dir) if csv_output_dir else base_output_dir / "csv"
        self.json_output_dir = Path(json_output_dir) if json_output_dir else base_output_dir / "json"
        self.csv_output_dir.mkdir(parents=True, exist_ok=True)
        self.json_output_dir.mkdir(parents=True, exist_ok=True)
        self.filename_pattern = filename_pattern or "{team_name}_{channel_name}_{meeting_date}_{meeting_id}_{report_id}_attendance"
        self.min_csv_report_duration_seconds = max(0, int(min_csv_report_duration_seconds))

    @staticmethod
    def _sanitize_filename(name: str) -> str:
        """
        Sanitize string for use in filename.

        Args:
            name: String to sanitize

        Returns:
            Sanitized string safe for filenames
        """
        # Replace invalid characters
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            name = name.replace(char, '_')

        # Replace spaces with underscores
        name = name.replace(' ', '_')

        # Remove leading/trailing dots and spaces
        name = name.strip('. ')

        # Truncate if too long
        if len(name) > 100:
            name = name[:100]

        return name

    def _build_filename(self, attendance_data: dict) -> str:
        """
        Build filename from attendance data and pattern.

        Args:
            attendance_data: Attendance data dictionary

        Returns:
            Sanitized filename (without extension)
        """
        meeting_info = attendance_data.get("meeting_info", {})

        # Extract data for filename
        team_name = "unknown_team"
        channel_name = "unknown_channel"

        # Try to get team/channel from context
        teams_context = attendance_data.get("teams_context", [])
        if teams_context:
            team_name = teams_context[0].get("team", {}).get("displayName", "unknown_team")
            channel_name = teams_context[0].get("channel", {}).get("displayName", "unknown_channel")

        # Get meeting date
        meeting_start = meeting_info.get("start", {})
        if isinstance(meeting_start, dict):
            date_str = meeting_start.get("dateTime", "")
        else:
            date_str = str(meeting_start)

        try:
            dt = datetime.fromisoformat(date_str.replace("Z", "+00:00"))
            meeting_date = dt.strftime("%Y%m%d_%H%M")
        except:
            meeting_date = "unknown_date"

        meeting_id = attendance_data.get("meeting_id", "unknown_meeting")[:8]
        report_id = attendance_data.get("report_id", "unknown_report")[:8]

        # Build filename from pattern
        filename = self.filename_pattern.format(
            team_name=self._sanitize_filename(team_name),
            channel_name=self._sanitize_filename(channel_name),
            meeting_date=meeting_date,
            meeting_id=meeting_id,
            report_id=report_id
        )

        return filename

    def export_to_json(self, attendance_data: dict, filename: str | None = None) -> Path:
        """
        Export attendance data to JSON file.

        Args:
            attendance_data: Attendance data dictionary
            filename: Custom filename (without extension), or None to auto-generate

        Returns:
            Path to created file
        """
        if not filename:
            filename = self._build_filename(attendance_data)

        filepath = self.json_output_dir / f"{filename}.json"

        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(attendance_data, f, indent=2, ensure_ascii=False)

        logger.info(f"Exported JSON to {filepath}")
        return filepath

    def export_to_csv(self, attendance_data: dict, filename: str | None = None) -> Path | None:
        """
        Export attendance records to CSV file.

        Args:
            attendance_data: Attendance data dictionary
            filename: Custom filename (without extension), or None to auto-generate

        Returns:
            Path to created file
        """
        if not filename:
            filename = self._build_filename(attendance_data)

        meetingStartDateTime = datetime.fromisoformat(attendance_data["report_data"]["meetingStartDateTime"])
        meetingEndDateTime = datetime.fromisoformat(attendance_data["report_data"]["meetingEndDateTime"])
        duration = meetingEndDateTime - meetingStartDateTime

        if (
            self.min_csv_report_duration_seconds > 0
            and duration.total_seconds() < self.min_csv_report_duration_seconds
        ):
            logger.info(
                "Skipping CSV for %s: report duration %ss is below minimum %ss",
                filename,
                duration.total_seconds(),
                self.min_csv_report_duration_seconds
            )
            return None

        filepath = self.csv_output_dir / f"{filename}.csv"

        with open(filepath, "w") as file:
            participant_count = attendance_data["report_data"]["totalParticipantCount"]
            total_time = sum(sum(ai["durationInSeconds"] for ai in ar["attendanceIntervals"])
                             for ar in attendance_data["attendance_records"])
            avg_time = round(total_time / participant_count) if participant_count > 0 else 0

            writer = csv.writer(file, delimiter='\t')
            writer.writerow(['1. Summary'])
            writer.writerow(['Meeting title', attendance_data["meeting_info"]["subject"]])
            writer.writerow(['Attended participants', participant_count])
            writer.writerow(['Start time', self._format_datetime(meetingStartDateTime)])
            writer.writerow(['End time', self._format_datetime(meetingEndDateTime)])
            writer.writerow(['Meeting duration', self._format_timedelta(duration)])
            writer.writerow(['Average attendance time', self._format_timedelta(timedelta(seconds=avg_time))])
            writer.writerow([])

            writer.writerow(['2. Participants'])
            writer.writerow(['Name', 'First Join', 'Last Leave', 'In-Meeting Duration',
                            'Email', 'Participant ID (UPN)', 'Role'])

            for ar in attendance_data["attendance_records"]:
                display_name = self._format_displayname(
                    ar["identity"], attendance_data["teams_context"][0]["team"]["tenantId"])
                email_address = ar["emailAddress"]
                upn = ar["emailAddress"]
                role = ar["role"]
                first_join = min(datetime.fromisoformat(ai["joinDateTime"]) for ai in ar["attendanceIntervals"])
                last_leave = max(datetime.fromisoformat(ai["leaveDateTime"]) for ai in ar["attendanceIntervals"])
                duration = timedelta(seconds=sum(ai["durationInSeconds"] for ai in ar["attendanceIntervals"]))
                writer.writerow([display_name, self._format_datetime(first_join), self._format_datetime(last_leave),
                                self._format_timedelta(duration), email_address,  upn, role])

            writer.writerow([])
            writer.writerow(['3. In-Meeting Activities'])
            writer.writerow(['Name', 'Join Time', 'Leave Time', 'Duration', 'Email', 'Role'])
            for ar in attendance_data["attendance_records"]:
                for ai in ar["attendanceIntervals"]:
                    display_name = self._format_displayname(
                        ar["identity"], attendance_data["teams_context"][0]["team"]["tenantId"])
                    email_address = ar["emailAddress"]
                    join_date_time = datetime.fromisoformat(ai["joinDateTime"])
                    leave_date_time = datetime.fromisoformat(ai["leaveDateTime"])
                    duration = timedelta(seconds=ai["durationInSeconds"])
                    role = ar["role"]
                    writer.writerow([display_name, self._format_datetime(join_date_time), self._format_datetime(leave_date_time),
                                    self._format_timedelta(duration), email_address, role])

        return filepath

    def export_batch(self, attendance_list: list[dict], format: str = "both") -> list[Path]:
        """
        Export multiple attendance data records.

        Args:
            attendance_list: List of attendance data dictionaries
            format: Export format - "csv", "json", or "both"

        Returns:
            List of created file paths
        """
        created_files = []

        for attendance_data in attendance_list:
            filename = self._build_filename(attendance_data)

            if format in ("json", "both"):
                json_path = self.export_to_json(attendance_data, filename)
                if json_path:
                    created_files.append(json_path)

            if format in ("csv", "both"):
                csv_path = self.export_to_csv(attendance_data, filename)
                if csv_path:
                    created_files.append(csv_path)

        logger.info(f"Exported {len(created_files)} files")
        return created_files

    @staticmethod
    def _format_datetime(dt: datetime) -> str:
        """Format datetime value to string."""
        return dt.astimezone().strftime("%-m/%d/%y, %-I:%M:%S %p")

    @staticmethod
    def _format_timedelta(td: timedelta) -> str:
        """Format timedelta value to string."""
        formatted = []
        total_seconds = int(td.total_seconds())
        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        if hours > 0:
            formatted.append(f"{hours}h")
        if minutes > 0:
            formatted.append(f"{minutes}m")
        if seconds > 0 or formatted == "":
            formatted.append(f"{seconds}s")
        return " ".join(formatted)

    @staticmethod
    def _format_displayname(identity: dict, team_id: str):
        name = identity["displayName"]
        id = identity["id"]
        if id.startswith("guest:"):
            return name + " (Non verificato)"
        if not re.match(r'^[0-9a-f-]+$', identity["id"]):
            return name
        if identity["tenantId"] is None:
            return name + " (Non verificato)"
        if identity["tenantId"] != team_id:
            return name + " (Esterno)"
        return name
        return name
