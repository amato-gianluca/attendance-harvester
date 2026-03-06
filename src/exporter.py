"""
Export attendance data to CSV and JSON formats.
"""
import csv
import json
import logging
from datetime import datetime
from pathlib import Path
from typing import Dict, List

logger = logging.getLogger(__name__)


class AttendanceExporter:
    """Exports attendance data to various formats."""

    def __init__(self, output_dir: str = "./output", filename_pattern: str = None):
        """
        Initialize exporter.

        Args:
            output_dir: Directory for output files
            filename_pattern: Pattern for output filenames
        """
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.filename_pattern = filename_pattern or "{team_name}_{channel_name}_{meeting_date}_{meeting_id}_{report_id}_attendance"

    def _sanitize_filename(self, name: str) -> str:
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

    def _build_filename(self, attendance_data: Dict) -> str:
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

    def export_to_json(self, attendance_data: Dict, filename: str = None) -> Path:
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

        filepath = self.output_dir / f"{filename}.json"

        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(attendance_data, f, indent=2, ensure_ascii=False)

        logger.info(f"Exported JSON to {filepath}")
        return filepath

    def export_to_csv(self, attendance_data: Dict, filename: str = None) -> Path:
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

        filepath = self.output_dir / f"{filename}.csv"

        records = attendance_data.get("attendance_records", [])

        if not records:
            logger.warning(f"No attendance records to export for {filename}")
            return None

        # Flatten records for CSV
        flattened_records = []
        for record in records:
            flat_record = {
                "email": record.get("emailAddress", ""),
                "display_name": record.get("identity", {}).get("displayName", ""),
                "role": record.get("role", ""),
                "total_attendance_duration": record.get("totalAttendanceInSeconds", 0),
                "join_datetime": self._format_datetime(record.get("attendanceIntervals", [{}])[0].get("joinDateTime")),
                "leave_datetime": self._format_datetime(record.get("attendanceIntervals", [{}])[0].get("leaveDateTime")),
            }

            # Add meeting info
            meeting_info = attendance_data.get("meeting_info", {})
            flat_record.update({
                "meeting_subject": meeting_info.get("subject", ""),
                "meeting_start": self._format_datetime(meeting_info.get("start", {}).get("dateTime")),
                "meeting_organizer": meeting_info.get("organizer", {}).get("emailAddress", {}).get("address", ""),
            })

            flattened_records.append(flat_record)

        # Write CSV
        if flattened_records:
            fieldnames = flattened_records[0].keys()
            with open(filepath, "w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(flattened_records)

            logger.info(f"Exported {len(flattened_records)} records to {filepath}")

        return filepath

    def _format_datetime(self, dt_value) -> str:
        """Format datetime value to string."""
        if not dt_value:
            return ""
        if isinstance(dt_value, str):
            return dt_value
        if isinstance(dt_value, dict):
            return dt_value.get("dateTime", "")
        return str(dt_value)

    def export_batch(self, attendance_list: List[Dict], format: str = "both") -> List[Path]:
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
