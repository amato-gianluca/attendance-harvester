"""
Meeting discovery and attendance extraction logic.
"""
import json
import logging
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Dict, List, Optional, Set

from .graph_client import GraphClient

logger = logging.getLogger(__name__)


class MeetingResolver:
    """Discovers meetings and extracts attendance reports."""

    def __init__(self, graph_client: GraphClient, checkpoint_file: Optional[str] = None):
        """
        Initialize meeting resolver.

        Args:
            graph_client: Graph API client instance
            checkpoint_file: Path to checkpoint file for tracking processed meetings
        """
        self.client = graph_client
        self.checkpoint_file = Path(checkpoint_file) if checkpoint_file else None
        self.processed_meetings = self._load_checkpoints()

    def _load_checkpoints(self) -> Set[str]:
        """Load processed meeting IDs from checkpoint file."""
        if not self.checkpoint_file or not self.checkpoint_file.exists():
            return set()

        try:
            with open(self.checkpoint_file, "r") as f:
                data = json.load(f)
                return set(data.get("processed_meetings", []))
        except Exception as e:
            logger.warning(f"Failed to load checkpoints: {e}")
            return set()

    def _save_checkpoints(self):
        """Save processed meeting IDs to checkpoint file."""
        if not self.checkpoint_file:
            return

        try:
            self.checkpoint_file.parent.mkdir(parents=True, exist_ok=True)
            with open(self.checkpoint_file, "w") as f:
                json.dump({"processed_meetings": list(self.processed_meetings)}, f, indent=2)
            logger.debug(f"Saved {len(self.processed_meetings)} processed meetings to checkpoint")
        except Exception as e:
            logger.warning(f"Failed to save checkpoints: {e}")

    def get_meetings_in_date_range(self, lookback_days: int) -> List[Dict]:
        """
        Get all Teams meetings in the specified date range.

        Args:
            lookback_days: Number of days to look back

        Returns:
            List of calendar event objects for Teams meetings
        """
        end_time = datetime.now(timezone.utc)
        start_time = end_time - timedelta(days=lookback_days)

        start_str = start_time.isoformat()
        end_str = end_time.isoformat()

        logger.info(f"Searching for meetings from {start_str} to {end_str}")
        return self.client.get_calendar_events(start_str, end_str)

    def resolve_online_meeting(self, event: Dict) -> Optional[Dict]:
        """
        Resolve calendar event to online meeting object.

        Args:
            event: Calendar event object

        Returns:
            Online meeting object with ID, or None if not found
        """
        online_meeting_info = event.get("onlineMeeting", {})
        join_url = online_meeting_info.get("joinUrl")

        if not join_url:
            logger.warning(f"No join URL for event: {event.get('subject', 'Unknown')}")
            return None

        # Try to get online meeting by join URL
        online_meeting = self.client.get_online_meeting_by_join_url(join_url)

        if not online_meeting:
            logger.warning(f"Could not resolve online meeting for: {event.get('subject', 'Unknown')}")
            return None

        # Enrich with event details
        online_meeting["_event"] = {
            "subject": event.get("subject"),
            "start": event.get("start"),
            "end": event.get("end"),
            "organizer": event.get("organizer")
        }

        return online_meeting

    def get_meeting_attendance(self, online_meeting: Dict, skip_processed: bool = True) -> List[Dict]:
        """
        Get all attendance reports and records for a meeting.

        Args:
            online_meeting: Online meeting object
            skip_processed: Skip meetings already processed (from checkpoint)

        Returns:
            List of dictionaries with report and records data
        """
        meeting_id = online_meeting.get("id")
        if not meeting_id:
            logger.warning("Meeting has no ID, skipping")
            return []

        # Check if already processed
        checkpoint_key = f"{meeting_id}"
        if skip_processed and checkpoint_key in self.processed_meetings:
            logger.debug(f"Skipping already processed meeting: {meeting_id}")
            return []

        # Get attendance reports
        reports = self.client.get_attendance_reports(meeting_id)

        if not reports:
            logger.info(f"No attendance reports for meeting: {online_meeting.get('_event', {}).get('subject', meeting_id)}")
            return []

        attendance_data = []
        for report in reports:
            report_id = report.get("id")
            if not report_id:
                continue

            # Get attendance records for this report
            records = self.client.get_attendance_records(meeting_id, report_id)

            attendance_data.append({
                "meeting_id": meeting_id,
                "meeting_info": online_meeting.get("_event", {}),
                "report_id": report_id,
                "report_data": report,
                "attendance_records": records
            })

            logger.info(f"Extracted {len(records)} attendance records from report {report_id}")

        # Mark as processed
        self.processed_meetings.add(checkpoint_key)
        self._save_checkpoints()

        return attendance_data

    def extract_all_attendance(self, teams_with_channels: List[Dict],
                              lookback_days: int) -> List[Dict]:
        """
        Extract attendance for all meetings across teams.

        Args:
            teams_with_channels: List of dicts with 'team' and 'channel' keys
            lookback_days: Number of days to look back for meetings

        Returns:
            List of attendance data dictionaries
        """
        # Get all meetings in date range
        all_events = self.get_meetings_in_date_range(lookback_days)

        if not all_events:
            logger.warning("No Teams meetings found in the specified date range")
            return []

        all_attendance = []

        # Process each meeting
        for event in all_events:
            event_subject = event.get("subject", "Unknown")
            logger.info(f"Processing meeting: {event_subject}")

            # Resolve to online meeting
            online_meeting = self.resolve_online_meeting(event)
            if not online_meeting:
                continue

            # Get attendance data
            attendance_list = self.get_meeting_attendance(online_meeting)

            # Enrich with team/channel info if possible
            # Note: Matching meetings to specific teams/channels is complex
            # For now, we include all meetings from user's calendar
            for attendance in attendance_list:
                attendance["teams_context"] = teams_with_channels
                all_attendance.append(attendance)

        logger.info(f"Extracted attendance data for {len(all_attendance)} reports across {len(all_events)} meetings")
        return all_attendance
