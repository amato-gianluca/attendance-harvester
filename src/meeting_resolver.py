"""
Meeting discovery and attendance extraction logic.
"""
import json
import logging
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

from .graph_client import GraphClient

logger = logging.getLogger(__name__)


class MeetingResolver:
    """Discovers meetings and extracts attendance reports."""

    def __init__(self, graph_client: GraphClient, checkpoint_file: str | None = None):
        """
        Initialize meeting resolver.

        Args:
            graph_client: Graph API client instance
            checkpoint_file: Path to checkpoint file for tracking processed meetings
        """
        self.client = graph_client
        self.checkpoint_file = Path(checkpoint_file) if checkpoint_file else None
        self.processed_meetings = self._load_checkpoints()

    def _load_checkpoints(self) -> set[str]:
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

    def get_meetings_in_date_range(self, lookback_days: int, lookahead_days: int = 0) -> list[dict]:
        """
        Get all Teams meetings in the specified date range.

        Args:
            lookback_days: Number of days to look back
            lookahead_days: Number of days to look ahead

        Returns:
            list of calendar event objects for Teams meetings
        """
        now = datetime.now(timezone.utc)
        start_time = now - timedelta(days=max(0, lookback_days))
        end_time = now + timedelta(days=max(0, lookahead_days))

        start_str = start_time.isoformat()
        end_str = end_time.isoformat()

        logger.info(
            "Searching for meetings from %s to %s (lookback=%d days, lookahead=%d days)",
            start_str,
            end_str,
            lookback_days,
            lookahead_days
        )
        return self.client.get_calendar_events(start_str, end_str)

    def resolve_online_meeting(self, event: dict) -> dict | None:
        """
        Resolve calendar event to online meeting object.

        For meetings you organize, we can look them up by join URL.
        For meetings you only attend, create a minimal online meeting object
        from calendar event data to enable attendance extraction attempts.

        Args:
            event: Calendar event object

        Returns:
            Online meeting object with ID, or None if unresolvable
        """
        online_meeting_info = event.get("onlineMeeting", {})
        join_url = online_meeting_info.get("joinUrl")

        if not join_url:
            logger.debug(
                f"No join URL for event: {event.get('subject', 'Unknown')}")
            return None

        # Try to get online meeting by join URL (works for organized meetings)
        online_meeting = self.client.get_online_meeting_by_join_url(join_url)

        if online_meeting:
            # Enrich with event details
            online_meeting["_event"] = {
                "subject": event.get("subject"),
                "start": event.get("start"),
                "end": event.get("end"),
                "organizer": event.get("organizer")
            }
            return online_meeting

        # Fallback: create minimal meeting object from calendar event
        # This allows attendance extraction to be attempted (will get 404 if not organized by user)
        logger.debug(
            "Could not resolve online meeting by join URL for '%s'. "
            "This is normal for meetings you don't organize.",
            event.get("subject", "Unknown")
        )

        # Construct a minimal online meeting object
        minimal_meeting = {
            "joinWebUrl": join_url,
            "chatInfo": {
                "threadId": self._extract_thread_id_from_join_url(join_url)
            },
            "_event": {
                "subject": event.get("subject"),
                "start": event.get("start"),
                "end": event.get("end"),
                "organizer": event.get("organizer")
            }
        }

        # Try to extract meeting ID from event ID if available
        # Teams meeting IDs sometimes embed the calendar event ID
        if "id" in event:
            minimal_meeting["_calendar_event_id"] = event["id"]

        return minimal_meeting

    def _extract_join_url(self, event: dict, online_meeting: dict | None = None) -> str:
        """Extract join URL from event or resolved online meeting."""
        event_url = event.get("onlineMeeting", {}).get("joinUrl")
        if event_url:
            return event_url

        if online_meeting:
            return online_meeting.get("joinWebUrl", "")

        return ""

    def _extract_thread_id_from_join_url(self, join_url: str) -> str:
        """Try to extract channel thread ID from Teams join URL context query parameter."""
        if not join_url:
            return ""

        try:
            parsed = urlparse(join_url)
            query = parse_qs(parsed.query)
            context_values = query.get("context")
            if not context_values:
                return ""

            raw_context = context_values[0]
            # Query value may be URL encoded JSON.
            decoded_context = unquote(raw_context)
            context_obj = json.loads(decoded_context)

            thread_id = context_obj.get("ThreadId") or context_obj.get("threadId")
            return thread_id if isinstance(thread_id, str) else ""
        except Exception:
            return ""

    def _build_event_haystack(self, event: dict, online_meeting: dict | None) -> str:
        """Build normalized text blob used for fuzzy team/channel matching."""
        parts = [
            event.get("subject", ""),
            event.get("bodyPreview", ""),
            event.get("location", {}).get("displayName", ""),
            " ".join(loc.get("displayName", "")
                     for loc in event.get("locations", [])),
            self._extract_join_url(event, online_meeting),
        ]
        return " ".join(p for p in parts if isinstance(p, str)).lower()

    def _match_event_contexts(self, event: dict, online_meeting: dict,
                              teams_with_channels: list[dict]) -> list[dict]:
        """
        Return matching team/channel contexts for the meeting.

        Matching strategy (best effort):
        1) Exact channel thread id match from meeting/chat or join URL context.
        2) Team/channel display name fuzzy match in event metadata.
        """
        if not teams_with_channels:
            return []

        meeting_thread_id = online_meeting.get(
            "chatInfo", {}).get("threadId", "")
        join_url = self._extract_join_url(event, online_meeting)
        url_thread_id = self._extract_thread_id_from_join_url(join_url)
        haystack = self._build_event_haystack(event, online_meeting)

        matched: list[dict] = []

        for context in teams_with_channels:
            team = context.get("team", {})
            channel = context.get("channel", {})

            team_id = str(team.get("id", ""))
            team_name = str(team.get("displayName", "")).lower()
            channel_id = str(channel.get("id", ""))
            channel_name = str(channel.get("displayName", "")).lower()

            id_match = bool(channel_id) and (
                channel_id == meeting_thread_id or channel_id == url_thread_id
            )

            name_match = bool(team_name and channel_name) and (
                team_name in haystack and channel_name in haystack
            )

            # Team-only fallback (useful when channel metadata is absent)
            team_only_match = bool(team_name) and team_name in haystack

            # URL fallback for identifiers
            identifier_in_url = bool(join_url) and (
                (channel_id and channel_id in join_url) or
                (team_id and team_id in join_url)
            )

            if id_match or name_match or (team_only_match and identifier_in_url):
                matched.append(context)

        # De-duplicate matches by team+channel id pair
        deduped: list[dict] = []
        seen = set()
        for context in matched:
            team_id = context.get("team", {}).get("id", "")
            channel_id = context.get("channel", {}).get("id", "")
            key = f"{team_id}:{channel_id}"
            if key not in seen:
                seen.add(key)
                deduped.append(context)

        return deduped

    def get_meeting_attendance(self, online_meeting: dict, skip_processed: bool = True) -> list[dict]:
        """
        Get all attendance reports and records for a meeting.

        Args:
            online_meeting: Online meeting object
            skip_processed: Skip meetings already processed (from checkpoint)

        Returns:
            list of dictionaries with report and records data
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
            logger.info(
                f"No attendance reports for meeting: {online_meeting.get('_event', {}).get('subject', meeting_id)}")
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

            logger.info(
                f"Extracted {len(records)} attendance records from report {report_id}")

        # Mark as processed
        self.processed_meetings.add(checkpoint_key)
        self._save_checkpoints()

        return attendance_data

    def extract_all_attendance(self, teams_with_channels: list[dict],
                               lookback_days: int,
                               lookahead_days: int = 0) -> list[dict]:
        """
        Extract attendance for all meetings across teams.

        Args:
            teams_with_channels: list of dicts with 'team' and 'channel' keys
            lookback_days: Number of days to look back for meetings
            lookahead_days: Number of days to look ahead for meetings

        Returns:
            list of attendance data dictionaries
        """
        # Get all meetings in date range
        all_events = self.get_meetings_in_date_range(
            lookback_days, lookahead_days)

        if not all_events:
            logger.warning(
                "No Teams meetings found in the specified date range")
            return []

        all_attendance = []
        matched_events = 0

        # Process each meeting
        for event in all_events:
            event_subject = event.get("subject", "Unknown")
            logger.info(f"Processing meeting: {event_subject}")

            # Resolve to online meeting
            online_meeting = self.resolve_online_meeting(event)
            if not online_meeting:
                continue


            matched_contexts = self._match_event_contexts(
                event, online_meeting, teams_with_channels)
            if not matched_contexts:
                logger.debug(f"Skipping meeting outside filtered teams/channels: {event_subject}")
                continue

            matched_events += 1

            logger.info("")
            # Get attendance data
            attendance_list = self.get_meeting_attendance(online_meeting)

            # Enrich with matched team/channel context
            for attendance in attendance_list:
                attendance["teams_context"] = matched_contexts
                all_attendance.append(attendance)

        logger.info(
            "Extracted attendance data for %d reports across %d matched meetings (out of %d calendar meetings)",
            len(all_attendance),
            matched_events,
            len(all_events)
        )
        return all_attendance
