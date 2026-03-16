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

    @staticmethod
    def _parse_datetime(value: dict | str | None) -> datetime | None:
        """Parse Graph datetime payloads into timezone-aware datetimes."""
        if isinstance(value, dict):
            value = value.get("dateTime")

        if not value or not isinstance(value, str):
            return None

        normalized = value.replace("Z", "+00:00")
        try:
            dt = datetime.fromisoformat(normalized)
        except ValueError:
            return None

        if dt.tzinfo is None:
            return dt.replace(tzinfo=timezone.utc)

        return dt

    def _report_checkpoint_key(self, report_id: str) -> str:
        """Build checkpoint key for a report."""
        return f"report:{report_id}"

    def _get_context_key(self, online_meeting: dict, matched_contexts: list[dict]) -> str:
        """Return a stable grouping key for a channel-scoped meeting."""
        if matched_contexts:
            channel_id = matched_contexts[0].get("channel", {}).get("id")
            if channel_id:
                return f"channel:{channel_id}"

        meeting_thread_id = online_meeting.get("chatInfo", {}).get("threadId")
        if meeting_thread_id:
            return f"thread:{meeting_thread_id}"

        meeting_id = online_meeting.get("id")
        if meeting_id:
            return f"meeting:{meeting_id}"

        event_id = online_meeting.get("_calendar_event_id")
        if event_id:
            return f"event:{event_id}"

        return "unknown"

    def _get_meeting_time_bounds(self, meeting_candidate: dict) -> tuple[datetime | None, datetime | None]:
        """Return parsed start/end times for a meeting candidate."""
        meeting_info = meeting_candidate.get("meeting_info", {})
        return (
            self._parse_datetime(meeting_info.get("start")),
            self._parse_datetime(meeting_info.get("end"))
        )

    def _get_report_time_bounds(self, report: dict) -> tuple[datetime | None, datetime | None]:
        """Return parsed start/end times for an attendance report."""
        return (
            self._parse_datetime(report.get("meetingStartDateTime")),
            self._parse_datetime(report.get("meetingEndDateTime"))
        )

    def _select_best_meeting_for_report(self, report: dict, meeting_candidates: list[dict]) -> dict | None:
        """Assign a report to the nearest meeting in time."""
        if not meeting_candidates:
            return None

        report_start, report_end = self._get_report_time_bounds(report)
        if not report_start and not report_end:
            return meeting_candidates[0]

        def distance(candidate: dict) -> float:
            meeting_start, meeting_end = self._get_meeting_time_bounds(candidate)
            total = 0.0

            if report_start and meeting_start:
                total += abs((report_start - meeting_start).total_seconds())
            elif report_start or meeting_start:
                total += 10 ** 12

            if report_end and meeting_end:
                total += abs((report_end - meeting_end).total_seconds())
            elif report_end or meeting_end:
                total += 10 ** 12

            return total

        return min(meeting_candidates, key=distance)

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

    def get_meeting_attendance(self, online_meeting: dict,
                               meeting_candidates: list[dict] | None = None,
                               skip_processed: bool = True) -> list[dict]:
        """
        Get attendance reports for a meeting or channel and map them to meetings.

        Args:
            online_meeting: Representative online meeting object used to query Graph
            meeting_candidates: Meetings in the same channel eligible for report mapping
            skip_processed: Skip reports already processed (from checkpoint)

        Returns:
            list of dictionaries with report and records data
        """
        meeting_id = online_meeting.get("id")
        if not meeting_id:
            logger.warning("Meeting has no ID, skipping")
            return []

        # Get attendance reports
        reports = self.client.get_attendance_reports(meeting_id)

        if not reports:
            logger.info(
                f"No attendance reports for meeting: {online_meeting.get('_event', {}).get('subject', meeting_id)}")
            return []

        candidates = meeting_candidates or [{
            "meeting_id": meeting_id,
            "meeting_info": online_meeting.get("_event", {}),
            "teams_context": []
        }]

        attendance_data = []
        for report in reports:
            report_id = report.get("id")
            if not report_id:
                continue

            checkpoint_key = self._report_checkpoint_key(report_id)
            if skip_processed and checkpoint_key in self.processed_meetings:
                logger.debug(f"Skipping already processed report: {report_id}")
                continue

            best_meeting = self._select_best_meeting_for_report(report, candidates)
            if not best_meeting:
                logger.debug(f"Could not map report {report_id} to a meeting")
                continue

            # Get attendance records for this report
            records = self.client.get_attendance_records(meeting_id, report_id)

            mapped_meeting_id = best_meeting.get("meeting_id", meeting_id)
            mapped_meeting_info = best_meeting.get("meeting_info", online_meeting.get("_event", {}))

            attendance_data.append({
                "meeting_id": mapped_meeting_id,
                "meeting_info": mapped_meeting_info,
                "report_id": report_id,
                "report_data": report,
                "attendance_records": records,
                "teams_context": best_meeting.get("teams_context", []),
                "source_meeting_id": meeting_id
            })

            logger.info(
                "Extracted %d attendance records from report %s mapped to '%s'",
                len(records),
                report_id,
                mapped_meeting_info.get("subject", mapped_meeting_id)
            )

            self.processed_meetings.add(checkpoint_key)

        if attendance_data:
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
        meetings_by_context: dict[str, dict] = {}

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

            context_key = self._get_context_key(online_meeting, matched_contexts)
            group = meetings_by_context.setdefault(context_key, {
                "online_meeting": online_meeting if online_meeting.get("id") else None,
                "meetings": []
            })

            if not group["online_meeting"] and online_meeting.get("id"):
                group["online_meeting"] = online_meeting

            group["meetings"].append({
                "meeting_id": online_meeting.get("id", online_meeting.get("_calendar_event_id", "")),
                "meeting_info": online_meeting.get("_event", {}),
                "teams_context": matched_contexts
            })

        for context_key, group in meetings_by_context.items():
            representative_meeting = group.get("online_meeting")
            if not representative_meeting:
                logger.debug(
                    "Skipping context %s because no resolvable meeting ID is available",
                    context_key
                )
                continue

            logger.info("")
            logger.info(
                "Processing channel attendance for %s using %d candidate meetings",
                context_key,
                len(group["meetings"])
            )
            attendance_list = self.get_meeting_attendance(
                representative_meeting,
                meeting_candidates=group["meetings"]
            )
            all_attendance.extend(attendance_list)

        logger.info(
            "Extracted attendance data for %d reports across %d matched meetings (out of %d calendar meetings)",
            len(all_attendance),
            matched_events,
            len(all_events)
        )
        return all_attendance
