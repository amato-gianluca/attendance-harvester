"""
Microsoft Graph API client for Teams, channels, meetings, and attendance.
"""
import logging
import time
from typing import Any, Dict, List, Optional
from urllib.parse import quote

import requests

logger = logging.getLogger(__name__)


class GraphAPIError(Exception):
    """Custom exception for Graph API errors."""
    pass


class GraphClient:
    """Client for Microsoft Graph API with retry logic and pagination."""

    BASE_URL = "https://graph.microsoft.com/v1.0"

    def __init__(self, access_token: str, max_retries: int = 3,
                 retry_backoff_factor: int = 2, timeout: int = 30):
        """
        Initialize Graph API client.

        Args:
            access_token: Bearer token for authentication
            max_retries: Maximum number of retries for failed requests
            retry_backoff_factor: Exponential backoff factor (seconds)
            timeout: Request timeout in seconds
        """
        self.access_token = access_token
        self.max_retries = max_retries
        self.retry_backoff_factor = retry_backoff_factor
        self.timeout = timeout

        self.session = requests.Session()
        self.session.headers.update({
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        })

    def _make_request(self, method: str, url: str, **kwargs) -> requests.Response:
        """
        Make HTTP request with retry logic for transient errors.

        Args:
            method: HTTP method (GET, POST, etc.)
            url: Full URL or relative path
            **kwargs: Additional arguments for requests

        Returns:
            Response object

        Raises:
            GraphAPIError: If request fails after all retries
        """
        if not url.startswith("http"):
            url = f"{self.BASE_URL}{url}"

        kwargs.setdefault("timeout", self.timeout)

        for attempt in range(self.max_retries):
            try:
                response = self.session.request(method, url, **kwargs)

                # Handle rate limiting (429)
                if response.status_code == 429:
                    retry_after = int(response.headers.get("Retry-After",
                                                           self.retry_backoff_factor ** (attempt + 1)))
                    logger.warning(f"Rate limited. Retrying after {retry_after}s")
                    time.sleep(retry_after)
                    continue

                # Handle server errors (5xx) with retry
                if 500 <= response.status_code < 600:
                    wait_time = self.retry_backoff_factor ** (attempt + 1)
                    logger.warning(f"Server error {response.status_code}. Retrying after {wait_time}s")
                    time.sleep(wait_time)
                    continue

                # Raise for client errors (4xx) - no retry
                if 400 <= response.status_code < 500:
                    error_msg = f"Client error {response.status_code}: {response.text}"
                    logger.error(error_msg)
                    raise GraphAPIError(error_msg)

                response.raise_for_status()
                return response

            except requests.exceptions.RequestException as e:
                if attempt == self.max_retries - 1:
                    raise GraphAPIError(f"Request failed after {self.max_retries} attempts: {e}")

                wait_time = self.retry_backoff_factor ** (attempt + 1)
                logger.warning(f"Request failed: {e}. Retrying after {wait_time}s")
                time.sleep(wait_time)

        raise GraphAPIError(f"Request failed after {self.max_retries} attempts")

    def _paginate(self, url: str, params: Optional[Dict] = None) -> List[Dict]:
        """
        Handle pagination for Graph API responses.

        Args:
            url: Initial URL
            params: Query parameters

        Returns:
            List of all items from paginated response
        """
        items = []
        next_link = url

        while next_link:
            # Only use params for the first request
            response = self._make_request("GET", next_link,
                                         params=params if next_link == url else None)
            data = response.json()

            # Add items from current page
            if "value" in data:
                items.extend(data["value"])
                logger.debug(f"Fetched {len(data['value'])} items, total: {len(items)}")

            # Get next page link
            next_link = data.get("@odata.nextLink")

        return items

    def get_joined_teams(self) -> List[Dict]:
        """
        Get all teams the user has joined.

        Returns:
            List of team objects
        """
        logger.info("Fetching joined teams")
        teams = self._paginate("/me/joinedTeams")
        logger.info(f"Found {len(teams)} joined teams")
        return teams

    def get_associated_teams(self) -> List[Dict]:
        """
        Get associated teams (includes shared channel hosts).

        Returns:
            List of associated team info objects
        """
        logger.info("Fetching associated teams")
        try:
            teams = self._paginate("/me/teamwork/associatedTeams")
            logger.info(f"Found {len(teams)} associated teams")
            return teams
        except GraphAPIError as e:
            logger.warning(f"Failed to fetch associated teams: {e}")
            return []

    def get_team_channels(self, team_id: str) -> List[Dict]:
        """
        Get all channels for a specific team.

        Args:
            team_id: Team ID

        Returns:
            List of channel objects
        """
        logger.debug(f"Fetching channels for team {team_id}")
        channels = self._paginate(f"/teams/{team_id}/channels")
        logger.debug(f"Found {len(channels)} channels in team {team_id}")
        return channels

    def get_calendar_events(self, start_datetime: str, end_datetime: str) -> List[Dict]:
        """
        Get calendar events in a time range.

        Args:
            start_datetime: Start time in ISO 8601 format
            end_datetime: End time in ISO 8601 format

        Returns:
            List of event objects with online meeting details
        """
        logger.info(f"Fetching calendar events from {start_datetime} to {end_datetime}")
        params = {
            "startDateTime": start_datetime,
            "endDateTime": end_datetime,
            "$select": "subject,start,end,isOnlineMeeting,onlineMeetingProvider,onlineMeeting,organizer"
        }
        events = self._paginate("/me/calendarView", params=params)

        # Filter to Teams meetings only
        teams_events = [e for e in events
                       if e.get("isOnlineMeeting") and
                       e.get("onlineMeetingProvider") == "teamsForBusiness"]

        logger.info(f"Found {len(teams_events)} Teams meetings out of {len(events)} total events")
        return teams_events

    def get_online_meeting_by_join_url(self, join_url: str) -> Optional[Dict]:
        """
        Get online meeting details by join URL.

        Args:
            join_url: Teams meeting join URL

        Returns:
            Online meeting object or None if not found
        """
        encoded_url = quote(join_url, safe='')
        filter_param = f"JoinWebUrl eq '{encoded_url}'"

        try:
            meetings = self._paginate("/me/onlineMeetings", params={"$filter": filter_param})
            if meetings:
                return meetings[0]
            return None
        except GraphAPIError as e:
            logger.warning(f"Failed to get online meeting by join URL: {e}")
            return None

    def get_attendance_reports(self, meeting_id: str) -> List[Dict]:
        """
        Get all attendance reports for an online meeting.

        Args:
            meeting_id: Online meeting ID

        Returns:
            List of attendance report objects
        """
        logger.debug(f"Fetching attendance reports for meeting {meeting_id}")
        try:
            reports = self._paginate(f"/me/onlineMeetings/{meeting_id}/attendanceReports")
            logger.debug(f"Found {len(reports)} attendance reports for meeting {meeting_id}")
            return reports
        except GraphAPIError as e:
            logger.warning(f"Failed to fetch attendance reports for meeting {meeting_id}: {e}")
            return []

    def get_attendance_records(self, meeting_id: str, report_id: str) -> List[Dict]:
        """
        Get attendance records for a specific report.

        Args:
            meeting_id: Online meeting ID
            report_id: Attendance report ID

        Returns:
            List of attendance record objects
        """
        logger.debug(f"Fetching attendance records for report {report_id}")
        try:
            records = self._paginate(
                f"/me/onlineMeetings/{meeting_id}/attendanceReports/{report_id}/attendanceRecords"
            )
            logger.debug(f"Found {len(records)} attendance records for report {report_id}")
            return records
        except GraphAPIError as e:
            logger.error(f"Failed to fetch attendance records for report {report_id}: {e}")
            return []
