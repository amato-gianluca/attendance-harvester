"""
Microsoft Graph API client for Teams, channels, meetings, and attendance.
"""
import json
import logging
import time
from pathlib import Path

import requests

logger = logging.getLogger(__name__)


class GraphAPIError(Exception):
    """Custom exception for Graph API errors."""
    pass


class GraphClient:
    """Client for Microsoft Graph API with retry logic and pagination."""

    BASE_URL = "https://graph.microsoft.com/v1.0"

    def __init__(self, access_token: str, max_retries: int = 3,
                 retry_backoff_factor: int = 2, timeout: int = 30,
                 user_id: str | None = None,
                 metadata_cache_file: str | None = None):
        """
        Initialize Graph API client.

        Args:
            access_token: Bearer token for authentication
            max_retries: Maximum number of retries for failed requests
            retry_backoff_factor: Exponential backoff factor (seconds)
            timeout: Request timeout in seconds
            user_id: User ID/UPN used for user-scoped APIs in confidential mode
            metadata_cache_file: Optional JSON file for caching team/channel metadata
        """
        self.access_token = access_token
        self.max_retries = max_retries
        self.retry_backoff_factor = retry_backoff_factor
        self.timeout = timeout
        self.user_id = user_id
        self.metadata_cache_file = Path(metadata_cache_file) if metadata_cache_file else None

        self.session = requests.Session()
        self.session.headers.update({
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        })
        self.metadata_cache = self._load_metadata_cache()

    def _load_metadata_cache(self) -> dict:
        """Load cached team/channel metadata from disk."""
        cache = {"teams": {}}

        if not self.metadata_cache_file or not self.metadata_cache_file.exists():
            return cache

        try:
            with open(self.metadata_cache_file, "r", encoding="utf-8") as f:
                payload = json.load(f)
        except (OSError, json.JSONDecodeError) as e:
            logger.warning("Failed to load metadata cache from %s: %s",
                           self.metadata_cache_file, e)
            return cache

        if not isinstance(payload, dict):
            logger.warning("Ignoring invalid metadata cache payload in %s",
                           self.metadata_cache_file)
            return cache

        teams_payload = payload.get("teams", {})
        if isinstance(teams_payload, dict):
            cache["teams"] = teams_payload

        logger.debug("Loaded metadata cache from %s", self.metadata_cache_file)
        return cache

    def _save_metadata_cache(self):
        """Persist cached team/channel metadata to disk."""
        if not self.metadata_cache_file:
            return

        try:
            self.metadata_cache_file.parent.mkdir(parents=True, exist_ok=True)
            with open(self.metadata_cache_file, "w", encoding="utf-8") as f:
                json.dump(self.metadata_cache, f, indent=2)
            logger.debug("Saved metadata cache to %s", self.metadata_cache_file)
        except OSError as e:
            logger.warning("Failed to save metadata cache to %s: %s",
                           self.metadata_cache_file, e)

    def clear_metadata_cache(self):
        """Clear in-memory and on-disk team/channel metadata cache."""
        self.metadata_cache = {"teams": {}}

        if self.metadata_cache_file and self.metadata_cache_file.exists():
            try:
                self.metadata_cache_file.unlink()
                logger.info("Metadata cache cleared: %s", self.metadata_cache_file)
            except OSError as e:
                logger.warning("Failed to clear metadata cache %s: %s",
                               self.metadata_cache_file, e)

    def sync_filtered_teams_cache(self, filtered_teams: list[dict]):
        """
        Keep cache entries only for teams matching the active filter.

        Team discovery is always fetched live; this cache only persists metadata
        for the currently matched teams.
        """
        filtered_team_ids = {
            team["id"] for team in filtered_teams if isinstance(team, dict) and team.get("id")
        }

        cached_teams = self.metadata_cache.setdefault("teams", {})
        removed_team_ids = [team_id for team_id in list(cached_teams) if team_id not in filtered_team_ids]
        for team_id in removed_team_ids:
            del cached_teams[team_id]

        for team in filtered_teams:
            team_id = team.get("id")
            if not team_id:
                continue

            team_cache = cached_teams.setdefault(team_id, {
                "team": team,
                "owners": None,
                "channels": None,
                "primary_channel": None,
                "owners_fetched": False,
                "channels_fetched": False,
                "primary_channel_fetched": False,
                "channels_by_id": {},
            })
            team_cache["team"] = team
            team_cache.setdefault("owners", None)
            team_cache.setdefault("channels", None)
            team_cache.setdefault("primary_channel", None)
            team_cache.setdefault("owners_fetched", False)
            team_cache.setdefault("channels_fetched", False)
            team_cache.setdefault("primary_channel_fetched", False)
            team_cache.setdefault("channels_by_id", {})

        self._save_metadata_cache()

    def _get_cached_team_entry(self, team_id: str) -> dict:
        """Return cached metadata bucket for a filtered team."""
        return self.metadata_cache.setdefault("teams", {}).setdefault(team_id, {
            "team": {"id": team_id},
            "owners": None,
            "channels": None,
            "primary_channel": None,
            "owners_fetched": False,
            "channels_fetched": False,
            "primary_channel_fetched": False,
            "channels_by_id": {},
        })

    def _user_path(self, relative_path: str) -> str:
        """
        Build a user-scoped Graph endpoint.

        In public delegated mode, this returns /me/{relative_path}.
        In confidential mode (user_id provided), this returns /users/{user_id}/{relative_path}.
        """
        normalized = relative_path.lstrip("/")
        if self.user_id:
            return f"/users/{self.user_id}/{normalized}"
        return f"/me/{normalized}"

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
                    logger.warning(
                        f"Rate limited. Retrying after {retry_after}s")
                    time.sleep(retry_after)
                    continue

                # Handle server errors (5xx) with retry
                if 500 <= response.status_code < 600:
                    wait_time = self.retry_backoff_factor ** (attempt + 1)
                    logger.warning(
                        f"Server error {response.status_code}. Retrying after {wait_time}s")
                    time.sleep(wait_time)
                    continue

                # Handle 404 (not found) gracefully - return without raising
                # 404 is expected for meetings without accessible attendance data
                if response.status_code == 404:
                    return response

                # Raise for other client errors (4xx)
                if 400 <= response.status_code < 500:
                    error_msg = f"Client error {response.status_code}: {response.text}"
                    logger.error(error_msg)
                    raise GraphAPIError(error_msg)

                response.raise_for_status()
                return response

            except requests.exceptions.RequestException as e:
                if attempt == self.max_retries - 1:
                    raise GraphAPIError(
                        f"Request failed after {self.max_retries} attempts: {e}")

                wait_time = self.retry_backoff_factor ** (attempt + 1)
                logger.warning(
                    f"Request failed: {e}. Retrying after {wait_time}s")
                time.sleep(wait_time)

        raise GraphAPIError(
            f"Request failed after {self.max_retries} attempts")

    def _paginate(self, url: str, params: dict | None = None) -> list[dict]:
        """
        Handle pagination for Graph API responses.

        Args:
            url: Initial URL
            params: Query parameters

        Returns:
            list of all items from paginated response, or empty list if 404
        """
        items = []
        next_link = url

        while next_link:
            # Only use params for the first request
            response = self._make_request("GET", next_link,
                                          params=params if next_link == url else None)

            # Handle 404 (not found) - return empty list
            if response.status_code == 404:
                return []

            data = response.json()

            # Add items from current page
            if "value" in data:
                items.extend(data["value"])
                logger.debug(
                    f"Fetched {len(data['value'])} items, total: {len(items)}")

            # Get next page link
            next_link = data.get("@odata.nextLink")

        return items

    def get_joined_teams(self) -> list[dict]:
        """
        Get all teams the user has joined.

        Returns:
            list of team objects
        """
        logger.info("Fetching joined teams")
        teams = self._paginate(self._user_path("joinedTeams"))
        logger.info(f"Found {len(teams)} joined teams")
        return teams

    def get_associated_teams(self) -> list[dict]:
        """
        Get associated teams (includes shared channel hosts).

        Returns:
            list of associated team info objects
        """
        logger.info("Fetching associated teams")
        try:
            teams = self._paginate(self._user_path("teamwork/associatedTeams"))
            logger.info(f"Found {len(teams)} associated teams")
            return teams
        except GraphAPIError as e:
            logger.warning(f"Failed to fetch associated teams: {e}")
            return []

    def get_team_channels(self, team_id: str) -> list[dict]:
        """
        Get all channels for a specific team.

        Args:
            team_id: Team ID

        Returns:
            list of channel objects
        """
        team_cache = self._get_cached_team_entry(team_id)
        if team_cache.get("channels_fetched"):
            cached_channels = team_cache.get("channels") or []
            logger.debug("Using cached channels for team %s", team_id)
            return cached_channels

        logger.debug(f"Fetching channels for team {team_id}")
        channels = self._paginate(f"/teams/{team_id}/channels")
        team_cache["channels"] = channels
        team_cache["channels_fetched"] = True
        channels_by_id = team_cache.setdefault("channels_by_id", {})
        for channel in channels:
            channel_id = channel.get("id")
            if channel_id:
                channels_by_id[channel_id] = channel
        self._save_metadata_cache()
        logger.debug(f"Found {len(channels)} channels in team {team_id}")
        return channels

    def get_team_owners(self, team_id: str) -> list[dict]:
        """
        Get owners for a specific team.

        Teams are backed by Microsoft 365 groups, so owners are read from the
        corresponding group object.

        Args:
            team_id: Team ID

        Returns:
            list of owner directory objects
        """
        team_cache = self._get_cached_team_entry(team_id)
        if team_cache.get("owners_fetched"):
            cached_owners = team_cache.get("owners") or []
            logger.debug("Using cached owners for team %s", team_id)
            return cached_owners

        logger.debug("Fetching owners for team %s", team_id)
        try:
            owners = self._paginate(f"/groups/{team_id}/owners")
        except GraphAPIError as e:
            logger.warning(
                "Failed to fetch owners for team %s: %s. "
                "This may require additional Microsoft Graph permissions.",
                team_id,
                e
            )
            return []

        team_cache["owners"] = owners
        team_cache["owners_fetched"] = True
        self._save_metadata_cache()
        logger.debug("Found %d owners for team %s", len(owners), team_id)
        return owners

    def get_team_channel(self, team_id: str, channel_id: str) -> dict | None:
        """
        Get a specific channel for a team.

        Args:
            team_id: Team ID
            channel_id: Channel ID

        Returns:
            Channel object or None if not found
        """
        team_cache = self._get_cached_team_entry(team_id)
        cached_team_channels = team_cache.setdefault("channels_by_id", {})
        if channel_id in cached_team_channels:
            logger.debug("Using cached channel %s for team %s", channel_id, team_id)
            return cached_team_channels[channel_id]

        logger.debug(f"Fetching channel {channel_id} for team {team_id}")
        response = self._make_request("GET", f"/teams/{team_id}/channels/{channel_id}")
        if response.status_code == 404:
            logger.warning("Channel %s not found in team %s", channel_id, team_id)
            cached_team_channels[channel_id] = None
            self._save_metadata_cache()
            return None
        channel = response.json()
        cached_team_channels[channel_id] = channel
        self._save_metadata_cache()
        return channel

    def get_team_primary_channel(self, team_id: str) -> dict | None:
        """
        Get the primary (General) channel for a specific team.

        Args:
            team_id: Team ID

        Returns:
            Channel object or None if not found
        """
        team_cache = self._get_cached_team_entry(team_id)
        if team_cache.get("primary_channel_fetched"):
            logger.info("Using cached primary channel for team %s", team_id)
            return team_cache.get("primary_channel")

        logger.info(f"Fetching primary channel for team {team_id}")
        response = self._make_request("GET", f"/teams/{team_id}/primaryChannel")
        if response.status_code == 404:
            logger.warning("Primary channel not found for team %s", team_id)
            team_cache["primary_channel"] = None
            team_cache["primary_channel_fetched"] = True
            self._save_metadata_cache()
            return None

        primary_channel = response.json()
        team_cache["primary_channel"] = primary_channel
        team_cache["primary_channel_fetched"] = True
        channel_id = primary_channel.get("id")
        if channel_id:
            team_cache.setdefault("channels_by_id", {})[channel_id] = primary_channel
        self._save_metadata_cache()

        # Re-fetch by id when possible to normalize the channel payload.
        # This is defensive: Microsoft documents /primaryChannel as returning a channel
        # object, but there are reports of renamed primary channels still coming back
        # with displayName="General".
        # channel_id = primary_channel.get("id")
        # if channel_id:
        #     full_channel = self.get_team_channel(team_id, channel_id)
        #     if full_channel:
        #         return full_channel

        return primary_channel

    def get_calendar_events(self, start_datetime: str, end_datetime: str) -> list[dict]:
        """
        Get calendar events in a time range.

        Args:
            start_datetime: Start time in ISO 8601 format
            end_datetime: End time in ISO 8601 format

        Returns:
            list of event objects with online meeting details
        """
        logger.info(
            f"Fetching calendar events from {start_datetime} to {end_datetime}")
        params = {
            "startDateTime": start_datetime,
            "endDateTime": end_datetime,
            "$select": "id,subject,start,end,isOnlineMeeting,onlineMeetingProvider,onlineMeeting,organizer,location,locations,bodyPreview"
        }
        events = self._paginate(self._user_path("calendarView"), params=params)

        # Filter to Teams meetings only
        teams_events = [e for e in events
                        if e.get("isOnlineMeeting") and
                        e.get("onlineMeetingProvider") == "teamsForBusiness"]

        logger.info(
            f"Found {len(teams_events)} Teams meetings out of {len(events)} total events")
        return teams_events

    def get_online_meeting_by_join_url(self, join_url: str) -> dict | None:
        """
        Get online meeting details by join URL.

        Args:
            join_url: Teams meeting join URL

        Returns:
            Online meeting object or None if not found (expected for non-organized meetings)
        """
        filter_param = f"JoinWebUrl eq '{join_url}'"
        try:
            meetings = self._paginate(self._user_path(
                "onlineMeetings"), params={"$filter": filter_param})
            if meetings:
                return meetings[0]
            logger.debug(
                f"No online meeting found for join URL (expected for non-organized meetings)")
            return None
        except GraphAPIError as e:
            logger.debug(f"Could not retrieve online meeting by join URL: {e}")
            return None

    def get_attendance_reports(self, meeting_id: str) -> list[dict]:
        """
        Get all attendance reports for an online meeting.

        Args:
            meeting_id: Online meeting ID

        Returns:
            list of attendance report objects, or empty list if not found/no access
        """
        logger.debug(f"Fetching attendance reports for meeting {meeting_id}")
        try:
            reports = self._paginate(self._user_path(
                f"onlineMeetings/{meeting_id}/attendanceReports"))
            if reports:
                logger.debug(
                    f"Found {len(reports)} attendance reports for meeting {meeting_id}")
            else:
                logger.debug(
                    f"No attendance reports for meeting {meeting_id}. "
                    "This is normal - reports are only available for meetings you organized."
                )
            return reports
        except GraphAPIError as e:
            logger.warning(
                f"Failed to fetch attendance reports for meeting {meeting_id}: {e}")
            return []

    def get_attendance_records(self, meeting_id: str, report_id: str) -> list[dict]:
        """
        Get attendance records for a specific report.

        Args:
            meeting_id: Online meeting ID
            report_id: Attendance report ID

        Returns:
            list of attendance record objects, or empty list if not found/no access
        """
        logger.debug(f"Fetching attendance records for report {report_id}")
        try:
            records = self._paginate(
                self._user_path(
                    f"onlineMeetings/{meeting_id}/attendanceReports/{report_id}/attendanceRecords")
            )
            logger.debug(
                f"Found {len(records)} attendance records for report {report_id}")
            return records
        except GraphAPIError as e:
            logger.debug(
                f"Failed to fetch attendance records for report {report_id}: {e}")
            return []
