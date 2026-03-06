"""
Team filtering logic using regular expressions.
"""
import re
import logging
from typing import Dict, List

logger = logging.getLogger(__name__)


class TeamFilter:
    """Filter teams based on regex patterns."""

    def __init__(self, pattern: str):
        """
        Initialize team filter with regex pattern.

        Args:
            pattern: Regular expression pattern to match team names
        """
        self.pattern = pattern
        try:
            self.regex = re.compile(pattern, re.IGNORECASE)
        except re.error as e:
            raise ValueError(f"Invalid regex pattern '{pattern}': {e}")

    def matches(self, team_name: str) -> bool:
        """
        Check if team name matches the filter pattern.

        Args:
            team_name: Name of the team

        Returns:
            True if team name matches pattern
        """
        return bool(self.regex.search(team_name))

    def filter_teams(self, teams: List[Dict]) -> List[Dict]:
        """
        Filter list of teams by regex pattern.

        Args:
            teams: List of team objects (must have 'displayName' field)

        Returns:
            List of teams matching the pattern
        """
        matched_teams = []
        for team in teams:
            team_name = team.get("displayName", "")
            if self.matches(team_name):
                matched_teams.append(team)
                logger.debug(f"Team matched: {team_name}")
            else:
                logger.debug(f"Team filtered out: {team_name}")

        logger.info(f"Filtered {len(matched_teams)} teams out of {len(teams)} total")
        return matched_teams

    def get_general_channel(self, channels: List[Dict]) -> Dict:
        """
        Get the General channel from a list of channels.

        Args:
            channels: List of channel objects

        Returns:
            General channel object or None if not found
        """
        for channel in channels:
            # General channel has a specific membership type
            if channel.get("membershipType") == "standard" and \
               channel.get("displayName", "").lower() == "general":
                return channel

        # Fallback: return first channel with "general" in name
        for channel in channels:
            if "general" in channel.get("displayName", "").lower():
                return channel

        logger.warning("General channel not found in team")
        return None
