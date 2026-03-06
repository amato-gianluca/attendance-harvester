"""
Authentication module using MSAL with device code flow and token caching.
"""
import os
import logging
from pathlib import Path
from typing import Dict, List, Optional

import msal

logger = logging.getLogger(__name__)


class AuthenticationError(Exception):
    """Custom exception for authentication errors."""
    pass


class Authenticator:
    """Handles authentication to Microsoft Graph API using device code flow."""

    def __init__(self, client_id: str, authority: str, scopes: List[str],
                 cache_dir: str = "./cache", cache_filename: str = "token_cache.bin"):
        """
        Initialize authenticator with Azure AD app credentials.

        Args:
            client_id: Azure AD application (client) ID
            authority: Authority URL (e.g., https://login.microsoftonline.com/tenant-id)
            scopes: List of Microsoft Graph API scopes
            cache_dir: Directory for token cache
            cache_filename: Token cache filename
        """
        self.client_id = client_id
        self.authority = authority
        self.scopes = scopes
        self.cache_path = Path(cache_dir) / cache_filename

        # Ensure cache directory exists
        self.cache_path.parent.mkdir(parents=True, exist_ok=True)

        # Initialize token cache
        self.cache = msal.SerializableTokenCache()
        if self.cache_path.exists():
            self.cache.deserialize(self.cache_path.read_text())

        # Create public client application
        self.app = msal.PublicClientApplication(
            client_id=self.client_id,
            authority=self.authority,
            token_cache=self.cache
        )

    def _save_cache(self):
        """Save token cache to disk if it has changed."""
        if self.cache.has_state_changed:
            self.cache_path.write_text(self.cache.serialize())
            logger.debug(f"Token cache saved to {self.cache_path}")

    def acquire_token(self) -> str:
        """
        Acquire access token using device code flow with caching.

        Returns:
            Access token string

        Raises:
            AuthenticationError: If authentication fails
        """
        # Try to get token silently from cache first
        accounts = self.app.get_accounts()
        if accounts:
            logger.info("Found cached account, attempting silent token acquisition")
            result = self.app.acquire_token_silent(self.scopes, account=accounts[0])
            if result and "access_token" in result:
                logger.info("Successfully acquired token from cache")
                self._save_cache()
                return result["access_token"]
            else:
                logger.debug("Silent token acquisition failed, will use device code flow")

        # Initiate device code flow
        logger.info("Starting device code authentication flow")
        flow = self.app.initiate_device_flow(scopes=self.scopes)

        if "user_code" not in flow:
            raise AuthenticationError(
                f"Failed to create device flow: {flow.get('error_description', 'Unknown error')}"
            )

        # Display user instructions
        print("\n" + "=" * 70)
        print("AUTHENTICATION REQUIRED")
        print("=" * 70)
        print(flow["message"])
        print("=" * 70 + "\n")

        # Wait for user to authenticate
        result = self.app.acquire_token_by_device_flow(flow)

        if "access_token" not in result:
            error_desc = result.get("error_description", "Unknown error")
            raise AuthenticationError(f"Authentication failed: {error_desc}")

        logger.info("Successfully authenticated via device code flow")
        self._save_cache()

        return result["access_token"]

    def clear_cache(self):
        """Clear the token cache file."""
        if self.cache_path.exists():
            self.cache_path.unlink()
            logger.info(f"Token cache cleared: {self.cache_path}")
        self.cache = msal.SerializableTokenCache()
