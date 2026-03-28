"""
Authentication module using MSAL with device code flow and token caching.
"""
import logging
from pathlib import Path

import msal

logger = logging.getLogger(__name__)


class AuthenticationError(Exception):
    """Custom exception for authentication errors."""
    pass


class Authenticator:
    """Handles authentication to Microsoft Graph API (public or confidential mode)."""

    def __init__(self, client_id: str, authority: str, scopes: list[str], cache_path: Path,
                 auth_mode: str = "public", client_secret: str | None = None):
        """
        Initialize authenticator with Azure AD app credentials.

        Args:
            client_id: Azure AD application (client) ID
            authority: Authority URL (e.g., https://login.microsoftonline.com/tenant-id)
            scopes: List of Microsoft Graph API scopes
            cache_path: Full path to token cache file
            auth_mode: Authentication mode, either "public" (device code) or "confidential" (client credentials)
            client_secret: Client secret for confidential mode
        """
        self.client_id = client_id
        self.authority = authority
        self.scopes = scopes
        self.auth_mode = auth_mode
        self.client_secret = client_secret
        self.cache_path = cache_path
        self.public_app: msal.PublicClientApplication | None = None
        self.confidential_app: msal.ConfidentialClientApplication | None = None

        # Initialize token cache
        self.cache = msal.SerializableTokenCache()
        if self.cache_path.exists():
            self.cache.deserialize(self.cache_path.read_text())

        if self.auth_mode == "confidential":
            if not self.client_secret:
                raise AuthenticationError("Confidential mode requires a client secret.")
            self.confidential_app = msal.ConfidentialClientApplication(
                client_id=self.client_id,
                authority=self.authority,
                client_credential=self.client_secret,
                token_cache=self.cache
            )
        else:
            self.public_app = msal.PublicClientApplication(
                client_id=self.client_id,
                authority=self.authority,
                token_cache=self.cache
            )

    def _save_cache(self):
        """Save token cache to disk if it has changed."""
        if self.cache.has_state_changed:
            self.cache_path.write_text(self.cache.serialize())

    def acquire_token(self) -> str:
        """
        Acquire access token using configured auth mode.

        Returns:
            Access token string

        Raises:
            AuthenticationError: If authentication fails
        """
        if self.auth_mode == "confidential":
            return self._acquire_token_confidential()
        else:
            return self._acquire_token_public()

    def _acquire_token_public(self) -> str:
        """
        Acquire access token using public client credentials flow.

        Returns:
            Access token string

        Raises:
            AuthenticationError: If authentication fails
        """
        if not self.public_app:
            raise AuthenticationError("Public authentication app is not initialized")

        # Try to get token silently from cache first
        accounts = self.public_app.get_accounts()
        if accounts:
            logger.debug("Found cached account, attempting silent token acquisition")
            result = self.public_app.acquire_token_silent(self.scopes, account=accounts[0])
            if result and "access_token" in result:
                logger.debug("Successfully acquired token from cache")
                self._save_cache()
                return result["access_token"]
            else:
                logger.debug("Silent token acquisition failed, will use device code flow")

        # Initiate device code flow
        logger.debug("Starting device code authentication flow")
        flow = self.public_app.initiate_device_flow(scopes=self.scopes)

        if "user_code" not in flow:
            raise AuthenticationError(
                f"Failed to create device flow: {flow.get('error_description', 'Unknown error')}"
            )

        # Display user instructions
        print("=" * 70)
        print("AUTHENTICATION REQUIRED")
        print("=" * 70)
        print(flow["message"])
        print("=" * 70 + "\n")

        # Wait for user to authenticate
        result = self.public_app.acquire_token_by_device_flow(flow)

        if "access_token" not in result:
            error_code = result.get("error")
            error_desc = result.get("error_description", "Unknown error")

            # Common misconfiguration for custom Azure app registrations
            if "AADSTS7000218" in error_desc or error_code == "invalid_client":
                raise AuthenticationError(
                    "Authentication failed: your Azure app is configured as a confidential client, "
                    "but device code flow requires a public client. "
                    "In Azure Portal -> App registration -> Authentication, enable 'Allow public client flows' = Yes. "
                    "If you prefer, you can use a well-known public client ID. "
                    f"Original error: {error_desc}"
                )
            else:
                raise AuthenticationError(f"Authentication failed: {error_desc}")

        logger.debug("Successfully authenticated via device code flow")
        self._save_cache()

        return result["access_token"]

    def _acquire_token_confidential(self) -> str:
        """
        Acquire access token using confidential client credentials flow.

        Returns:
            Access token string

        Raises:
            AuthenticationError: If authentication fails
        """
        if not self.confidential_app:
            raise AuthenticationError("Confidential authentication app is not initialized")

        logger.debug("Starting confidential client authentication flow")
        result = self.confidential_app.acquire_token_for_client(scopes=self.scopes)
        if not result:
            raise AuthenticationError("Authentication failed: empty token response")

        if "access_token" not in result:
            error = result.get("error")
            error_desc = result.get("error_description", "Unknown error")
            raise AuthenticationError(f"Authentication failed: {error} {error_desc}")

        logger.debug("Successfully authenticated via confidential client credentials")
        self._save_cache()
        return result["access_token"]

    def clear_cache(self):
        """Clear the token cache file."""
        if self.cache_path.exists():
            self.cache_path.unlink()
            logger.debug(f"Token cache cleared: {self.cache_path}")
        self.cache = msal.SerializableTokenCache()
