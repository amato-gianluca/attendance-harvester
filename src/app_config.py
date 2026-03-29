"""
Application configuration loading and validation.
"""
from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml


PUBLIC_SCOPES_DEFAULT = [
    "Team.ReadBasic.All",
    "Channel.ReadBasic.All",
    "Calendars.Read",
    "OnlineMeetings.Read",
    "OnlineMeetingArtifact.Read.All",
]

SHAREPOINT_PUBLIC_SCOPES_DEFAULT = [
    "Files.ReadWrite.All",
    "Sites.ReadWrite.All",
]

OUTPUT_FORMATS = {"both", "csv", "json"}
AUTH_MODES = {"public", "confidential"}


def _ensure_mapping(value: Any, field_name: str) -> dict[str, Any]:
    if value is None:
        return {}
    if not isinstance(value, dict):
        raise ValueError(f"{field_name} must be a mapping")
    return value


def _require_string(value: Any, field_name: str) -> str:
    if not isinstance(value, str):
        raise ValueError(f"{field_name} must be a string")
    return value


def _optional_string(value: Any, default: str = "") -> str:
    if value is None:
        return default
    if not isinstance(value, str):
        raise ValueError("Expected a string value")
    return value


def _require_non_empty_string(value: Any, field_name: str) -> str:
    parsed = _require_string(value, field_name).strip()
    if not parsed:
        raise ValueError(f"{field_name} is required")
    return parsed


def _optional_non_negative_int(value: Any, default: int, field_name: str) -> int:
    if value is None:
        return default
    parsed = int(value)
    if parsed < 0:
        raise ValueError(f"{field_name} must be >= 0")
    return parsed


def _require_positive_int(value: Any, field_name: str) -> int:
    parsed = int(value)
    if parsed <= 0:
        raise ValueError(f"{field_name} must be > 0")
    return parsed


def _require_bool(value: Any, field_name: str) -> bool:
    if not isinstance(value, bool):
        raise ValueError(f"{field_name} must be a boolean")
    return value


def _read_scopes(value: Any, default: list[str], field_name: str) -> list[str]:
    if value is None:
        return list(default)
    if not isinstance(value, list) or any(not isinstance(item, str) or not item.strip() for item in value):
        raise ValueError(f"{field_name} must be a list of non-empty strings")
    return [item.strip() for item in value]


@dataclass(frozen=True)
class AuthConfig:
    mode: str
    client_id: str
    authority: str
    client_secret: str | None
    scopes: list[str]
    target_user_id: str
    token_cache: str
    clear_cache: bool

    @classmethod
    def from_mapping(
        cls,
        raw: dict[str, Any],
        *,
        default_scopes: list[str],
        clear_cache: bool,
        inherited_mode: str | None = None,
        inherited_client_id: str | None = None,
        inherited_authority: str | None = None,
        inherited_client_secret: str | None = None,
        inherited_target_user_id: str = "",
        default_token_cache: str = "token_cache.bin",
    ) -> "AuthConfig":
        mode = _optional_string(raw.get("mode"), "").strip().lower() or (inherited_mode or "public")
        if mode not in AUTH_MODES:
            raise ValueError("auth.mode must be either 'public' or 'confidential'")

        client_id = _optional_string(raw.get("client_id"), "").strip() or (inherited_client_id or "")
        authority = _optional_string(raw.get("authority"), "").strip() or (inherited_authority or "")
        client_secret_raw = raw.get("client_secret")
        client_secret = (
            (_optional_string(client_secret_raw).strip() or None)
            if client_secret_raw is not None
            else inherited_client_secret
        )
        if client_secret is None:
            client_secret = inherited_client_secret
        target_user_id = _optional_string(raw.get("target_user_id"), "").strip() or inherited_target_user_id
        token_cache = _optional_string(raw.get("token_cache"), "").strip()
        if not token_cache:
            token_cache = _optional_string(raw.get("cache_filename"), "").strip() or default_token_cache

        if not client_id:
            raise ValueError("auth.client_id is required")
        if not authority:
            raise ValueError("auth.authority is required")
        if not token_cache:
            raise ValueError("auth.token_cache is required")

        scopes = _read_scopes(raw.get("scopes"), default_scopes, "auth.scopes")
        if mode == "confidential":
            scopes = ["https://graph.microsoft.com/.default"]

        return cls(
            mode=mode,
            client_id=client_id,
            authority=authority,
            client_secret=client_secret or None,
            scopes=scopes,
            target_user_id=target_user_id,
            token_cache=token_cache,
            clear_cache=clear_cache,
        )


@dataclass(frozen=True)
class TeamFilterConfig:
    regex: str

    @classmethod
    def from_mapping(cls, raw: dict[str, Any], *, regex_override: str | None) -> "TeamFilterConfig":
        regex = regex_override if regex_override is not None else raw.get("regex", ".*")
        return cls(regex=_require_string(regex, "team_filter.regex"))


@dataclass(frozen=True)
class MeetingsConfig:
    lookback_days: int
    lookahead_days: int
    include_associated_teams: bool
    general_channel_only: bool

    @classmethod
    def from_mapping(
        cls,
        raw: dict[str, Any],
        *,
        lookback_days_override: int | None,
        lookahead_days_override: int | None,
    ) -> "MeetingsConfig":
        return cls(
            lookback_days=_optional_non_negative_int(
                lookback_days_override if lookback_days_override is not None else raw.get("lookback_days"),
                30,
                "meetings.lookback_days",
            ),
            lookahead_days=_optional_non_negative_int(
                lookahead_days_override if lookahead_days_override is not None else raw.get("lookahead_days"),
                0,
                "meetings.lookahead_days",
            ),
            include_associated_teams=_require_bool(
                raw.get("include_associated_teams", True),
                "meetings.include_associated_teams",
            ),
            general_channel_only=_require_bool(
                raw.get("general_channel_only", True),
                "meetings.general_channel_only",
            ),
        )


@dataclass(frozen=True)
class SharePointCSVConfig:
    auto_upload: bool
    site_id: str
    site_hostname: str
    site_path: str
    drive_id: str
    drive_name: str
    folder_path: str
    auth: AuthConfig

    @classmethod
    def from_mapping(
        cls,
        raw: dict[str, Any],
        *,
        base_auth: AuthConfig,
        clear_cache: bool,
    ) -> "SharePointCSVConfig":
        auth_raw = _ensure_mapping(raw.get("auth"), "output.sharepoint_csv.auth")
        sharepoint_auth = AuthConfig.from_mapping(
            auth_raw,
            default_scopes=SHAREPOINT_PUBLIC_SCOPES_DEFAULT,
            clear_cache=clear_cache,
            inherited_mode=base_auth.mode,
            inherited_client_id=base_auth.client_id,
            inherited_authority=base_auth.authority,
            inherited_client_secret=base_auth.client_secret,
            default_token_cache="sharepoint_token_cache.bin",
        )

        return cls(
            auto_upload=_require_bool(raw.get("auto_upload", False), "output.sharepoint_csv.auto_upload"),
            site_id=_optional_string(raw.get("site_id")).strip(),
            site_hostname=_optional_string(raw.get("site_hostname")).strip(),
            site_path=_optional_string(raw.get("site_path")).strip(),
            drive_id=_optional_string(raw.get("drive_id")).strip(),
            drive_name=_optional_string(raw.get("drive_name")).strip(),
            folder_path=_optional_string(raw.get("folder_path")).strip(),
            auth=sharepoint_auth,
        )


@dataclass(frozen=True)
class OutputConfig:
    directory: Path
    csv_directory: Path
    json_directory: Path
    team_directories_file: str | None
    format: str
    min_csv_report_duration_seconds: int
    filename_pattern: str
    sharepoint_csv: SharePointCSVConfig

    @classmethod
    def from_mapping(
        cls,
        raw: dict[str, Any],
        *,
        base_auth: AuthConfig,
        min_csv_report_duration_seconds_override: int | None,
        clear_cache: bool,
    ) -> "OutputConfig":
        directory = Path(_require_non_empty_string(raw.get("directory", "./output"), "output.directory"))
        csv_directory_raw = raw.get("csv_directory")
        json_directory_raw = raw.get("json_directory")
        export_format = _require_string(raw.get("format", "both"), "output.format").strip().lower()
        if export_format not in OUTPUT_FORMATS:
            raise ValueError("output.format must be one of: both, csv, json")

        min_csv_report_duration_seconds = _optional_non_negative_int(
            min_csv_report_duration_seconds_override
            if min_csv_report_duration_seconds_override is not None
            else raw.get("min_csv_report_duration_seconds"),
            0,
            "output.min_csv_report_duration_seconds",
        )

        return cls(
            directory=directory,
            csv_directory=Path(csv_directory_raw) if csv_directory_raw else directory / "csv",
            json_directory=Path(json_directory_raw) if json_directory_raw else directory / "json",
            team_directories_file=_optional_string(raw.get("team_directories_file")).strip() or None,
            format=export_format,
            min_csv_report_duration_seconds=min_csv_report_duration_seconds,
            filename_pattern=_require_non_empty_string(
                raw.get("filename_pattern", "{team_name}_{channel_name}_{meeting_date}_{meeting_id}_{report_id}_attendance"),
                "output.filename_pattern",
            ),
            sharepoint_csv=SharePointCSVConfig.from_mapping(
                _ensure_mapping(raw.get("sharepoint_csv"), "output.sharepoint_csv"),
                base_auth=base_auth,
                clear_cache=clear_cache,
            ),
        )


@dataclass(frozen=True)
class CacheConfig:
    directory: Path
    metadata_cache: str

    @property
    def metadata_cache_file(self) -> Path:
        return self.directory / self.metadata_cache

    @classmethod
    def from_mapping(cls, raw: dict[str, Any]) -> "CacheConfig":
        return cls(
            directory=Path(_require_non_empty_string(raw.get("directory", "./cache"), "cache.directory")),
            metadata_cache=_require_non_empty_string(raw.get("metadata_cache", "teams_channels.json"), "cache.metadata_cache"),
        )


@dataclass(frozen=True)
class APIConfig:
    max_retries: int
    retry_backoff_factor: int
    timeout: int

    @classmethod
    def from_mapping(cls, raw: dict[str, Any]) -> "APIConfig":
        return cls(
            max_retries=_optional_non_negative_int(raw.get("max_retries"), 3, "api.max_retries"),
            retry_backoff_factor=_require_positive_int(raw.get("retry_backoff_factor", 2), "api.retry_backoff_factor"),
            timeout=_require_positive_int(raw.get("timeout", 30), "api.timeout"),
        )


@dataclass(frozen=True)
class AppConfig:
    auth: AuthConfig
    team_filter: TeamFilterConfig
    meetings: MeetingsConfig
    output: OutputConfig
    cache: CacheConfig
    api: APIConfig

    @classmethod
    def from_mapping(cls, raw: dict[str, Any], args: argparse.Namespace) -> "AppConfig":
        auth = AuthConfig.from_mapping(
            _ensure_mapping(raw.get("auth"), "auth"),
            default_scopes=PUBLIC_SCOPES_DEFAULT,
            clear_cache=args.clear_cache,
        )
        cache = CacheConfig.from_mapping(_ensure_mapping(raw.get("cache"), "cache"))

        return cls(
            auth=auth,
            team_filter=TeamFilterConfig.from_mapping(
                _ensure_mapping(raw.get("team_filter"), "team_filter"),
                regex_override=args.team_regex,
            ),
            meetings=MeetingsConfig.from_mapping(
                _ensure_mapping(raw.get("meetings"), "meetings"),
                lookback_days_override=args.lookback_days,
                lookahead_days_override=args.lookahead_days,
            ),
            output=OutputConfig.from_mapping(
                _ensure_mapping(raw.get("output"), "output"),
                base_auth=auth,
                min_csv_report_duration_seconds_override=args.min_csv_report_duration_seconds,
                clear_cache=args.clear_cache,
            ),
            cache=cache,
            api=APIConfig.from_mapping(_ensure_mapping(raw.get("api"), "api")),
        )


def load_app_config(config_path: str, args: argparse.Namespace) -> AppConfig:
    """Load, validate, and resolve the application configuration."""
    config_file = Path(config_path)
    if not config_file.exists():
        raise FileNotFoundError(f"Configuration file not found: {config_path}")

    with open(config_file, "r", encoding="utf-8") as file:
        raw = yaml.safe_load(file) or {}

    if not isinstance(raw, dict):
        raise ValueError("Configuration root must be a mapping")

    return AppConfig.from_mapping(raw, args)
