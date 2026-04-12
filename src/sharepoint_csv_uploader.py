"""
Upload CSV exports to a SharePoint document library via Microsoft Graph.
"""
import logging
from pathlib import Path, PurePosixPath
from urllib.parse import quote

from .graph_client import GraphAPIError, GraphClient

logger = logging.getLogger(__name__)


class SharePointCSVUploader:
    """Uploads CSV files to a SharePoint drive."""

    def __init__(
        self,
        graph_client: GraphClient,
        site_id: str | None = None,
        site_hostname: str | None = None,
        site_path: str | None = None,
        drive_id: str | None = None,
        drive_name: str | None = None,
        folder_path: str | None = None,
    ):
        self.client = graph_client
        self.site_id = site_id.strip() if site_id else ""
        self.site_hostname = site_hostname.strip() if site_hostname else ""
        self.site_path = site_path.strip("/") if site_path else ""
        self.drive_id = drive_id.strip() if drive_id else ""
        self.drive_name = (drive_name or "Documents").strip()
        self.folder_path = folder_path.strip("/") if folder_path else ""

    @staticmethod
    def _apply_top_level_folder_suffix(relative_path: Path, folder_suffix: str) -> PurePosixPath:
        """Append a suffix to the top-level folder in a relative upload path."""
        remote_relative_path = PurePosixPath(relative_path.as_posix())
        if not folder_suffix or not remote_relative_path.parts:
            return remote_relative_path

        first_part, *remaining_parts = remote_relative_path.parts
        suffixed_first_part = f"{first_part}{folder_suffix}"
        return PurePosixPath(suffixed_first_part, *remaining_parts)

    def _resolve_site_id(self) -> str:
        """Resolve SharePoint site ID from config."""
        if self.site_id:
            return self.site_id

        if not self.site_hostname or not self.site_path:
            raise ValueError("SharePoint CSV upload requires either site_id or both site_hostname and site_path")

        response = self.client._make_request("GET", f"/sites/{self.site_hostname}:/{self.site_path}")
        site = response.json()
        site_id = site.get("id")
        if not site_id:
            raise GraphAPIError("SharePoint site lookup returned no site id")

        self.site_id = site_id
        return site_id

    def _resolve_drive_id(self) -> str:
        """Resolve document library drive ID from config."""
        if self.drive_id:
            return self.drive_id

        site_id = self._resolve_site_id()
        drives = self.client._paginate(f"/sites/{site_id}/drives")
        for drive in drives:
            if str(drive.get("name", "")).strip().lower() == self.drive_name.lower():
                drive_id = drive.get("id")
                if drive_id:
                    self.drive_id = drive_id
                    return drive_id

        available_drive_names = sorted(
            str(drive.get("name", "")).strip() for drive in drives if str(drive.get("name", "")).strip()
        )
        available_fragment = ", ".join(available_drive_names) if available_drive_names else "none"
        raise GraphAPIError(
            f"SharePoint drive '{self.drive_name}' not found in site {site_id}. "
            f"Available drives: {available_fragment}"
        )

    def _resolve_root_folder(self) -> tuple[str, PurePosixPath]:
        """Resolve the configured SharePoint root folder for uploads and scans."""
        drive_id = self._resolve_drive_id()
        root_folder = PurePosixPath(self.folder_path)
        return drive_id, root_folder

    def _list_children(self, drive_id: str, parent_item_id: str) -> list[dict]:
        """List children of a drive item."""
        if parent_item_id == "root":
            return self.client._paginate(f"/drives/{drive_id}/root/children")
        return self.client._paginate(f"/drives/{drive_id}/items/{parent_item_id}/children")

    def _get_item_by_path(self, drive_id: str, relative_path: PurePosixPath) -> dict | None:
        """Return a drive item by relative path, or None when not found."""
        normalized = relative_path.as_posix().strip("/")
        if not normalized:
            response = self.client._make_request("GET", f"/drives/{drive_id}/root")
            return response.json()

        response = self.client._make_request(
            "GET",
            f"/drives/{drive_id}/root:/{quote(normalized, safe='/')}"
        )
        if response.status_code == 404:
            return None
        return response.json()

    def _create_folder(self, drive_id: str, parent_item_id: str, folder_name: str) -> dict:
        """Create a folder inside the specified drive item."""
        if parent_item_id == "root":
            endpoint = f"/drives/{drive_id}/root/children"
        else:
            endpoint = f"/drives/{drive_id}/items/{parent_item_id}/children"

        response = self.client._make_request(
            "POST",
            endpoint,
            json={
                "name": folder_name,
                "folder": {},
                "@microsoft.graph.conflictBehavior": "fail"
            }
        )
        return response.json()

    def _ensure_folder_path(self, drive_id: str, relative_folder: PurePosixPath) -> str:
        """Ensure a nested folder path exists and return the final item id."""
        parent_item_id = "root"

        for part in relative_folder.parts:
            if part in {"", "."}:
                continue

            children = self._list_children(drive_id, parent_item_id)
            existing = next(
                (
                    child for child in children
                    if child.get("name") == part and isinstance(child.get("folder"), dict)
                ),
                None
            )

            if existing:
                parent_item_id = existing["id"]
                continue

            try:
                created = self._create_folder(drive_id, parent_item_id, part)
            except GraphAPIError:
                # Handle concurrent or pre-existing folder creation by refetching once.
                children = self._list_children(drive_id, parent_item_id)
                existing = next(
                    (
                        child for child in children
                        if child.get("name") == part and isinstance(child.get("folder"), dict)
                    ),
                    None
                )
                if not existing:
                    raise
                parent_item_id = existing["id"]
                continue

            parent_item_id = created["id"]

        return parent_item_id

    def find_files_by_name(self, filename: str) -> list[dict]:
        """Recursively list files with the given name under the configured SharePoint folder."""
        drive_id, root_folder = self._resolve_root_folder()
        root_item = self._get_item_by_path(drive_id, root_folder)
        if not root_item:
            return []

        matches: list[dict] = []
        stack: list[tuple[dict, PurePosixPath]] = [(root_item, root_folder)]

        while stack:
            current_item, current_path = stack.pop()
            if not isinstance(current_item.get("folder"), dict):
                continue

            children = self._list_children(drive_id, current_item["id"])
            for child in children:
                child_path = current_path / child["name"]
                if isinstance(child.get("folder"), dict):
                    stack.append((child, child_path))
                    continue

                if child.get("name") == filename:
                    child_copy = dict(child)
                    child_copy["_relative_path"] = child_path
                    child_copy["_parent_relative_path"] = current_path
                    matches.append(child_copy)

        return matches

    def folder_contains_name(self, folder_relative_path: PurePosixPath, filename: str) -> bool:
        """Return True if the specified SharePoint folder contains a child with the given name."""
        drive_id, _ = self._resolve_root_folder()
        folder_item = self._get_item_by_path(drive_id, folder_relative_path)
        if not folder_item:
            return False

        children = self._list_children(drive_id, folder_item["id"])
        return any(child.get("name") == filename for child in children)

    def download_file_content(self, relative_path: PurePosixPath) -> bytes:
        """Download a SharePoint file by path."""
        drive_id, _ = self._resolve_root_folder()
        normalized = relative_path.as_posix().strip("/")
        response = self.client._make_request(
            "GET",
            f"/drives/{drive_id}/root:/{quote(normalized, safe='/')}:/content"
        )
        return response.content

    def create_empty_file(self, folder_relative_path: PurePosixPath, filename: str) -> str:
        """Create an empty file in the specified SharePoint folder."""
        drive_id, _ = self._resolve_root_folder()
        parent_item_id = self._ensure_folder_path(drive_id, folder_relative_path)
        encoded_name = quote(filename, safe="")
        response = self.client._make_request(
            "PUT",
            f"/drives/{drive_id}/items/{parent_item_id}:/{encoded_name}:/content",
            headers={"Content-Type": "application/octet-stream"},
            data=b""
        )
        item = response.json()
        return item.get("webUrl", "")

    def upload_file(
        self,
        local_path: Path,
        relative_path: Path,
        *,
        top_level_folder_suffix: str = "",
        create_remote_folders: bool = True,
    ) -> str:
        """Upload a local CSV file to SharePoint."""
        drive_id = self._resolve_drive_id()
        remote_relative_path = self._apply_top_level_folder_suffix(
            relative_path,
            top_level_folder_suffix
        )
        remote_path = PurePosixPath(self.folder_path) / remote_relative_path
        parent_folder = remote_path.parent
        if create_remote_folders:
            parent_item_id = self._ensure_folder_path(drive_id, parent_folder)
        else:
            parent_item = self._get_item_by_path(drive_id, parent_folder)
            if not parent_item or not isinstance(parent_item.get("folder"), dict):
                logger.warning(
                    "Skipping SharePoint upload for %s: target folder does not exist: %s",
                    local_path,
                    parent_folder.as_posix()
                )
                return ""
            parent_item_id = parent_item["id"]

        with open(local_path, "rb") as f:
            content = f.read()

        encoded_name = quote(remote_path.name, safe="")
        response = self.client._make_request(
            "PUT",
            f"/drives/{drive_id}/items/{parent_item_id}:/{encoded_name}:/content",
            headers={"Content-Type": "text/csv"},
            data=content
        )
        item = response.json()
        web_url = item.get("webUrl", "")
        logger.debug("Uploaded CSV to SharePoint: %s", web_url or remote_path.as_posix())

        return web_url

    def upload_files(
        self,
        file_paths: list[Path],
        local_csv_root: Path,
        *,
        top_level_folder_suffix: str = "",
        create_remote_folders: bool = True,
    ) -> list[str]:
        """Upload multiple CSV files preserving their relative directory structure."""
        uploaded_urls: list[str] = []
        root = local_csv_root.resolve()

        for file_path in file_paths:
            resolved = file_path.resolve()
            try:
                relative_path = resolved.relative_to(root)
            except ValueError:
                relative_path = Path(resolved.name)

            uploaded_url = self.upload_file(
                resolved,
                relative_path,
                top_level_folder_suffix=top_level_folder_suffix,
                create_remote_folders=create_remote_folders
            )
            if uploaded_url:
                uploaded_urls.append(uploaded_url)

        return uploaded_urls
