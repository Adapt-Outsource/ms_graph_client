from __future__ import annotations

from collections.abc import Mapping
from io import BytesIO
from os import PathLike, fspath
from pathlib import Path
from types import TracebackType
from typing import Any, Protocol, cast
from urllib.parse import quote

from azure.core.credentials import TokenCredential
from azure.identity import DefaultAzureCredential
from msgraph_core import APIVersion, GraphClientFactory

from .models import DriveEntry
from .utils import parse_sharepoint_url, resolve_drive


class GraphHttpResponse(Protocol):
    status_code: int
    headers: Mapping[str, str]
    content: bytes

    def raise_for_status(self) -> None: ...

    def json(self) -> dict[str, Any]: ...


class GraphHttpClient(Protocol):
    async def get(
        self,
        url: str,
        *,
        headers: dict[str, str] | None = None,
    ) -> GraphHttpResponse: ...

    async def put(
        self,
        url: str,
        *,
        headers: dict[str, str] | None = None,
        content: bytes = b"",
    ) -> GraphHttpResponse: ...

    async def aclose(self) -> None: ...


class SharePointGraphClient:
    """Reusable Microsoft Graph client for SharePoint document libraries."""

    def __init__(
        self,
        sharepoint_url: str,
        credential: TokenCredential | None = None,
    ) -> None:
        self.sharepoint_url = sharepoint_url
        self._credential: TokenCredential | None = credential
        self._owns_credential = credential is None
        self._client: GraphHttpClient | None = None
        self._headers: dict[str, str] | None = None
        self._drive_id: str | None = None

    async def __aenter__(self) -> SharePointGraphClient:
        if self._credential is None:
            self._credential = DefaultAzureCredential()

        assert self._credential is not None
        token = self._credential.get_token("https://graph.microsoft.com/.default")
        self._headers = {"Authorization": f"Bearer {token.token}"}
        self._client = cast(
            GraphHttpClient,
            GraphClientFactory.create_with_default_middleware(api_version=APIVersion.v1),
        )
        await self._initialize_drive()
        return self

    async def __aexit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        tb: TracebackType | None,
    ) -> None:
        if self._client:
            await self._client.aclose()
        if self._credential and self._owns_credential:
            close_method = getattr(self._credential, "close", None)
            if callable(close_method):
                close_method()

    async def list_items(self) -> list[DriveEntry]:
        self._ensure_ready()
        assert self._drive_id is not None
        return await self._walk_drive(self._drive_id)

    async def download_file_by_path(
        self,
        file_path: str,
        output_dir: str | PathLike[str],
        preserve_path: bool = False,
    ) -> Path:
        self._ensure_ready()

        output_dir_path = Path(fspath(output_dir))
        content = await self.download_file_bytes_by_path(file_path)

        if preserve_path:
            normalized = file_path.strip("/\\").replace("\\", "/")
            local_path = output_dir_path.joinpath(*normalized.split("/"))
        else:
            local_path = output_dir_path / Path(file_path).name
        local_path.parent.mkdir(parents=True, exist_ok=True)
        local_path.write_bytes(content)
        return local_path

    async def download_file_bytes_by_path(self, file_path: str) -> bytes:
        """Download a SharePoint file by path and return its content as bytes."""
        self._ensure_ready()
        assert self._drive_id is not None

        encoded_path = quote(file_path.strip("/"), safe="/")
        return await self._download_content(
            f"/drives/{self._drive_id}/root:/{encoded_path}:/content"
        )

    async def download_file_stream_by_path(self, file_path: str) -> BytesIO:
        """Download a SharePoint file by path and return it as an in-memory stream."""
        content = await self.download_file_bytes_by_path(file_path)
        return BytesIO(content)

    async def upload_file_by_path(
        self,
        local_file: str | PathLike[str],
        remote_path: str,
    ) -> dict[str, Any]:
        self._ensure_ready()

        local_file_path = Path(fspath(local_file))

        if not local_file_path.exists() or not local_file_path.is_file():
            raise FileNotFoundError(f"Local file not found: {local_file_path}")

        return await self.upload_file_bytes_by_path(
            content=local_file_path.read_bytes(),
            remote_path=remote_path,
        )

    async def upload_file_bytes_by_path(
        self,
        content: bytes,
        remote_path: str,
    ) -> dict[str, Any]:
        """Upload raw bytes to a SharePoint path."""
        self._ensure_ready()
        assert self._client is not None
        assert self._headers is not None
        assert self._drive_id is not None

        encoded_path = quote(remote_path.strip("/"), safe="/")
        headers = {**self._headers, "Content-Type": "application/octet-stream"}
        response = await self._client.put(
            f"/drives/{self._drive_id}/root:/{encoded_path}:/content",
            headers=headers,
            content=content,
        )
        response.raise_for_status()
        return response.json()

    async def upload_file_stream_by_path(
        self,
        stream: BytesIO,
        remote_path: str,
    ) -> dict[str, Any]:
        """Upload a BytesIO stream to a SharePoint path."""
        stream.seek(0)
        return await self.upload_file_bytes_by_path(
            content=stream.read(),
            remote_path=remote_path,
        )

    async def _initialize_drive(self) -> None:
        assert self._client is not None
        assert self._headers is not None

        host, site_path, library_name = parse_sharepoint_url(self.sharepoint_url)
        site = await self._get_json(f"/sites/{host}:{site_path}")
        site_id = site["id"]

        drives = await self._get_all_pages(f"/sites/{site_id}/drives?$top=200")
        drive = resolve_drive(drives, library_name)
        self._drive_id = drive["id"]

    async def _walk_drive(
        self,
        drive_id: str,
        item_id: str | None = None,
        parent_path: str = "",
    ) -> list[DriveEntry]:
        if item_id:
            children_url = f"/drives/{drive_id}/items/{item_id}/children?$top=200"
        else:
            children_url = f"/drives/{drive_id}/root/children?$top=200"

        children = await self._get_all_pages(children_url)
        results: list[DriveEntry] = []

        for item in children:
            name = item.get("name", "")
            full_path = f"{parent_path}/{name}" if parent_path else name
            is_folder = "folder" in item

            results.append(
                DriveEntry(
                    type="folder" if is_folder else "file",
                    path=full_path,
                    size=item.get("size", 0),
                    id=item.get("id", ""),
                    web_url=item.get("webUrl", ""),
                )
            )

            if is_folder:
                child_id = item.get("id")
                if isinstance(child_id, str) and child_id:
                    results.extend(
                        await self._walk_drive(drive_id, child_id, full_path)
                    )

        return results

    async def _get_json(self, url: str) -> dict[str, Any]:
        assert self._client is not None
        assert self._headers is not None

        response = await self._client.get(url, headers=self._headers)
        response.raise_for_status()
        return response.json()

    async def _get_all_pages(self, url: str) -> list[dict[str, Any]]:
        items: list[dict[str, Any]] = []
        next_url: str | None = url

        while next_url:
            data = await self._get_json(next_url)
            value = data.get("value", [])
            if isinstance(value, list):
                items.extend(x for x in value if isinstance(x, dict))
            next_link = data.get("@odata.nextLink")
            next_url = next_link if isinstance(next_link, str) else None

        return items

    async def _download_content(self, url: str) -> bytes:
        assert self._client is not None
        assert self._headers is not None

        response = await self._client.get(url, headers=self._headers)

        if (
            response.status_code in {301, 302, 303, 307, 308}
            and response.headers.get("Location")
        ):
            response = await self._client.get(response.headers["Location"])

        response.raise_for_status()
        return response.content

    def _ensure_ready(self) -> None:
        if not self._client or not self._headers:
            raise RuntimeError(
                "Client not initialized. "
                "Use 'async with SharePointGraphClient(...)'."
            )
