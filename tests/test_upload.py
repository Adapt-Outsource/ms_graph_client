from pathlib import Path
from typing import cast

from adapt_sharepoint.graph import GraphHttpClient, SharePointGraphClient


class FakeResponse:
    def __init__(
        self,
        status_code: int,
        payload: dict[str, str] | None = None,
    ) -> None:
        self.status_code = status_code
        self._payload = payload or {}
        self.headers: dict[str, str] = {}
        self.content: bytes = b""

    def raise_for_status(self) -> None:
        if self.status_code >= 400:
            raise RuntimeError(f"HTTP {self.status_code}")

    def json(self) -> dict[str, str]:
        return self._payload


class FakeClient:
    def __init__(self) -> None:
        self.calls: list[tuple[str, bytes]] = []

    async def get(
        self,
        url: str,
        *,
        headers: dict[str, str] | None = None,
    ) -> FakeResponse:
        _ = (url, headers)
        raise RuntimeError("not used in this test")

    async def put(
        self,
        url: str,
        *,
        headers: dict[str, str] | None = None,
        content: bytes = b"",
    ) -> FakeResponse:
        _ = headers
        self.calls.append((url, content))
        return FakeResponse(
            201,
            {"name": "uploaded.pdf", "webUrl": "https://example/uploaded.pdf"},
        )

    async def aclose(self) -> None:
        return None


async def test_upload_file_by_path_sends_content(tmp_path: Path) -> None:
    source = tmp_path / "source.pdf"
    source.write_bytes(b"hello")

    subject = SharePointGraphClient("https://atc2021.sharepoint.com/sites/ATC-3RPA/Shared%20Documents/Forms/AllItems.aspx")
    fake_client = FakeClient()
    subject._client = cast(GraphHttpClient, fake_client)
    subject._headers = {"Authorization": "Bearer token"}
    subject._drive_id = "drive-id"

    result = await subject.upload_file_by_path(source, "target/folder/source.pdf")

    assert result["name"] == "uploaded.pdf"
    assert len(fake_client.calls) == 1
    assert fake_client.calls[0][0].endswith("/root:/target/folder/source.pdf:/content")
    assert fake_client.calls[0][1] == b"hello"
