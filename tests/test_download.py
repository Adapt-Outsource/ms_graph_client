from os import PathLike
from pathlib import Path
from typing import cast

from adapt_sharepoint.graph import GraphHttpClient, SharePointGraphClient


class FakeResponse:
    def __init__(
        self,
        status_code: int,
        content: bytes = b"",
        location: str | None = None,
    ) -> None:
        self.status_code = status_code
        self.content = content
        self.headers: dict[str, str] = {}
        if location:
            self.headers["Location"] = location

    def raise_for_status(self) -> None:
        if self.status_code >= 400:
            raise RuntimeError(f"HTTP {self.status_code}")

    def json(self) -> dict[str, str]:
        return {}


class FakeClient:
    def __init__(self) -> None:
        self.calls: list[str] = []

    async def get(
        self,
        url: str,
        *,
        headers: dict[str, str] | None = None,
    ) -> FakeResponse:
        _ = headers
        self.calls.append(url)
        if url.endswith(":/content"):
            return FakeResponse(302, location="https://download.example/file")
        if url == "https://download.example/file":
            return FakeResponse(200, content=b"pdf-bytes")
        raise RuntimeError(f"Unexpected URL: {url}")

    async def put(
        self,
        url: str,
        *,
        headers: dict[str, str] | None = None,
        content: bytes = b"",
    ) -> FakeResponse:
        _ = (url, headers, content)
        raise RuntimeError("not used in this test")

    async def aclose(self) -> None:
        return None


async def test_download_file_by_path_follows_redirect(tmp_path: Path) -> None:
    subject = SharePointGraphClient("https://atc2021.sharepoint.com/sites/ATC-3RPA/Shared%20Documents/Forms/AllItems.aspx")
    fake_client = FakeClient()
    subject._client = cast(GraphHttpClient, fake_client)
    subject._headers = {"Authorization": "Bearer token"}
    subject._drive_id = "drive-id"

    class CustomPath(PathLike[str]):
        def __fspath__(self) -> str:
            return str(tmp_path)

    output = await subject.download_file_by_path(
        "test/เขัาห้องสอบ.pdf",
        CustomPath(),
    )

    assert output.exists()
    assert output.read_bytes() == b"pdf-bytes"
    assert output.name == "เขัาห้องสอบ.pdf"
    assert output.parent == tmp_path
    assert len(fake_client.calls) == 2


async def test_download_file_by_path_preserve_path_true(tmp_path: Path) -> None:
    subject = SharePointGraphClient("https://atc2021.sharepoint.com/sites/ATC-3RPA/Shared%20Documents/Forms/AllItems.aspx")
    fake_client = FakeClient()
    subject._client = cast(GraphHttpClient, fake_client)
    subject._headers = {"Authorization": "Bearer token"}
    subject._drive_id = "drive-id"

    output = await subject.download_file_by_path(
        "test/เขัาห้องสอบ.pdf",
        tmp_path,
        preserve_path=True,
    )

    assert output.exists()
    assert output.read_bytes() == b"pdf-bytes"
    assert output == tmp_path / "test" / "เขัาห้องสอบ.pdf"
