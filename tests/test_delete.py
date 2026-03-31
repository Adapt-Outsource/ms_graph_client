from typing import Any, cast

from adapt_sharepoint.graph import GraphHttpClient, SharePointGraphClient


class FakeResponse:
    def __init__(self, status_code: int = 204) -> None:
        self.status_code = status_code
        self.headers: dict[str, str] = {}
        self.content: bytes = b""

    def raise_for_status(self) -> None:
        if self.status_code >= 400:
            raise RuntimeError(f"HTTP {self.status_code}")

    def json(self) -> dict[str, Any]:
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
        _ = (url, headers)
        raise RuntimeError("not used in this test")

    async def put(
        self,
        url: str,
        *,
        headers: dict[str, str] | None = None,
        content: bytes = b"",
    ) -> FakeResponse:
        _ = (url, headers, content)
        raise RuntimeError("not used in this test")

    async def delete(
        self,
        url: str,
        *,
        headers: dict[str, str] | None = None,
    ) -> FakeResponse:
        _ = headers
        self.calls.append(url)
        return FakeResponse(204)

    async def aclose(self) -> None:
        return None


async def test_delete_file_by_path_calls_graph_delete() -> None:
    subject = SharePointGraphClient("https://atc2021.sharepoint.com/sites/ATC-3RPA/Shared%20Documents/Forms/AllItems.aspx")
    fake_client = FakeClient()
    subject._client = cast(GraphHttpClient, fake_client)
    subject._headers = {"Authorization": "Bearer token"}
    subject._drive_id = "drive-id"

    await subject.delete_file_by_path("docs/report.pdf")

    assert fake_client.calls == ["/drives/drive-id/root:/docs/report.pdf"]


async def test_delete_directory_by_path_encodes_unicode_and_nested_path() -> None:
    subject = SharePointGraphClient("https://atc2021.sharepoint.com/sites/ATC-3RPA/Shared%20Documents/Forms/AllItems.aspx")
    fake_client = FakeClient()
    subject._client = cast(GraphHttpClient, fake_client)
    subject._headers = {"Authorization": "Bearer token"}
    subject._drive_id = "drive-id"

    await subject.delete_directory_by_path("test/เขัาห้องสอบ")

    expected_suffix = (
        "/drives/drive-id/root:/test/"
        "%E0%B9%80%E0%B8%82%E0%B8%B1%E0%B8%B2%E0%B8%AB%E0%B9%89"
        "%E0%B8%AD%E0%B8%87%E0%B8%AA%E0%B8%AD%E0%B8%9A"
    )
    assert fake_client.calls[0] == expected_suffix


async def test_delete_item_by_path_rejects_empty_path() -> None:
    subject = SharePointGraphClient("https://atc2021.sharepoint.com/sites/ATC-3RPA/Shared%20Documents/Forms/AllItems.aspx")
    fake_client = FakeClient()
    subject._client = cast(GraphHttpClient, fake_client)
    subject._headers = {"Authorization": "Bearer token"}
    subject._drive_id = "drive-id"

    try:
        await subject.delete_item_by_path("/")
        raise AssertionError("Expected ValueError for empty path")
    except ValueError as exc:
        assert "must not be empty" in str(exc)

    assert fake_client.calls == []
