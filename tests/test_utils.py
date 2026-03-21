from adapt_sharepoint.utils import parse_sharepoint_url, resolve_drive


def test_parse_sharepoint_url() -> None:
    host, site_path, library_name = parse_sharepoint_url(
        "https://atc2021.sharepoint.com/sites/ATC-3RPA/Shared%20Documents/Forms/AllItems.aspx"
    )

    assert host == "atc2021.sharepoint.com"
    assert site_path == "/sites/ATC-3RPA"
    assert library_name == "Shared Documents"


def test_resolve_drive_prefers_requested_library() -> None:
    drives = [
        {"name": "Documents", "id": "1"},
        {"name": "Shared Documents", "id": "2"},
    ]

    drive = resolve_drive(drives, "Shared Documents")
    assert drive["id"] == "2"
