from typing import Any
from urllib.parse import unquote, urlparse


def parse_sharepoint_url(sharepoint_url: str) -> tuple[str, str, str]:
    parsed = urlparse(sharepoint_url)
    host = parsed.netloc
    parts = [p for p in unquote(parsed.path).split("/") if p]

    if len(parts) < 3 or parts[0].lower() != "sites":
        raise ValueError("Expected URL like /sites/<site>/<library>/...")

    site_path = f"/{parts[0]}/{parts[1]}"
    library_name = parts[2]
    return host, site_path, library_name


def resolve_drive(drives: list[dict[str, Any]], library_name: str) -> dict[str, Any]:
    preferred_name = library_name.lower()
    drive = next(
        (d for d in drives if d.get("name", "").lower() == preferred_name),
        None,
    )

    if not drive:
        fallback_names = {"documents", "shared documents"}
        drive = next(
            (d for d in drives if d.get("name", "").lower() in fallback_names),
            None,
        )

    if not drive:
        available = ", ".join(d.get("name", "<unknown>") for d in drives)
        raise RuntimeError(
            "Document library "
            f"'{library_name}' was not found. Available drives: {available}"
        )

    return drive
