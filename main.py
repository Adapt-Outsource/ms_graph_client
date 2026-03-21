import asyncio
from pathlib import Path
from typing import Any
from urllib.parse import quote
from urllib.parse import unquote, urlparse

from azure.identity import AzureCliCredential
from msgraph_core import APIVersion, GraphClientFactory


SHAREPOINT_URL = "https://atc2021.sharepoint.com/sites/ATC-3RPA/Shared%20Documents/Forms/AllItems.aspx"
TEST_FILE_PATH = "test/เขัาห้องสอบ.pdf"
DOWNLOAD_DIR = Path("downloads")


def parse_sharepoint_url(sharepoint_url: str) -> tuple[str, str, str]:
	parsed = urlparse(sharepoint_url)
	host = parsed.netloc
	parts = [p for p in unquote(parsed.path).split("/") if p]

	if len(parts) < 3 or parts[0].lower() != "sites":
		raise ValueError("Expected a SharePoint site URL like /sites/<site>/<library>/...")

	site_path = f"/{parts[0]}/{parts[1]}"
	library_name = parts[2]
	return host, site_path, library_name


async def get_json(client, url: str, headers: dict[str, str]) -> dict[str, Any]:
	response = await client.get(url, headers=headers)
	response.raise_for_status()
	return response.json()


async def get_all_pages(client, url: str, headers: dict[str, str]) -> list[dict[str, Any]]:
	items: list[dict[str, Any]] = []
	next_url = url

	while next_url:
		data = await get_json(client, next_url, headers)
		items.extend(data.get("value", []))
		next_url = data.get("@odata.nextLink")

	return items


async def walk_drive(client, drive_id: str, headers: dict[str, str], item_id: str | None = None, parent_path: str = "") -> list[dict[str, Any]]:
	if item_id:
		children_url = f"/drives/{drive_id}/items/{item_id}/children?$top=200"
	else:
		children_url = f"/drives/{drive_id}/root/children?$top=200"

	children = await get_all_pages(client, children_url, headers)
	results: list[dict[str, Any]] = []

	for item in children:
		name = item.get("name", "")
		full_path = f"{parent_path}/{name}" if parent_path else name
		is_folder = "folder" in item

		results.append(
			{
				"type": "folder" if is_folder else "file",
				"path": full_path,
				"size": item.get("size", 0),
				"id": item.get("id", ""),
				"webUrl": item.get("webUrl", ""),
			}
		)

		if is_folder:
			results.extend(
				await walk_drive(client, drive_id, headers, item.get("id"), full_path)
			)

	return results


def resolve_drive(drives: list[dict[str, Any]], library_name: str) -> dict[str, Any]:
	drive = next(
		(
			d
			for d in drives
			if d.get("name", "").lower()
			in {library_name.lower(), "documents", "shared documents"}
		),
		None,
	)

	if not drive:
		available = ", ".join(d.get("name", "<unknown>") for d in drives)
		raise RuntimeError(
			f"Document library '{library_name}' was not found. Available drives: {available}"
		)

	return drive


async def download_file_by_path(
	client,
	drive_id: str,
	headers: dict[str, str],
	file_path: str,
	output_dir: Path,
) -> Path:
	encoded_path = quote(file_path.strip("/"), safe="/")
	response = await client.get(
		f"/drives/{drive_id}/root:/{encoded_path}:/content",
		headers=headers,
	)
	if response.status_code in {301, 302, 303, 307, 308} and response.headers.get("Location"):
		response = await client.get(response.headers["Location"])
	response.raise_for_status()

	output_dir.mkdir(parents=True, exist_ok=True)
	local_path = output_dir / Path(file_path).name
	local_path.write_bytes(response.content)
	return local_path


async def main() -> None:
	host, site_path, library_name = parse_sharepoint_url(SHAREPOINT_URL)

	credential = AzureCliCredential()
	token = credential.get_token("https://graph.microsoft.com/.default")
	headers = {"Authorization": f"Bearer {token.token}"}

	client = GraphClientFactory.create_with_default_middleware(api_version=APIVersion.v1)

	try:
		site = await get_json(client, f"/sites/{host}:{site_path}", headers)
		site_id = site["id"]

		drives = await get_all_pages(client, f"/sites/{site_id}/drives?$top=200", headers)
		drive = resolve_drive(drives, library_name)

		items = await walk_drive(client, drive["id"], headers)

		for item in items:
			print(f"{item['type']:6} {item['path']}")

		print(f"\nTotal items: {len(items)}")

		downloaded_file = await download_file_by_path(
			client=client,
			drive_id=drive["id"],
			headers=headers,
			file_path=TEST_FILE_PATH,
			output_dir=DOWNLOAD_DIR,
		)
		print(f"Downloaded test file to: {downloaded_file}")

	finally:
		await client.aclose()
		credential.close()


if __name__ == "__main__":
	asyncio.run(main())