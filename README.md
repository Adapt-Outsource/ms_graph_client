# adapt-sharepoint

Reusable Python library for SharePoint via Microsoft Graph.

This package is designed for import and use inside other Python projects.
The caller provides all runtime parameters (site URL, file paths, output paths) directly in code.

## What This Package Does

- Connects to SharePoint document libraries using Microsoft Entra credentials (DefaultAzureCredential by default).
- Lists folders and files recursively.
- Downloads a file from a SharePoint path.
- Downloads a file from a SharePoint path directly into memory.
- Uploads a local file, raw bytes, or a BytesIO stream to a SharePoint path.

## Authentication Model

This package supports credential injection.

Default behavior:

- If caller does not pass credential, package uses DefaultAzureCredential.

Recommended usage:

- Azure Function App: pass DefaultAzureCredential or rely on default behavior.
- Local machine: sign in with az login, then DefaultAzureCredential can use Azure CLI identity.

## Prerequisites

1. Python 3.11+
2. Azure CLI installed
3. Logged in to Azure CLI with an account that can access the SharePoint site

Login command:

```bash
az login
```

## Installation

### Option 1: Install from package index

```bash
uv add adapt-sharepoint
```

### Option 2: Install directly from GitHub

```bash
uv add git+https://github.com/Adapt-Outsource/ms_graph_client.git
```

### Option 3: Local development install (inside this repo)

```bash
uv sync
```

## Import Path

Use this import in your project code:

```python
from adapt_sharepoint import SharePointGraphClient
```

Constructor:

- `SharePointGraphClient(sharepoint_url: str, credential: TokenCredential | None = None)`

## Parameters You Must Pass

The caller must provide:

1. `sharepoint_url`: URL of the target SharePoint document library page
2. `file_path`: path in SharePoint library for download
3. `output_dir`: local folder for downloaded file
4. `preserve_path`: whether to keep SharePoint folder structure on local disk (default: `False`)
5. `local_file`: local file path for upload when using `upload_file_by_path`
6. `content`: raw bytes for upload when using `upload_file_bytes`
7. `stream`: `BytesIO` upload stream when using `upload_file_stream`
8. `remote_path`: destination path in SharePoint library for upload

### Example URL format

```text
https://your-tenant.sharepoint.com/sites/YourSite/Shared%20Documents/Forms/AllItems.aspx
```

### Example SharePoint file path format

```text
folder/subfolder/report.pdf
```

## Download Path Behavior

When calling `download_file_by_path`, behavior depends on `preserve_path`:

- `preserve_path=False` (default): save by filename only into `output_dir`
- `preserve_path=True`: preserve full SharePoint sub-path under `output_dir`

Example:

- `file_path="incoming/monthly/report.pdf"`
- `output_dir="downloads"`

Results:

- `preserve_path=False` -> `downloads/report.pdf`
- `preserve_path=True` -> `downloads/incoming/monthly/report.pdf`

## Step-By-Step Usage In Another Project

### Step 1: Add package dependency

```bash
uv add adapt-sharepoint
```

### Step 2: Write code and pass parameters from your app

```python
import asyncio
from io import BytesIO
from pathlib import Path

from azure.identity import DefaultAzureCredential
from adapt_sharepoint import SharePointGraphClient


async def main() -> None:
    sharepoint_url = "https://your-tenant.sharepoint.com/sites/YourSite/Shared%20Documents/Forms/AllItems.aspx"
    remote_download_path = "incoming/monthly/report.pdf"
    local_download_dir = Path("downloads")
    local_upload_file = Path("downloads/report.pdf")
    remote_upload_path = "archive/2026/report-copy.pdf"
    credential = DefaultAzureCredential()

    async with SharePointGraphClient(sharepoint_url, credential=credential) as client:
        # 1) List all entries
        entries = await client.list_items()
        for entry in entries:
            print(entry.type, entry.path)

        # 2) Download one file to local disk
        downloaded_file = await client.download_file_by_path(
            file_path=remote_download_path,
            output_dir=local_download_dir,
            preserve_path=False,
        )
        print("Downloaded:", downloaded_file)

        # 3) Download one file as raw bytes (in memory)
        file_bytes = await client.download_file_bytes_by_path(remote_download_path)
        print("Downloaded bytes:", len(file_bytes))

        # 4) Download one file as BytesIO stream (in memory)
        stream = await client.download_file_stream_by_path(remote_download_path)
        print("Stream size:", len(stream.getvalue()))

        # 5) Upload one local file
        uploaded = await client.upload_file_by_path(
            local_file=local_upload_file,
            remote_path=remote_upload_path,
        )
        print("Uploaded file:", uploaded.get("name", "<unknown>"))

        # 6) Upload raw bytes
        uploaded_bytes = await client.upload_file_bytes(
            content=b"hello from memory",
            remote_path="archive/2026/bytes-example.txt",
        )
        print("Uploaded bytes:", uploaded_bytes.get("name", "<unknown>"))

        # 7) Upload BytesIO stream
        uploaded_stream = await client.upload_file_stream(
            stream=BytesIO(b"hello from stream"),
            remote_path="archive/2026/stream-example.txt",
        )
        print("Uploaded stream:", uploaded_stream.get("name", "<unknown>"))


if __name__ == "__main__":
    asyncio.run(main())
```

## Azure Function App Pattern

Minimal pattern:

1. Enable Managed Identity on Function App.
2. Grant Microsoft Graph permissions required for SharePoint access.
3. Create DefaultAzureCredential in function code.
4. Pass credential into SharePointGraphClient.

Example snippet:

```python
from azure.identity import DefaultAzureCredential
from adapt_sharepoint import SharePointGraphClient

credential = DefaultAzureCredential()
client = SharePointGraphClient(sharepoint_url, credential=credential)
```

### Step 3: Run your script

```bash
uv run app.py
```

## API Reference (Core)

Class:

- `SharePointGraphClient(sharepoint_url: str, credential: TokenCredential | None = None)`

Methods:

- `await list_items() -> list[DriveEntry]`
- `await download_file_by_path(file_path: str, output_dir: str | PathLike[str], preserve_path: bool = False) -> Path`
- `await download_file_bytes_by_path(file_path: str) -> bytes`
- `await download_file_stream_by_path(file_path: str) -> BytesIO`
- `await upload_file_by_path(local_file: str | PathLike[str], remote_path: str) -> dict[str, Any]`
- `await upload_file_bytes(content: bytes, remote_path: str) -> dict[str, Any]`
- `await upload_file_stream(stream: BytesIO, remote_path: str) -> dict[str, Any]`
- `await delete_item_by_path(item_path: str) -> None`
- `await delete_file_by_path(file_path: str) -> None`
- `await delete_directory_by_path(directory_path: str) -> None`

Download behavior:

- `preserve_path=False` (default): save file directly in `output_dir` using filename only.
- `preserve_path=True`: keep full SharePoint subfolder path under `output_dir`.
- Deleting a directory removes the folder and all nested files/folders.

## Error Notes

- If login is missing or expired: run `az login` again.
- If access is denied (403): your account lacks permission to that site/library.
- If file is not found (404): verify `file_path` or `remote_path` is correct.

## Quality Checks (Recommended Before Release)

Run lint:

```bash
uv run --with ruff ruff check .
```

Run type check:

```bash
uv run --with mypy mypy src tests
```

Run tests:

```bash
uv run --with pytest --with pytest-asyncio pytest -q
```

## Build And Publish

Build package artifacts:

```bash
uv build
```

Publish package:

```bash
uv publish
```
