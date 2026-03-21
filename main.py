import asyncio
import os
from pathlib import Path

from adapt_sharepoint import SharePointGraphClient


async def run_example() -> None:
    sharepoint_url = os.getenv(
        "SHAREPOINT_URL",
        "https://contoso.sharepoint.com/sites/TeamSite/Shared%20Documents/Forms/AllItems.aspx",
    )
    remote_file_path = os.getenv("REMOTE_FILE_PATH", "samples/report.pdf")
    upload_target_path = os.getenv("UPLOAD_TARGET_PATH", "uploads/report-copy.pdf")

    async with SharePointGraphClient(sharepoint_url) as client:
        items = await client.list_items()
        for item in items:
            print(f"{item.type:6} {item.path}")

        local_file = await client.download_file_by_path(
            file_path=remote_file_path,
            output_dir=Path("downloads"),
        )
        print(f"Downloaded to: {local_file}")

        uploaded = await client.upload_file_by_path(
            local_file=local_file,
            remote_path=upload_target_path,
        )
        print(f"Uploaded: {uploaded.get('name', '<unknown>')}")


if __name__ == "__main__":
    asyncio.run(run_example())