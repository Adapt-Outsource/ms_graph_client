from __future__ import annotations

import argparse
import asyncio
from pathlib import Path

from .graph import SharePointGraphClient


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="SharePoint Graph helper")
    parser.add_argument("--url", required=True, help="SharePoint documents URL")

    subparsers = parser.add_subparsers(dest="command", required=True)

    list_parser = subparsers.add_parser(
        "list",
        help="List all files and folders",
    )
    list_parser.add_argument(
        "--only-files",
        action="store_true",
        help="Print files only",
    )

    download_parser = subparsers.add_parser("download", help="Download file by path")
    download_parser.add_argument(
        "--file-path",
        required=True,
        help="Path inside document library",
    )
    download_parser.add_argument(
        "--output-dir",
        default="downloads",
        help="Directory to save downloaded file",
    )

    upload_parser = subparsers.add_parser(
        "upload",
        help="Upload local file to SharePoint path",
    )
    upload_parser.add_argument(
        "--local-file",
        required=True,
        help="Local file path to upload",
    )
    upload_parser.add_argument(
        "--remote-path",
        required=True,
        help="Destination path inside document library",
    )

    return parser


async def _run(args: argparse.Namespace) -> int:
    async with SharePointGraphClient(args.url) as client:
        if args.command == "list":
            items = await client.list_items()
            for item in items:
                if args.only_files and item.type != "file":
                    continue
                print(f"{item.type:6} {item.path}")
            print(f"Total items: {len(items)}")
            return 0

        if args.command == "download":
            local_file = await client.download_file_by_path(
                file_path=args.file_path,
                output_dir=Path(args.output_dir),
            )
            print(f"Downloaded: {local_file}")
            return 0

        if args.command == "upload":
            uploaded = await client.upload_file_by_path(
                local_file=Path(args.local_file),
                remote_path=args.remote_path,
            )
            print(
                "Uploaded: "
                f"{uploaded.get('name', '<unknown>')}"
                f" -> {uploaded.get('webUrl', '')}"
            )
            return 0

    return 1


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    return asyncio.run(_run(args))


if __name__ == "__main__":
    raise SystemExit(main())
