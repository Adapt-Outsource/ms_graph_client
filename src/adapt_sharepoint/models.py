from dataclasses import dataclass


@dataclass(slots=True)
class DriveEntry:
    type: str
    path: str
    size: int
    id: str
    web_url: str
