"""Runtime configuration, read from environment variables."""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

# Fixed column values, copied character-for-character from the existing
# "Work activity" sheet so new rows look exactly like the historical ones.
DEFAULT_CDL_PERSONS = (
    "ANIL KUMAR M",
    "NIRANJAN",
    "ABHILASH K R",
    "NAVEEN KUMAR N",
    "MILAN TUDU",
    "TRINATH DALAI",
    "GIRIDHARI SAHU",
    "BHUMESH CHOUDHARI",
)
DEFAULT_SHIFT_DRI = "NAGENDRA.C      GOKUL.R  "
WORK_STATUS = "Completed"
REMARKS = "CLOSED"


def _env_int(name: str, default: int) -> int:
    raw = os.environ.get(name)
    return int(raw) if raw else default


def _cdl_persons() -> tuple[str, ...]:
    # CDL_PERSONS="NAME ONE;NAME TWO;..." overrides the default list.
    raw = os.environ.get("CDL_PERSONS")
    if not raw:
        return DEFAULT_CDL_PERSONS
    return tuple(p.strip() for p in raw.split(";") if p.strip())


@dataclass(frozen=True)
class Settings:
    data_dir: Path = field(default_factory=lambda: Path(os.environ.get("DATA_DIR", "/data")))
    master_name: str = field(default_factory=lambda: os.environ.get("MASTER_WORKBOOK_NAME", "GOKUL 26-27.xlsx"))
    seed_path: Path | None = field(
        default_factory=lambda: Path(p) if (p := os.environ.get("SEED_WORKBOOK_PATH")) else None
    )
    lock_timeout: float = field(default_factory=lambda: float(os.environ.get("LOCK_TIMEOUT_SECONDS", "30")))
    backups_to_keep: int = field(default_factory=lambda: _env_int("BACKUPS_TO_KEEP", 60))
    cdl_persons: tuple[str, ...] = field(default_factory=_cdl_persons)
    shift_dri: str = field(default_factory=lambda: os.environ.get("SHIFT_DRI", DEFAULT_SHIFT_DRI))
    max_activities: int = field(default_factory=lambda: _env_int("MAX_ACTIVITIES", 100))
    max_activity_length: int = field(default_factory=lambda: _env_int("MAX_ACTIVITY_LENGTH", 1000))

    @property
    def master_path(self) -> Path:
        return self.data_dir / self.master_name

    @property
    def cdl_text(self) -> str:
        return "\n".join(self.cdl_persons)
