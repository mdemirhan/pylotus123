"""Application configuration management.

Handles loading and saving user preferences like theme, default column width,
and recent files list.
"""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from pathlib import Path

CONFIG_DIR = Path.home() / ".config" / "lotus123"
CONFIG_FILE = CONFIG_DIR / "config.json"


@dataclass
class AppConfig:
    """Application configuration."""

    theme: str = "LOTUS"
    default_col_width: int = 10
    recent_files: list[str] = field(default_factory=list)

    @classmethod
    def load(cls) -> "AppConfig":
        """Load configuration from file."""
        try:
            if CONFIG_FILE.exists():
                with open(CONFIG_FILE, "r") as f:
                    data = json.load(f)
                return cls(**data)
        except (OSError, IOError, json.JSONDecodeError, TypeError, KeyError):
            pass
        return cls()

    def save(self) -> None:
        """Save configuration to file."""
        try:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            with open(CONFIG_FILE, "w") as f:
                json.dump(asdict(self), f, indent=2)
        except (OSError, IOError, TypeError):
            pass

    def add_recent_file(self, filepath: str) -> None:
        """Add a file to the recent files list."""
        if filepath in self.recent_files:
            self.recent_files.remove(filepath)
        self.recent_files.insert(0, filepath)
        self.recent_files = self.recent_files[:10]  # Keep max 10
