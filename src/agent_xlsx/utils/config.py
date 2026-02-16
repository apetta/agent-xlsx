"""Persistent configuration for agent-xlsx (~/.agent-xlsx/config.json)."""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

CONFIG_DIR = Path.home() / ".agent-xlsx"
CONFIG_FILE = CONFIG_DIR / "config.json"


def load_config() -> dict[str, Any]:
    """Load config from disk. Returns empty dict if file missing."""
    if not CONFIG_FILE.exists():
        return {}
    try:
        return json.loads(CONFIG_FILE.read_text())
    except (json.JSONDecodeError, OSError):
        return {}


def save_config(config: dict[str, Any]) -> None:
    """Write config to disk, creating directory if needed."""
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    CONFIG_FILE.write_text(json.dumps(config, indent=2) + "\n")


def get_aspose_license_path() -> str | None:
    """Resolve Aspose licence path. Priority: env var > config file.

    Returns:
        - File path string from ASPOSE_LICENSE_PATH env var
        - "base64:<data>" string from ASPOSE_LICENSE_DATA env var
        - File path string from config file
        - None if no licence configured
    """
    # 1. ASPOSE_LICENSE_PATH env var (file path)
    env_path = os.environ.get("ASPOSE_LICENSE_PATH")
    if env_path:
        return env_path

    # 2. ASPOSE_LICENSE_DATA env var (base64-encoded .lic content)
    env_data = os.environ.get("ASPOSE_LICENSE_DATA")
    if env_data:
        return f"base64:{env_data}"

    # 3. Config file
    config = load_config()
    return config.get("aspose_license_path")
