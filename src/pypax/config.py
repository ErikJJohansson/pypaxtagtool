"""Application configuration: settings, TOML files, per-PLC profiles.

Three layers, one source of truth each (see PLAN.md section 5):

- ``AppSettings`` — app-wide policy (template path, bind host/port, write
  chunk size, REAL read-back tolerance, retries). Values come from
  environment variables with the ``PYPAX_`` prefix first, then the
  ``[app]`` section of a TOML config file, then built-in defaults.
- ``PlcProfile`` — one entry per PLC/site: comm path, the PLC name the
  app expects to find there, the PlantPAX AOI version, and a safety
  level. ``test`` profiles write freely; ``production`` profiles require
  the PLC-name double-check on every write (enforced by the webapp).
- TOML files are local only and gitignored — this repo is public, so
  real PLC addresses and PLC names must never be committed. Only
  ``config/config.example.toml`` (placeholders) lives in the repo.
"""

from __future__ import annotations

import ipaddress
import tomllib
from pathlib import Path
from typing import Literal

from pydantic import BaseModel, field_validator
from pydantic_settings import BaseSettings, SettingsConfigDict

DEFAULT_AOI_VERSION = "4.10.6"


class PlcProfile(BaseModel):
    """One PLC/site: how to reach it and what to expect there."""

    name: str
    commpath: str
    expected_plc_name: str | None = None
    aoi_version: str = DEFAULT_AOI_VERSION
    safety_level: Literal["test", "production"] = "test"


class AppSettings(BaseSettings):
    """App-wide policy. Env vars (``PYPAX_*``) beat the TOML file."""

    model_config = SettingsConfigDict(env_prefix="PYPAX_")

    template_path: str | None = None
    host: str = "127.0.0.1"
    port: int = 8000
    write_chunk_size: int = 100
    real_tolerance: float = 1e-6
    read_back_retries: int = 1
    config_file: str | None = None
    # Directory scanned for Studio 5000 ``.L5X`` exports to import as
    # custom AOIs (Phase 8). Relative paths resolve against the repo
    # root; env override ``PYPAX_AOI_IMPORT_DIR``.
    aoi_import_dir: str = "aoi_imports"
    # Directory holding PLC snapshots (Phase 9a): manual snapshots at
    # the top level, auto-snapshots under ``auto/``. Relative paths
    # resolve against the repo root; env override
    # ``PYPAX_SNAPSHOT_DIR``.
    snapshot_dir: str = "snapshots"

    @field_validator("host")
    @classmethod
    def _host_must_be_loopback(cls, value: str) -> str:
        """The webapp has no authentication and can write to a PLC, so
        it must only ever listen on loopback. Anything else (0.0.0.0,
        a LAN address) would silently expose the write API to the
        network — refuse it at startup instead."""
        host = value.strip().lower()
        if host == "localhost":
            return value
        try:
            address = ipaddress.ip_address(host)
        except ValueError:
            raise ValueError(f"host must be a loopback address, got {value!r}")
        if not address.is_loopback:
            raise ValueError(f"host must be a loopback address, got {value!r}")
        return value


def load_config_file(path: str | Path) -> dict:
    """Parse a TOML config file. Raises ``FileNotFoundError`` / ``ValueError``."""
    with open(path, "rb") as handle:
        data = tomllib.load(handle)
    if not isinstance(data, dict):
        raise ValueError(f"Config file {path} must contain a TOML table at the top level")
    return data


def load_profiles(data: dict) -> dict[str, PlcProfile]:
    """``[[profiles]]`` entries from a parsed config file, keyed by name."""
    profiles: dict[str, PlcProfile] = {}
    entries = data.get("profiles", [])
    if not isinstance(entries, list):
        raise ValueError("Config key 'profiles' must be a list of [[profiles]] tables")
    for entry in entries:
        profile = PlcProfile(**entry)
        if profile.name in profiles:
            raise ValueError(f"Duplicate PLC profile name {profile.name!r} in config file")
        profiles[profile.name] = profile
    return profiles


def load_app_config(
    path: str | Path | None = None,
) -> tuple[AppSettings, dict[str, PlcProfile]]:
    """Load settings + profiles from a TOML file (if any) and the environment.

    Precedence: ``PYPAX_*`` environment variables, then the file's ``[app]``
    section, then built-in defaults. Unknown ``[app]`` keys fail fast.
    ``path`` defaults to the ``PYPAX_CONFIG_FILE`` setting, if set.
    """
    env_settings = AppSettings()
    if path is None:
        path = env_settings.config_file
    profiles: dict[str, PlcProfile] = {}
    file_app: dict = {}
    if path is not None:
        data = load_config_file(path)
        profiles = load_profiles(data)
        file_app = data.get("app", {})
        if not isinstance(file_app, dict):
            raise ValueError("Config key 'app' must be a [app] table")
        for key in file_app:
            if key not in AppSettings.model_fields:
                raise ValueError(f"Unknown [app] setting {key!r} in config file {path}")
    overrides = {
        key: value for key, value in file_app.items() if key not in env_settings.model_fields_set
    }
    settings = AppSettings(**overrides)
    return settings, profiles
