"""Tests for config loading: defaults, env overrides, TOML files, profiles."""

from __future__ import annotations

import pytest
from pydantic import ValidationError

from pypax.config import AppSettings, load_app_config, load_config_file, load_profiles


@pytest.fixture(autouse=True)
def _clean_env(monkeypatch):
    """Settings read PYPAX_* from the environment; keep tests hermetic."""
    for key in list(__import__("os").environ):
        if key.startswith("PYPAX_"):
            monkeypatch.delenv(key, raising=False)


def test_defaults():
    settings, profiles = load_app_config()
    assert settings.host == "127.0.0.1"
    assert settings.port == 8000
    assert settings.write_chunk_size == 100
    assert settings.real_tolerance == 1e-6
    assert settings.read_back_retries == 1
    assert settings.template_path is None
    assert profiles == {}


def test_env_overrides(monkeypatch):
    monkeypatch.setenv("PYPAX_PORT", "9000")
    monkeypatch.setenv("PYPAX_HOST", "127.0.0.2")
    settings, _ = load_app_config()
    assert settings.port == 9000
    assert settings.host == "127.0.0.2"


def test_toml_file(tmp_path):
    config_path = tmp_path / "config.toml"
    config_path.write_text(
        "[app]\nport = 9001\nwrite_chunk_size = 25\n\n"
        '[[profiles]]\nname = "line1"\ncommpath = "192.0.2.10/0"\n'
        'expected_plc_name = "LINE1_PLC"\nsafety_level = "production"\n'
    )
    settings, profiles = load_app_config(config_path)
    assert settings.port == 9001
    assert settings.write_chunk_size == 25
    assert settings.host == "127.0.0.1"  # default, not in file
    assert list(profiles) == ["line1"]
    profile = profiles["line1"]
    assert profile.commpath == "192.0.2.10/0"
    assert profile.expected_plc_name == "LINE1_PLC"
    assert profile.safety_level == "production"
    assert profile.aoi_version == "4.10.6"  # default


def test_env_beats_toml_file(tmp_path, monkeypatch):
    config_path = tmp_path / "config.toml"
    config_path.write_text("[app]\nport = 9001\n")
    monkeypatch.setenv("PYPAX_PORT", "9002")
    settings, _ = load_app_config(config_path)
    assert settings.port == 9002


def test_config_file_from_env(tmp_path, monkeypatch):
    config_path = tmp_path / "config.toml"
    config_path.write_text("[app]\nport = 9003\n")
    monkeypatch.setenv("PYPAX_CONFIG_FILE", str(config_path))
    settings, _ = load_app_config()
    assert settings.port == 9003


def test_unknown_app_key_fails_fast(tmp_path):
    config_path = tmp_path / "config.toml"
    config_path.write_text("[app]\nno_such_setting = 1\n")
    with pytest.raises(ValueError, match="Unknown \\[app\\] setting"):
        load_app_config(config_path)


def test_bad_safety_level_rejected():
    with pytest.raises(ValidationError):
        load_profiles(
            {"profiles": [{"name": "x", "commpath": "1.2.3.4/0", "safety_level": "yolo"}]}
        )


def test_duplicate_profile_name_rejected():
    with pytest.raises(ValueError, match="Duplicate PLC profile name"):
        load_profiles(
            {
                "profiles": [
                    {"name": "dup", "commpath": "192.0.2.1/0"},
                    {"name": "dup", "commpath": "192.0.2.2/0"},
                ]
            }
        )


def test_missing_file_raises(tmp_path):
    with pytest.raises(FileNotFoundError):
        load_config_file(tmp_path / "nope.toml")


def test_profile_defaults():
    profiles = load_profiles({"profiles": [{"name": "plain", "commpath": "192.0.2.5/0"}]})
    profile = profiles["plain"]
    assert profile.expected_plc_name is None
    assert profile.safety_level == "test"
    assert profile.aoi_version == "4.10.6"


def test_settings_prefix_is_pypax(monkeypatch):
    # An unprefixed variable must NOT be picked up.
    monkeypatch.setenv("PORT", "9999")
    assert AppSettings().port == 8000
