"""FastAPI backend for the PyPaxTagTool webapp."""

from __future__ import annotations

from .app import create_app, default_plc_factory

__all__ = ["create_app", "default_plc_factory"]
