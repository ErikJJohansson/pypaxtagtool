"""Shared pytest fixtures and path setup for the test suite.

Adds ``src/`` (the ``pypax`` package) and ``tests/`` (the ``fakes`` package)
to ``sys.path`` so tests can import both without an installed package.
Cross-platform: uses ``pathlib`` only, no OS-specific behaviour.
"""

import sys
from pathlib import Path

TESTS_DIR = Path(__file__).resolve().parent
REPO_ROOT = TESTS_DIR.parent

for _path in (REPO_ROOT / "src", TESTS_DIR):
    if str(_path) not in sys.path:
        sys.path.insert(0, str(_path))

# The repo root also holds pypax.py (the legacy CLI script). It must never
# shadow the src/pypax package in tests: drop '' and the repo root from
# sys.path so `import pypax` always resolves to the package. The legacy
# script is executed as __main__ or loaded by file path (golden test),
# never imported as `pypax`.
for _entry in ("", str(REPO_ROOT)):
    while _entry in sys.path:
        sys.path.remove(_entry)
sys.path.insert(0, str(REPO_ROOT / "src"))
