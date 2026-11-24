"""Thin CLI shim over the src/pypax core library.

Usage (unchanged)::

    pypax.py 10.10.17.10/4 read [ProcessLibraryOnlineConfigTool.xlsm]
    pypax.py 10.10.17.10/4 write PLCNAME_ConfigTags.xlsx

The implementation lives in :mod:`pypax.cli`. This file only puts
``src/`` on the path and runs that module as ``__main__``; it must be
executed as a script, not imported (``import pypax`` resolves to the
installed ``src/pypax`` package, not to this file).
"""

import runpy
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

if __name__ == "__main__":
    runpy.run_module("pypax.cli", run_name="__main__")
