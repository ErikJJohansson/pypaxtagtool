"""Generate config/aoi/v4.10.6.yaml from the template spreadsheet.

Reads every AOI sheet's subtag list (rows 7+8) out of
ProcessLibraryOnlineConfigTool.xlsm with openpyxl (read-only,
data-only — the workbook is data, its macros are never touched) and
writes the ordered ``{AOI: [subtag, ...]}`` mapping via
``pypax.aoi.dump_aoi_definitions``. The YAML is the versioned record
of the AOI definitions; ``AOI_definitions.py`` stays the runtime
source for now.

Usage (from the repo root)::

    python scripts/extract_aoi_definitions.py
    python scripts/extract_aoi_definitions.py --template X.xlsm --output config/aoi/v4.10.6.yaml
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import openpyxl

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT / "src"))

from pypax.aoi import dump_aoi_definitions  # noqa: E402
from pypax.layout import get_aoi_list  # noqa: E402
from pypax.spreadsheet import get_subtag_list  # noqa: E402

DEFAULT_TEMPLATE = REPO_ROOT / "ProcessLibraryOnlineConfigTool.xlsm"
DEFAULT_OUTPUT = REPO_ROOT / "config" / "aoi" / "v4.10.6.yaml"


def extract_definitions(template_path):
    """Ordered ``{AOI: [subtag, ...]}`` from the template workbook."""
    book = openpyxl.load_workbook(template_path, read_only=True, data_only=True)
    try:
        return {aoi: get_subtag_list(book[aoi]) for aoi in get_aoi_list(book)}
    finally:
        book.close()


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Generate the AOI definitions YAML from the template spreadsheet."
    )
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args(argv)

    definitions = extract_definitions(args.template)
    dump_aoi_definitions(definitions, args.output)
    print(f"Wrote {len(definitions)} AOIs to {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
