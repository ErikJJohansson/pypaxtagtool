"""The AOI model the webapp shows: columns and instances per AOI.

The column model (subtag, data type, description) comes from the
PlantPAX template workbook — the same source of truth the CLI uses
(rows 4/6/7/8 of each AOI sheet). Instances come from the connected
PLC: base instance names grouped by AOI type. The template is required;
without it the app cannot know the columns, and the AOI endpoints
answer 503 saying so.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import openpyxl

from ..custom_aoi import load_custom_aois
from ..layout import DATA_TYPE_ROW, START_COL, get_aoi_list
from ..l5x import Column as L5xColumn
from ..plc import get_aoi_tag_instances, make_tag_list, read_plc_row
from ..spreadsheet import get_subtag_list


@dataclass
class Column:
    """One grid column: a subtag with its data type and description."""

    subtag: str
    data_type: str
    description: str = ""


@dataclass
class AoiModel:
    """One AOI tab: its columns and its PLC instance names, in order."""

    name: str
    columns: list[Column] = field(default_factory=list)
    instances: list[str] = field(default_factory=list)


def _read_row_values(sheet, row, count):
    """First ``count`` values of ``row`` from ``START_COL`` onward."""
    (values,) = sheet.iter_rows(
        min_row=row,
        max_row=row,
        min_col=START_COL,
        max_col=START_COL + count - 1,
        values_only=True,
    )
    return values


def load_template_model(template_path: str | Path) -> dict[str, AoiModel]:
    """AOI columns for every AOI sheet in the template workbook.

    Raises ``FileNotFoundError`` / ``ValueError`` (via openpyxl) when the
    workbook cannot be opened — callers translate that to a 503.
    """
    book = openpyxl.load_workbook(template_path, read_only=True, data_only=True)
    try:
        model: dict[str, AoiModel] = {}
        for aoi in get_aoi_list(book):
            sheet = book[aoi]
            sub_tags = get_subtag_list(sheet)
            data_types = _read_row_values(sheet, DATA_TYPE_ROW, len(sub_tags))
            descriptions = _read_row_values(sheet, 4, len(sub_tags))
            columns = [
                Column(
                    subtag=sub_tag,
                    data_type=str(data_type) if data_type is not None else "",
                    description=str(description) if description is not None else "",
                )
                for sub_tag, data_type, description in zip(sub_tags, data_types, descriptions)
            ]
            model[aoi] = AoiModel(name=aoi, columns=columns)
        return model
    finally:
        book.close()


def load_combined_model(
    template_path: str | Path,
    custom_dir: str | Path | None,
    column_meta: dict[str, dict[str, L5xColumn]] | None = None,
) -> tuple[dict[str, AoiModel], set[str]]:
    """Template AOI model ∪ confirmed custom AOIs.

    Custom AOIs come from ``config/aoi/custom/*.yaml`` (subtag lists
    only); their column types/descriptions are enriched from
    ``column_meta`` (``{aoi: {subtag: Column}}``, cached from import
    proposals) when available, and fall back to empty strings — a
    custom AOI whose source ``.L5X`` is gone still serves its columns,
    just untyped. Returns ``(model, custom_names)``. A custom AOI whose
    name collides with a template sheet never wins: the template entry
    stands (confirm-time collision checks prevent the collision, this
    is only a defensive second layer). Raises ``ValueError`` when the
    custom store is corrupt (e.g. one AOI in two files).
    """
    model = load_template_model(template_path)
    custom_names: set[str] = set()
    template_lower = {name.lower() for name in model}
    for aoi, sub_tags in load_custom_aois(custom_dir).items():
        custom_names.add(aoi)
        if aoi.lower() in template_lower:
            continue
        meta = (column_meta or {}).get(aoi, {})
        columns = [
            Column(
                subtag=sub_tag,
                data_type=str(meta[sub_tag].data_type)
                if sub_tag in meta and meta[sub_tag].data_type is not None
                else "",
                description=str(meta[sub_tag].description)
                if sub_tag in meta and meta[sub_tag].description
                else "",
            )
            for sub_tag in sub_tags
        ]
        model[aoi] = AoiModel(name=aoi, columns=columns)
    return model, custom_names


def discover_instances(plc, aoi: str) -> list[str]:
    """Base instance names of ``aoi`` in the PLC, in discovery order.

    ``get_aoi_tag_instances`` returns base tags for a real PLC and
    member tags for the FakePLC (which seeds one ``tags`` entry per
    subtag); stripping at the first ``.`` gives the base instance name
    in both cases. Array elements (``PUMP[0]``) stay distinct instances.
    """
    instances: list[str] = []
    for tag in get_aoi_tag_instances(plc, aoi):
        base = tag.split(".", 1)[0]
        if base not in instances:
            instances.append(base)
    return instances


def read_aoi_rows(plc, aoi_model: AoiModel) -> dict[str, dict[str, object]]:
    """Read every instance row of an AOI: instance -> {subtag: value}."""
    sub_tags = [column.subtag for column in aoi_model.columns]
    rows: dict[str, dict[str, object]] = {}
    for instance in aoi_model.instances:
        tag_list = make_tag_list(instance, sub_tags)
        tag_data, _raw = read_plc_row(plc, tag_list)
        values = {tag: value for tag, value in tag_data}
        rows[instance] = {sub_tag: values.get(instance + sub_tag) for sub_tag in sub_tags}
    return rows
