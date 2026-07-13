"""Build Road pre-carriage for Air MAWB and HAWB output tabs."""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

from build_air_rates import _filter_air_by_service_type
from rate_layout_common import (
    ROAD_PRECARRIAGE_AIR_BOLD_COLUMNS,
    ROAD_PRECARRIAGE_AIR_SHIPMENT_COLUMNS,
    build_road_precarriage_cost_blocks,
    build_road_precarriage_shipment_from_air,
    get_rate_card_name,
    load_air_dataframe,
    output_workbook_path,
    prepare_air_for_precarriage_costs,
    save_output_sheet,
    write_matrix_sheet,
)

ROAD_PRE_CARRIAGE_AIR_MAWB_SHEET = "Road pre-carriage for Air MAWB"
ROAD_PRE_CARRIAGE_AIR_HAWB_SHEET = "Road pre-carriage for Air HAWB"


def build_road_precarriage_air(
    processing_path: Path | None = None,
    output_path: Path | None = None,
) -> Path:
    air_df = load_air_dataframe(processing_path)
    rc_name = get_rate_card_name()
    path = output_path or output_workbook_path()

    tab_specs = (
        (ROAD_PRE_CARRIAGE_AIR_MAWB_SHEET, "MAWB"),
        (ROAD_PRE_CARRIAGE_AIR_HAWB_SHEET, "HAWB"),
    )

    for sheet_name, service_type in tab_specs:
        filtered_air_df = _filter_air_by_service_type(air_df, service_type)
        shipment_df = build_road_precarriage_shipment_from_air(filtered_air_df, rc_name=rc_name)
        cost_source_df = prepare_air_for_precarriage_costs(filtered_air_df)
        cost_blocks = build_road_precarriage_cost_blocks(cost_source_df, for_air=True)

        def write_sheet(
            worksheet,
            source_df: pd.DataFrame = cost_source_df,
            shipments: pd.DataFrame = shipment_df,
            blocks: list = cost_blocks,
        ) -> None:
            write_matrix_sheet(
                worksheet,
                source_df,
                shipments,
                ROAD_PRECARRIAGE_AIR_SHIPMENT_COLUMNS,
                blocks,
                highlight_min_gt_max=True,
                transport_mode="air",
                bold_shipment_columns=ROAD_PRECARRIAGE_AIR_BOLD_COLUMNS,
            )

        save_output_sheet(sheet_name, write_sheet, output_path=path)
        print(
            f"Built '{sheet_name}' with {len(shipment_df)} rows and "
            f"{len(cost_blocks)} cost blocks"
        )

    print(f"Saved air pre-carriage tabs -> {path}")
    return path


def main() -> int:
    try:
        build_road_precarriage_air()
        return 0
    except KeyboardInterrupt:
        print("\nOperation cancelled.")
        return 130
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
