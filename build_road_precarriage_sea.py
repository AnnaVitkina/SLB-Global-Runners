"""Build the Road pre-carriage for Sea output tab from Ocean processing data."""

from __future__ import annotations

import sys
from pathlib import Path

from rate_layout_common import (
    ROAD_PRECARRIAGE_SEA_BOLD_COLUMNS,
    ROAD_PRECARRIAGE_SHIPMENT_COLUMNS,
    build_road_precarriage_cost_blocks,
    build_road_precarriage_shipment_from_ocean,
    load_ocean_dataframe,
    save_output_sheet,
    write_matrix_sheet,
)

ROAD_PRE_CARRIAGE_SHEET = "Road pre-carriage for Sea"


def build_road_precarriage_sea(
    processing_path: Path | None = None,
    output_path: Path | None = None,
) -> Path:
    ocean_df = load_ocean_dataframe(processing_path)
    shipment_df = build_road_precarriage_shipment_from_ocean(ocean_df)
    cost_blocks = build_road_precarriage_cost_blocks(ocean_df)

    def write_sheet(worksheet) -> None:
        write_matrix_sheet(
            worksheet,
            ocean_df,
            shipment_df,
            ROAD_PRECARRIAGE_SHIPMENT_COLUMNS,
            cost_blocks,
            highlight_min_gt_max=True,
            bold_shipment_columns=ROAD_PRECARRIAGE_SEA_BOLD_COLUMNS,
        )

    saved_path = save_output_sheet(
        ROAD_PRE_CARRIAGE_SHEET,
        write_sheet,
        output_path=output_path,
    )
    print(
        f"Built '{ROAD_PRE_CARRIAGE_SHEET}' with {len(shipment_df)} rows and "
        f"{len(cost_blocks)} cost blocks -> {saved_path}"
    )
    return saved_path


def main() -> int:
    try:
        build_road_precarriage_sea()
        return 0
    except KeyboardInterrupt:
        print("\nOperation cancelled.")
        return 130
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
