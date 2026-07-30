"""Build Air MAWB and Air HAWB output tabs from Air processing data."""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

from carrier_groups import CARRIER_GROUP_HOU, CARRIER_GROUP_MEA, CARRIER_GROUP_RTM, CARRIER_GROUP_SC
from rate_layout_common import (
    CostBlock,
    CostValueColumn,
    DZS_AEI_SUPPLIER_NAME,
    apply_grouped_second_column_fill,
    capitalize_type_label,
    cell_text,
    display_text,
    format_date,
    get_rate_card_name,
    load_air_dataframe,
    output_workbook_path,
    save_output_sheet,
    trim_lane_uid_suffix,
    write_matrix_sheet,
)
from supplier_name_lookup import map_fred_supplier_names

AIR_MAWB_SHEET = "Air MAWB rate"
AIR_HAWB_SHEET = "Air HAWB rate"

AIR_SHIPMENT_COLUMNS = [
    "Rate Card",
    "Scope of work",
    "Service type",
    "Lane UID",
    "Origin country code",
    "Origin base city",
    "Departure airport city",
    "Departure airport code",
    "Destination country code",
    "Destination airport city",
    "Destination airport code",
    "Destination base city",
    "Destination base city (second)",
    "Supplier name (Q)",
    "Supplier name",
    "Type",
    "Valid from",
    "Valid to",
]

FREIGHT_AIR_WEIGHT_COLUMNS: tuple[tuple[str, str], ...] = (
    ("MIN", "Min per AWB"),
    ("<100 Chargeable KG", "<100 Chargeable KG"),
    ("101-250 Chargeable KG", "101-250 Chargeable KG"),
    ("251-300 Chargeable KG", "251-300 Chargeable KG"),
    ("301-500 Chargeable KG", "301-500 Chargeable KG"),
    ("501-1000 Chargeable KG", "501-1000 Chargeable KG"),
    ("1001-3000 Chargeable KG", "1001-3000 Chargeable KG"),
    ("3001-5000 Chargeable KG", "3001-5000 Chargeable KG"),
    (">5000 Chargeable KG", ">5000 Chargeable KG"),
)

CERTIFICATE_OF_ORIGIN_COL = "Certificate of Origin"
EXPORT_CLEARANCE_COL = "Export Clearance"
SCREENING_FEE_COL = "Screening per actual KG"
FUMIGATION_COL = "Fumigation cost per crate/box/pallet"
DGR_DECLARATION_COL = "DGR Declaration"

AIR_BOLD_SHIPMENT_COLUMNS = frozenset(
    {
        "Origin country code",
        "Origin base city",
        "Destination country code",
        "Destination airport code",
        "Supplier name",
        "Type",
        "Valid from",
        "Valid to",
    }
)

AIR_DEST_BASE_CITY_GROUP_COLUMNS = tuple(
    column
    for column in AIR_SHIPMENT_COLUMNS
    if column not in ("Destination base city", "Destination base city (second)")
)


def is_dhl_us_to_ec_lane(air_row: pd.Series, supplier_name: object) -> bool:
    if cell_text(air_row.get("Origin Country Code")).upper() != "US":
        return False
    if cell_text(air_row.get("Destination Country Code")).upper() != "EC":
        return False
    supplier_q = cell_text(air_row.get("Supplier Name (Q)")).lower()
    mapped_name = cell_text(supplier_name).lower()
    return "dhl" in supplier_q or "dhl" in mapped_name


def duplicate_dhl_us_ec_dzs_rows(
    source_df: pd.DataFrame,
    shipment_df: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    source_rows: list[pd.Series] = []
    shipment_rows: list[pd.Series] = []

    for idx in range(len(shipment_df)):
        source_row = source_df.iloc[idx]
        shipment_row = shipment_df.iloc[idx]
        if not is_dhl_us_to_ec_lane(source_row, shipment_row["Supplier name"]):
            continue
        dzs_shipment = shipment_row.copy()
        dzs_shipment["Supplier name"] = DZS_AEI_SUPPLIER_NAME
        source_rows.append(source_row)
        shipment_rows.append(dzs_shipment)

    if not source_rows:
        return source_df, shipment_df

    return (
        pd.concat([source_df, pd.DataFrame(source_rows)], ignore_index=True),
        pd.concat([shipment_df, pd.DataFrame(shipment_rows)], ignore_index=True),
    )


def build_air_shipment_df(
    air_df: pd.DataFrame,
    *,
    rc_name: str,
    type_value: str,
) -> pd.DataFrame:
    empty = pd.Series([""] * len(air_df), index=air_df.index, dtype="object")
    type_series = pd.Series(
        [capitalize_type_label(type_value)] * len(air_df),
        index=air_df.index,
        dtype="object",
    )

    return pd.DataFrame(
        {
            "Rate Card": rc_name,
            "Scope of work": air_df["Scope of Work"].map(display_text),
            "Service type": air_df["Service Type"].map(display_text),
            "Lane UID": air_df["Lane UID"].map(trim_lane_uid_suffix),
            "Origin country code": air_df["Origin Country Code"].map(display_text),
            "Origin base city": air_df["Origin Base City"].map(display_text),
            "Departure airport city": air_df["Departure Airport City"].map(display_text),
            "Departure airport code": air_df["Departure Aiport Code"].map(display_text),
            "Destination country code": air_df["Destination Country Code"].map(display_text),
            "Destination airport city": air_df["Destination Airport City"].map(display_text),
            "Destination airport code": air_df["Destination Airport Code"].map(display_text),
            "Destination base city": air_df["Destination Base City"].map(display_text),
            "Destination base city (second)": empty,
            "Supplier name (Q)": air_df["Supplier Name (Q)"].map(display_text),
            "Supplier name": map_fred_supplier_names(
                air_df["Origin Country Code"],
                air_df["Supplier Name (Q)"],
                transport_mode="air",
                destination_countries=air_df["Destination Country Code"],
            ),
            "Type": type_series,
            "Valid from": air_df["Valid from"].map(format_date),
            "Valid to": air_df["Valid to"].map(format_date),
        },
        columns=AIR_SHIPMENT_COLUMNS,
    )


def build_air_carrier_group_cost_blocks(
    *,
    include_arx_ae_export: bool = True,
) -> list[CostBlock]:
    blocks: list[CostBlock] = []

    export_specs = (
        (
            "Air Export Clearance (HOU, MX)",
            EXPORT_CLEARANCE_COL,
            (CARRIER_GROUP_HOU,),
            True,
        ),
        (
            "Air Export Clearance (RTM, MEA, SC)",
            EXPORT_CLEARANCE_COL,
            (CARRIER_GROUP_RTM, CARRIER_GROUP_MEA, CARRIER_GROUP_SC),
            True,
        ),
        (
            "Air Export Clearance (ARX AE)",
            EXPORT_CLEARANCE_COL,
        ),
        (
            "Air Export Clearance (DZS AEI)",
            EXPORT_CLEARANCE_COL,
        ),
    )
    for title, source_column, carrier_groups, exclude_dzs_aei in export_specs[:2]:
        if carrier_groups == (CARRIER_GROUP_HOU,):
            apply_if = "Apply if: Carrier group in {'HOU', 'MX'}"
        else:
            apply_if = "Apply if: Carrier group in {'RTM', 'MEA', 'SC'}"
        blocks.append(
            CostBlock(
                title=title,
                apply_if=apply_if,
                rate_by="Rate by: Per shipment",
                value_columns=(CostValueColumn("p/unit", source_column),),
                match_carrier_groups=carrier_groups,
                exclude_dzs_aei_carrier_lane=exclude_dzs_aei,
            )
        )

    if include_arx_ae_export:
        blocks.append(
            CostBlock(
                title=export_specs[2][0],
                apply_if="Apply if: Supplier name equals 'ARX'",
                rate_by="Rate by: Per shipment",
                value_columns=(CostValueColumn("p/unit", export_specs[2][1]),),
                match_arx_carrier_lane=True,
            )
        )
    blocks.append(
        CostBlock(
            title=export_specs[3][0],
            apply_if="Apply if: Supplier name equals 'DZS AEI'",
            rate_by="Rate by: Per shipment",
            value_columns=(CostValueColumn("p/unit", export_specs[3][1]),),
            match_supplier_name=DZS_AEI_SUPPLIER_NAME,
        )
    )

    screening_specs = (
        (
            "Air Screening fee (HOU, MEA, MX)",
            SCREENING_FEE_COL,
            (CARRIER_GROUP_HOU, CARRIER_GROUP_MEA),
        ),
        (
            "Air Screening fee (SC, RTM)",
            SCREENING_FEE_COL,
            (CARRIER_GROUP_SC, CARRIER_GROUP_RTM),
        ),
    )
    for title, source_column, carrier_groups in screening_specs:
        if carrier_groups == (CARRIER_GROUP_HOU, CARRIER_GROUP_MEA):
            apply_if = "Apply if: Carrier group in {'HOU', 'MEA', 'MX'}"
        else:
            apply_if = "Apply if: Carrier group in {'SC', 'RTM'}"
        blocks.append(
            CostBlock(
                title=title,
                apply_if=apply_if,
                rate_by="Rate by: Per shipment",
                value_columns=(CostValueColumn("p/unit", source_column),),
                match_carrier_groups=carrier_groups,
                exclude_dzs_aei_carrier_lane=True,
            )
        )

    blocks.append(
        CostBlock(
            title="Air Screening fee (DZS AEI)",
            apply_if="Apply if: Supplier name equals 'DZS AEI'",
            rate_by="Rate by: Per shipment",
            value_columns=(CostValueColumn("p/unit", SCREENING_FEE_COL),),
            match_supplier_name=DZS_AEI_SUPPLIER_NAME,
        )
    )

    return blocks


def build_air_cost_blocks(*, include_arx_ae_export: bool = True) -> list[CostBlock]:
    freight_columns = tuple(
        CostValueColumn(header, source_column)
        for header, source_column in FREIGHT_AIR_WEIGHT_COLUMNS
    )

    return [
        CostBlock(
            title="Freight Air Intl",
            apply_if="",
            rate_by="Rate by: Chargeable kg",
            value_columns=freight_columns,
        ),
        CostBlock(
            title="Air Certificate of origin",
            apply_if="",
            rate_by="Rate by: Per shipment",
            value_columns=(CostValueColumn("p/unit", CERTIFICATE_OF_ORIGIN_COL),),
        ),
        CostBlock(
            title="Air Fumigation",
            apply_if="",
            rate_by="Rate by: Per shipment",
            value_columns=(CostValueColumn("p/unit", FUMIGATION_COL),),
        ),
        CostBlock(
            title="DGR Surcharge",
            apply_if="",
            rate_by="Rate by: Per shipment",
            value_columns=(CostValueColumn("p/unit", DGR_DECLARATION_COL),),
        ),
        *build_air_carrier_group_cost_blocks(include_arx_ae_export=include_arx_ae_export),
    ]


def _filter_air_by_service_type(air_df: pd.DataFrame, service_type: str) -> pd.DataFrame:
    mask = air_df["Service Type"].map(lambda value: cell_text(value).upper()) == service_type.upper()
    return air_df.loc[mask].reset_index(drop=True)


def build_air_rates(
    processing_path: Path | None = None,
    output_path: Path | None = None,
    rc_name: str | None = None,
) -> Path:
    air_df = load_air_dataframe(processing_path)
    rate_card_name = get_rate_card_name(rc_name)
    path = output_path or output_workbook_path(rate_card_name)

    tab_specs = (
        (AIR_MAWB_SHEET, "MAWB", "direct"),
        (AIR_HAWB_SHEET, "HAWB", "Consol"),
    )

    for sheet_name, service_type, type_value in tab_specs:
        cost_blocks = build_air_cost_blocks(include_arx_ae_export=service_type == "HAWB")
        filtered_air_df = _filter_air_by_service_type(air_df, service_type)
        shipment_df = build_air_shipment_df(
            filtered_air_df,
            rc_name=rate_card_name,
            type_value=type_value,
        )
        shipment_df = apply_grouped_second_column_fill(
            shipment_df,
            first_column="Destination base city",
            second_column="Destination base city (second)",
            group_columns=AIR_DEST_BASE_CITY_GROUP_COLUMNS,
        )
        filtered_air_df, shipment_df = duplicate_dhl_us_ec_dzs_rows(filtered_air_df, shipment_df)

        def write_sheet(
            worksheet,
            source_df: pd.DataFrame = filtered_air_df,
            shipments: pd.DataFrame = shipment_df,
        ) -> None:
            write_matrix_sheet(
                worksheet,
                source_df,
                shipments,
                AIR_SHIPMENT_COLUMNS,
                cost_blocks,
                transport_mode="air",
                bold_shipment_columns=AIR_BOLD_SHIPMENT_COLUMNS,
                highlight_missing_freight_block="Freight Air Intl",
            )

        save_output_sheet(sheet_name, write_sheet, output_path=path)
        print(
            f"Built '{sheet_name}' with {len(shipment_df)} rows and "
            f"{len(cost_blocks)} cost blocks"
        )

    print(f"Saved air tabs -> {path}")
    return path


def main() -> int:
    try:
        build_air_rates()
        return 0
    except KeyboardInterrupt:
        print("\nOperation cancelled.")
        return 130
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
