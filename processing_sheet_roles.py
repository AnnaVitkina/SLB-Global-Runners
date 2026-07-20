"""Map custom processing workbook tabs to standard Ocean / Air names."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from rate_layout_common import AIR_SOURCE_SHEET, OCEAN_SOURCE_SHEET

ROLE_OCEAN = "ocean"
ROLE_AIR = "air"
ROLE_BOTH = "both"
ROLE_SKIP = "skip"


def _find_sheet_name(sheet_names: list[str], standard_name: str) -> str | None:
    return next((name for name in sheet_names if name.lower() == standard_name.lower()), None)


def processing_has_sheet(processing_path: Path, standard_name: str) -> bool:
    workbook = pd.ExcelFile(processing_path)
    return _find_sheet_name(workbook.sheet_names, standard_name) is not None


def prompt_transport_role(sheet_name: str) -> str:
    print(f"\nSheet '{sheet_name}' is not named Ocean or Air.")
    print("How should this sheet be used for output generation?")
    print("  1 - Treat as Ocean data (Sea Rates + Road pre-carriage for Sea)")
    print("  2 - Treat as Air data (Air MAWB/HAWB + Road pre-carriage for Air)")
    print("  3 - Treat as both Ocean and Air")
    print("  4 - Skip (do not use for rate tabs)")

    mapping = {
        "1": ROLE_OCEAN,
        "2": ROLE_AIR,
        "3": ROLE_BOTH,
        "4": ROLE_SKIP,
        "ocean": ROLE_OCEAN,
        "air": ROLE_AIR,
        "both": ROLE_BOTH,
        "skip": ROLE_SKIP,
    }

    while True:
        answer = input("Enter choice (1-4): ").strip().lower()
        if answer in mapping:
            return mapping[answer]
        print("Please enter 1, 2, 3, or 4.")


def resolve_processing_sheets(
    processing_path: Path,
    *,
    interactive: bool = True,
) -> Path:
    """
    Ensure standard Ocean/Air tabs exist by asking how to treat non-standard sheets.
    Adds Ocean and/or Air sheets (copies) without removing existing tabs.
    """
    workbook = pd.ExcelFile(processing_path)
    sheet_names = list(workbook.sheet_names)

    ocean_tab = _find_sheet_name(sheet_names, OCEAN_SOURCE_SHEET)
    air_tab = _find_sheet_name(sheet_names, AIR_SOURCE_SHEET)

    other_tabs = [
        name
        for name in sheet_names
        if name != ocean_tab and name != air_tab
    ]

    if not other_tabs and ocean_tab and air_tab:
        return processing_path

    if other_tabs and not interactive:
        raise RuntimeError(
            f"Processing workbook '{processing_path.name}' has no Ocean/Air tabs "
            f"(found: {', '.join(sheet_names)}). "
            "Re-run without --auto to choose how to process each sheet."
        )

    ocean_source = ocean_tab
    air_source = air_tab

    for tab_name in other_tabs:
        role = prompt_transport_role(tab_name)
        if role == ROLE_SKIP:
            continue
        if role in (ROLE_OCEAN, ROLE_BOTH):
            if ocean_source is None:
                ocean_source = tab_name
            elif ocean_source != tab_name:
                print(
                    f"  Note: Ocean data already set from '{ocean_source}'; "
                    f"ignoring '{tab_name}' for Ocean."
                )
        if role in (ROLE_AIR, ROLE_BOTH):
            if air_source is None:
                air_source = tab_name
            elif air_source != tab_name:
                print(
                    f"  Note: Air data already set from '{air_source}'; "
                    f"ignoring '{tab_name}' for Air."
                )

    if ocean_source is None and air_source is None:
        print("\nNo sheets mapped to Ocean or Air. Rate output tabs will not be built.")
        return processing_path

    needs_write = (ocean_source is not None and ocean_tab is None) or (
        air_source is not None and air_tab is None
    )
    if not needs_write:
        return processing_path

    frames: dict[str, pd.DataFrame] = {
        name: pd.read_excel(processing_path, sheet_name=name) for name in sheet_names
    }

    if ocean_source is not None and ocean_tab is None:
        frames[OCEAN_SOURCE_SHEET] = frames[ocean_source].copy()
        print(f"  Using '{ocean_source}' as Ocean data -> tab '{OCEAN_SOURCE_SHEET}'")

    if air_source is not None and air_tab is None:
        frames[AIR_SOURCE_SHEET] = frames[air_source].copy()
        print(f"  Using '{air_source}' as Air data -> tab '{AIR_SOURCE_SHEET}'")

    with pd.ExcelWriter(processing_path, engine="openpyxl") as writer:
        for name, frame in frames.items():
            frame.to_excel(writer, sheet_name=name, index=False)

    return processing_path
