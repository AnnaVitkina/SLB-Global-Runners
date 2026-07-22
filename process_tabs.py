"""
Select Excel files from input/, confirm Ocean/Air tabs (or pick custom tabs),
and write one combined workbook to processing/.
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd

from project_paths import INPUT_DIR, PROCESSING_DIR, ensure_workspace_dirs

EXCEL_SUFFIXES = {".xlsx", ".xls", ".xlsm"}
DEFAULT_TABS = ("Ocean", "Air")

COLUMN_FILTERS: dict[str, dict[str, list[str] | str]] = {
    "ocean": {
        "drop_if_contains": ["detention", "dg classes", "demurrage"],
        "cutoff_after": "import customs clearance fee",
    },
    "air": {
        "drop_if_contains": ["detention"],
        "cutoff_after": "import customs clearance fee",
    },
}


@dataclass(frozen=True)
class SheetSelection:
    file_path: Path
    sheet_name: str

    @property
    def label(self) -> str:
        return f"{self.file_path.name} -> {self.sheet_name}"


@dataclass(frozen=True)
class CombineResult:
    path: Path
    rate_card_name: str


def rate_card_name_from_selections(selections: list[SheetSelection]) -> str:
    """Use the selected source workbook name(s); not every file in input/."""
    stems = list(dict.fromkeys(selection.file_path.stem for selection in selections))
    if not stems:
        raise ValueError("No source files in sheet selections.")
    return stems[0]


def ensure_directories() -> None:
    ensure_workspace_dirs()


def list_input_files() -> list[Path]:
    return [
        path
        for path in sorted(INPUT_DIR.iterdir())
        if path.is_file()
        and path.suffix.lower() in EXCEL_SUFFIXES
        and not path.name.startswith("~$")
    ]


def parse_selection(raw: str, max_index: int) -> list[int]:
    """Parse '1,3-5' into zero-based indices."""
    raw = raw.strip().lower()
    if raw in {"all", "*"}:
        return list(range(max_index))

    indices: set[int] = set()
    for part in re.split(r"\s*,\s*", raw):
        if not part:
            continue
        if "-" in part:
            start_s, end_s = part.split("-", 1)
            start = int(start_s) - 1
            end = int(end_s) - 1
            if start > end or start < 0 or end >= max_index:
                raise ValueError(f"Invalid range: {part}")
            indices.update(range(start, end + 1))
        else:
            idx = int(part) - 1
            if idx < 0 or idx >= max_index:
                raise ValueError(f"Invalid index: {part}")
            indices.add(idx)
    return sorted(indices)


def prompt_selection(
    title: str,
    items: list[str],
    allow_empty: bool = False,
) -> list[int]:
    if not items:
        return []

    print(f"\n{title}")
    for i, item in enumerate(items, start=1):
        print(f"  {i}. {item}")

    hint = "Enter numbers (e.g. 1,3 or 1-3) or 'all'"

    while True:
        raw = input(f"{hint}: ").strip()
        if not raw and allow_empty:
            return []
        if not raw:
            print("Please enter at least one choice.")
            continue
        try:
            chosen = parse_selection(raw, len(items))
            if chosen or allow_empty:
                return chosen
            print("Please enter at least one choice.")
        except ValueError as exc:
            print(f"Invalid input: {exc}")


def ask_yes_no(prompt: str) -> bool:
    while True:
        answer = input(f"{prompt} (y/n): ").strip().lower()
        if answer in {"y", "yes"}:
            return True
        if answer in {"n", "no"}:
            return False
        print("Please enter y or n.")


def select_input_files(files: list[Path]) -> list[Path]:
    if not files:
        print(f"No Excel files found in: {INPUT_DIR}")
        sys.exit(1)

    print("\nAvailable input files:")
    for i, path in enumerate(files, start=1):
        print(f"  {i}. {path.name}")

    indices = prompt_selection(
        "Select file(s) to process (multiple allowed):",
        [f.name for f in files],
    )
    return [files[i] for i in indices]


def default_tabs_for_workbook(sheet_names: list[str]) -> list[str]:
    """Return Ocean and Air when present (case-insensitive match)."""
    by_lower = {name.lower(): name for name in sheet_names}
    selected: list[str] = []
    for tab in DEFAULT_TABS:
        actual = by_lower.get(tab.lower())
        if actual is not None:
            selected.append(actual)
    return selected


def select_sheets(file_path: Path, sheet_names: list[str]) -> list[str]:
    default_tabs = default_tabs_for_workbook(sheet_names)

    if default_tabs:
        print(f"\nDefault tabs for {file_path.name}: {', '.join(default_tabs)}")
        if ask_yes_no("Use these tabs?"):
            return default_tabs

    indices = prompt_selection(
        f"Choose sheet(s) for {file_path.name}:",
        sheet_names,
    )
    return [sheet_names[i] for i in indices]


def gather_default_sheet_selections(files: list[Path]) -> list[SheetSelection]:
    selections: list[SheetSelection] = []

    for file_path in files:
        try:
            workbook = pd.ExcelFile(file_path)
        except Exception as exc:
            print(f"Skipping {file_path.name}: could not open file ({exc})")
            continue

        selected_sheets = default_tabs_for_workbook(workbook.sheet_names)
        if not selected_sheets:
            print(f"No default Ocean/Air tabs in {file_path.name}; skipping.")
            continue

        for sheet_name in selected_sheets:
            selections.append(SheetSelection(file_path, sheet_name))

    return selections


def gather_sheet_selections(files: list[Path]) -> list[SheetSelection]:
    selections: list[SheetSelection] = []

    for file_path in files:
        try:
            workbook = pd.ExcelFile(file_path)
        except Exception as exc:
            print(f"Skipping {file_path.name}: could not open file ({exc})")
            continue

        selected_sheets = select_sheets(file_path, workbook.sheet_names)
        if not selected_sheets:
            print(f"No sheets selected for {file_path.name}; skipping.")
            continue

        for sheet_name in selected_sheets:
            selections.append(SheetSelection(file_path, sheet_name))

    return selections


MAX_SHEET_NAME_LEN = 31


def sanitize_sheet_part(text: str) -> str:
    return re.sub(r"[\[\]:*?/\\]", "_", text).strip() or "Sheet"


def combine_file_and_sheet(file_stem: str, sheet_name: str, max_len: int = MAX_SHEET_NAME_LEN) -> str:
    file_part = sanitize_sheet_part(file_stem)
    sheet_part = sanitize_sheet_part(sheet_name)
    separator = "_"

    combined = f"{file_part}{separator}{sheet_part}"
    if len(combined) <= max_len:
        return combined

    max_file_len = max_len - len(separator) - len(sheet_part)
    if max_file_len >= 1:
        return f"{file_part[:max_file_len]}{separator}{sheet_part}"

    return sheet_part[:max_len]


def output_sheet_name(
    sheet_name: str,
    used: set[str],
    file_stem: str | None = None,
) -> str:
    if file_stem:
        base = combine_file_and_sheet(file_stem, sheet_name)
    else:
        base = sanitize_sheet_part(sheet_name)

    if base not in used:
        used.add(base)
        return base

    for n in range(2, 1000):
        suffix = f"_{n}"
        if file_stem:
            trimmed = combine_file_and_sheet(
                file_stem,
                sheet_name,
                max_len=MAX_SHEET_NAME_LEN - len(suffix),
            )
            candidate = trimmed + suffix
        else:
            trimmed = sanitize_sheet_part(sheet_name)[: MAX_SHEET_NAME_LEN - len(suffix)]
            candidate = trimmed + suffix
        if candidate not in used:
            used.add(candidate)
            return candidate

    raise RuntimeError(f"Could not create a unique sheet name for: {file_stem or ''} / {sheet_name}")

HEADER_MARKER = "status"


def _cell_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def find_header_row_index(df_raw: pd.DataFrame, max_scan_rows: int = 50) -> int | None:
    scan_limit = min(len(df_raw), max_scan_rows)
    for row_idx in range(scan_limit):
        first_cell = _cell_text(df_raw.iloc[row_idx, 0])
        if first_cell.lower() == HEADER_MARKER:
            return row_idx
    return None


def _normalize_headers(headers: list[object]) -> list[str]:
    cleaned: list[str] = []
    seen: dict[str, int] = {}

    for idx, header in enumerate(headers, start=1):
        value = _cell_text(header)
        if not value or value.lower() == "nan":
            value = f"column_{idx}"

        base = value
        count = seen.get(base, 0)
        if count:
            value = f"{base}_{count + 1}"
        seen[base] = count + 1
        cleaned.append(value)

    return cleaned


def _drop_empty_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    non_empty_mask = df.apply(
        lambda row: any(_cell_text(value) for value in row),
        axis=1,
    )
    return df.loc[non_empty_mask].copy()


def _drop_empty_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    keep_columns = [
        column
        for column in df.columns
        if any(_cell_text(value) for value in df[column])
    ]
    return df.loc[:, keep_columns].copy()


def _normalize_column_name(name: object) -> str:
    return _cell_text(name).lower()


def filter_columns_by_tab(df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    """Drop tab-specific unwanted columns."""
    if df.empty:
        return df

    rules = COLUMN_FILTERS.get(sheet_name.strip().lower())
    if not rules:
        return df

    drop_keywords = [str(keyword).lower() for keyword in rules["drop_if_contains"]]
    cutoff_marker = str(rules["cutoff_after"]).lower()

    columns = list(df.columns)
    cutoff_idx: int | None = None
    for idx, column in enumerate(columns):
        if cutoff_marker in _normalize_column_name(column):
            cutoff_idx = idx
            break

    columns_to_keep = columns if cutoff_idx is None else columns[:cutoff_idx]
    if cutoff_idx is None:
        print(
            f"  Warning: cutoff column containing '{rules['cutoff_after']}' "
            f"not found in '{sheet_name}'."
        )

    kept_columns = [
        column
        for column in columns_to_keep
        if not any(keyword in _normalize_column_name(column) for keyword in drop_keywords)
    ]
    return df.loc[:, kept_columns].copy()


def clean_tab_df(df_raw: pd.DataFrame) -> pd.DataFrame:
    """Remove blank rows above the header and fully empty rows/columns."""
    if df_raw.empty:
        return df_raw.copy()

    header_row_idx = find_header_row_index(df_raw)
    if header_row_idx is None:
        cleaned = _drop_empty_rows(df_raw)
        cleaned.columns = _normalize_headers(list(cleaned.columns))
        return cleaned.reset_index(drop=True)

    headers = _normalize_headers(df_raw.iloc[header_row_idx].tolist())
    df = df_raw.iloc[header_row_idx + 1 :].copy()
    df.columns = headers
    df = _drop_empty_rows(df)
    df = _drop_empty_columns(df)
    return df.reset_index(drop=True)


def tab_to_df(file_path: Path, sheet_name: str) -> pd.DataFrame:
    raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    df = clean_tab_df(raw)
    return filter_columns_by_tab(df, sheet_name)


def collect_selected_frames(
    selections: list[SheetSelection],
) -> list[tuple[str, pd.DataFrame]]:
    frames: list[tuple[str, pd.DataFrame]] = []
    used_names: set[str] = set()
    single_file = len({selection.file_path for selection in selections}) == 1

    for selection in selections:
        try:
            raw = pd.read_excel(selection.file_path, sheet_name=selection.sheet_name, header=None)
            df = clean_tab_df(raw)
            removed_rows = len(raw) - len(df)
            before_cols = len(df.columns)
            df = filter_columns_by_tab(df, selection.sheet_name)
            removed_cols = before_cols - len(df.columns)
        except Exception as exc:
            print(f"Skipping {selection.label}: could not read sheet ({exc})")
            continue

        if single_file:
            sheet_label = output_sheet_name(selection.sheet_name, used_names)
        else:
            sheet_label = output_sheet_name(
                selection.sheet_name,
                used_names,
                file_stem=selection.file_path.stem,
            )
        frames.append((sheet_label, df))
        print(
            f"  Loaded: {selection.label} -> output tab '{sheet_label}' "
            f"({len(df)} rows, {len(df.columns)} columns; "
            f"removed {removed_rows} rows, {removed_cols} columns)"
        )

    return frames


def save_combined_workbook(
    frames: list[tuple[str, pd.DataFrame]],
    *,
    rate_card_name: str,
) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = re.sub(r'[\\/*?:\[\]]+', "_", rate_card_name).strip()
    output_path = PROCESSING_DIR / f"combined_{safe_name}_{timestamp}.xlsx"

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in frames:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"  Wrote tab: {sheet_name}")

    return output_path


def run_combine_tabs(*, auto: bool = False) -> CombineResult:
    """Combine selected input tabs into one processing workbook."""
    ensure_directories()
    files = list_input_files()

    if auto:
        if not files:
            raise FileNotFoundError(f"No Excel files found in: {INPUT_DIR}")
        print(f"Auto mode: processing {len(files)} file(s) with default Ocean/Air tabs.")
        selections = gather_default_sheet_selections(files)
    else:
        selected_files = select_input_files(files)
        selections = gather_sheet_selections(selected_files)

    if not selections:
        raise RuntimeError("Nothing to save. No sheets were selected.")

    rate_card_name = rate_card_name_from_selections(selections)
    frames = collect_selected_frames(selections)

    if not frames:
        raise RuntimeError("Nothing to save. No sheets could be loaded.")

    output_path = save_combined_workbook(frames, rate_card_name=rate_card_name)
    print(f"\nSaved {len(frames)} sheet(s) to: {output_path}")
    return CombineResult(path=output_path, rate_card_name=rate_card_name)


def main() -> int:
    try:
        run_combine_tabs()
        return 0
    except KeyboardInterrupt:
        print("\nOperation cancelled.")
        return 130
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
