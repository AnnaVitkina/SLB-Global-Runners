"""
Global Runners — end-to-end rate pipeline.

Steps:
  1. Select Excel files and Ocean/Air tabs from input/
  2. Clean and combine tabs into processing/
  3. Build all output rate tabs into output/

Usage:
  python run_pipeline.py          # interactive file/tab selection
  python run_pipeline.py --auto   # all input files, default Ocean/Air tabs

Google Colab:
  from google.colab import drive
  drive.mount('/content/drive')

  import os
  exec(open('/content/SLB-Global-Runners/run_pipeline.py').read())

  Ensure the repo lives at /content/SLB-Global-Runners (or set SLB_CODE_DIR).
  Data folders on Drive (input / processing / output) are picked up automatically.
"""

from __future__ import annotations

import argparse
import os
import sys
from dataclasses import dataclass
from pathlib import Path

COLAB_CODE_DIR = Path("/content/SLB-Global-Runners")


def _ensure_code_on_path() -> None:
    """Make package imports work when run_pipeline.py is exec()'d in Colab."""
    env_code_dir = os.environ.get("SLB_CODE_DIR")
    candidates: list[Path] = []
    if env_code_dir:
        candidates.append(Path(env_code_dir))
    candidates.append(COLAB_CODE_DIR)
    script_file = globals().get("__file__")
    if script_file:
        candidates.append(Path(script_file).resolve().parent)

    for candidate in candidates:
        if (candidate / "project_paths.py").is_file():
            code_dir = str(candidate.resolve())
            if code_dir not in sys.path:
                sys.path.insert(0, code_dir)
            os.chdir(code_dir)
            return


_ensure_code_on_path()

from project_paths import bootstrap_colab_runtime, is_colab_environment, print_path_config

bootstrap_colab_runtime()

from build_air_rates import build_air_rates
from build_conditions import CONDITIONS_SHEET, finalize_output_workbook
from build_road_precarriage_air import build_road_precarriage_air
from build_road_precarriage_sea import build_road_precarriage_sea
from build_sea_rates import build_sea_rates
from process_tabs import run_combine_tabs
from rate_layout_common import output_workbook_path

OUTPUT_TAB_NAMES = (
    "Sea Rates",
    "Road pre-carriage for Sea",
    "Air MAWB rate",
    "Air HAWB rate",
    "Road pre-carriage for Air MAWB",
    "Road pre-carriage for Air HAWB",
    CONDITIONS_SHEET,
)


@dataclass(frozen=True)
class PipelineResult:
    processing_path: Path
    output_path: Path
    output_tabs: tuple[str, ...]


def build_all_output_tabs(
    processing_path: Path,
    output_path: Path | None = None,
) -> Path:
    path = output_path or output_workbook_path()

    build_sea_rates(processing_path=processing_path, output_path=path)
    build_road_precarriage_sea(processing_path=processing_path, output_path=path)
    build_air_rates(processing_path=processing_path, output_path=path)
    build_road_precarriage_air(processing_path=processing_path, output_path=path)

    print("\n=== Applying conditions and building Conditions tab ===")
    finalize_output_workbook(path)

    print(f"\nAll {len(OUTPUT_TAB_NAMES)} output tabs saved -> {path}")
    return path


def run_pipeline(
    *,
    auto: bool = False,
    output_path: Path | None = None,
) -> PipelineResult:
    print_path_config()
    print("=== Step 1/2: Select and combine input tabs ===")
    processing_path = run_combine_tabs(auto=auto)

    print("\n=== Step 2/2: Build output rate tabs ===")
    saved_output_path = build_all_output_tabs(processing_path, output_path=output_path)

    print("\n=== Pipeline complete ===")
    print(f"  Processing workbook: {processing_path}")
    print(f"  Output workbook:     {saved_output_path}")
    print("  Output tabs:")
    for tab_name in OUTPUT_TAB_NAMES:
        print(f"    - {tab_name}")

    return PipelineResult(
        processing_path=processing_path,
        output_path=saved_output_path,
        output_tabs=OUTPUT_TAB_NAMES,
    )


def _running_in_notebook() -> bool:
    if is_colab_environment():
        return True
    if not sys.argv:
        return False
    argv0 = Path(sys.argv[0]).name.lower()
    return any(
        marker in argv0
        for marker in ("ipykernel_launcher", "colab_kernel_launcher", "jupyter")
    )


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run the Global Runners end-to-end pipeline.")
    parser.add_argument(
        "--auto",
        action="store_true",
        help="Use all input files and default Ocean/Air tabs without prompts.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Optional output workbook path (defaults to output/<rate card>_output.xlsx).",
    )
    return parser.parse_args()


def main() -> int:
    try:
        if _running_in_notebook():
            run_pipeline(auto=True)
            return 0

        args = _parse_args()
        run_pipeline(auto=args.auto, output_path=args.output)
        return 0
    except KeyboardInterrupt:
        print("\nOperation cancelled.")
        return 130
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    exit_code = main()
    if not _running_in_notebook():
        raise SystemExit(exit_code)
