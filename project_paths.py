"""Project folder paths for local and Google Colab runs."""

from __future__ import annotations

import os
import sys
from pathlib import Path

_LOCAL_ROOT = Path(__file__).resolve().parent
CODE_DIR = _LOCAL_ROOT

COLAB_CODE_DIR = Path("/content/SLB-Global-Runners")

_DRIVE_DATA_ROOT = Path(
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team "
    "/Documents/AI Adoption RMT/RMT_SLB/RMT_Global_Runners"
)


def is_colab_environment() -> bool:
    if os.environ.get("SLB_COLAB") == "1":
        return True
    if COLAB_CODE_DIR.is_dir():
        return True
    try:
        import google.colab  # type: ignore[import-not-found]

        return True
    except ImportError:
        return False


def _resolve_data_paths() -> tuple[Path, Path, Path]:
    input_override = os.environ.get("SLB_INPUT_DIR")
    processing_override = os.environ.get("SLB_PROCESSING_DIR")
    output_override = os.environ.get("SLB_OUTPUT_DIR")

    if input_override and processing_override and output_override:
        return Path(input_override), Path(processing_override), Path(output_override)

    data_root_override = os.environ.get("SLB_DATA_ROOT")
    if data_root_override:
        data_root = Path(data_root_override)
        return data_root / "input", data_root / "processing", data_root / "output"

    if _DRIVE_DATA_ROOT.is_dir():
        return (
            _DRIVE_DATA_ROOT / "input",
            _DRIVE_DATA_ROOT / "processing",
            _DRIVE_DATA_ROOT / "output",
        )

    return _LOCAL_ROOT / "input", _LOCAL_ROOT / "processing", _LOCAL_ROOT / "output"


INPUT_DIR, PROCESSING_DIR, OUTPUT_DIR = _resolve_data_paths()


def ensure_workspace_dirs() -> None:
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    PROCESSING_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def bootstrap_colab_runtime() -> None:
    """Prepare sys.path and cwd when running from /content/SLB-Global-Runners."""
    if not COLAB_CODE_DIR.is_dir():
        return

    code_dir = str(COLAB_CODE_DIR)
    if code_dir not in sys.path:
        sys.path.insert(0, code_dir)
    os.chdir(code_dir)


def print_path_config() -> None:
    print("Global Runners paths:")
    print(f"  Code:       {CODE_DIR}")
    print(f"  Input:      {INPUT_DIR}")
    print(f"  Processing: {PROCESSING_DIR}")
    print(f"  Output:     {OUTPUT_DIR}")
