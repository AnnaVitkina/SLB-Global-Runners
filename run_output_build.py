"""Build all Global Runners output tabs from the latest processing workbook."""

from __future__ import annotations

import os
import sys
from pathlib import Path

COLAB_CODE_DIR = Path("/content/SLB-Global-Runners")


def _ensure_code_on_path() -> None:
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

from rate_layout_common import latest_processing_file

from run_pipeline import build_all_output_tabs


def main() -> int:
    try:
        build_all_output_tabs(latest_processing_file())
        return 0
    except KeyboardInterrupt:
        print("\nOperation cancelled.")
        return 130
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
