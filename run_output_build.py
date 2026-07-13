"""Build all Global Runners output tabs from the latest processing workbook."""

from __future__ import annotations

import sys

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
