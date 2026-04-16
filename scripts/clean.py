from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Clean unreferenced files in unpacked PPTX"
    )
    parser.add_argument("unpacked_dir", help="Unpacked PPTX directory")
    args = parser.parse_args()

    from pptx_agent.structure import clean_unreferenced_files

    removed = clean_unreferenced_files(Path(args.unpacked_dir))
    if not removed:
        print("No unreferenced files found")
        return 0

    print(f"Removed {len(removed)} unreferenced file(s):")
    for item in removed:
        print(f"  {item}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
