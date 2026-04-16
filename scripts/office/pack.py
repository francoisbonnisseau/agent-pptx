from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


def main() -> int:
    parser = argparse.ArgumentParser(description="Pack an unpacked PPTX directory")
    parser.add_argument("input_directory", help="Unpacked directory")
    parser.add_argument("output_pptx", help="Output PPTX file")
    args = parser.parse_args()

    from pptx_agent.structure import pack_pptx

    pack_pptx(Path(args.input_directory), Path(args.output_pptx))
    print(f"Packed: {args.input_directory} -> {args.output_pptx}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
