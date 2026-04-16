from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


def main() -> int:
    parser = argparse.ArgumentParser(description="Unpack a PPTX to editable XML")
    parser.add_argument("input_pptx", help="Input .pptx file")
    parser.add_argument("output_dir", help="Output unpacked directory")
    args = parser.parse_args()

    from pptx_agent.structure import unpack_pptx

    unpack_pptx(Path(args.input_pptx), Path(args.output_dir))
    print(f"Unpacked: {args.input_pptx} -> {args.output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
