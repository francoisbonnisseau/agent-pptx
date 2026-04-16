from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


def _parse_layout_index(source: str) -> int | None:
    match = re.fullmatch(r"slideLayout(\d+)\.xml", source)
    if not match:
        return None
    return int(match.group(1))


def main() -> int:
    parser = argparse.ArgumentParser(description="Add slide to unpacked PPTX")
    parser.add_argument("unpacked_dir", help="Unpacked PPTX directory")
    parser.add_argument(
        "source",
        help="slideN.xml to duplicate OR slideLayoutN.xml to create from layout",
    )
    parser.add_argument("--target-index", type=int, help="Insert position (1-based)")
    args = parser.parse_args()

    from pptx_agent.structure import (
        add_slide_from_layout,
        duplicate_slide,
        list_slide_sequence,
    )

    unpacked = Path(args.unpacked_dir)
    target_index = args.target_index

    layout_index = _parse_layout_index(args.source)
    if layout_index is not None:
        created = add_slide_from_layout(
            unpacked, layout_index=layout_index, target_index=target_index
        )
        print(f"Created {created} from layout slideLayout{layout_index}.xml")
        return 0

    sequence = list_slide_sequence(unpacked)
    if args.source not in sequence:
        raise SystemExit(
            f"Slide source introuvable dans l'ordre courant: {args.source}. "
            f"Slides disponibles: {', '.join(sequence)}"
        )

    slide_index = sequence.index(args.source) + 1
    created = duplicate_slide(
        unpacked, slide_index=slide_index, target_index=target_index
    )
    print(f"Duplicated {args.source} -> {created}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
