from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


def _read_instruction(args: argparse.Namespace) -> str:
    if args.instruction:
        return args.instruction.strip()
    if args.instruction_file:
        return Path(args.instruction_file).read_text(encoding="utf-8").strip()
    raise ValueError("Fournir --instruction ou --instruction-file")


def main() -> int:
    parser = argparse.ArgumentParser(description="Run autonomous PPTX editing agent")
    parser.add_argument("--input", required=True, help="Input PPTX template")
    parser.add_argument("--output", required=True, help="Output PPTX path")
    parser.add_argument("--instruction", help="Instruction text")
    parser.add_argument(
        "--instruction-file", help="Path to instruction markdown/text file"
    )
    parser.add_argument("--report-out", help="Optional explicit report JSON path")
    args = parser.parse_args()

    from pptx_agent.config import load_settings
    from pptx_agent.pipeline import PPTXEditingPipeline, print_human_report, save_report

    instruction = _read_instruction(args)
    settings = load_settings()
    pipeline = PPTXEditingPipeline(settings)

    try:
        report = pipeline.run(
            input_pptx=Path(args.input),
            output_pptx=Path(args.output),
            instruction=instruction,
        )
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    if args.report_out:
        save_report(report, Path(args.report_out))

    print(print_human_report(report))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
