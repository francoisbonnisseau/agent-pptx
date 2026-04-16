from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from .analysis import analyze_template
from .config import AgentSettings
from .content import apply_content_plan
from .genai import plan_content, plan_content_fixes, plan_structure
from .models import QAReport, RunArtifacts, RunReport
from .qa import merge_issue_counts, run_content_qa, run_visual_qa_with_gemini
from .structure import (
    apply_structure_plan,
    clean_unreferenced_files,
    pack_pptx,
    unpack_pptx,
)


def _count_major_issues(reports: list[QAReport]) -> int:
    return sum(
        1
        for report in reports
        for issue in report.issues
        if issue.severity.value in {"high", "medium"}
    )


class PPTXEditingPipeline:
    def __init__(self, settings: AgentSettings):
        self.settings = settings

    def _new_run_dir(self) -> Path:
        now = datetime.now().strftime("%Y%m%d-%H%M%S")
        run_dir = self.settings.workdir / f"run-{now}"
        run_dir.mkdir(parents=True, exist_ok=True)
        return run_dir

    def _run_qa(
        self, pptx_path: Path, run_dir: Path, pass_index: int
    ) -> list[QAReport]:
        reports: list[QAReport] = []
        reports.append(run_content_qa(pptx_path))

        if self.settings.enable_visual_qa:
            visual_dir = run_dir / f"qa-pass-{pass_index}"
            reports.append(
                run_visual_qa_with_gemini(
                    pptx_path,
                    output_dir=visual_dir,
                    gemini_api_key=self.settings.gemini_api_key,
                    model_name=self.settings.gemini_vision_model,
                )
            )

        return reports

    def run(self, input_pptx: Path, output_pptx: Path, instruction: str) -> RunReport:
        input_pptx = input_pptx.resolve()
        output_pptx = output_pptx.resolve()
        if not input_pptx.exists():
            raise FileNotFoundError(f"Input PPTX introuvable: {input_pptx}")

        run_dir = self._new_run_dir()
        analysis_dir = run_dir / "analysis"
        analysis = analyze_template(
            input_pptx, analysis_dir, default_language=self.settings.default_language
        )

        structure_plan = plan_structure(self.settings, analysis, instruction)

        unpacked_dir = run_dir / "unpacked"
        unpack_pptx(input_pptx, unpacked_dir)
        apply_structure_plan(unpacked_dir, structure_plan)

        structured_pptx = run_dir / "structured.pptx"
        pack_pptx(unpacked_dir, structured_pptx)

        post_structure_analysis = analyze_template(
            structured_pptx,
            run_dir / "analysis-post-structure",
            default_language=analysis.detected_language
            or self.settings.default_language,
        )

        language = (
            post_structure_analysis.detected_language or self.settings.default_language
        )
        content_plan = plan_content(
            self.settings, post_structure_analysis, instruction, language=language
        )

        current_pptx = run_dir / "edited.pptx"
        apply_content_plan(structured_pptx, content_plan, current_pptx)

        qa_reports: list[QAReport] = []
        latest_reports = self._run_qa(current_pptx, run_dir, pass_index=0)
        qa_reports.extend(latest_reports)

        major_issues = _count_major_issues(latest_reports)
        zero_defect_verified = False

        if major_issues == 0:
            verification_reports = self._run_qa(current_pptx, run_dir, pass_index=1)
            qa_reports.extend(verification_reports)
            latest_reports = verification_reports
            major_issues = _count_major_issues(verification_reports)
            zero_defect_verified = major_issues == 0

        for loop_index in range(1, self.settings.max_fix_loops + 1):
            if major_issues == 0:
                zero_defect_verified = True
                break

            fix_plan = plan_content_fixes(
                self.settings,
                post_structure_analysis,
                instruction,
                language=language,
                reports=latest_reports,
            )

            if not fix_plan.slides:
                break

            apply_content_plan(current_pptx, fix_plan, current_pptx)

            latest_reports = self._run_qa(
                current_pptx, run_dir, pass_index=loop_index + 1
            )
            qa_reports.extend(latest_reports)
            major_issues = _count_major_issues(latest_reports)

        if not zero_defect_verified:
            raise RuntimeError(
                "QA gate not satisfied: no zero-defect verification pass was reached. "
                "Increase max fix loops or review generated content."
            )

        final_unpack = run_dir / "final-unpacked"
        unpack_pptx(current_pptx, final_unpack)
        clean_unreferenced_files(final_unpack)
        pack_pptx(final_unpack, output_pptx)

        rendered_paths = [str(path) for path in sorted(run_dir.rglob("qa-slide-*.png"))]

        report = RunReport(
            input_pptx=str(input_pptx),
            output_pptx=str(output_pptx),
            instruction=instruction,
            structure_plan=structure_plan,
            content_plan=content_plan,
            qa_reports=qa_reports,
            final_issue_count=merge_issue_counts(latest_reports),
            artifacts=RunArtifacts(
                workdir=str(run_dir),
                unpacked_dir=str(unpacked_dir),
                analysis_text_path=str(analysis_dir / "analysis.txt"),
                rendered_image_paths=rendered_paths,
                report_json_path=str(run_dir / "run-report.json"),
            ),
        )

        report_path = run_dir / "run-report.json"
        report_path.write_text(report.model_dump_json(indent=2), encoding="utf-8")

        return report


def save_report(report: RunReport, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(report.model_dump_json(indent=2), encoding="utf-8")


def print_human_report(report: RunReport) -> str:
    qa_lines = [
        f"- {item.mode}: {len(item.issues)} issue(s)" for item in report.qa_reports
    ]
    payload = {
        "input": report.input_pptx,
        "output": report.output_pptx,
        "final_issue_count": report.final_issue_count,
        "qa": qa_lines,
        "report_json": report.artifacts.report_json_path,
        "workdir": report.artifacts.workdir,
    }
    return json.dumps(payload, ensure_ascii=False, indent=2)
