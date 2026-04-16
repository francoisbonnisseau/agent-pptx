from __future__ import annotations

from enum import Enum

from pydantic import BaseModel, Field, model_validator


class SlideSummary(BaseModel):
    slide_index: int = Field(ge=1)
    layout_name: str | None = None
    placeholder_names: list[str] = Field(default_factory=list)
    text_preview: str = ""


class DeckAnalysis(BaseModel):
    source_path: str
    slide_count: int = Field(ge=0)
    detected_language: str = "fr"
    extracted_text: str = ""
    thumbnail_paths: list[str] = Field(default_factory=list)
    slides: list[SlideSummary] = Field(default_factory=list)


class StructureOperationType(str, Enum):
    delete_slide = "delete_slide"
    duplicate_slide = "duplicate_slide"
    add_layout_slide = "add_layout_slide"
    reorder_slides = "reorder_slides"


class StructureOperation(BaseModel):
    op: StructureOperationType
    slide_index: int | None = Field(default=None, ge=1)
    target_index: int | None = Field(default=None, ge=1)
    layout_index: int | None = Field(default=None, ge=1)
    new_order: list[int] = Field(default_factory=list)
    reason: str = ""

    @model_validator(mode="after")
    def validate_by_op(self) -> "StructureOperation":
        if self.op == StructureOperationType.delete_slide and self.slide_index is None:
            raise ValueError("delete_slide requires slide_index")
        if (
            self.op == StructureOperationType.duplicate_slide
            and self.slide_index is None
        ):
            raise ValueError("duplicate_slide requires slide_index")
        if (
            self.op == StructureOperationType.add_layout_slide
            and self.layout_index is None
        ):
            raise ValueError("add_layout_slide requires layout_index")
        if self.op == StructureOperationType.reorder_slides and not self.new_order:
            raise ValueError("reorder_slides requires new_order")
        return self


class StructurePlan(BaseModel):
    rationale: str = ""
    operations: list[StructureOperation] = Field(default_factory=list)


class SlideContentUpdate(BaseModel):
    slide_index: int = Field(ge=1)
    title: str | None = None
    subtitle: str | None = None
    bullets: list[str] = Field(default_factory=list)
    body_paragraphs: list[str] = Field(default_factory=list)
    notes: str | None = None
    remove_empty_placeholders: bool = True


class ContentPlan(BaseModel):
    language: str = "fr"
    slides: list[SlideContentUpdate] = Field(default_factory=list)


class QAIssueSeverity(str, Enum):
    low = "low"
    medium = "medium"
    high = "high"


class QAIssueCategory(str, Enum):
    placeholder = "placeholder"
    typo = "typo"
    overflow = "overflow"
    overlap = "overlap"
    alignment = "alignment"
    contrast = "contrast"
    spacing = "spacing"
    other = "other"


class QAIssue(BaseModel):
    slide_index: int = Field(ge=1)
    severity: QAIssueSeverity
    category: QAIssueCategory
    description: str
    fix_hint: str = ""


class QAReport(BaseModel):
    mode: str
    issues: list[QAIssue] = Field(default_factory=list)
    notes: list[str] = Field(default_factory=list)


class RunArtifacts(BaseModel):
    workdir: str
    unpacked_dir: str
    analysis_text_path: str | None = None
    rendered_image_paths: list[str] = Field(default_factory=list)
    report_json_path: str | None = None


class RunReport(BaseModel):
    input_pptx: str
    output_pptx: str
    instruction: str
    structure_plan: StructurePlan
    content_plan: ContentPlan
    qa_reports: list[QAReport]
    final_issue_count: int
    artifacts: RunArtifacts
