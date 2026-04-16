from __future__ import annotations

from pathlib import Path

from .config import AgentSettings, load_settings
from .models import RunReport
from .pipeline import PPTXEditingPipeline


def run_autonomous_agent(
    input_pptx: str | Path,
    output_pptx: str | Path,
    instruction: str,
    settings: AgentSettings | None = None,
) -> RunReport:
    cfg = settings or load_settings()
    pipeline = PPTXEditingPipeline(cfg)
    return pipeline.run(Path(input_pptx), Path(output_pptx), instruction)
