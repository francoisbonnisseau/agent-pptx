from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


def _load_env_file() -> None:
    project_root = Path(__file__).resolve().parents[1]
    env_path = project_root / ".env"
    if not env_path.exists():
        return

    try:
        from dotenv import load_dotenv
    except Exception:
        return

    load_dotenv(dotenv_path=env_path, override=False)


_load_env_file()


def _env_bool(name: str, default: bool) -> bool:
    raw = os.getenv(name)
    if raw is None:
        return default
    return raw.strip().lower() in {"1", "true", "yes", "on"}


def _env_int(name: str, default: int) -> int:
    raw = os.getenv(name)
    if raw is None:
        return default
    try:
        return int(raw)
    except ValueError:
        return default


@dataclass(slots=True)
class AgentSettings:
    groq_api_key: str | None
    groq_model: str
    gemini_api_key: str | None
    gemini_vision_model: str
    workdir: Path
    max_fix_loops: int
    default_language: str
    enable_visual_qa: bool


def load_settings() -> AgentSettings:
    return AgentSettings(
        groq_api_key=os.getenv("GROQ_API_KEY"),
        groq_model=os.getenv("GROQ_MODEL", "gpt-oss-120b"),
        gemini_api_key=os.getenv("GEMINI_API_KEY"),
        gemini_vision_model=os.getenv("GEMINI_VISION_MODEL", "gemini-2.0-flash"),
        workdir=Path(os.getenv("PPTX_AGENT_WORKDIR", ".pptx_agent_work")).resolve(),
        max_fix_loops=max(1, _env_int("PPTX_AGENT_MAX_FIX_LOOPS", 2)),
        default_language=os.getenv("PPTX_AGENT_DEFAULT_LANGUAGE", "fr").strip().lower(),
        enable_visual_qa=_env_bool("PPTX_AGENT_ENABLE_VISUAL_QA", True),
    )
