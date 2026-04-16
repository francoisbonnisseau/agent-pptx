from __future__ import annotations

import json
import re
import shutil
import subprocess
from pathlib import Path


def command_exists(command: str) -> bool:
    return shutil.which(command) is not None


def run_command(
    command: list[str], cwd: Path | None = None
) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        command,
        cwd=str(cwd) if cwd else None,
        text=True,
        capture_output=True,
        check=False,
    )


def extract_json_object(raw_text: str) -> dict:
    raw_text = raw_text.strip()
    if not raw_text:
        return {}

    try:
        candidate = raw_text
        if candidate.startswith("```"):
            candidate = re.sub(r"^```(?:json)?", "", candidate).strip()
            candidate = re.sub(r"```$", "", candidate).strip()
        return json.loads(candidate)
    except json.JSONDecodeError:
        pass

    match = re.search(r"\{[\s\S]*\}", raw_text)
    if not match:
        return {}

    try:
        return json.loads(match.group(0))
    except json.JSONDecodeError:
        return {}


def natural_sort_key(path: Path) -> tuple:
    parts: list[str] = re.split(r"(\d+)", path.name)
    key: list[object] = []
    for part in parts:
        if part.isdigit():
            key.append(int(part))
        else:
            key.append(part.lower())
    return tuple(key)
