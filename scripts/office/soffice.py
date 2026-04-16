from __future__ import annotations

import os
import shutil
import subprocess
import sys


def _resolve_soffice() -> str | None:
    env = os.getenv("SOFFICE_BIN")
    if env:
        return env
    for candidate in ("soffice", "libreoffice"):
        found = shutil.which(candidate)
        if found:
            return found
    return None


def main() -> int:
    soffice = _resolve_soffice()
    if soffice is None:
        print("Erreur: soffice/libreoffice introuvable dans PATH", file=sys.stderr)
        return 1

    process = subprocess.run([soffice, *sys.argv[1:]], check=False)
    return process.returncode


if __name__ == "__main__":
    raise SystemExit(main())
