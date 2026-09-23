# -*- coding: utf-8 -*-
"""Etapa Google Ads do chatbot Radar.

Roda APENAS a coleta Google Ads Transparency (sem Meta/LinkedIn/redes).

Uso:
    python "5a - google ads.py" chatguru
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

FINAL_DIR = Path(__file__).resolve().parent
CORE = FINAL_DIR / "5 - metricas concorrentes.py"


def main() -> None:
    if len(sys.argv) < 2:
        print('Uso: python "5a - google ads.py" <cliente>')
        sys.exit(1)
    client = sys.argv[1].strip()
    cmd = [sys.executable, str(CORE), client, "google_ads"]
    raise SystemExit(subprocess.call(cmd, cwd=str(FINAL_DIR)))


if __name__ == "__main__":
    main()
