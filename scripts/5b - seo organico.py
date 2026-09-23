# -*- coding: utf-8 -*-
"""Etapa SEO organico do chatbot Radar.

Monta o ranking SEO a partir do Semrush ja coletado no script 3
(Trafego SEO Marca). Nao chama ScrapingBee.

Uso:
    python "5b - seo organico.py" chatguru
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

FINAL_DIR = Path(__file__).resolve().parent
CORE = FINAL_DIR / "5 - metricas concorrentes.py"


def main() -> None:
    if len(sys.argv) < 2:
        print('Uso: python "5b - seo organico.py" <cliente>')
        sys.exit(1)
    client = sys.argv[1].strip()
    cmd = [sys.executable, str(CORE), client, "seo"]
    raise SystemExit(subprocess.call(cmd, cwd=str(FINAL_DIR)))


if __name__ == "__main__":
    main()
