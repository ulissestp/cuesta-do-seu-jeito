"""CLI: gera ficha de anamnese (legado). Prefira src.anamnese_pdf pelo aplicativo."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.anamnese_pdf import anamnese_pdf_filename, build_anamnese_pdf


def main() -> None:
    out_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "output"
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / anamnese_pdf_filename("Modelo")
    build_anamnese_pdf(path, nome_grupo="")
    print(f"Gerado: {path}")


if __name__ == "__main__":
    main()
