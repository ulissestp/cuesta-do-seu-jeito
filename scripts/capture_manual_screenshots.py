"""Captura telas do app desktop para o manual em PDF (janela CUESTA DO SEU JEITO)."""
from __future__ import annotations

import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "docs" / "screenshots"

if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

TAB_SPECS: list[tuple[str, str]] = [
    ("01_perfil.png", "1. Perfil"),
    ("02_inventario.png", "2. Inventário"),
    ("03_roteiro_pdf.png", "3. Roteiro PDF"),
    ("04_ficha_anamnese.png", "4. Ficha anamnese"),
]


def _find_window():
    import pygetwindow as gw

    for w in gw.getAllWindows():
        title = w.title or ""
        if "CUESTA DO SEU JEITO" in title or "RoteiroTurismo" in title:
            return w
    raise RuntimeError(
        "Abra o programa (RoteiroTurismo.exe ou python desktop_main.py) e rode este script de novo."
    )


def _grab_window(path: Path, left: int, top: int, width: int, height: int) -> None:
    import mss
    import mss.tools

    path.parent.mkdir(parents=True, exist_ok=True)
    with mss.MSS() as sct:
        shot = sct.grab({"left": left, "top": top, "width": width, "height": height})
        mss.tools.to_png(shot.rgb, shot.size, output=str(path))


def main() -> None:
    import pyautogui

    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.15

    win = _find_window()
    try:
        win.activate()
    except Exception:
        pass
    time.sleep(0.8)

    left, top, width, height = win.left, win.top, win.width, win.height
    if width < 400 or height < 300:
        raise RuntimeError(f"Janela muito pequena: {width}x{height}")

    # Abas CTkTabview (abaixo do cabeçalho ~72px)
    tab_y = top + 118
    margin = 130
    step = max(118, (width - 2 * margin) // 4)
    tab_x = [left + margin + i * step for i in range(4)]

    for (filename, _label), tx in zip(TAB_SPECS, tab_x):
        pyautogui.click(tx, tab_y)
        time.sleep(1.0)
        _grab_window(OUT / filename, left, top, width, height)
        print(f"Capturado: {OUT / filename}")

    print("Concluído. Rode: python scripts/gerar_manual_pdf.py")


if __name__ == "__main__":
    main()
