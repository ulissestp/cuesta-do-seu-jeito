"""Gera PDF do manual do usuário a partir de docs/MANUAL_DO_USUARIO.md."""
from __future__ import annotations

import html
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from src.app_paths import logo_path
from src.branding import (
    APP_NAME,
    APP_TAGLINE,
    APP_VERSION,
    desktop_window_title,
    pdf_version_label,
)

MANUAL_MD = ROOT / "docs" / "MANUAL_DO_USUARIO.md"
SCREENSHOTS = ROOT / "docs" / "screenshots"
OUTPUT = ROOT / "docs" / "Manual_CUESTA_RT_v1.1.pdf"
PLACEHOLDER_H = 10.5 * cm

PAGE_W, PAGE_H = A4
MARGIN = 2 * cm

# Figura inserida antes destes capítulos (após o texto do capítulo anterior).
SCREENSHOT_BEFORE_HEADING: dict[str, tuple[str, str]] = {
    "5. Aba 2": ("01_perfil.png", "Aba 1 — Perfil do turista"),
    "6. Aba 3": ("02_inventario.png", "Aba 2 — Inventário turístico"),
    "7. Aba 4": ("03_roteiro_pdf.png", "Aba 3 — Geração do roteiro em PDF"),
    "8. Pasta": ("04_ficha_anamnese.png", "Aba 4 — Ficha de anamnese"),
}


def _styles():
    base = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            "DocTitle",
            parent=base["Title"],
            fontSize=20,
            textColor=colors.HexColor("#1a365d"),
            spaceAfter=12,
            alignment=TA_CENTER,
        ),
        "h1": ParagraphStyle(
            "H1",
            parent=base["Heading1"],
            fontSize=14,
            textColor=colors.HexColor("#2c5282"),
            spaceBefore=16,
            spaceAfter=8,
        ),
        "h2": ParagraphStyle(
            "H2",
            parent=base["Heading2"],
            fontSize=11,
            textColor=colors.HexColor("#1a365d"),
            spaceBefore=10,
            spaceAfter=6,
        ),
        "body": ParagraphStyle(
            "Body",
            parent=base["BodyText"],
            fontSize=10,
            leading=14,
            alignment=TA_JUSTIFY,
        ),
        "bullet": ParagraphStyle(
            "Bullet",
            parent=base["BodyText"],
            fontSize=10,
            leading=13,
            leftIndent=14,
            bulletIndent=0,
        ),
        "caption": ParagraphStyle(
            "Caption",
            parent=base["Normal"],
            fontSize=9,
            textColor=colors.HexColor("#4a5568"),
            alignment=TA_CENTER,
            spaceAfter=10,
        ),
        "code": ParagraphStyle(
            "CodeBlock",
            parent=base["Code"],
            fontSize=9,
            leading=12,
            fontName="Courier",
            leftIndent=8,
            backColor=colors.HexColor("#EDF2F7"),
        ),
    }


def md_inline(text: str) -> str:
    text = html.escape(text, quote=False)
    text = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", text)
    text = re.sub(
        r"`([^`]+)`",
        r'<font face="Courier" size="9">\1</font>',
        text,
    )
    text = text.replace("&lt;b&gt;", "<b>").replace("&lt;/b&gt;", "</b>")
    return text


def _cover(st) -> list:
    story: list = []
    logo = logo_path()
    if logo.is_file():
        story.append(Spacer(1, 1.2 * cm))
        story.append(Image(str(logo), width=5 * cm, height=5 * cm, kind="proportional"))
        story.append(Spacer(1, 0.5 * cm))
    story.append(Paragraph(desktop_window_title(), st["title"]))
    story.append(Paragraph(f"{APP_TAGLINE}<br/>{pdf_version_label()}", st["body"]))
    story.append(Spacer(1, 0.8 * cm))
    story.append(
        Paragraph(
            "<b>Manual do usuário</b> gerado a partir de "
            "<font face='Courier'>docs/MANUAL_DO_USUARIO.md</font>.",
            st["body"],
        )
    )
    story.append(PageBreak())
    return story


def _screenshot_placeholder(st, caption: str, hint: str = "") -> list:
    box_w = PAGE_W - 2 * MARGIN
    label = Paragraph(
        f"<font color='#718096'><i>Inserir captura de tela aqui</i></font>"
        f"{f'<br/><font size=\"8\">{hint}</font>' if hint else ''}",
        ParagraphStyle(
            "PlaceholderInner",
            parent=st["caption"],
            alignment=TA_CENTER,
            leading=12,
        ),
    )
    table = Table([[label]], colWidths=[box_w], rowHeights=[PLACEHOLDER_H])
    table.setStyle(
        TableStyle(
            [
                ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#CBD5E0")),
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#EDF2F7")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ]
        )
    )
    return [
        Spacer(1, 0.2 * cm),
        table,
        Paragraph(f"<b>Figura:</b> {caption}", st["caption"]),
        Spacer(1, 0.35 * cm),
    ]


def _screenshot(st, filename: str, caption: str, *, use_image: bool = False) -> list:
    if not use_image:
        return _screenshot_placeholder(st, caption, hint=filename)
    path = SCREENSHOTS / filename
    if not path.is_file():
        return _screenshot_placeholder(st, caption, hint=filename)
    img = Image(str(path))
    max_w = PAGE_W - 2 * MARGIN
    max_h = PLACEHOLDER_H
    iw, ih = img.imageWidth, img.imageHeight
    scale = min(max_w / iw, max_h / ih, 1.0)
    img.drawWidth = iw * scale
    img.drawHeight = ih * scale
    return [
        Spacer(1, 0.2 * cm),
        img,
        Paragraph(f"<b>Figura:</b> {caption}", st["caption"]),
        Spacer(1, 0.35 * cm),
    ]


def _screenshot_for_heading(st, heading: str, *, use_image: bool) -> list:
    for key, (filename, caption) in SCREENSHOT_BEFORE_HEADING.items():
        if heading.startswith(key) or key in heading:
            return _screenshot(st, filename, caption, use_image=use_image)
    return []


def _parse_table(lines: list[str], start: int) -> tuple[list[list[str]], int]:
    rows: list[list[str]] = []
    i = start
    while i < len(lines) and lines[i].strip().startswith("|"):
        raw = lines[i].strip()
        inner = raw.strip("|").replace(" ", "")
        if inner and set(inner) <= set("-:|"):
            i += 1
            continue
        cells = [cell.strip() for cell in raw.strip("|").split("|")]
        rows.append(cells)
        i += 1
    return rows, i


def _markdown_table(st, rows: list[list[str]]) -> list:
    if not rows:
        return []
    ncols = max(len(r) for r in rows)
    normalized = [row + [""] * (ncols - len(row)) for row in rows]
    data = [[Paragraph(md_inline(cell), st["body"]) for cell in row] for row in normalized]
    avail = PAGE_W - 2 * MARGIN
    col_w = avail / ncols
    table = Table(data, colWidths=[col_w] * ncols, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E2E8F0")),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E0")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    return [Spacer(1, 0.15 * cm), table, Spacer(1, 0.25 * cm)]


def _story_from_markdown(md_text: str, st, *, embed_screenshots: bool) -> list:
    lines = md_text.splitlines()
    story: list = []
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if stripped.startswith("# ") and i == 0:
            i += 1
            continue

        if stripped == "---":
            story.append(Spacer(1, 0.25 * cm))
            i += 1
            continue

        if stripped.startswith("## "):
            title = stripped[3:].strip()
            story.extend(_screenshot_for_heading(st, title, use_image=embed_screenshots))
            story.append(Paragraph(md_inline(title), st["h1"]))
            i += 1
            continue

        if stripped.startswith("### "):
            story.append(Paragraph(md_inline(stripped[4:].strip()), st["h2"]))
            i += 1
            continue

        if stripped.startswith("|"):
            rows, i = _parse_table(lines, i)
            story.extend(_markdown_table(st, rows))
            continue

        if stripped.startswith("```"):
            i += 1
            code_lines: list[str] = []
            while i < len(lines) and not lines[i].strip().startswith("```"):
                code_lines.append(lines[i])
                i += 1
            if i < len(lines):
                i += 1
            code_html = "<br/>".join(html.escape(part) for part in code_lines)
            story.append(Paragraph(f'<font face="Courier" size="9">{code_html}</font>', st["code"]))
            story.append(Spacer(1, 0.15 * cm))
            continue

        if stripped.startswith("- "):
            story.append(
                Paragraph(md_inline(stripped[2:].strip()), st["bullet"], bulletText="•")
            )
            i += 1
            continue

        numbered = re.match(r"^(\d+)\.\s+(.*)$", stripped)
        if numbered:
            story.append(
                Paragraph(
                    md_inline(numbered.group(2)),
                    st["bullet"],
                    bulletText=f"{numbered.group(1)}.",
                )
            )
            i += 1
            continue

        if not stripped:
            i += 1
            continue

        story.append(Paragraph(md_inline(stripped), st["body"]))
        story.append(Spacer(1, 0.08 * cm))
        i += 1

    return story


def _footer(canvas, doc) -> None:
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.HexColor("#718096"))
    canvas.drawCentredString(
        PAGE_W / 2,
        1 * cm,
        f"{APP_NAME} · v{APP_VERSION} · Página {doc.page}",
    )
    canvas.restoreState()


def _screenshots_available() -> bool:
    return all((SCREENSHOTS / name).is_file() for name, _ in SCREENSHOT_BEFORE_HEADING.values())


def build_manual_pdf(output_path: Path | None = None, *, embed_screenshots: bool = False) -> Path:
    output_path = output_path or OUTPUT
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if not MANUAL_MD.is_file():
        raise FileNotFoundError(f"Manual Markdown não encontrado: {MANUAL_MD}")

    st = _styles()
    md_text = MANUAL_MD.read_text(encoding="utf-8")

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        leftMargin=MARGIN,
        rightMargin=MARGIN,
        topMargin=MARGIN,
        bottomMargin=1.6 * cm,
        title=f"Manual — {APP_NAME}",
        author=APP_NAME,
        subject=f"{APP_TAGLINE} · v{APP_VERSION}",
    )

    story: list = []
    story.extend(_cover(st))
    story.extend(_story_from_markdown(md_text, st, embed_screenshots=embed_screenshots))

    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    return output_path


def main() -> None:
    if "--sem-capturas" in sys.argv:
        embed = False
    elif "--com-capturas" in sys.argv:
        embed = True
    else:
        embed = _screenshots_available()
    out = build_manual_pdf(embed_screenshots=embed)
    print(f"PDF gerado: {out}")
    print(f"Capturas embutidas: {'sim' if embed else 'não'}")


if __name__ == "__main__":
    main()
