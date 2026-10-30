"""Gera PDF do manual do usuário com capturas de tela."""
from __future__ import annotations

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

SCREENSHOTS = ROOT / "docs" / "screenshots"
OUTPUT = ROOT / "docs" / "Manual_CUESTA_RT_v1.0_Beta.pdf"
PLACEHOLDER_H = 10.5 * cm

PAGE_W, PAGE_H = A4
MARGIN = 2 * cm


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
    }


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
            "<b>Manual do usuário</b> do aplicativo desktop (Windows). "
            "Os retângulos cinza reservam espaço para colar ou imprimir "
            "<b>capturas de tela</b> de cada aba.",
            st["body"],
        )
    )
    story.append(PageBreak())
    return story


def _section(st, title: str, paragraphs: list[str]) -> list:
    story = [Paragraph(title, st["h1"])]
    for p in paragraphs:
        story.append(Paragraph(p, st["body"]))
        story.append(Spacer(1, 0.15 * cm))
    return story


def _screenshot_placeholder(st, caption: str, hint: str = "") -> list:
    """Área reservada para captura de tela (sem imagem embutida)."""
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
    story: list = []
    if not path.is_file():
        return _screenshot_placeholder(st, caption, hint=filename)
    img = Image(str(path))
    max_w = PAGE_W - 2 * MARGIN
    max_h = PLACEHOLDER_H
    iw, ih = img.imageWidth, img.imageHeight
    scale = min(max_w / iw, max_h / ih, 1.0)
    img.drawWidth = iw * scale
    img.drawHeight = ih * scale
    story.append(Spacer(1, 0.2 * cm))
    story.append(img)
    story.append(Paragraph(f"<b>Figura:</b> {caption}", st["caption"]))
    story.append(Spacer(1, 0.35 * cm))
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


def build_manual_pdf(output_path: Path | None = None, *, embed_screenshots: bool = False) -> Path:
    output_path = output_path or OUTPUT
    output_path.parent.mkdir(parents=True, exist_ok=True)
    st = _styles()

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

    story.extend(
        _section(
            st,
            "1. Introdução",
            [
                f"O <b>{APP_NAME}</b> monta roteiros turísticos personalizados com base no "
                "perfil do grupo (eixos Plog e Cohen), inventário em Excel e regras de "
                "horário. Não utiliza inteligência artificial.",
                "Fluxo recomendado: preencher <b>Perfil</b> → importar <b>Inventário</b> → "
                "gerar <b>Roteiro PDF</b> e, se necessário, <b>Ficha anamnese</b>. "
                "Arquivos ficam na pasta <b>output</b> (ao lado do executável ou na raiz do projeto).",
            ],
        )
    )

    story.extend(
        _section(
            st,
            "2. Aba Perfil — dados do grupo",
            [
                "Informe nome do grupo, número de pessoas, duração em dias, perfil predominante "
                "e ritmo (Relaxado, Moderado ou Intenso). Marque ao menos uma cidade "
                "(Botucatu, Pardinho, Bofete).",
                "Responda às seis perguntas Plog (1–5) e Cohen (1–5). Clique em "
                "<b>Gravar na planilha</b>. A seção 4 exibe Média Plog, Média Cohen e "
                "fator restritivo.",
            ],
        )
    )
    def shot(name: str, cap: str) -> list:
        return _screenshot(st, name, cap, use_image=embed_screenshots)

    story.extend(shot("01_perfil.png", "Aba 1 — Perfil do turista"))

    story.extend(
        _section(
            st,
            "3. Aba Inventário",
            [
                "Use <b>Importar inventário (.xlsx)</b> com o modelo "
                "(<i>Modelo_Inventario_Turistico.xlsx</i>). Colunas essenciais: horários, "
                "tempo médio, pesos Plog/Cohen, cidade e link <b>Maps</b>.",
                "Links curtos do Google Maps permitem calcular distância e tempo entre "
                "atrativos (requer internet na primeira geração).",
            ],
        )
    )
    story.extend(shot("02_inventario.png", "Aba 2 — Inventário turístico"))

    story.extend(
        _section(
            st,
            "4. Aba Roteiro PDF",
            [
                "<b>Gerar roteiro e PDF</b> valida o perfil e produz o arquivo na pasta output. "
                "Primeira visita do dia inicia até <b>9h30</b>; almoço 12h–13h30; fim até "
                "<b>18h</b>. Deslocamentos usam distância entre coordenadas dos links Maps.",
                "O PDF inclui capa, quadro-resumo, linha do tempo por dia, miniaturas de mapa "
                "e análise técnica.",
            ],
        )
    )
    story.extend(shot("03_roteiro_pdf.png", "Aba 3 — Geração do roteiro em PDF"))

    story.extend(
        _section(
            st,
            "5. Aba Ficha anamnese",
            [
                "Gera PDF para arquivo do grupo: dados clínicos com checkboxes para marcar à "
                "mão e <b>termo em página separada</b>. Cabeçalho com logo e nome do grupo.",
            ],
        )
    )
    story.extend(shot("04_ficha_anamnese.png", "Aba 4 — Ficha de anamnese"))

    story.append(PageBreak())
    story.extend(
        _section(
            st,
            "6. Instalação e compilação",
            [
                "<b>Executável:</b> <font face='Courier'>dist\\RoteiroTurismo.exe</font> "
                "(ícone do logo, versão 1.0 Beta nas propriedades do arquivo).",
                "<b>Compilar:</b> feche o app e execute "
                "<font face='Courier'>powershell -File scripts\\build_desktop.ps1</font>.",
                "<b>Chave Google Maps (opcional):</b> arquivo "
                "<font face='Courier'>google_maps_api_key.txt</font> para miniaturas Static Maps; "
                "sem chave, usa OpenStreetMap.",
            ],
        )
    )
    story.extend(
        _section(
            st,
            "7. Solução de problemas",
            [
                "• <b>Campos obrigatórios:</b> complete a aba Perfil e grave na planilha.<br/>"
                "• <b>Inventário vazio:</b> importe na aba 2.<br/>"
                "• <b>Sem atrativos na cidade:</b> confira a coluna Cidade no Excel.<br/>"
                "• <b>Deslocamentos iguais:</b> inclua links Maps válidos e use internet.<br/>"
                "• <b>Erro ao compilar .exe:</b> feche o programa antes do build.",
            ],
        )
    )

    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    return output_path


def main() -> None:
    embed = "--com-capturas" in sys.argv
    out = build_manual_pdf(embed_screenshots=embed)
    print(f"PDF gerado: {out}")


if __name__ == "__main__":
    main()
