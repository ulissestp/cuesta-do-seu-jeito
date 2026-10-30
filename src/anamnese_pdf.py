from __future__ import annotations

from datetime import datetime
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from src.app_paths import logo_path
from src.branding import APP_NAME_PDF, APP_VERSION
from src.pdf_export import COLOR_MUTED, COLOR_PRIMARY, FOOTER_H, HEADER_H, PAGE_H, PAGE_W

# Espaço para marcar opções à mão (nbsp evita colapso no Paragraph do ReportLab).
_CHECK = "(&nbsp;&nbsp;&nbsp;)"


def anamnese_pdf_filename(nome_grupo: str, *, when: datetime | None = None) -> str:
    when = when or datetime.now()
    stamp = when.strftime("%H%M")
    safe = (
        "".join(ch for ch in (nome_grupo or "grupo") if ch.isalnum() or ch in (" ", "-", "_")).strip()
        or "grupo"
    )
    safe = safe.replace(" ", "_")
    return f"Ficha_Anamnese_{safe}_{stamp}.pdf"


def _draw_anamnese_page_chrome(canvas, doc, nome_grupo: str) -> None:
    grupo = (nome_grupo or "").strip() or "—"
    canvas.saveState()
    canvas.setFillColor(COLOR_PRIMARY)
    canvas.rect(0, PAGE_H - HEADER_H, PAGE_W, HEADER_H, fill=1, stroke=0)

    logo = logo_path()
    if logo.exists():
        canvas.drawImage(
            str(logo),
            1.5 * cm,
            PAGE_H - HEADER_H + 0.25 * cm,
            width=1.8 * cm,
            height=1.8 * cm,
            preserveAspectRatio=True,
            mask="auto",
        )

    canvas.setFillColor(colors.white)
    canvas.setFont("Helvetica-Bold", 13)
    canvas.drawString(3.6 * cm, PAGE_H - 1.15 * cm, APP_NAME_PDF)
    canvas.setFont("Helvetica-Bold", 10)
    grupo_line = f"Grupo: {grupo}"
    max_w = PAGE_W - 3.6 * cm - 1.2 * cm
    while canvas.stringWidth(grupo_line, "Helvetica-Bold", 10) > max_w and len(grupo_line) > 12:
        grupo_line = grupo_line[:-4] + "…"
    canvas.drawString(3.6 * cm, PAGE_H - 1.55 * cm, grupo_line)

    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(COLOR_MUTED)
    canvas.drawCentredString(PAGE_W / 2, 0.65 * cm, f"Página {doc.page} · v{APP_VERSION}")
    canvas.restoreState()


def _form_story(
    styles: dict[str, ParagraphStyle],
    *,
    nome_grupo: str,
    year: int,
) -> list:
    title = styles["title"]
    h2 = styles["h2"]
    body = styles["body"]
    small = styles["small"]
    term = styles["term"]

    def p(text: str) -> Paragraph:
        return Paragraph(text.replace("( )", _CHECK), body)

    story: list = []
    story.append(
        Paragraph("Imprimir uma cópia por participante.", small)
    )
    story.append(Spacer(1, 0.15 * cm))

    story.extend(
        [
            Paragraph("FICHA DE ANAMNESE MÉDICA", title),
            Spacer(1, 0.15 * cm),
            Paragraph("1. IDENTIFICAÇÃO DO PARTICIPANTE", h2),
            p("Nome completo: _________________________________________________________________"),
            p("Data de nascimento: ___/___/______ | Idade: ______ anos | Sexo: ( ) M ( ) F"),
            p("Documento (RG ou CPF): ______________________________________________________"),
            p("Endereço: ____________________________________________________________________"),
            p("Telefone celular: ( ) ______________ | E-mail: _________________________________"),
            Paragraph("2. CONTATOS E INFORMAÇÕES DE EMERGÊNCIA", h2),
            p("Contato principal: _________________________ | Parentesco: __________________"),
            p("Telefones: ( ) ______________ / ( ) ______________"),
            p("Contato secundário: ________________________ | Parentesco: __________________"),
            p("Telefones: ( ) ______________ / ( ) ______________"),
            p("Tipo sanguíneo: ( ) A ( ) B ( ) AB ( ) O | Fator Rh: ( ) + ( ) −"),
            p("Possui seguro viagem ou plano de saúde? ( ) Não ( ) Sim"),
            p("Operadora: __________________________ | Nº carteirinha: _____________________"),
            Paragraph("3. ANAMNESE CLÍNICA DETALHADA", h2),
            Paragraph("<b>A. Condições cardiovasculares e respiratórias</b>", body),
            p("Hipertensão ou hipotensão? ( ) Não ( ) Sim"),
            p("Palpitações, dor no peito ou falta de ar ao esforço? ( ) Não ( ) Sim"),
            p("Asma, bronquite ou outra doença respiratória? ( ) Não ( ) Sim"),
            p("Usa inalador? ( ) Não ( ) Sim. Qual? _________________________________________"),
            Paragraph("<b>B. Condições metabólicas e gástricas</b>", body),
            p("Diabetes? ( ) Não ( ) Sim ( ) Tipo I ( ) Tipo II"),
            p("Uso de insulina? ( ) Não ( ) Sim"),
            p("Hipoglicemia ou desmaios frequentes? ( ) Não ( ) Sim"),
            p("Restrição alimentar severa? ( ) Não ( ) Sim: ________________________________"),
            Paragraph("<b>C. Condições osteoarticulares e neurológicas</b>", body),
            p("Lesões/dores coluna, joelhos, tornozelos ou quadris? ( ) Não ( ) Sim"),
            p("Cirurgia ou internação nos últimos 12 meses? ( ) Não ( ) Sim"),
            p("Especifique: ________________________________________________________________"),
            p("Convulsões, epilepsia, labirintite ou vertigem? ( ) Não ( ) Sim"),
            Paragraph("<b>D. Alergias graves</b>", body),
            p("Medicamentos? ( ) Não ( ) Sim: ____________________________________________"),
            p("Alimentos? ( ) Não ( ) Sim: _______________________________________________"),
            p("Picadas de insetos? ( ) Não ( ) Sim: ______________________________________"),
            p("Outras (látex, poeira)? ( ) Não ( ) Sim: __________________________________"),
            p("Medicamento antialérgico de emergência? ( ) Não ( ) Sim"),
            Paragraph("<b>E. Medicamentos e outras condições</b>", body),
            p("Medicamentos de uso contínuo:"),
        ]
    )

    med_box = Table([[""]], colWidths=[17 * cm], rowHeights=[1.2 * cm])
    med_box.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.5, colors.black)]))
    obs_box = Table([[""]], colWidths=[17 * cm], rowHeights=[1.2 * cm])
    obs_box.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.5, colors.black)]))
    story.extend(
        [
            med_box,
            Spacer(1, 0.1 * cm),
            p("Fobia grave? ( ) Não ( ) Sim: __________________________________"),
            p("[Mulheres] Gestante? ( ) Não ( ) Sim. Semanas: ______"),
            p("Outra condição de saúde a informar:"),
            obs_box,
            PageBreak(),
            Paragraph("TERMO DE CONHECIMENTO DE RISCO, RESPONSABILIDADE E DECLARAÇÃO", title),
            Spacer(1, 0.2 * cm),
            Paragraph(
                "<b>Conhecimento dos riscos:</b> Declaro estar ciente de que a atividade de "
                "turismo/aventura contratada envolve deslocamentos em ambientes naturais, exposição a "
                "variações climáticas (chuva, calor ou frio intensos), terrenos irregulares, trilhas e "
                "esforço físico compatível com o roteiro apresentado. Reconheço riscos inerentes "
                "(quedas, escorregões, picadas de animais peçonhentos e escoriações).",
                term,
            ),
            Paragraph(
                "<b>Capacidade física e psicológica:</b> Afirmo que escolhi este roteiro de livre e "
                "espontânea vontade e que me considero física e psicologicamente apto(a). Assumo a "
                "responsabilidade por omitir informações sobre a minha saúde.",
                term,
            ),
            Paragraph(
                "<b>Orientações de segurança:</b> Comprometo-me a seguir instruções dos guias, "
                "condutores e instrutores e utilizar EPIs recomendados.",
                term,
            ),
            Paragraph(
                "<b>Autorização de atendimento médico:</b> Autorizo primeiros socorros e evacuação "
                "para centro médico, se necessário.",
                term,
            ),
            Spacer(1, 0.3 * cm),
            p(f"________________________________________, ______ de __________________ de {year}."),
            Spacer(1, 0.8 * cm),
        ]
    )

    sign = Table(
        [
            ["Assinatura por extenso do participante"],
            ["(Se menor de 18 anos, assinatura do responsável legal abaixo)"],
            ["Nome e assinatura do responsável legal", "CPF do responsável"],
        ],
        colWidths=[8.5 * cm, 8.5 * cm],
        rowHeights=[1.5 * cm, 0.6 * cm, 1.5 * cm],
    )
    sign.setStyle(
        TableStyle(
            [
                ("SPAN", (0, 0), (1, 0)),
                ("SPAN", (0, 1), (1, 1)),
                ("LINEABOVE", (0, 0), (1, 0), 0.5, colors.black),
                ("LINEABOVE", (0, 2), (0, 2), 0.5, colors.black),
                ("LINEABOVE", (1, 2), (1, 2), 0.5, colors.black),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "BOTTOM"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("TEXTCOLOR", (0, 1), (1, 1), colors.grey),
            ]
        )
    )
    story.extend([sign])
    return story


def build_anamnese_pdf(
    output_path: Path,
    *,
    nome_grupo: str = "",
    year: int | None = None,
) -> Path:
    """Gera um PDF imprimível (A4): anamnese + termo em páginas separadas."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    year = year or datetime.now().year
    grupo_header = (nome_grupo or "").strip()

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        leftMargin=1.8 * cm,
        rightMargin=1.8 * cm,
        topMargin=HEADER_H + 0.5 * cm,
        bottomMargin=FOOTER_H + 0.5 * cm,
        title="Ficha de Anamnese Médica e Termo de Compromisso e Risco",
    )
    base = getSampleStyleSheet()
    styles = {
        "title": ParagraphStyle(
            "AnTitle",
            parent=base["Heading1"],
            fontSize=13,
            alignment=TA_CENTER,
            spaceAfter=10,
            textColor=colors.HexColor("#1a365d"),
        ),
        "h2": ParagraphStyle(
            "AnH2",
            parent=base["Heading2"],
            fontSize=10.5,
            spaceBefore=8,
            spaceAfter=4,
            textColor=colors.HexColor("#2c5282"),
            backColor=colors.HexColor("#eef2f7"),
            borderPadding=4,
        ),
        "body": ParagraphStyle("AnBody", parent=base["Normal"], fontSize=9.5, leading=12),
        "small": ParagraphStyle("AnSmall", parent=base["Normal"], fontSize=8.5, textColor=colors.grey),
        "term": ParagraphStyle(
            "AnTerm", parent=base["Normal"], fontSize=9.5, leading=12, alignment=TA_JUSTIFY, spaceAfter=6
        ),
    }

    story = _form_story(
        styles,
        nome_grupo=nome_grupo,
        year=year,
    )

    doc.build(
        story,
        onFirstPage=lambda c, d: _draw_anamnese_page_chrome(c, d, grupo_header),
        onLaterPages=lambda c, d: _draw_anamnese_page_chrome(c, d, grupo_header),
    )
    return output_path
