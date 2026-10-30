from __future__ import annotations

from datetime import datetime
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    Image,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from src.app_paths import logo_path
from src.attraction_images import attraction_thumbnail_flowable
from src.branding import APP_NAME_PDF, APP_TAGLINE, APP_VERSION, pdf_version_label
from src.itinerary_engine import DayPlan, ItineraryResult, ScheduleBlock, _fmt_time, _lodging_options_summary
from src.maps_links import maps_pdf_html

PAGE_W, PAGE_H = A4
HEADER_H = 2.4 * cm
FOOTER_H = 1.0 * cm

COLOR_PRIMARY = colors.HexColor("#1a365d")
COLOR_ACCENT = colors.HexColor("#2c5282")
COLOR_MUTED = colors.HexColor("#4a5568")
COLOR_VISIT = colors.HexColor("#EBF8FF")
COLOR_LUNCH = colors.HexColor("#FAF5FF")
COLOR_TRANSIT = colors.HexColor("#EDF2F7")
COLOR_END = colors.HexColor("#F0FFF4")
COLOR_BORDER_VISIT = colors.HexColor("#3182CE")
COLOR_BORDER_LUNCH = colors.HexColor("#805AD5")
COLOR_BORDER_TRANSIT = colors.HexColor("#A0AEC0")
COLOR_BORDER_END = colors.HexColor("#38A169")
COLOR_LODGING = colors.HexColor("#FFFAF0")
COLOR_BORDER_LODGING = colors.HexColor("#DD6B20")


def _kind_colors(kind: str) -> tuple[colors.Color, colors.Color]:
    if kind == "lunch":
        return COLOR_LUNCH, COLOR_BORDER_LUNCH
    if kind == "transit":
        return COLOR_TRANSIT, COLOR_BORDER_TRANSIT
    if kind == "end":
        return COLOR_END, COLOR_BORDER_END
    if kind == "lodging":
        return COLOR_LODGING, COLOR_BORDER_LODGING
    return COLOR_VISIT, COLOR_BORDER_VISIT


def _time_label(block: ScheduleBlock) -> str:
    if block.kind == "lodging" and block.start == block.end:
        return "Noite"
    if block.kind == "end" and block.start == block.end:
        return _fmt_time(block.start)
    return f"{_fmt_time(block.start)} – {_fmt_time(block.end)}"


def _draw_page_chrome(canvas, doc, client_name: str) -> None:
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
    canvas.setFont("Helvetica", 9)
    canvas.drawString(
        3.6 * cm, PAGE_H - 1.55 * cm, f"{APP_TAGLINE} · {client_name} · {pdf_version_label()}"
    )

    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(COLOR_MUTED)
    canvas.drawCentredString(
        PAGE_W / 2, 0.65 * cm, f"Página {doc.page} · v{APP_VERSION}"
    )
    canvas.restoreState()


def _cover_block(client_name: str, duracao: int, body: ParagraphStyle, brand: ParagraphStyle) -> Table:
    logo_cell: list = []
    path = logo_path()
    if path.exists():
        logo_cell.append(Image(str(path), width=4.2 * cm, height=4.2 * cm))
    else:
        logo_cell.append(Spacer(1, 4.2 * cm))

    text = Paragraph(
        f"<b>{APP_NAME_PDF}</b><br/>{APP_TAGLINE}<br/>{pdf_version_label()}<br/><br/>"
        f"<b>Cliente:</b> {client_name}<br/>"
        f"<b>Duração:</b> {duracao} dia(s)",
        brand,
    )

    cover = Table([[logo_cell, text]], colWidths=[4.8 * cm, 12 * cm])
    cover.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("BOX", (0, 0), (-1, -1), 1, COLOR_ACCENT),
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F7FAFC")),
                ("LEFTPADDING", (0, 0), (-1, -1), 12),
                ("RIGHTPADDING", (0, 0), (-1, -1), 12),
                ("TOPPADDING", (0, 0), (-1, -1), 12),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 12),
            ]
        )
    )
    return cover


def _profile_summary_table(result: ItineraryResult, body: ParagraphStyle, label: ParagraphStyle) -> Table:
    p = result.profile
    s = result.scores
    ritmo = p.ritmo.replace("(Equilibrado)", "/ Equilibrado")

    data = [
        [
            Paragraph("<b>Cliente principal</b>", label),
            Paragraph("<b>Participantes</b>", label),
            Paragraph("<b>Duração</b>", label),
        ],
        [
            Paragraph(p.nome_grupo, body),
            Paragraph(f"{p.numero_pessoas} pessoas<br/>({p.perfil_predominante})", body),
            Paragraph(f"{p.duracao_dias} dia(s)", body),
        ],
        [
            Paragraph("<b>Ritmo desejado</b>", label),
            Paragraph("<b>Cidades do roteiro</b>", label),
            Paragraph("<b>Score do grupo (seção 4)</b>", label),
        ],
        [
            Paragraph(ritmo, body),
            Paragraph(", ".join(p.cidades) if p.cidades else "—", body),
            Paragraph(
                f"Plog: {s.media_plog:.2f}<br/>Cohen: {s.media_cohen:.2f}<br/>{s.fator_restritivo}",
                body,
            ),
        ],
    ]

    if result.lodging_suggestions:
        summary_html = _lodging_options_summary(result.lodging_suggestions).replace("\n", "<br/>")
        data.append(
            [
                Paragraph("<b>Sugestões de hospedagem (pernoite)</b>", label),
                Paragraph("", body),
                Paragraph("", body),
            ]
        )
        data.append(
            [
                Paragraph(summary_html, body),
                Paragraph("", body),
                Paragraph("", body),
            ]
        )

    table = Table(data, colWidths=[5.6 * cm, 5.6 * cm, 5.6 * cm])
    style_commands = [
        ("BACKGROUND", (0, 0), (-1, 0), COLOR_ACCENT),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, 2), (-1, 2), colors.HexColor("#E2E8F0")),
        ("BOX", (0, 0), (-1, -1), 1, COLOR_ACCENT),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E0")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]
    if result.lodging_suggestions:
        lodging_header_row = len(data) - 2
        lodging_body_row = len(data) - 1
        style_commands.extend(
            [
                ("SPAN", (0, lodging_header_row), (2, lodging_header_row)),
                ("SPAN", (0, lodging_body_row), (2, lodging_body_row)),
                ("BACKGROUND", (0, lodging_header_row), (2, lodging_header_row), colors.HexColor("#E2E8F0")),
            ]
        )
    table.setStyle(TableStyle(style_commands))
    return table


def _timeline_row(block: ScheduleBlock, body: ParagraphStyle, title: ParagraphStyle) -> Table:
    bg, border = _kind_colors(block.kind)
    marker = "●"
    if block.kind == "lunch":
        marker = "◆"
    elif block.kind == "transit":
        marker = "→"
    elif block.kind == "end":
        marker = "■"
    elif block.kind == "lodging":
        marker = "☾"

    time_cell: list = [Paragraph(f"<b>{_time_label(block)}</b>", body)]
    if block.kind in ("visit", "lunch"):
        thumb = attraction_thumbnail_flowable(
            maps_url=block.maps_url,
            place_title=block.title,
        )
        if thumb:
            time_cell.append(Spacer(1, 0.06 * cm))
            time_cell.append(thumb)

    time_cell = time_cell if len(time_cell) > 1 else time_cell[0]
    marker_colors = {
        "visit": "#3182CE",
        "lunch": "#805AD5",
        "transit": "#718096",
        "end": "#38A169",
        "lodging": "#DD6B20",
    }
    mc = marker_colors.get(block.kind, "#3182CE")
    marker_cell = Paragraph(f'<font color="{mc}"><b>{marker}</b></font>', body)
    if block.kind == "lodging":
        content = Paragraph(f"<b>{block.title}</b>", title)
    else:
        detail_html = block.detail.replace("\n", "<br/>")
        if block.maps_url:
            detail_html += maps_pdf_html(block.maps_url)
        content = Paragraph(f"<b>{block.title}</b><br/>{detail_html}", title)

    row = Table(
        [[time_cell, marker_cell, content]],
        colWidths=[3.2 * cm, 0.8 * cm, 12.5 * cm],
    )
    row.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), bg),
                ("BOX", (0, 0), (-1, -1), 0.75, border),
                ("LINEBEFORE", (1, 0), (1, -1), 2, border),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ALIGN", (1, 0), (1, -1), "CENTER"),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    return row


def _day_timeline_section(day: DayPlan, h2: ParagraphStyle, body: ParagraphStyle, title: ParagraphStyle) -> list:
    elements: list = [
        Paragraph(f"DIA {day.day_number} — {day.theme}", h2),
        Paragraph("Linha do tempo da programação", body),
        Spacer(1, 0.15 * cm),
    ]
    for block in day.blocks:
        elements.append(_timeline_row(block, body, title))
        elements.append(Spacer(1, 0.12 * cm))
    return elements


def roteiro_pdf_filename(nome_grupo: str, *, when: datetime | None = None) -> str:
    """Nome do arquivo PDF: Roteiro_<grupo>_HHMM.pdf (hora e minuto da geração)."""
    when = when or datetime.now()
    stamp = when.strftime("%H%M")
    safe_name = (
        "".join(ch for ch in nome_grupo if ch.isalnum() or ch in (" ", "-", "_")).strip()
        or "roteiro"
    )
    return f"Roteiro_{safe_name.replace(' ', '_')}_{stamp}.pdf"


def build_pdf(result: ItineraryResult, output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    client_name = result.profile.nome_grupo or "Cliente"

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        rightMargin=1.8 * cm,
        leftMargin=1.8 * cm,
        topMargin=HEADER_H + 0.5 * cm,
        bottomMargin=FOOTER_H + 0.5 * cm,
        title=f"{APP_NAME_PDF} — {client_name}",
        author=APP_NAME_PDF,
        subject=f"{APP_TAGLINE} · v{APP_VERSION}",
    )

    styles = getSampleStyleSheet()
    brand = ParagraphStyle(
        "Brand",
        parent=styles["Normal"],
        fontSize=12,
        leading=16,
        textColor=COLOR_PRIMARY,
    )
    h2 = ParagraphStyle(
        "H2",
        parent=styles["Heading2"],
        textColor=COLOR_ACCENT,
        spaceBefore=14,
        spaceAfter=6,
        fontSize=13,
    )
    label = ParagraphStyle("Label", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=9)
    body = ParagraphStyle("Body", parent=styles["Normal"], fontSize=9.5, leading=13)
    block_title = ParagraphStyle(
        "BlockTitle",
        parent=body,
        fontName="Helvetica-Bold",
        fontSize=10,
        leading=14,
    )
    section = ParagraphStyle(
        "Section",
        parent=styles["Heading3"],
        textColor=COLOR_PRIMARY,
        spaceBefore=8,
        spaceAfter=8,
    )

    p = result.profile
    story = [
        _cover_block(client_name, p.duracao_dias, body, brand),
        Spacer(1, 0.35 * cm),
        Paragraph("Quadro-resumo do perfil do cliente", section),
        _profile_summary_table(result, body, label),
        Spacer(1, 0.35 * cm),
        Paragraph("Programação em linha do tempo", section),
    ]

    for i, day in enumerate(result.days):
        story.extend(_day_timeline_section(day, h2, body, block_title))
        if i < len(result.days) - 1:
            story.append(Spacer(1, 0.2 * cm))

    story.extend(
        [
            PageBreak(),
            Paragraph("Análise e validação técnica", h2),
            Paragraph(
                f"<b>Índices calculados:</b> Aventura (Plog) {result.media_plog:.2f} · "
                f"Independência (Cohen) {result.media_cohen:.2f}",
                body,
            ),
            Spacer(1, 0.1 * cm),
            Paragraph(
                f"<b>Fator restritivo:</b> {result.scores.fator_restritivo}",
                body,
            ),
            Spacer(1, 0.15 * cm),
            Paragraph(result.justification, body),
        ]
    )

    doc.build(
        story,
        onFirstPage=lambda c, d: _draw_page_chrome(c, d, client_name),
        onLaterPages=lambda c, d: _draw_page_chrome(c, d, client_name),
    )
    return output_path
