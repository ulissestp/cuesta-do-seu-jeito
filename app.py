"""Gerador de Roteiros Turísticos — Projeto Integrador (sem IA)."""

from __future__ import annotations

import tempfile
from datetime import date, datetime
from pathlib import Path

import streamlit as st

from src.branding import desktop_window_title
from src.form_constants import (
    FORM_BLANK,
    PERFIL_FORM_OPCOES,
    RITMO_FORM_OPCOES,
    SCORE_FORM_OPCOES,
    is_form_blank,
    normalize_form_choice,
    parse_score_choice,
)
from src.excel_io import (
    COHEN_QUESTIONS,
    CIDADES_ROTEIRO,
    PLOG_QUESTIONS,
    TouristProfile,
    compute_group_scores,
    create_inventory_template,
    default_workbook_path,
    initialize_session_workbook,
    clear_profile_only,
    import_inventory_sheet,
    inventory_template_path,
    read_inventory,
    read_profile,
    save_profile,
    validate_profile_for_roteiro,
)
from src.itinerary_engine import itinerary_to_markdown
from src.pdf_export import build_pdf, roteiro_pdf_filename
from src.anamnese_pdf import anamnese_pdf_filename, build_anamnese_pdf
from src.roteiro_service import gerar_roteiro_da_planilha
from src.session_store import load_last_workbook, remember_workbook
from src.help_docs import load_help_markdown, user_manual_md_path, user_manual_pdf_path
from src.inventory_sync import check_and_apply_inventory_update, load_local_inventory_version

PROJECT_ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = PROJECT_ROOT / "output"
TEMPLATE = default_workbook_path(PROJECT_ROOT)
create_inventory_template(inventory_template_path(PROJECT_ROOT))


def _init_state() -> None:
    if "workbook_path" not in st.session_state:
        last = load_last_workbook()
        if last is not None:
            st.session_state.workbook_path = last
        else:
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            path = OUTPUT_DIR / f"Projeto_Integrador_{stamp}.xlsx"
            initialize_session_workbook(path, TEMPLATE)
            st.session_state.workbook_path = path
        remember_workbook(st.session_state.workbook_path)
    if "itinerary_result" not in st.session_state:
        st.session_state.itinerary_result = None
    if "last_anamnese_pdf" not in st.session_state:
        st.session_state.last_anamnese_pdf = None
    if "inventory_sync_checked" not in st.session_state:
        st.session_state.inventory_sync_checked = False


def _run_inventory_sync_once() -> None:
    if st.session_state.inventory_sync_checked:
        return
    st.session_state.inventory_sync_checked = True
    with st.spinner("Verificando atualização do inventário no GitHub…"):
        result = check_and_apply_inventory_update(st.session_state.workbook_path)
    if result.status == "updated":
        st.success(result.message)
    elif result.status == "error":
        st.warning(result.message)


def _show_scores_panel(profile: TouristProfile) -> None:
    st.markdown("#### 4. Score do grupo e fatores restritivos (cálculo automático)")
    missing = validate_profile_for_roteiro(profile)
    if missing:
        st.info("Preencha todo o perfil (incluindo Plog e Cohen) para exibir os scores.")
        return
    scores = compute_group_scores(profile)
    c1, c2, c3 = st.columns(3)
    c1.metric("Média Plog (Aventura)", f"{scores['media_plog']:.2f}")
    c2.metric("Média Cohen (Independência)", f"{scores['media_cohen']:.2f}")
    c3.metric("Fator restritivo (tamanho)", scores["fator_restritivo"])


def step_profile() -> None:
    st.subheader("1. Perfil do Turista (Grupos)")
    st.caption("Dados gerais, eixos Plog e Cohen (1–5). O roteiro é montado a partir do score calculado.")

    current = read_profile(st.session_state.workbook_path)

    with st.form("perfil_form"):
        st.markdown("#### 1. Dados gerais do grupo")
        c1, c2 = st.columns(2)
        nome = c1.text_input("Nome do Grupo / Contato Principal", value=current.nome_grupo)
        pessoas = c2.number_input(
            "Número de Pessoas",
            min_value=0,
            max_value=500,
            value=max(0, int(current.numero_pessoas)),
        )
        check_in = c1.date_input(
            "Check-in",
            value=current.check_in or date.today(),
            format="DD/MM/YYYY",
        )
        check_out = c2.date_input(
            "Check-out",
            value=current.check_out or current.check_in or date.today(),
            format="DD/MM/YYYY",
        )
        st.caption(
            "O roteiro inclui os dois dias e usa só atrativos em funcionamento "
            "(segunda a domingo ou feriado)."
        )
        perfil = c2.selectbox(
            "Perfil predominante",
            PERFIL_FORM_OPCOES,
            index=_index_or_zero(PERFIL_FORM_OPCOES, current.perfil_predominante),
        )
        ritmo = c1.selectbox(
            "Ritmo de viagem", RITMO_FORM_OPCOES, index=_index_or_zero(RITMO_FORM_OPCOES, current.ritmo)
        )
        cidades = c2.multiselect(
            "Cidades do roteiro (escolha uma ou mais)",
            list(CIDADES_ROTEIRO),
            default=[c for c in current.cidades if c in CIDADES_ROTEIRO],
            format_func=lambda c: f"{list(CIDADES_ROTEIRO).index(c) + 1} - {c}",
        )

        st.markdown("#### 2. Eixo psicográfico (Plog) — avalie de 1 a 5")
        st.caption("Cada pergunta começa em branco (—). Nota de 1 a 5 é obrigatória para gerar o roteiro.")
        plog_scores: dict[str, int] = {}
        for q in PLOG_QUESTIONS:
            raw = st.selectbox(
                f"{q.id}: {q.pergunta}",
                SCORE_FORM_OPCOES,
                index=_score_form_index(current.plog_scores.get(q.id, 0)),
                help=f"1 = {q.significado_nota_1}\n5 = {q.significado_nota_5}",
            )
            plog_scores[q.id] = parse_score_choice(raw)
            st.markdown(
                f"<span style='color:#2B6CB0'>**1:** {q.significado_nota_1}</span> · "
                f"<span style='color:#2F855A'>**5:** {q.significado_nota_5}</span>",
                unsafe_allow_html=True,
            )

        st.markdown("#### 3. Eixo de institucionalização (Cohen) — avalie de 1 a 5")
        st.caption("Mesma regra: início em branco; nota 1 a 5 obrigatória para o PDF.")
        cohen_scores: dict[str, int] = {}
        for q in COHEN_QUESTIONS:
            raw = st.selectbox(
                f"{q.id}: {q.pergunta}",
                SCORE_FORM_OPCOES,
                index=_score_form_index(current.cohen_scores.get(q.id, 0)),
                help=f"1 = {q.significado_nota_1}\n5 = {q.significado_nota_5}",
            )
            cohen_scores[q.id] = parse_score_choice(raw)
            st.markdown(
                f"<span style='color:#2B6CB0'>**1:** {q.significado_nota_1}</span> · "
                f"<span style='color:#2F855A'>**5:** {q.significado_nota_5}</span>",
                unsafe_allow_html=True,
            )

        submitted = st.form_submit_button("Gravar na planilha", type="primary")

    if submitted:
        profile = TouristProfile(
            nome_grupo=nome.strip(),
            numero_pessoas=int(pessoas),
            duracao_dias=0,
            check_in=check_in,
            check_out=check_out,
            perfil_predominante=normalize_form_choice(perfil),
            ritmo=normalize_form_choice(ritmo),
            cidades=list(cidades),
            plog_scores=plog_scores,
            cohen_scores=cohen_scores,
        )
        save_profile(st.session_state.workbook_path, profile)
        remember_workbook(st.session_state.workbook_path)
        st.success(f"Dados gravados em `{st.session_state.workbook_path.name}` (aba Perfil do Turista).")

    refreshed = read_profile(st.session_state.workbook_path)
    _show_scores_panel(refreshed)


def step_inventory() -> None:
    st.subheader("2. Inventário turístico")
    ver = load_local_inventory_version()
    st.caption(
        f"Catálogo local: versão **{ver or 'ainda não sincronizado'}**. "
        "Duplicatas (mesmo ID ou mesmo nome+cidade) são removidas ao atualizar."
    )
    if st.button("Verificar atualização do inventário"):
        with st.spinner("Baixando catálogo e mesclando…"):
            result = check_and_apply_inventory_update(st.session_state.workbook_path)
        if result.status == "updated":
            st.success(result.message)
        elif result.status == "up_to_date":
            st.info(result.message)
        else:
            st.warning(result.message)
        st.rerun()
    st.caption(
        "Modelo com **11 colunas** (… descrição e **Cidade** na coluna K). "
        "O roteiro usa só atrações das cidades marcadas no perfil."
    )
    model_path = inventory_template_path(PROJECT_ROOT)
    create_inventory_template(model_path)
    st.download_button(
        "Baixar modelo de inventário (.xlsx)",
        model_path.read_bytes(),
        file_name="Modelo_Inventario_Turistico.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    uploaded = st.file_uploader("Planilha de inventário", type=["xlsx"])
    if uploaded and st.button("Importar inventário"):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded.getvalue())
            tmp_path = Path(tmp.name)
        count = import_inventory_sheet(st.session_state.workbook_path, tmp_path)
        remember_workbook(st.session_state.workbook_path)
        st.success(f"Inventário importado: **{count}** atrações.")

    items = read_inventory(st.session_state.workbook_path)
    if items:
        st.dataframe(
            [
                {
                    "ID": a.id,
                    "Nome da atração": a.nome,
                    "Categoria": a.categoria,
                    "Descrição breve": a.descricao,
                }
                for a in items
            ],
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.warning("Nenhuma atração encontrada. Use o modelo ou importe uma planilha.")


def step_roteiro_pdf() -> None:
    st.subheader("3. Roteiro turístico (PDF)")
    st.caption(
        "O roteiro é gerado automaticamente cruzando **Média Plog**, **Média Cohen**, "
        "**fator restritivo** e o inventário — sem uso de IA."
    )

    profile = read_profile(st.session_state.workbook_path)
    _show_scores_panel(profile)

    if st.button("Calcular roteiro e gerar PDF", type="primary"):
        try:
            result = gerar_roteiro_da_planilha(st.session_state.workbook_path)
            st.session_state.itinerary_result = result

            pdf_path = OUTPUT_DIR / roteiro_pdf_filename(result.profile.nome_grupo)
            build_pdf(result, pdf_path)
            st.session_state.last_pdf = pdf_path
            st.success(f"PDF gerado: `{pdf_path}`")
        except ValueError as exc:
            st.error(str(exc))

    if st.session_state.itinerary_result:
        st.markdown("#### Pré-visualização do roteiro")
        st.markdown(itinerary_to_markdown(st.session_state.itinerary_result))

    if st.session_state.get("last_pdf") and Path(st.session_state.last_pdf).exists():
        st.download_button(
            "Baixar PDF",
            Path(st.session_state.last_pdf).read_bytes(),
            file_name=Path(st.session_state.last_pdf).name,
            mime="application/pdf",
        )

    st.download_button(
        "Baixar planilha completa (.xlsx)",
        st.session_state.workbook_path.read_bytes(),
        file_name=st.session_state.workbook_path.name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def step_anamnese_pdf() -> None:
    st.subheader("4. Ficha de anamnese e termo de risco (PDF)")
    st.caption(
        "Modelo imprimível em A4 — **uma ficha por participante** (imprima quantas cópias precisar)."
    )
    profile = read_profile(st.session_state.workbook_path)

    if st.button("Gerar PDF da ficha de anamnese", type="primary"):
        pdf_path = OUTPUT_DIR / anamnese_pdf_filename(profile.nome_grupo or "grupo")
        build_anamnese_pdf(
            pdf_path,
            nome_grupo=profile.nome_grupo,
        )
        st.session_state.last_anamnese_pdf = pdf_path
        st.success(f"PDF gerado: `{pdf_path}`")

    if st.session_state.get("last_anamnese_pdf") and Path(st.session_state.last_anamnese_pdf).exists():
        st.download_button(
            "Baixar ficha (PDF)",
            Path(st.session_state.last_anamnese_pdf).read_bytes(),
            file_name=Path(st.session_state.last_anamnese_pdf).name,
            mime="application/pdf",
        )


def step_help() -> None:
    st.subheader("Ajuda")
    st.caption("Guia rápido. Manual completo disponível para download abaixo.")
    st.markdown(load_help_markdown())

    pdf_path = user_manual_pdf_path()
    md_path = user_manual_md_path()
    c1, c2 = st.columns(2)
    if pdf_path.is_file():
        c1.download_button(
            "Baixar manual do usuário (PDF)",
            pdf_path.read_bytes(),
            file_name=pdf_path.name,
            mime="application/pdf",
        )
    else:
        c1.warning("Manual PDF não encontrado em `docs/`.")
    if md_path.is_file():
        c2.download_button(
            "Baixar manual do usuário (Markdown)",
            md_path.read_text(encoding="utf-8").encode("utf-8"),
            file_name=md_path.name,
            mime="text/markdown",
        )
    else:
        c2.warning("Manual Markdown não encontrado em `docs/`.")


def _index_or_zero(options: list[str], value: str) -> int:
    if is_form_blank(value):
        return 0
    try:
        return options.index(value)
    except ValueError:
        for i, opt in enumerate(options):
            if value.lower() in opt.lower():
                return i
        return 0


def _score_form_index(value: int) -> int:
    if value in range(1, 6):
        return _index_or_zero(SCORE_FORM_OPCOES, str(value))
    return 0


def main() -> None:
    st.set_page_config(page_title=desktop_window_title(), page_icon="🧭", layout="wide")
    _init_state()
    _run_inventory_sync_once()

    st.title(f"🧭 {desktop_window_title()}")
    st.markdown("Perfil → inventário → **score automático (seção 4)** → PDF do roteiro → **ficha anamnese**.")

    st.sidebar.header("Arquivo de trabalho")
    st.sidebar.code(str(st.session_state.workbook_path), language=None)
    if st.sidebar.button("Limpar informações"):
        clear_profile_only(st.session_state.workbook_path)
        remember_workbook(st.session_state.workbook_path)
        st.session_state.itinerary_result = None
        st.session_state.pop("last_pdf", None)
        st.rerun()

    if st.sidebar.button("Nova planilha"):
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = OUTPUT_DIR / f"Projeto_Integrador_{stamp}.xlsx"
        initialize_session_workbook(path, TEMPLATE)
        st.session_state.workbook_path = path
        remember_workbook(path)
        st.session_state.itinerary_result = None
        st.session_state.pop("last_pdf", None)
        st.rerun()

    st.sidebar.markdown("---")
    st.sidebar.header("Ajuda")
    st.sidebar.caption("Abra a aba **Ajuda** para o guia rápido e downloads do manual.")

    tab1, tab2, tab3, tab4, tab5 = st.tabs(
        ["1. Perfil", "2. Inventário", "3. Roteiro PDF", "4. Ficha anamnese", "Ajuda"]
    )
    with tab1:
        step_profile()
    with tab2:
        step_inventory()
    with tab3:
        step_roteiro_pdf()
    with tab4:
        step_anamnese_pdf()
    with tab5:
        step_help()


if __name__ == "__main__":
    main()
