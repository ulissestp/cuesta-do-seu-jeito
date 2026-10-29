"""Gerador de Roteiros Turísticos — Projeto Integrador (sem IA)."""

from __future__ import annotations

import tempfile
from datetime import datetime
from pathlib import Path

import streamlit as st

from src.branding import APP_NAME
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
from src.roteiro_service import gerar_roteiro_da_planilha
from src.session_store import load_last_workbook, remember_workbook

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
        duracao = c1.number_input(
            "Duração da Viagem (dias)",
            min_value=0,
            max_value=30,
            value=max(0, int(current.duracao_dias)),
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
            duracao_dias=int(duracao),
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
                    "Nome": a.nome,
                    "Categoria": a.categoria,
                    "Abertura": a.horario_abertura.strftime("%H:%M") if a.horario_abertura else "—",
                    "Fechamento": a.horario_fechamento.strftime("%H:%M") if a.horario_fechamento else "—",
                    "Capacidade": a.capacidade,
                    "Tempo (h)": a.tempo_horas,
                    "Plog": a.peso_plog,
                    "Cohen": a.peso_cohen,
                    "Cidade": a.cidade,
                    "Descrição": a.descricao,
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
    st.set_page_config(page_title=APP_NAME, page_icon="🧭", layout="wide")
    _init_state()

    st.title(f"🧭 {APP_NAME}")
    st.markdown("Perfil → inventário → **score automático (seção 4)** → PDF do roteiro.")

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

    tab1, tab2, tab3 = st.tabs(["1. Perfil", "2. Inventário", "3. Roteiro PDF"])
    with tab1:
        step_profile()
    with tab2:
        step_inventory()
    with tab3:
        step_roteiro_pdf()


if __name__ == "__main__":
    main()
