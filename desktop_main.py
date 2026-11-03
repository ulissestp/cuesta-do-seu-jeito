"""
Gerador de Roteiros Turísticos — aplicativo desktop (Windows).
Execute: python desktop_main.py
"""

from __future__ import annotations

import os
import shutil
import sys
import threading
import tkinter as tk
from datetime import date, datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from src.app_paths import output_dir, project_root
from src.app_icon import ensure_app_icon
from src.branding import desktop_header_text, desktop_window_title
from src.excel_io import (
    COHEN_QUESTIONS,
    CIDADES_ROTEIRO,
    PLOG_QUESTIONS,
    TouristProfile,
    compute_group_scores,
    create_inventory_template,
    default_workbook_path,
    inventory_template_path,
    initialize_session_workbook,
    clear_profile_only,
    filter_inventory_by_cities,
    import_inventory_sheet,
    inventory_maps_coverage,
    read_inventory,
    read_profile,
    save_profile,
    validate_profile_for_roteiro,
    format_profile_validation_errors,
)
from src.form_constants import (
    FORM_BLANK,
    PERFIL_FORM_OPCOES,
    RITMO_FORM_OPCOES,
    SCORE_FORM_OPCOES,
    is_form_blank,
    normalize_form_choice,
    parse_score_choice,
)
from src.session_store import load_last_workbook, remember_workbook
from src.itinerary_engine import itinerary_to_markdown
from src.pdf_export import build_pdf, roteiro_pdf_filename
from src.anamnese_pdf import anamnese_pdf_filename, build_anamnese_pdf
from src.roteiro_service import gerar_roteiro_da_planilha

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

ROOT = project_root()
OUT = output_dir()
TEMPLATE = default_workbook_path(ROOT)
create_inventory_template(inventory_template_path(ROOT))


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


def _parse_int_field(text: str) -> int:
    text = text.strip()
    if not text:
        return 0
    return int(text)


def _parse_score_field(text: str) -> int:
    return parse_score_choice(text)


def _score_display(value: int) -> str:
    return str(value) if value in range(1, 6) else FORM_BLANK


_MONTHS_PT = (
    "janeiro",
    "fevereiro",
    "março",
    "abril",
    "maio",
    "junho",
    "julho",
    "agosto",
    "setembro",
    "outubro",
    "novembro",
    "dezembro",
)
_WEEK_LETTERS = ("Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom")


class TripCalendar(ctk.CTkFrame):
    """Calendário de check-in e check-out na aba de perfil."""

    def __init__(self, master) -> None:
        super().__init__(master)
        self.check_in: date | None = None
        self.check_out: date | None = None
        today = date.today()
        self._visible = date(today.year, today.month, 1)

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=8, pady=(8, 4))
        ctk.CTkButton(header, text="◀", width=36, command=self._prev_month).pack(side="left")
        self._title = ctk.CTkLabel(header, text="", font=ctk.CTkFont(weight="bold"))
        self._title.pack(side="left", expand=True)
        ctk.CTkButton(header, text="▶", width=36, command=self._next_month).pack(side="right")

        self._grid = ctk.CTkFrame(self, fg_color="transparent")
        self._grid.pack(padx=8, pady=(0, 4))
        for col, letter in enumerate(_WEEK_LETTERS):
            ctk.CTkLabel(self._grid, text=letter, width=44, font=ctk.CTkFont(size=11)).grid(
                row=0, column=col, padx=1, pady=2
            )

        self._summary = ctk.CTkLabel(self, text="", anchor="w", justify="left")
        self._summary.pack(fill="x", padx=8, pady=(0, 8))
        self._redraw()

    def set_dates(self, check_in: date | None, check_out: date | None) -> None:
        self.check_in = check_in
        self.check_out = check_out
        anchor = check_in or date.today()
        self._visible = date(anchor.year, anchor.month, 1)
        self._redraw()

    def clear(self) -> None:
        self.set_dates(None, None)

    def _prev_month(self) -> None:
        self._shift_month(-1)

    def _next_month(self) -> None:
        self._shift_month(1)

    def _shift_month(self, delta: int) -> None:
        month = self._visible.month + delta
        year = self._visible.year
        while month < 1:
            month += 12
            year -= 1
        while month > 12:
            month -= 12
            year += 1
        self._visible = date(year, month, 1)
        self._redraw()

    def _pick(self, chosen: date) -> None:
        if self.check_in is None or (self.check_in and self.check_out):
            self.check_in = chosen
            self.check_out = None
        elif chosen < self.check_in:
            self.check_out = self.check_in
            self.check_in = chosen
        else:
            self.check_out = chosen
        self._redraw()

    def _redraw(self) -> None:
        for child in self._grid.winfo_children():
            info = child.grid_info()
            if int(info.get("row", 0)) > 0:
                child.destroy()
        self._title.configure(text=f"{_MONTHS_PT[self._visible.month - 1]} {self._visible.year}")
        first = self._visible
        start = first - timedelta(days=first.weekday())
        for cell in range(42):
            current = start + timedelta(days=cell)
            row, col = divmod(cell, 7)
            in_month = current.month == self._visible.month
            selected = current in (self.check_in, self.check_out)
            in_range = (
                self.check_in is not None
                and self.check_out is not None
                and self.check_in < current < self.check_out
            )
            if selected:
                color = "#2B6CB0"
                text_color = "white"
            elif in_range:
                color = "#BEE3F8"
                text_color = "#1A365D"
            elif in_month:
                color = "#E2E8F0"
                text_color = "#1A202C"
            else:
                color = "#EDF2F7"
                text_color = "#A0AEC0"
            ctk.CTkButton(
                self._grid,
                text=str(current.day),
                width=44,
                height=28,
                fg_color=color,
                text_color=text_color,
                hover_color="#3182CE",
                command=lambda chosen=current: self._pick(chosen),
            ).grid(row=row + 1, column=col, padx=1, pady=1)
        if self.check_in and self.check_out:
            days = (self.check_out - self.check_in).days + 1
            self._summary.configure(
                text=(
                    f"Check-in: {self.check_in.strftime('%d/%m/%Y')}    "
                    f"Check-out: {self.check_out.strftime('%d/%m/%Y')}    "
                    f"{days} dia(s)"
                )
            )
        elif self.check_in:
            self._summary.configure(
                text=f"Check-in: {self.check_in.strftime('%d/%m/%Y')}    Clique a data de check-out."
            )
        else:
            self._summary.configure(text="Clique a data de check-in e depois a de check-out.")


class RoteiroDesktopApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title(desktop_window_title())
        self.geometry("1120x780")
        self.minsize(960, 640)
        icon = ensure_app_icon()
        if icon is not None:
            try:
                self.iconbitmap(default=str(icon))
            except tk.TclError:
                pass

        last = load_last_workbook()
        if last is not None:
            self.workbook_path = last
        else:
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            self.workbook_path = initialize_session_workbook(
                OUT / f"Projeto_Integrador_{stamp}.xlsx", TEMPLATE
            )
        remember_workbook(self.workbook_path)
        self.last_pdf: Path | None = None
        self.last_anamnese_pdf: Path | None = None
        self._pdf_job_running = False
        self.itinerary_result = None
        self._used_order: list[str] = []
        self._novo_ids: dict[str, str] = {}
        self._edit_entry: tk.Widget | None = None
        self._edit_row: str | None = None

        self.plog_combos: dict[str, ctk.CTkComboBox] = {}
        self.cohen_combos: dict[str, ctk.CTkComboBox] = {}
        self.cidade_vars: dict[str, tk.BooleanVar] = {}

        self._build_header()
        self._build_tabs()
        self._load_profile_into_form()

    def _build_header(self) -> None:
        header = ctk.CTkFrame(self, corner_radius=0)
        header.pack(fill="x", padx=0, pady=0)

        ctk.CTkLabel(
            header,
            text=desktop_header_text(),
            font=ctk.CTkFont(size=18, weight="bold"),
        ).pack(side="left", padx=16, pady=12)

        self.path_label = ctk.CTkLabel(
            header,
            text=str(self.workbook_path.name),
            font=ctk.CTkFont(size=11),
            text_color="gray",
        )
        self.path_label.pack(side="right", padx=8)

        ctk.CTkButton(header, text="Limpar informações", width=140, command=self._clear_all).pack(
            side="right", padx=8, pady=8
        )
        ctk.CTkButton(header, text="Nova planilha", width=120, command=self._new_workbook).pack(
            side="right", padx=8, pady=8
        )

    def _build_tabs(self) -> None:
        self.tabs = ctk.CTkTabview(self)
        self.tabs.pack(fill="both", expand=True, padx=12, pady=12)
        self.tabs.add("1. Perfil")
        self.tabs.add("2. Inventário")
        self.tabs.add("3. Roteiro PDF")
        self.tabs.add("4. Ficha anamnese")

        self._build_profile_tab(self.tabs.tab("1. Perfil"))
        self._build_inventory_tab(self.tabs.tab("2. Inventário"))
        self._build_pdf_tab(self.tabs.tab("3. Roteiro PDF"))
        self._build_anamnese_tab(self.tabs.tab("4. Ficha anamnese"))

    def _build_profile_tab(self, parent: ctk.CTkFrame) -> None:
        scroll = ctk.CTkScrollableFrame(parent)
        scroll.pack(fill="both", expand=True)

        ctk.CTkLabel(scroll, text="1. Dados gerais do grupo", font=ctk.CTkFont(weight="bold")).pack(
            anchor="w", pady=(4, 8)
        )

        grid = ctk.CTkFrame(scroll, fg_color="transparent")
        grid.pack(fill="x")
        grid.columnconfigure(1, weight=1)
        grid.columnconfigure(3, weight=1)

        self.nome_var = tk.StringVar()
        self.pessoas_var = tk.StringVar(value="")

        self._labeled_entry(grid, 0, "Nome do grupo / contato", self.nome_var)
        self._labeled_entry(grid, 0, "Número de pessoas", self.pessoas_var, col=2)
        self.perfil_combo = self._labeled_combo(grid, 1, "Perfil predominante", PERFIL_FORM_OPCOES)
        self.ritmo_combo = self._labeled_combo(grid, 1, "Ritmo de viagem", RITMO_FORM_OPCOES, col=2)

        ctk.CTkLabel(scroll, text="Período da viagem", font=ctk.CTkFont(weight="bold")).pack(
            anchor="w", pady=(12, 4)
        )
        ctk.CTkLabel(
            scroll,
            text=(
                "Escolha o check-in e o check-out. O roteiro usa só atrativos em funcionamento "
                "naquele dia (segunda a domingo ou feriado)."
            ),
            font=ctk.CTkFont(size=12),
            text_color="gray",
            wraplength=900,
        ).pack(anchor="w", pady=(0, 6))
        self.trip_calendar = TripCalendar(scroll)
        self.trip_calendar.pack(anchor="w", pady=(0, 8))

        ctk.CTkLabel(scroll, text="Cidades do roteiro", font=ctk.CTkFont(weight="bold")).pack(
            anchor="w", pady=(12, 4)
        )
        ctk.CTkLabel(
            scroll,
            text="Selecione uma ou mais cidades. O roteiro usará apenas atrações dessas cidades (coluna Cidade do inventário).",
            font=ctk.CTkFont(size=12),
            text_color="gray",
            wraplength=900,
        ).pack(anchor="w", pady=(0, 6))
        cities_row = ctk.CTkFrame(scroll, fg_color="transparent")
        cities_row.pack(fill="x", pady=(0, 8))
        for idx, city in enumerate(CIDADES_ROTEIRO, start=1):
            var = tk.BooleanVar(value=False)
            self.cidade_vars[city] = var
            ctk.CTkCheckBox(
                cities_row,
                text=f"{idx} - {city}",
                variable=var,
            ).pack(side="left", padx=(0, 20))

        ctk.CTkLabel(scroll, text="2. Eixo Plog (1 a 5)", font=ctk.CTkFont(weight="bold")).pack(
            anchor="w", pady=(16, 6)
        )
        ctk.CTkLabel(
            scroll,
            text="Cada pergunta inicia em branco (—). Todas são obrigatórias para gerar o roteiro.",
            font=ctk.CTkFont(size=12),
            text_color="gray",
        ).pack(anchor="w", pady=(0, 4))
        for q in PLOG_QUESTIONS:
            self._score_row(scroll, q, self.plog_combos)

        ctk.CTkLabel(scroll, text="3. Eixo Cohen (1 a 5)", font=ctk.CTkFont(weight="bold")).pack(
            anchor="w", pady=(16, 6)
        )
        ctk.CTkLabel(
            scroll,
            text="Mesma regra: início em branco; nota 1 a 5 obrigatória para o PDF.",
            font=ctk.CTkFont(size=12),
            text_color="gray",
        ).pack(anchor="w", pady=(0, 4))
        for q in COHEN_QUESTIONS:
            self._score_row(scroll, q, self.cohen_combos)

        ctk.CTkButton(scroll, text="Gravar na planilha", command=self._save_profile).pack(
            anchor="w", pady=16
        )

        self.score_frame = ctk.CTkFrame(scroll)
        self.score_frame.pack(fill="x", pady=8)
        self._refresh_scores_panel()

    def _labeled_entry(self, parent, row, label, var, col=0):
        ctk.CTkLabel(parent, text=label).grid(row=row, column=col, sticky="w", padx=(0, 8), pady=6)
        entry = ctk.CTkEntry(parent, textvariable=var)
        entry.grid(row=row, column=col + 1, sticky="ew", pady=6)
        return entry

    def _labeled_combo(self, parent, row, label, values, col=0):
        ctk.CTkLabel(parent, text=label).grid(row=row, column=col, sticky="w", padx=(0, 8), pady=6)
        combo = ctk.CTkComboBox(parent, values=values, width=280, state="readonly")
        combo.set(FORM_BLANK)
        combo.grid(row=row, column=col + 1, sticky="ew", pady=6)
        return combo

    def _score_row(self, parent, question, store: dict[str, ctk.CTkComboBox]) -> None:
        block = ctk.CTkFrame(parent, corner_radius=8)
        block.pack(fill="x", pady=6, padx=2)

        qid = question.id
        top = ctk.CTkFrame(block, fg_color="transparent")
        top.pack(fill="x", padx=10, pady=(8, 4))
        ctk.CTkLabel(
            top,
            text=f"{qid}: {question.pergunta}",
            anchor="w",
            wraplength=780,
            font=ctk.CTkFont(weight="bold"),
        ).pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(top, text="Nota:", width=40).pack(side="right", padx=(8, 4))
        combo = ctk.CTkComboBox(top, values=SCORE_FORM_OPCOES, width=70, state="readonly")
        combo.set(FORM_BLANK)
        combo.pack(side="right")
        store[qid] = combo

        hints = ctk.CTkFrame(block, fg_color="transparent")
        hints.pack(fill="x", padx=10, pady=(0, 10))
        ctk.CTkLabel(
            hints,
            text=f"1 → {question.significado_nota_1}",
            anchor="w",
            wraplength=880,
            font=ctk.CTkFont(size=11),
            text_color="#2B6CB0",
        ).pack(fill="x")
        ctk.CTkLabel(
            hints,
            text=f"5 → {question.significado_nota_5}",
            anchor="w",
            wraplength=880,
            font=ctk.CTkFont(size=11),
            text_color="#2F855A",
        ).pack(fill="x")

    def _build_inventory_tab(self, parent: ctk.CTkFrame) -> None:
        top = ctk.CTkFrame(parent, fg_color="transparent")
        top.pack(fill="x", pady=8)
        ctk.CTkButton(top, text="Importar inventário (.xlsx)", command=self._import_inventory).pack(
            side="left"
        )
        ctk.CTkButton(top, text="Salvar modelo de inventário", command=self._save_inventory_template).pack(
            side="left", padx=8
        )
        self.inv_count_label = ctk.CTkLabel(top, text="")
        self.inv_count_label.pack(side="left", padx=12)

        ctk.CTkLabel(
            parent,
            text=(
                "Depois de gerar o roteiro, as linhas amarelas são os IDs usados. "
                "Clique em Novo ID e escolha outro destino na lista. Depois gere o PDF de novo."
            ),
            wraplength=980,
            justify="left",
        ).pack(anchor="w", pady=(0, 4))

        table_frame = ctk.CTkFrame(parent)
        table_frame.pack(fill="both", expand=True, pady=8)

        cols = ("novo_id", "id", "nome", "categoria", "descricao")
        style = ttk.Style(self)

        def _tag_friendly(option: str) -> list:
            try:
                mapped = style.map("Treeview", query_opt=option)
            except tk.TclError:
                return []
            return [item for item in mapped if item[:2] != ("!disabled", "!selected")]

        try:
            style.map(
                "Treeview",
                foreground=_tag_friendly("foreground"),
                background=_tag_friendly("background"),
            )
        except tk.TclError:
            pass
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", height=18)
        self.tree.tag_configure("usado", background="#F6E05E")
        headings = {
            "novo_id": "Novo ID",
            "id": "ID",
            "nome": "Nome da atração",
            "categoria": "Categoria",
            "descricao": "Descrição breve",
        }
        widths = {"novo_id": 90, "id": 80, "nome": 260, "categoria": 150, "descricao": 420}
        for c, title in headings.items():
            self.tree.heading(c, text=title)
            self.tree.column(c, width=widths[c], minwidth=60, stretch=c == "descricao", anchor="w")
        self.tree.bind("<ButtonRelease-1>", self._start_novo_id_edit)
        scroll_y = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")
        self._refresh_inventory_table()

    def _build_pdf_tab(self, parent: ctk.CTkFrame) -> None:
        top = ctk.CTkFrame(parent, fg_color="transparent")
        top.pack(fill="x", pady=8)

        self.btn_generate_pdf = ctk.CTkButton(
            top,
            text="Calcular roteiro e gerar PDF",
            height=40,
            font=ctk.CTkFont(weight="bold"),
            command=self._generate_pdf,
        )
        self.btn_generate_pdf.pack(side="left")

        ctk.CTkButton(top, text="Abrir PDF", command=self._open_pdf).pack(side="left", padx=8)
        ctk.CTkButton(top, text="Abrir pasta de saída", command=self._open_output_folder).pack(
            side="left", padx=8
        )

        self.pdf_status = ctk.CTkLabel(parent, text="", anchor="w")
        self.pdf_status.pack(fill="x", pady=4)

        self.preview = ctk.CTkTextbox(parent, font=ctk.CTkFont(family="Consolas", size=11))
        self.preview.pack(fill="both", expand=True, pady=8)

    def _build_anamnese_tab(self, parent: ctk.CTkFrame) -> None:
        ctk.CTkLabel(
            parent,
            text="Ficha de anamnese médica e termo de compromisso e risco",
            font=ctk.CTkFont(size=16, weight="bold"),
        ).pack(anchor="w", pady=(8, 4))
        ctk.CTkLabel(
            parent,
            text=(
                "Gera um PDF em branco para impressão (A4): anamnese e termo em páginas "
                "separadas. Imprima uma cópia por participante."
            ),
            wraplength=900,
            justify="left",
        ).pack(anchor="w", pady=(0, 12))

        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", pady=4)
        ctk.CTkButton(
            row,
            text="Gerar PDF da ficha",
            height=40,
            font=ctk.CTkFont(weight="bold"),
            command=self._generate_anamnese_pdf,
        ).pack(side="left")
        ctk.CTkButton(row, text="Abrir PDF", command=self._open_anamnese_pdf).pack(side="left", padx=8)
        ctk.CTkButton(row, text="Abrir pasta de saída", command=self._open_output_folder).pack(
            side="left", padx=8
        )

        self.anamnese_status = ctk.CTkLabel(parent, text="", anchor="w")
        self.anamnese_status.pack(fill="x", pady=8)

    def _refresh_scores_panel(self) -> None:
        for w in self.score_frame.winfo_children():
            w.destroy()
        profile = read_profile(self.workbook_path)
        missing = validate_profile_for_roteiro(profile)
        ctk.CTkLabel(
            self.score_frame,
            text="4. Score do grupo (cálculo automático)",
            font=ctk.CTkFont(weight="bold"),
        ).pack(anchor="w", padx=8, pady=(8, 4))
        if missing:
            ctk.CTkLabel(
                self.score_frame,
                text="Preencha todo o perfil (incluindo Plog e Cohen) para exibir os scores.",
                text_color="gray",
                wraplength=900,
            ).pack(anchor="w", padx=8, pady=(0, 8))
            return
        scores = compute_group_scores(profile)
        row = ctk.CTkFrame(self.score_frame, fg_color="transparent")
        row.pack(fill="x", padx=8, pady=(0, 8))
        ctk.CTkLabel(row, text=f"Média Plog: {scores['media_plog']:.2f}").pack(side="left", padx=(0, 24))
        ctk.CTkLabel(row, text=f"Média Cohen: {scores['media_cohen']:.2f}").pack(side="left", padx=(0, 24))
        ctk.CTkLabel(row, text=f"Fator: {scores['fator_restritivo']}").pack(side="left")

    def _apply_empty_form(self) -> None:
        """Zera a tela (inclui Plog e Cohen em —), sem reler valores antigos da planilha."""
        self.nome_var.set("")
        self.pessoas_var.set("")
        self.trip_calendar.clear()
        self.perfil_combo.set(FORM_BLANK)
        self.ritmo_combo.set(FORM_BLANK)
        for var in self.cidade_vars.values():
            var.set(False)
        for q in PLOG_QUESTIONS:
            self.plog_combos[q.id].set(FORM_BLANK)
        for q in COHEN_QUESTIONS:
            self.cohen_combos[q.id].set(FORM_BLANK)
        self.update_idletasks()
        self._refresh_scores_panel()

    def _load_profile_into_form(self) -> None:
        profile = read_profile(self.workbook_path)
        self.nome_var.set(profile.nome_grupo)
        self.pessoas_var.set(str(profile.numero_pessoas) if profile.numero_pessoas else "")
        self.trip_calendar.set_dates(profile.check_in, profile.check_out)
        self.perfil_combo.set(
            PERFIL_FORM_OPCOES[_index_or_zero(PERFIL_FORM_OPCOES, profile.perfil_predominante)]
        )
        self.ritmo_combo.set(RITMO_FORM_OPCOES[_index_or_zero(RITMO_FORM_OPCOES, profile.ritmo)])
        for city, var in self.cidade_vars.items():
            var.set(city in profile.cidades)
        for qid, combo in self.plog_combos.items():
            combo.set(_score_display(int(profile.plog_scores.get(qid, 0))))
        for qid, combo in self.cohen_combos.items():
            combo.set(_score_display(int(profile.cohen_scores.get(qid, 0))))
        self.update_idletasks()
        self._refresh_scores_panel()

    def _profile_from_form(self) -> TouristProfile:
        try:
            pessoas = _parse_int_field(self.pessoas_var.get())
        except ValueError:
            pessoas = 0
        return TouristProfile(
            nome_grupo=self.nome_var.get().strip(),
            numero_pessoas=pessoas,
            duracao_dias=0,
            check_in=self.trip_calendar.check_in,
            check_out=self.trip_calendar.check_out,
            perfil_predominante=normalize_form_choice(self.perfil_combo.get()),
            ritmo=normalize_form_choice(self.ritmo_combo.get()),
            cidades=[city for city, var in self.cidade_vars.items() if var.get()],
            plog_scores={qid: _parse_score_field(combo.get()) for qid, combo in self.plog_combos.items()},
            cohen_scores={qid: _parse_score_field(combo.get()) for qid, combo in self.cohen_combos.items()},
        )

    def _save_profile(self) -> None:
        try:
            profile = self._profile_from_form()
        except ValueError:
            messagebox.showerror("Dados inválidos", "Número de pessoas deve ser um inteiro.")
            return
        if self.pessoas_var.get().strip() and profile.numero_pessoas < 1:
            messagebox.showerror("Dados inválidos", "Número de pessoas deve ser pelo menos 1.")
            return
        if profile.check_in and profile.check_out and profile.check_out < profile.check_in:
            messagebox.showerror("Dados inválidos", "O check-out deve ser na data do check-in ou depois.")
            return

        save_profile(self.workbook_path, profile)
        remember_workbook(self.workbook_path)
        self._refresh_scores_panel()
        messagebox.showinfo("Salvo", "Perfil gravado na planilha.")

    def _save_inventory_template(self) -> None:
        create_inventory_template(inventory_template_path(ROOT))
        dest = filedialog.asksaveasfilename(
            title="Salvar modelo de inventário",
            defaultextension=".xlsx",
            initialfile="Modelo_Inventario_Turistico.xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not dest:
            return
        shutil.copy2(inventory_template_path(ROOT), dest)
        messagebox.showinfo("Modelo", f"Modelo salvo em:\n{dest}")

    def _import_inventory(self) -> None:
        path = filedialog.askopenfilename(
            title="Selecione o inventário turístico",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not path:
            return
        try:
            count = import_inventory_sheet(self.workbook_path, Path(path))
            remember_workbook(self.workbook_path)
            self._used_order = []
            self._novo_ids = {}
            self._refresh_inventory_table()
            with_maps, total = inventory_maps_coverage(self.workbook_path)
            maps_line = (
                f"\nLinks Google Maps: {with_maps}/{total}."
                if total
                else ""
            )
            messagebox.showinfo("Importado", f"{count} atrações importadas.{maps_line}")
        except Exception as exc:
            messagebox.showerror("Erro", str(exc))

    def _refresh_inventory_table(self) -> None:
        self._close_novo_id_edit(save=False)
        for item in self.tree.get_children():
            self.tree.delete(item)
        items = read_inventory(self.workbook_path)
        used = set(self._used_order)
        items.sort(key=lambda attraction: (attraction.id not in used, attraction.id))
        for attraction in items:
            self.tree.insert(
                "",
                "end",
                iid=attraction.id,
                tags=("usado",) if attraction.id in used else (),
                values=(
                    self._novo_ids.get(attraction.id, ""),
                    attraction.id,
                    attraction.nome,
                    attraction.categoria or "—",
                    attraction.descricao or "—",
                ),
            )
        if used:
            self.inv_count_label.configure(
                text=f"{len(items)} atração(ões) · {len(used)} no roteiro"
            )
        else:
            self.inv_count_label.configure(text=f"{len(items)} atração(ões)")

    def _novo_id_choices(self, current_id: str) -> list[str]:
        options = ["— manter o ID —"]
        for attraction in read_inventory(self.workbook_path):
            if attraction.id == current_id:
                continue
            if "hospedagem" in attraction.categoria.lower():
                continue
            options.append(f"{attraction.id} — {attraction.nome}")
        return options

    def _start_novo_id_edit(self, event: tk.Event) -> None:
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        if self.tree.identify_column(event.x) != "#1":
            return
        row = self.tree.identify_row(event.y)
        if not row or row not in self._used_order:
            return
        if self._edit_row == row:
            return
        bbox = self.tree.bbox(row, "novo_id")
        if not bbox:
            return
        self._close_novo_id_edit(save=True)
        x, y, width, height = bbox
        choices = self._novo_id_choices(row)
        current = self._novo_ids.get(row, "")
        selected = "— manter o ID —"
        for option in choices:
            if option.startswith(f"{current} —"):
                selected = option
                break
        combo = ttk.Combobox(self.tree, values=choices, state="readonly")
        combo.place(x=x, y=y, width=max(width, 280), height=height)
        combo.set(selected)
        combo.bind("<<ComboboxSelected>>", lambda _event: self._close_novo_id_edit(save=True))
        combo.bind("<Escape>", lambda _event: self._close_novo_id_edit(save=False))
        self._edit_entry = combo
        self._edit_row = row
        combo.focus_set()
        combo.after(40, lambda widget=combo: widget.tk.call("ttk::combobox::Post", widget._w))

    def _close_novo_id_edit(self, save: bool) -> None:
        entry = self._edit_entry
        row = self._edit_row
        if entry is None:
            return
        self._edit_entry = None
        self._edit_row = None
        raw = entry.get().strip() if save else ""
        entry.destroy()
        if not save or not row:
            return
        if not raw or raw.startswith("—"):
            self._novo_ids.pop(row, None)
            value = ""
        else:
            value = raw.split(" — ", 1)[0].strip()
            self._novo_ids[row] = value
        if self.tree.exists(row):
            self.tree.set(row, "novo_id", value)

    def _pinned_ids_from_table(self) -> list[str] | None:
        if not self._used_order:
            return None
        self._close_novo_id_edit(save=True)
        profile = self._profile_from_form()
        inventory = filter_inventory_by_cities(read_inventory(self.workbook_path), profile.cidades)
        by_id = {attraction.id: attraction for attraction in inventory}
        chosen: list[str] = []
        seen: set[str] = set()
        for original in self._used_order:
            target = self._novo_ids.get(original, "").strip() or original
            attraction = by_id.get(target)
            if attraction is None:
                raise ValueError(
                    f"O ID '{target}' não está no inventário das cidades marcadas no perfil."
                )
            if "hospedagem" in attraction.categoria.lower():
                raise ValueError(
                    f"{target} é hospedagem. Informe o ID de um atrativo ou de uma gastronomia."
                )
            if target in seen:
                raise ValueError(f"O ID {target} ficou repetido na nova seleção.")
            seen.add(target)
            chosen.append(target)
        return chosen

    def _generate_pdf(self) -> None:
        if self._pdf_job_running:
            return
        try:
            profile = self._profile_from_form()
            missing = validate_profile_for_roteiro(profile)
            if missing:
                messagebox.showerror("Perfil incompleto", format_profile_validation_errors(missing))
                return
            save_profile(self.workbook_path, profile)
            remember_workbook(self.workbook_path)
            with_maps, total = inventory_maps_coverage(self.workbook_path)
            if total and with_maps == 0:
                proceed = messagebox.askyesno(
                    "Links do Google Maps ausentes",
                    "O inventário desta sessão não tem links na coluna Maps.\n\n"
                    "Na etapa 2, importe de novo o arquivo "
                    "Modelo_Inventario_Turistico.xlsx (Desktop → PROJETO INTEGRADOR).\n\n"
                    "Gerar o PDF mesmo assim, sem links no roteiro?",
                )
                if not proceed:
                    return
        except ValueError as exc:
            messagebox.showerror("Inventário", str(exc))
            return
        except Exception as exc:
            messagebox.showerror("Erro", str(exc))
            return

        try:
            pinned_ids = self._pinned_ids_from_table()
        except ValueError as exc:
            messagebox.showerror("Novo ID", str(exc))
            return

        workbook = self.workbook_path
        self._pdf_job_running = True
        self.btn_generate_pdf.configure(state="disabled")
        self.pdf_status.configure(text="Calculando roteiro e gerando PDF… Aguarde.")
        self.update_idletasks()

        def worker() -> None:
            err: tuple[str, BaseException] | None = None
            result = None
            pdf_path: Path | None = None
            preview_text = ""
            try:
                result = gerar_roteiro_da_planilha(workbook, pinned_ids=pinned_ids)
                pdf_path = OUT / roteiro_pdf_filename(result.profile.nome_grupo)
                build_pdf(result, pdf_path)
                preview_text = itinerary_to_markdown(result)
            except ValueError as exc:
                err = ("inventario", exc)
            except Exception as exc:
                err = ("erro", exc)

            def finish() -> None:
                self._pdf_job_running = False
                self.btn_generate_pdf.configure(state="normal")
                if err:
                    kind, exc = err
                    if kind == "inventario":
                        messagebox.showerror("Inventário", str(exc))
                    else:
                        messagebox.showerror("Erro", str(exc))
                    self.pdf_status.configure(text="")
                    return
                assert result is not None and pdf_path is not None
                self.itinerary_result = result
                self.last_pdf = pdf_path
                self._used_order = [attraction.id for attraction in result.selected_attractions]
                self._novo_ids = {}
                self._refresh_inventory_table()
                self.preview.delete("0.0", "end")
                self.preview.insert("0.0", preview_text)
                self.pdf_status.configure(text=f"PDF gerado: {pdf_path}")
                note = ""
                if pinned_ids:
                    kept = {attraction.id for attraction in result.selected_attractions}
                    left_out = [item_id for item_id in pinned_ids if item_id not in kept]
                    if left_out:
                        note = (
                            "\n\nNão couberam neste roteiro (horário, deslocamento ou ritmo): "
                            + ", ".join(left_out)
                        )
                messagebox.showinfo("Concluído", f"Roteiro PDF salvo em:\n{pdf_path}{note}")

            self.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    def _open_pdf(self) -> None:
        if not self.last_pdf or not self.last_pdf.exists():
            messagebox.showwarning("PDF", "Gere o PDF primeiro na aba 3.")
            return
        os.startfile(self.last_pdf)

    def _generate_anamnese_pdf(self) -> None:
        try:
            profile = read_profile(self.workbook_path)
            pdf_path = OUT / anamnese_pdf_filename(profile.nome_grupo or "grupo")
            build_anamnese_pdf(
                pdf_path,
                nome_grupo=profile.nome_grupo,
            )
            self.last_anamnese_pdf = pdf_path
            self.anamnese_status.configure(text=f"PDF gerado: {pdf_path}")
            messagebox.showinfo(
                "Ficha de anamnese",
                f"PDF salvo em:\n{pdf_path}\n\nAnamnese e termo em páginas separadas.",
            )
        except Exception as exc:
            messagebox.showerror("Erro", str(exc))

    def _open_anamnese_pdf(self) -> None:
        if not self.last_anamnese_pdf or not self.last_anamnese_pdf.exists():
            messagebox.showwarning("PDF", "Gere o PDF primeiro na aba 4.")
            return
        os.startfile(self.last_anamnese_pdf)

    def _open_output_folder(self) -> None:
        os.startfile(OUT)

    def _reset_session(self, *, new_file: bool) -> None:
        if new_file:
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            self.workbook_path = initialize_session_workbook(
                OUT / f"Projeto_Integrador_{stamp}.xlsx", TEMPLATE
            )
            self.path_label.configure(text=self.workbook_path.name)
        else:
            clear_profile_only(self.workbook_path)
        remember_workbook(self.workbook_path)
        self.last_pdf = None
        self.itinerary_result = None
        self._used_order = []
        self._novo_ids = {}
        self.preview.delete("0.0", "end")
        self.pdf_status.configure(text="")
        self._apply_empty_form()
        self._refresh_inventory_table()
        self.tabs.set("1. Perfil")

    def _clear_all(self) -> None:
        if not messagebox.askyesno(
            "Limpar informações",
            "Apagar o perfil e o roteiro gerado nesta sessão?\n(O inventário importado será mantido.)",
        ):
            return
        self._reset_session(new_file=False)
        messagebox.showinfo("Limpo", "Perfil resetado. O inventário foi mantido.")

    def _new_workbook(self) -> None:
        if not messagebox.askyesno("Nova planilha", "Criar nova planilha e limpar a sessão atual?"):
            return
        self._reset_session(new_file=True)


def main() -> None:
    if sys.platform == "win32":
        try:
            from ctypes import windll

            windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    app = RoteiroDesktopApp()
    app.mainloop()


if __name__ == "__main__":
    main()
