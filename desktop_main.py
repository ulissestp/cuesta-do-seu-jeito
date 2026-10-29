"""
Gerador de Roteiros Turísticos — aplicativo desktop (Windows).
Execute: python desktop_main.py
"""

from __future__ import annotations

import os
import shutil
import sys
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from src.app_paths import output_dir, project_root
from src.branding import APP_NAME, APP_TAGLINE
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
    import_inventory_sheet,
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


class RoteiroDesktopApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1120x780")
        self.minsize(960, 640)

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
        self.itinerary_result = None

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
            text=f"{APP_NAME} · {APP_TAGLINE}",
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

        self._build_profile_tab(self.tabs.tab("1. Perfil"))
        self._build_inventory_tab(self.tabs.tab("2. Inventário"))
        self._build_pdf_tab(self.tabs.tab("3. Roteiro PDF"))

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
        self.duracao_var = tk.StringVar(value="")

        self._labeled_entry(grid, 0, "Nome do grupo / contato", self.nome_var)
        self._labeled_entry(grid, 0, "Número de pessoas", self.pessoas_var, col=2)
        self._labeled_entry(grid, 1, "Duração (dias)", self.duracao_var)
        self.perfil_combo = self._labeled_combo(grid, 1, "Perfil predominante", PERFIL_FORM_OPCOES, col=2)
        self.ritmo_combo = self._labeled_combo(grid, 2, "Ritmo de viagem", RITMO_FORM_OPCOES)

        ctk.CTkLabel(scroll, text="Cidades do roteiro", font=ctk.CTkFont(weight="bold")).pack(
            anchor="w", pady=(12, 4)
        )
        ctk.CTkLabel(
            scroll,
            text="Selecione uma ou mais cidades. O roteiro usará apenas atrações dessas cidades (coluna K do inventário).",
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

        table_frame = ctk.CTkFrame(parent)
        table_frame.pack(fill="both", expand=True, pady=8)

        cols = ("id", "nome", "cidade", "abertura", "fechamento", "cap", "plog", "cohen")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", height=18)
        headings = {
            "id": "ID",
            "nome": "Atração",
            "cidade": "Cidade",
            "abertura": "Abre",
            "fechamento": "Fecha",
            "cap": "Cap.",
            "plog": "Plog",
            "cohen": "Cohen",
        }
        for c, title in headings.items():
            self.tree.heading(c, text=title)
            self.tree.column(c, width=90 if c == "cidade" else (100 if c != "nome" else 240))
        scroll_y = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")
        self._refresh_inventory_table()

    def _build_pdf_tab(self, parent: ctk.CTkFrame) -> None:
        top = ctk.CTkFrame(parent, fg_color="transparent")
        top.pack(fill="x", pady=8)

        ctk.CTkButton(
            top,
            text="Calcular roteiro e gerar PDF",
            height=40,
            font=ctk.CTkFont(weight="bold"),
            command=self._generate_pdf,
        ).pack(side="left")

        ctk.CTkButton(top, text="Abrir PDF", command=self._open_pdf).pack(side="left", padx=8)
        ctk.CTkButton(top, text="Abrir pasta de saída", command=self._open_output_folder).pack(
            side="left", padx=8
        )

        self.pdf_status = ctk.CTkLabel(parent, text="", anchor="w")
        self.pdf_status.pack(fill="x", pady=4)

        self.preview = ctk.CTkTextbox(parent, font=ctk.CTkFont(family="Consolas", size=11))
        self.preview.pack(fill="both", expand=True, pady=8)

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
        self.duracao_var.set("")
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
        self.duracao_var.set(str(profile.duracao_dias) if profile.duracao_dias else "")
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
            duracao = _parse_int_field(self.duracao_var.get())
        except ValueError:
            pessoas = 0
            duracao = 0
        return TouristProfile(
            nome_grupo=self.nome_var.get().strip(),
            numero_pessoas=pessoas,
            duracao_dias=duracao,
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
            messagebox.showerror("Dados inválidos", "Número de pessoas e duração devem ser inteiros.")
            return
        if self.pessoas_var.get().strip() and profile.numero_pessoas < 1:
            messagebox.showerror("Dados inválidos", "Número de pessoas deve ser pelo menos 1.")
            return
        if self.duracao_var.get().strip() and profile.duracao_dias < 1:
            messagebox.showerror("Dados inválidos", "Duração deve ser pelo menos 1 dia.")
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
            self._refresh_inventory_table()
            messagebox.showinfo("Importado", f"{count} atrações importadas.")
        except Exception as exc:
            messagebox.showerror("Erro", str(exc))

    def _refresh_inventory_table(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        items = read_inventory(self.workbook_path)
        for a in items:
            self.tree.insert(
                "",
                "end",
                values=(
                    a.id,
                    a.nome,
                    a.cidade or "—",
                    a.horario_abertura.strftime("%H:%M") if a.horario_abertura else "—",
                    a.horario_fechamento.strftime("%H:%M") if a.horario_fechamento else "—",
                    a.capacidade,
                    a.peso_plog,
                    a.peso_cohen,
                ),
            )
        self.inv_count_label.configure(text=f"{len(items)} atração(ões)")

    def _generate_pdf(self) -> None:
        try:
            profile = self._profile_from_form()
            missing = validate_profile_for_roteiro(profile)
            if missing:
                messagebox.showerror("Perfil incompleto", format_profile_validation_errors(missing))
                return
            save_profile(self.workbook_path, profile)
            remember_workbook(self.workbook_path)
            result = gerar_roteiro_da_planilha(self.workbook_path)
            self.itinerary_result = result
            profile = result.profile
            pdf_path = OUT / roteiro_pdf_filename(profile.nome_grupo)
            build_pdf(result, pdf_path)
            self.last_pdf = pdf_path
            self.preview.delete("0.0", "end")
            self.preview.insert("0.0", itinerary_to_markdown(result))
            self.pdf_status.configure(text=f"PDF gerado: {pdf_path}")
            messagebox.showinfo("Concluído", f"Roteiro PDF salvo em:\n{pdf_path}")
        except ValueError as exc:
            messagebox.showerror("Inventário", str(exc))
        except Exception as exc:
            messagebox.showerror("Erro", str(exc))

    def _open_pdf(self) -> None:
        if not self.last_pdf or not self.last_pdf.exists():
            messagebox.showwarning("PDF", "Gere o PDF primeiro na aba 3.")
            return
        os.startfile(self.last_pdf)

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
