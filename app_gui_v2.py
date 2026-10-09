# -*- coding: utf-8 -*-
"""Interface principal do Leitor de XML NF-e.

Fluxo orientado à tarefa: carregar XMLs, selecionar campos e exportar.
"""

import os
import threading
import unicodedata
from datetime import datetime
from typing import Any, Dict, List, Optional

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from excel_exporter import (
    ITEM_COLUMNS,
    NOTA_COLUMNS,
    export_nfe_to_excel,
)
from nfe_parser import NFeParser, format_brazilian_number
from xml_field_labels import describe_xml_group, describe_xml_path


COLORS = {
    "nav": "#0b1220",
    "panel": "#111c2f",
    "panel_alt": "#17243a",
    "line": "#263751",
    "primary": "#2f81f7",
    "primary_hover": "#1f6feb",
    "success": "#22c55e",
    "success_hover": "#16a34a",
    "muted": "#94a3b8",
    "text": "#f8fafc",
    "warning": "#f59e0b",
}


def normalize_text(value: Any) -> str:
    normalized = unicodedata.normalize("NFD", str(value).casefold())
    return "".join(char for char in normalized if unicodedata.category(char) != "Mn")


class NotaDetailWindow(ctk.CTkToplevel):
    """Resumo da nota selecionada, sem interferir na seleção de exportação."""

    def __init__(self, parent, nota: Dict[str, Any]):
        super().__init__(parent)
        self.nota = nota
        self.title(f"NF-e {nota.get('numero_nota', '')} — detalhes")
        self.geometry("1180x760")
        self.minsize(900, 620)
        self.transient(parent)
        self.grab_set()
        self._build()

    def _build(self):
        header = ctk.CTkFrame(self, fg_color=COLORS["panel"], corner_radius=12)
        header.pack(fill="x", padx=18, pady=(18, 8))
        ctk.CTkLabel(
            header,
            text=f"NF-e {self.nota.get('numero_nota', '')}  •  Série {self.nota.get('serie', '')}",
            font=ctk.CTkFont(size=20, weight="bold"),
        ).pack(anchor="w", padx=16, pady=(13, 2))
        ctk.CTkLabel(
            header,
            text=f"Chave: {self.nota.get('chave', '')}",
            font=ctk.CTkFont(family="Consolas", size=12),
            text_color=COLORS["muted"],
        ).pack(anchor="w", padx=16, pady=(0, 13))

        tabs = ctk.CTkTabview(self)
        tabs.pack(fill="both", expand=True, padx=18, pady=8)
        self._build_summary(tabs.add("Resumo"))
        self._build_items(tabs.add("Itens"))
        self._build_xml(tabs.add("Tags XML"))

        ctk.CTkButton(self, text="Fechar", width=110, command=self.destroy).pack(
            anchor="e", padx=18, pady=(2, 16)
        )

    def _build_summary(self, parent):
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=6, pady=6)
        sections = [
            ("Identificação", [
                ("Emissão", self.nota.get("data_emissao", "")),
                ("Natureza", self.nota.get("natureza_operacao", "")),
                ("Valor total", f"R$ {self.nota.get('total_nota_fmt', '0,00')}"),
            ]),
            ("Emitente", [
                ("Razão social", self.nota.get("emit_nome", "")),
                ("CNPJ/CPF", self.nota.get("emit_cnpj_formatado", "")),
                ("Município/UF", f"{self.nota.get('emit_municipio', '')}/{self.nota.get('emit_uf', '')}"),
            ]),
            ("Destinatário", [
                ("Razão social", self.nota.get("dest_nome", "")),
                ("CNPJ/CPF", self.nota.get("dest_cnpj_formatado", "")),
                ("Município/UF", f"{self.nota.get('dest_municipio', '')}/{self.nota.get('dest_uf', '')}"),
            ]),
            ("Transporte", [
                ("Transportadora", self.nota.get("transp_nome", "") or "Não informada"),
                ("Placa", self.nota.get("placa_veiculo", "") or "Não informada"),
                ("Peso bruto", f"{self.nota.get('peso_bruto_fmt', '0,000')} kg"),
            ]),
        ]
        for title, rows in sections:
            card = ctk.CTkFrame(scroll, fg_color=COLORS["panel_alt"], corner_radius=10)
            card.pack(fill="x", pady=5)
            ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=14, weight="bold"),
                         text_color="#7dd3fc").pack(anchor="w", padx=14, pady=(10, 5))
            for label, value in rows:
                ctk.CTkLabel(card, text=f"{label}: {value}", anchor="w", justify="left").pack(
                    fill="x", padx=14, pady=2
                )
            ctk.CTkLabel(card, text="", height=6).pack()

    def _build_items(self, parent):
        columns = ("item", "codigo", "descricao", "ncm", "qtd", "valor")
        tree = ttk.Treeview(parent, columns=columns, show="headings")
        headings = {
            "item": "Item", "codigo": "Código", "descricao": "Descrição",
            "ncm": "NCM", "qtd": "Quantidade", "valor": "Valor",
        }
        widths = {"item": 55, "codigo": 110, "descricao": 330, "ncm": 90, "qtd": 100, "valor": 110}
        for key in columns:
            tree.heading(key, text=headings[key])
            tree.column(key, width=widths[key], anchor="w" if key in ("codigo", "descricao") else "center")
        for item in self.nota.get("itens", []):
            tree.insert("", "end", values=(
                item.get("n_item", ""), item.get("c_prod", ""), item.get("x_prod", ""),
                item.get("ncm", ""), item.get("q_com_fmt", ""), item.get("v_prod_fmt", ""),
            ))
        tree.pack(fill="both", expand=True, padx=8, pady=8)

    def _build_xml(self, parent):
        search = ctk.CTkEntry(parent, placeholder_text="Localizar caminho ou valor...")
        search.pack(fill="x", padx=8, pady=(8, 4))
        box = ctk.CTkTextbox(parent, font=ctk.CTkFont(family="Consolas", size=11))
        box.pack(fill="both", expand=True, padx=8, pady=(4, 8))

        def render(_event=None):
            query = normalize_text(search.get().strip())
            lines = []
            for path, values in sorted(self.nota.get("tags_xml", {}).items(), key=lambda pair: normalize_text(pair[0])):
                line = f"{path} = {' | '.join(values)}"
                if not query or query in normalize_text(line):
                    lines.append(line)
            box.configure(state="normal")
            box.delete("1.0", "end")
            box.insert("1.0", "\n".join(lines))
            box.configure(state="disabled")

        search.bind("<KeyRelease>", render)
        render()


class AppNFe(ctk.CTk):
    """Aplicação redesenhada com seleção integrada de campos."""

    CATEGORY_ALL = "Todos os campos"
    CATEGORY_NOTE = "Dados da nota"
    CATEGORY_ITEM = "Itens e produtos"
    CATEGORY_XML = "Tags encontradas no XML"

    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")
        self.title("Extrator NF-e — XML para Excel")
        self.geometry("1600x920")
        self.minsize(1280, 760)

        self.notas: List[Dict[str, Any]] = []
        self.selected: Dict[str, List[str]] = {"notas": [], "itens": [], "xml": [], "ordem": []}
        self.field_variables: Dict[str, tk.BooleanVar] = {}
        self.field_catalog: List[Dict[str, str]] = []
        self.visible_field_ids: List[str] = []
        self.discovered_paths = set()
        self.last_excel: Optional[str] = None
        self.loading = False
        self.note_page = 0
        self.notes_per_page = 250
        self.field_page = 0
        self.fields_per_page = 30
        self._field_search_job = None
        self._note_search_job = None

        self._configure_tree_style()
        self._build_layout()
        self._rebuild_field_catalog()
        self.after(0, self._maximize_window)

    def _maximize_window(self):
        try:
            self.state("zoomed")
        except tk.TclError:
            pass

    def _configure_tree_style(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background=COLORS["panel"], foreground=COLORS["text"],
                        fieldbackground=COLORS["panel"], rowheight=36, borderwidth=0,
                        font=("Segoe UI", 11))
        style.configure("Treeview.Heading", background=COLORS["nav"], foreground="#bae6fd",
                        font=("Segoe UI", 11, "bold"), padding=(9, 10), borderwidth=0)
        style.map("Treeview", background=[("selected", COLORS["primary"])],
                  foreground=[("selected", "#ffffff")])

    def _build_layout(self):
        self.configure(fg_color="#0f172a")
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self._build_sidebar()
        self._build_workspace()

    def _build_sidebar(self):
        sidebar = ctk.CTkFrame(self, width=275, corner_radius=0, fg_color=COLORS["nav"])
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)

        ctk.CTkLabel(sidebar, text="EXTRATOR NF-e", font=ctk.CTkFont(size=23, weight="bold"),
                     text_color="#7dd3fc").pack(anchor="w", padx=22, pady=(26, 2))
        ctk.CTkLabel(sidebar, text="XML  →  Excel", font=ctk.CTkFont(size=12),
                     text_color=COLORS["muted"]).pack(anchor="w", padx=22, pady=(0, 24))

        self.btn_files = self._sidebar_button(sidebar, "1   Carregar arquivos XML", self._choose_files, COLORS["primary"])
        self.btn_folder = self._sidebar_button(sidebar, "    Carregar uma pasta", self._choose_folder, COLORS["panel_alt"])

        separator = ctk.CTkFrame(sidebar, height=1, fg_color=COLORS["line"])
        separator.pack(fill="x", padx=20, pady=18)

        self.step_notes = self._step_label(sidebar, "1", "XMLs carregados", "Aguardando arquivos")
        self.step_fields = self._step_label(sidebar, "2", "Campos escolhidos", "Somente a chave")
        self.step_export = self._step_label(sidebar, "3", "Exportar planilha", "Pronto após carregar")

        self.btn_clear = self._sidebar_button(sidebar, "Limpar sessão", self._clear_all, "#334155")
        self.btn_clear.pack(side="bottom", fill="x", padx=18, pady=(0, 10))
        ctk.CTkLabel(sidebar, text="Leitor de documentos fiscais", font=ctk.CTkFont(size=10),
                     text_color="#64748b").pack(side="bottom", pady=(8, 18))

    def _sidebar_button(self, parent, text, command, color):
        button = ctk.CTkButton(parent, text=text, command=command, height=42, anchor="w",
                               font=ctk.CTkFont(size=12, weight="bold"), fg_color=color,
                               hover_color=COLORS["primary_hover"], corner_radius=8)
        button.pack(fill="x", padx=18, pady=4)
        return button

    def _step_label(self, parent, number, title, detail):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", padx=18, pady=7)
        ctk.CTkLabel(frame, text=number, width=28, height=28, corner_radius=14,
                     fg_color=COLORS["panel_alt"], font=ctk.CTkFont(weight="bold")).pack(side="left")
        text = ctk.CTkFrame(frame, fg_color="transparent")
        text.pack(side="left", padx=9)
        ctk.CTkLabel(text, text=title, font=ctk.CTkFont(size=11, weight="bold")).pack(anchor="w")
        label = ctk.CTkLabel(text, text=detail, font=ctk.CTkFont(size=10), text_color=COLORS["muted"])
        label.pack(anchor="w")
        return label

    def _build_workspace(self):
        workspace = ctk.CTkFrame(self, fg_color="transparent")
        workspace.grid(row=0, column=1, sticky="nsew", padx=26, pady=20)
        workspace.grid_columnconfigure(0, weight=1)
        workspace.grid_rowconfigure(2, weight=1)

        header = ctk.CTkFrame(workspace, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        ctk.CTkLabel(header, text="Transforme XMLs em uma planilha sob medida",
                     font=ctk.CTkFont(size=28, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(header, text="A chave da NF-e é fixa. Você decide todas as demais colunas.",
                     text_color=COLORS["muted"], font=ctk.CTkFont(size=12)).pack(anchor="w", pady=(3, 0))

        self._build_metrics(workspace)

        self.tabs = ctk.CTkTabview(workspace, corner_radius=12, fg_color=COLORS["panel"],
                                   segmented_button_selected_color=COLORS["primary"])
        self.tabs.grid(row=2, column=0, sticky="nsew", pady=8)
        self._build_notes_tab(self.tabs.add("Notas carregadas"))
        self._build_fields_tab(self.tabs.add("Campos para exportar"))

        self._build_footer(workspace)

    def _build_metrics(self, parent):
        metrics = ctk.CTkFrame(parent, fg_color="transparent")
        metrics.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        metrics.grid_columnconfigure((0, 1, 2, 3), weight=1)
        self.metric_notes = self._metric(metrics, 0, "NOTAS", "0", "#7dd3fc")
        self.metric_items = self._metric(metrics, 1, "ITENS", "0", "#c4b5fd")
        self.metric_value = self._metric(metrics, 2, "VALOR TOTAL", "R$ 0,00", "#86efac")
        self.metric_tags = self._metric(metrics, 3, "TAGS DESCOBERTAS", "0", "#fcd34d")

    def _metric(self, parent, column, title, value, color):
        card = ctk.CTkFrame(parent, fg_color=COLORS["panel"], corner_radius=10)
        card.grid(row=0, column=column, sticky="ew", padx=5)
        ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=10, weight="bold"),
                     text_color=COLORS["muted"]).pack(anchor="w", padx=14, pady=(10, 1))
        label = ctk.CTkLabel(card, text=value, font=ctk.CTkFont(size=22, weight="bold"), text_color=color)
        label.pack(anchor="w", padx=14, pady=(0, 10))
        return label

    def _build_notes_tab(self, parent):
        toolbar = ctk.CTkFrame(parent, fg_color="transparent")
        toolbar.pack(fill="x", padx=8, pady=(8, 4))
        ctk.CTkLabel(toolbar, text="Arquivos processados", font=ctk.CTkFont(size=15, weight="bold")).pack(side="left")
        self.note_search = ctk.CTkEntry(toolbar, placeholder_text="Buscar nota, empresa ou chave...", width=330, height=36)
        self.note_search.pack(side="right")
        self.note_search.bind("<KeyRelease>", self._schedule_note_search)

        pagination = ctk.CTkFrame(toolbar, fg_color="transparent")
        pagination.pack(side="right", padx=12)
        self.btn_prev_page = ctk.CTkButton(pagination, text="‹", width=34, height=32,
                                           fg_color="#334155", command=lambda: self._change_note_page(-1))
        self.btn_prev_page.pack(side="left")
        self.page_label = ctk.CTkLabel(pagination, text="Página 1 de 1", width=105)
        self.page_label.pack(side="left", padx=5)
        self.btn_next_page = ctk.CTkButton(pagination, text="›", width=34, height=32,
                                           fg_color="#334155", command=lambda: self._change_note_page(1))
        self.btn_next_page.pack(side="left")

        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=8, pady=(4, 8))
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(0, weight=1)
        columns = ("numero", "emissao", "emitente", "destinatario", "itens", "valor", "arquivo")
        self.notes_tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="browse")
        config = {
            "numero": ("NF", 80, "center"), "emissao": ("Emissão", 140, "center"),
            "emitente": ("Emitente", 220, "w"), "destinatario": ("Destinatário", 220, "w"),
            "itens": ("Itens", 70, "center"), "valor": ("Valor total", 115, "e"),
            "arquivo": ("Arquivo", 180, "w"),
        }
        for key, (title, width, anchor) in config.items():
            self.notes_tree.heading(key, text=title)
            self.notes_tree.column(key, width=width, anchor=anchor)
        self.notes_tree.bind("<Double-1>", self._open_note_details)
        y_scroll = ttk.Scrollbar(frame, orient="vertical", command=self.notes_tree.yview)
        self.notes_tree.configure(yscrollcommand=y_scroll.set)
        self.notes_tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")

    def _build_fields_tab(self, parent):
        info = ctk.CTkFrame(parent, fg_color="#102a43", corner_radius=9)
        info.pack(fill="x", padx=8, pady=(8, 6))
        ctk.CTkLabel(info, text="✓  Chave da NF-e", font=ctk.CTkFont(size=13, weight="bold"),
                     text_color="#7dd3fc").pack(side="left", padx=14, pady=10)
        ctk.CTkLabel(info, text="Primeira coluna obrigatória em todas as exportações",
                     text_color=COLORS["muted"]).pack(side="left", padx=8)

        filters = ctk.CTkFrame(parent, fg_color="transparent")
        filters.pack(fill="x", padx=8, pady=4)
        self.field_search = ctk.CTkEntry(
            filters,
            placeholder_text="Localizar pelo significado: pedido, município, valor, imposto...",
            width=440,
            height=38,
        )
        self.field_search.pack(side="left")
        self.field_search.bind("<KeyRelease>", self._schedule_field_search)
        self.category_filter = ctk.CTkOptionMenu(
            filters,
            values=[self.CATEGORY_ALL, self.CATEGORY_NOTE, self.CATEGORY_ITEM, self.CATEGORY_XML],
            width=210,
            command=self._category_changed,
        )
        self.category_filter.set(self.CATEGORY_ALL)
        self.category_filter.pack(side="left", padx=8)
        ctk.CTkButton(filters, text="Marcar página", width=115, command=lambda: self._set_visible(True)).pack(side="right")
        ctk.CTkButton(filters, text="Desmarcar página", width=130, fg_color="#334155",
                      hover_color="#475569", command=lambda: self._set_visible(False)).pack(side="right", padx=8)

        body = ctk.CTkFrame(parent, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=8, pady=(4, 8))
        body.grid_columnconfigure(0, weight=4)
        body.grid_columnconfigure(1, weight=2)
        body.grid_rowconfigure(0, weight=1)

        available = ctk.CTkFrame(body, fg_color=COLORS["panel_alt"], corner_radius=10)
        available.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        available_header = ctk.CTkFrame(available, fg_color="transparent")
        available_header.pack(fill="x", padx=12, pady=(10, 2))
        self.available_title = ctk.CTkLabel(available_header, text="Campos disponíveis",
                                            font=ctk.CTkFont(size=15, weight="bold"))
        self.available_title.pack(side="left")
        self.btn_next_field_page = ctk.CTkButton(
            available_header, text="›", width=34, height=30, fg_color="#334155",
            command=lambda: self._change_field_page(1),
        )
        self.btn_next_field_page.pack(side="right")
        self.field_page_label = ctk.CTkLabel(available_header, text="Página 1 de 1", width=105)
        self.field_page_label.pack(side="right", padx=5)
        self.btn_prev_field_page = ctk.CTkButton(
            available_header, text="‹", width=34, height=30, fg_color="#334155",
            command=lambda: self._change_field_page(-1),
        )
        self.btn_prev_field_page.pack(side="right")
        self.field_list = ctk.CTkScrollableFrame(available, fg_color="transparent")
        self.field_list.pack(fill="both", expand=True, padx=4, pady=(2, 6))
        self.field_list.grid_columnconfigure((0, 1), weight=1)

        chosen = ctk.CTkFrame(body, fg_color=COLORS["panel_alt"], corner_radius=10)
        chosen.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        ctk.CTkLabel(chosen, text="Colunas do Excel", font=ctk.CTkFont(size=15, weight="bold")).pack(
            anchor="w", padx=12, pady=(10, 2)
        )
        self.selected_list = ctk.CTkTextbox(
            chosen,
            fg_color="transparent",
            font=ctk.CTkFont(size=12),
            wrap="word",
        )
        self.selected_list.pack(fill="both", expand=True, padx=4, pady=(2, 6))

    def _build_footer(self, parent):
        footer = ctk.CTkFrame(parent, fg_color=COLORS["panel"], corner_radius=10)
        footer.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        status_area = ctk.CTkFrame(footer, fg_color="transparent")
        status_area.pack(side="left", fill="x", expand=True, padx=14, pady=10)
        self.status_label = ctk.CTkLabel(status_area, text="Carregue os XMLs para começar.", text_color=COLORS["muted"])
        self.status_label.pack(anchor="w")
        self.progress = ctk.CTkProgressBar(status_area, height=7)
        self.progress.set(0)
        self.progress.pack(fill="x", pady=(5, 0))

        self.btn_open = ctk.CTkButton(footer, text="Abrir Excel", width=110, state="disabled",
                                      fg_color="#334155", hover_color="#475569", command=self._open_excel)
        self.btn_open.pack(side="right", padx=(6, 14), pady=12)
        self.btn_export = ctk.CTkButton(footer, text="Exportar planilha", width=155,
                                       fg_color=COLORS["success"], hover_color=COLORS["success_hover"],
                                       font=ctk.CTkFont(weight="bold"), command=self._export)
        self.btn_export.pack(side="right", padx=6, pady=12)

    def _choose_files(self):
        if self.loading:
            return
        files = filedialog.askopenfilenames(title="Selecione os XMLs", filetypes=[("Arquivos XML", "*.xml")])
        if files:
            self._start_loading(list(files))

    def _choose_folder(self):
        if self.loading:
            return
        folder = filedialog.askdirectory(title="Selecione uma pasta com XMLs")
        if not folder:
            return
        files = [os.path.join(root, name) for root, _, names in os.walk(folder)
                 for name in names if name.lower().endswith(".xml")]
        if not files:
            messagebox.showwarning("Nenhum XML", "A pasta selecionada não contém arquivos XML.")
            return
        self._start_loading(files)

    def _start_loading(self, files: List[str]):
        self.loading = True
        self.btn_files.configure(state="disabled")
        self.btn_folder.configure(state="disabled")
        self.progress.set(0)
        self.status_label.configure(text=f"Lendo {len(files)} arquivo(s)...")
        threading.Thread(target=self._load_worker, args=(files,), daemon=True).start()

    def _load_worker(self, files: List[str]):
        existing = {note.get("chave") for note in self.notas if note.get("chave")}
        loaded_with_index, errors = [], []
        progress_step = max(1, len(files) // 100)
        for original_index, path in enumerate(files):
            try:
                data = NFeParser.parse_file(path)
            except Exception as exc:
                data = {"sucesso": False, "erro": str(exc)}
            key = data.get("chave")
            if data.get("sucesso") and (not key or key not in existing):
                loaded_with_index.append((original_index, data))
                if key:
                    existing.add(key)
            elif not data.get("sucesso"):
                errors.append(f"{os.path.basename(path)}: {data.get('erro', 'erro desconhecido')}")
            completed = original_index + 1
            if completed == len(files) or completed % progress_step == 0:
                self.after(0, self._loading_progress, completed, len(files))
        loaded = [data for _, data in sorted(loaded_with_index, key=lambda pair: pair[0])]
        self.after(0, self._finish_loading, loaded, errors)

    def _loading_progress(self, current: int, total: int):
        self.progress.set(current / total)
        self.status_label.configure(text=f"Processando {current} de {total} XMLs...")

    def _finish_loading(self, loaded: List[Dict[str, Any]], errors: List[str]):
        self.notas.extend(loaded)
        for note in loaded:
            self.discovered_paths.update(note.get("tags_xml", {}))
        self.note_page = 0
        self.loading = False
        self.btn_files.configure(state="normal")
        self.btn_folder.configure(state="normal")
        self.progress.set(1 if self.notas else 0)
        self._rebuild_field_catalog()
        self._render_notes()
        self._update_metrics()
        self.step_notes.configure(text=f"{len(self.notas)} nota(s) pronta(s)")
        tag_count = len(self.discovered_paths)
        message = f"{len(loaded)} nota(s) adicionada(s). {tag_count} caminhos XML disponíveis."
        if errors:
            message += f" {len(errors)} arquivo(s) com erro."
        self.status_label.configure(text=message)
        if loaded:
            self.tabs.set("Campos para exportar")
            self.category_filter.set(self.CATEGORY_XML)
            self._render_fields()

    def _rebuild_field_catalog(self):
        old_values = {field_id: var.get() for field_id, var in self.field_variables.items()}
        catalog: List[Dict[str, str]] = []
        for key, title, *_ in NOTA_COLUMNS:
            if key != "chave":
                catalog.append({"id": f"notas:{key}", "section": "notas", "key": key,
                                "category": self.CATEGORY_NOTE, "label": title,
                                "detail": "Informação consolidada da nota fiscal"})
        for key, title, *_ in ITEM_COLUMNS:
            if key != "chave":
                catalog.append({"id": f"itens:{key}", "section": "itens", "key": key,
                                "category": self.CATEGORY_ITEM, "label": f"Itens - {title}",
                                "detail": "Valores dos produtos; múltiplos itens ficam separados por |"})
        for path in self.discovered_paths:
            label = describe_xml_path(path)
            catalog.append({"id": f"xml:{path}", "section": "xml", "key": path,
                            "category": self.CATEGORY_XML, "label": label,
                            "detail": describe_xml_group(path)})
        self.field_catalog = sorted(catalog, key=lambda field: (normalize_text(field["label"]), normalize_text(field["detail"])))
        self.field_variables = {
            field["id"]: tk.BooleanVar(value=old_values.get(field["id"], False)) for field in self.field_catalog
        }
        self._sync_selected_from_variables()
        if hasattr(self, "field_list"):
            self._render_fields()

    def _schedule_field_search(self, _event=None):
        if self._field_search_job is not None:
            self.after_cancel(self._field_search_job)
        self._field_search_job = self.after(180, self._run_field_search)

    def _run_field_search(self):
        self._field_search_job = None
        self.field_page = 0
        self._render_fields()

    def _category_changed(self, _value=None):
        self.field_page = 0
        self._render_fields()

    def _change_field_page(self, direction: int):
        self.field_page = max(0, self.field_page + direction)
        self._render_fields()

    def _schedule_note_search(self, _event=None):
        if self._note_search_job is not None:
            self.after_cancel(self._note_search_job)
        self._note_search_job = self.after(180, self._run_note_search)

    def _run_note_search(self):
        self._note_search_job = None
        self.note_page = 0
        self._render_notes()

    def _filtered_fields(self):
        query = normalize_text(self.field_search.get().strip()) if hasattr(self, "field_search") else ""
        category = self.category_filter.get() if hasattr(self, "category_filter") else self.CATEGORY_ALL
        result = []
        for field in self.field_catalog:
            if category != self.CATEGORY_ALL and field["category"] != category:
                continue
            searchable = normalize_text(f"{field['label']} {field['detail']} {field['key']}")
            compact_query = "".join(char for char in query if char.isalnum())
            compact_searchable = "".join(char for char in searchable if char.isalnum())
            if query and query not in searchable and compact_query not in compact_searchable:
                continue
            result.append(field)
        return result

    def _render_fields(self):
        if not hasattr(self, "field_list"):
            return
        for child in self.field_list.winfo_children():
            child.destroy()
        fields = self._filtered_fields()
        total_pages = max(1, (len(fields) + self.fields_per_page - 1) // self.fields_per_page)
        self.field_page = min(self.field_page, total_pages - 1)
        start = self.field_page * self.fields_per_page
        page_fields = fields[start:start + self.fields_per_page]
        self.visible_field_ids = [field["id"] for field in page_fields]
        self.available_title.configure(text=f"Campos disponíveis  •  {len(fields)}")
        self.field_page_label.configure(text=f"Página {self.field_page + 1} de {total_pages}")
        self.btn_prev_field_page.configure(state="normal" if self.field_page > 0 else "disabled")
        self.btn_next_field_page.configure(state="normal" if self.field_page + 1 < total_pages else "disabled")
        if not fields:
            ctk.CTkLabel(self.field_list, text="Nenhum campo encontrado.", text_color=COLORS["muted"]).pack(pady=30)
        for index, field in enumerate(page_fields):
            row = ctk.CTkFrame(self.field_list, fg_color=COLORS["panel"], corner_radius=8)
            row.grid(row=index // 2, column=index % 2, sticky="nsew", padx=5, pady=5)
            checkbox = ctk.CTkCheckBox(
                row,
                text=field["label"],
                variable=self.field_variables[field["id"]],
                command=self._selection_changed,
                font=ctk.CTkFont(size=14, weight="bold"),
                checkbox_width=24,
                checkbox_height=24,
            )
            checkbox.pack(anchor="w", padx=12, pady=(12, 3))
            ctk.CTkLabel(row, text=field["detail"], font=ctk.CTkFont(size=11),
                         text_color=COLORS["muted"], wraplength=430, justify="left").pack(
                anchor="w", padx=46, pady=(0, 12)
            )
        self._render_selected()

    def _render_selected(self):
        chosen = [field for field in self.field_catalog if self.field_variables[field["id"]].get()]
        lines = ["1. Chave da NF-e  (obrigatória)"]
        lines.extend(f"{index}. {field['label']}" for index, field in enumerate(chosen, start=2))
        self.selected_list.configure(state="normal")
        self.selected_list.delete("1.0", "end")
        self.selected_list.insert("1.0", "\n\n".join(lines))
        self.selected_list.configure(state="disabled")
        total = 1 + len(chosen)
        self.step_fields.configure(text=f"{total} coluna(s) no Excel")

    def _selection_changed(self):
        self._sync_selected_from_variables()
        self._render_selected()

    def _sync_selected_from_variables(self):
        self.selected = {"notas": [], "itens": [], "xml": [], "ordem": []}
        for field in self.field_catalog:
            variable = self.field_variables.get(field["id"])
            if variable is not None and variable.get():
                self.selected[field["section"]].append(field["key"])
                self.selected["ordem"].append(field["id"])

    def _set_visible(self, checked: bool):
        for field_id in self.visible_field_ids:
            self.field_variables[field_id].set(checked)
        self._selection_changed()
        self._render_fields()

    def _remove_field(self, field_id: str):
        variable = self.field_variables.get(field_id)
        if variable:
            variable.set(False)
            self._selection_changed()
            self._render_fields()

    def _render_notes(self):
        self.notes_tree.delete(*self.notes_tree.get_children())
        query = normalize_text(self.note_search.get().strip())
        filtered = []
        for index, note in enumerate(self.notas):
            searchable = normalize_text(" ".join(str(note.get(key, "")) for key in
                ("numero_nota", "chave", "emit_nome", "dest_nome", "arquivo_origem")))
            if query and query not in searchable:
                continue
            filtered.append((index, note))

        total_pages = max(1, (len(filtered) + self.notes_per_page - 1) // self.notes_per_page)
        self.note_page = min(self.note_page, total_pages - 1)
        start = self.note_page * self.notes_per_page
        end = start + self.notes_per_page
        for index, note in filtered[start:end]:
            self.notes_tree.insert("", "end", iid=str(index), values=(
                note.get("numero_nota", ""), note.get("data_emissao", ""),
                note.get("emit_nome", ""), note.get("dest_nome", ""),
                note.get("quantidade_itens_distintos", 0), f"R$ {note.get('total_nota_fmt', '0,00')}",
                os.path.basename(note.get("arquivo_origem", "")),
            ))
        self.page_label.configure(text=f"Página {self.note_page + 1} de {total_pages}")
        self.btn_prev_page.configure(state="normal" if self.note_page > 0 else "disabled")
        self.btn_next_page.configure(state="normal" if self.note_page + 1 < total_pages else "disabled")

    def _change_note_page(self, direction: int):
        self.note_page = max(0, self.note_page + direction)
        self._render_notes()

    def _open_note_details(self, _event=None):
        item_id = self.notes_tree.focus()
        if item_id:
            try:
                NotaDetailWindow(self, self.notas[int(item_id)])
            except (ValueError, IndexError):
                pass

    def _update_metrics(self):
        item_count = sum(len(note.get("itens", [])) for note in self.notas)
        total_value = sum(note.get("total_nota", 0.0) for note in self.notas)
        tag_count = len(self.discovered_paths)
        self.metric_notes.configure(text=str(len(self.notas)))
        self.metric_items.configure(text=str(item_count))
        self.metric_value.configure(text=f"R$ {format_brazilian_number(total_value)}")
        self.metric_tags.configure(text=str(tag_count))

    def _export(self):
        if not self.notas:
            messagebox.showwarning("Nenhuma nota", "Carregue ao menos um XML antes de exportar.")
            return
        self._sync_selected_from_variables()
        default_name = f"Dados_NFe_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path = filedialog.asksaveasfilename(title="Salvar planilha", defaultextension=".xlsx",
                                            initialfile=default_name, filetypes=[("Planilha Excel", "*.xlsx")])
        if not path:
            return
        selection = {key: list(values) for key, values in self.selected.items()}
        notes_snapshot = list(self.notas)
        self.btn_export.configure(state="disabled", text="Gerando...")
        self.btn_open.configure(state="disabled")
        self.status_label.configure(text=f"Gerando planilha com {len(notes_snapshot)} nota(s)...")
        self.progress.configure(mode="indeterminate")
        self.progress.start()
        threading.Thread(
            target=self._export_worker,
            args=(notes_snapshot, path, selection),
            daemon=True,
        ).start()

    def _export_worker(self, notes, path, selection):
        try:
            generated = export_nfe_to_excel(notes, path, selection)
            self.after(0, self._finish_export, generated, None)
        except Exception as exc:
            self.after(0, self._finish_export, None, str(exc))

    def _finish_export(self, generated_path, error):
        self.progress.stop()
        self.progress.configure(mode="determinate")
        self.progress.set(1 if self.notas else 0)
        self.btn_export.configure(state="normal", text="Exportar planilha")
        if error:
            self.status_label.configure(text="Falha ao gerar a planilha.")
            messagebox.showerror("Erro na exportação", error)
            return
        self.last_excel = generated_path
        self.btn_open.configure(state="normal", fg_color=COLORS["primary"])
        self.step_export.configure(text="Planilha gerada")
        self.status_label.configure(text=f"Arquivo salvo: {os.path.basename(generated_path)}")
        if messagebox.askyesno("Exportação concluída", "A planilha foi criada. Deseja abri-la agora?"):
            self._open_excel()

    def _open_excel(self):
        if not self.last_excel or not os.path.exists(self.last_excel):
            messagebox.showwarning("Arquivo indisponível", "O último Excel gerado não foi encontrado.")
            return
        try:
            os.startfile(self.last_excel)
        except Exception as exc:
            messagebox.showerror("Não foi possível abrir", str(exc))

    def _clear_all(self):
        if not self.notas:
            return
        if not messagebox.askyesno("Limpar sessão", "Remover todas as notas e seleções atuais?"):
            return
        self.notas.clear()
        self.discovered_paths.clear()
        self.note_page = 0
        self.selected = {"notas": [], "itens": [], "xml": [], "ordem": []}
        self.last_excel = None
        self.btn_open.configure(state="disabled", fg_color="#334155")
        self.progress.set(0)
        self._rebuild_field_catalog()
        self._render_notes()
        self._update_metrics()
        self.step_notes.configure(text="Aguardando arquivos")
        self.step_export.configure(text="Pronto após carregar")
        self.status_label.configure(text="Sessão limpa. Carregue novos XMLs.")


def iniciar_aplicacao():
    app = AppNFe()
    app.mainloop()


if __name__ == "__main__":
    iniciar_aplicacao()
