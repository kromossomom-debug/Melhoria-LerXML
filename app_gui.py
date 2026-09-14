# -*- coding: utf-8 -*-
"""
Interface Gráfica Moderna (GUI) para Leitura de XMLs de NF-e e Exportação para Excel.
Desenvolvida com CustomTkinter, oferecendo visual profissional, moderno e intuitivo.
"""

import os
import sys
import threading
from datetime import datetime
from typing import List, Dict, Any, Optional

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import customtkinter as ctk

from nfe_parser import NFeParser, format_brazilian_number
from excel_exporter import export_nfe_to_excel


# Configuração padrão de tema
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class DetalhesNotaModal(ctk.CTkToplevel):
    """Janela modal para exibição detalhada de uma NF-e selecionada."""

    def __init__(self, parent, nota: Dict[str, Any]):
        super().__init__(parent)
        self.nota = nota
        self.title(f"Detalhes da NF-e Nº {nota.get('numero_nota')} - Série {nota.get('serie')}")
        self.geometry("900x650")
        self.minsize(780, 500)
        self.transient(parent)
        self.grab_set()

        self._criar_layout()

    def _criar_layout(self):
        # Header com Chave
        header_frame = ctk.CTkFrame(self, corner_radius=10, fg_color=("gray90", "#1e293b"))
        header_frame.pack(fill="x", padx=16, pady=(16, 8))

        lbl_titulo = ctk.CTkLabel(
            header_frame, 
            text=f"NF-e: {self.nota.get('numero_nota', '')} | Série: {self.nota.get('serie', '')} | Emissão: {self.nota.get('data_emissao', '')}",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        lbl_titulo.pack(anchor="w", padx=16, pady=(10, 4))

        chave_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        chave_frame.pack(fill="x", padx=16, pady=(0, 10))

        lbl_chave = ctk.CTkLabel(
            chave_frame,
            text=f"Chave: {self.nota.get('chave', 'N/A')}",
            font=ctk.CTkFont(family="Consolas", size=12),
            text_color=("gray30", "gray70")
        )
        lbl_chave.pack(side="left")

        btn_copiar = ctk.CTkButton(
            chave_frame,
            text="📋 Copiar Chave",
            width=110,
            height=26,
            font=ctk.CTkFont(size=11),
            command=self._copiar_chave
        )
        btn_copiar.pack(side="right", padx=6)

        # Tabview para organizar detalhes
        self.tabview = ctk.CTkTabview(self, corner_radius=10)
        self.tabview.pack(fill="both", expand=True, padx=16, pady=8)

        tab_geral = self.tabview.add("Geral & Partes")
        tab_transp = self.tabview.add("Transporte & Placa")
        tab_trib = self.tabview.add("Tributos & Totais")
        tab_adic = self.tabview.add("Dados Adicionais")
        tab_itens = self.tabview.add("Itens / Produtos")

        self._build_tab_geral(tab_geral)
        self._build_tab_transp(tab_transp)
        self._build_tab_trib(tab_trib)
        self._build_tab_adic(tab_adic)
        self._build_tab_itens(tab_itens)

        # Botão Fechar no rodapé
        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.pack(fill="x", padx=16, pady=(4, 14))
        btn_fechar = ctk.CTkButton(footer, text="Fechar", width=120, command=self.destroy)
        btn_fechar.pack(side="right")

    def _copiar_chave(self):
        chave = self.nota.get('chave', '')
        self.clipboard_clear()
        self.clipboard_append(chave)
        messagebox.showinfo("Copiado", "Chave de acesso copiada para a área de transferência!")

    def _build_tab_geral(self, parent):
        scroll = ctk.CTkScrollableFrame(parent)
        scroll.pack(fill="both", expand=True, padx=8, pady=8)

        # Emitente
        lbl_sec1 = ctk.CTkLabel(scroll, text="🏢 EMITENTE", font=ctk.CTkFont(size=14, weight="bold"), text_color="#38bdf8")
        lbl_sec1.pack(anchor="w", pady=(4, 2))
        
        info_emit = (
            f"Razão Social: {self.nota.get('emit_nome')}\n"
            f"Nome Fantasia: {self.nota.get('emit_fantasia') or 'N/A'}\n"
            f"CNPJ/CPF: {self.nota.get('emit_cnpj_formatado')}   |   IE: {self.nota.get('emit_ie')}\n"
            f"Endereço: {self.nota.get('emit_endereco')} - {self.nota.get('emit_municipio')}/{self.nota.get('emit_uf')} - CEP: {self.nota.get('emit_cep')}"
        )
        ctk.CTkLabel(scroll, text=info_emit, justify="left", font=ctk.CTkFont(size=12)).pack(anchor="w", padx=10, pady=(0, 10))

        # Destinatário
        lbl_sec2 = ctk.CTkLabel(scroll, text="🎯 DESTINATÁRIO", font=ctk.CTkFont(size=14, weight="bold"), text_color="#38bdf8")
        lbl_sec2.pack(anchor="w", pady=(6, 2))

        info_dest = (
            f"Razão Social: {self.nota.get('dest_nome')}\n"
            f"CNPJ/CPF: {self.nota.get('dest_cnpj_formatado')}   |   IE: {self.nota.get('dest_ie')}\n"
            f"Endereço: {self.nota.get('dest_endereco')} - {self.nota.get('dest_municipio')}/{self.nota.get('dest_uf')} - CEP: {self.nota.get('dest_cep')}"
        )
        ctk.CTkLabel(scroll, text=info_dest, justify="left", font=ctk.CTkFont(size=12)).pack(anchor="w", padx=10, pady=(0, 10))

        # Operação
        lbl_sec3 = ctk.CTkLabel(scroll, text="📄 OPERAÇÃO", font=ctk.CTkFont(size=14, weight="bold"), text_color="#38bdf8")
        lbl_sec3.pack(anchor="w", pady=(6, 2))
        info_op = (
            f"Natureza da Operação: {self.nota.get('natureza_operacao')}\n"
            f"Tipo de Operação: {self.nota.get('tipo_operacao')}   |   Modelo: {self.nota.get('modelo')}\n"
            f"Arquivo de Origem: {self.nota.get('arquivo_origem')}"
        )
        ctk.CTkLabel(scroll, text=info_op, justify="left", font=ctk.CTkFont(size=12)).pack(anchor="w", padx=10, pady=(0, 10))

    def _build_tab_transp(self, parent):
        scroll = ctk.CTkScrollableFrame(parent)
        scroll.pack(fill="both", expand=True, padx=8, pady=8)

        # Transportadora
        lbl_t1 = ctk.CTkLabel(scroll, text="🚚 TRANSPORTADORA & FRETE", font=ctk.CTkFont(size=14, weight="bold"), text_color="#38bdf8")
        lbl_t1.pack(anchor="w", pady=(4, 2))
        info_transp = (
            f"Modalidade do Frete: {self.nota.get('modalidade_frete')}\n"
            f"Nome/Razão Social: {self.nota.get('transp_nome') or 'N/A'}\n"
            f"CNPJ/CPF: {self.nota.get('transp_cnpj_formatado') or 'N/A'}   |   IE: {self.nota.get('transp_ie') or 'N/A'}\n"
            f"Município/UF: {self.nota.get('transp_municipio') or 'N/A'} / {self.nota.get('transp_uf') or 'N/A'}"
        )
        ctk.CTkLabel(scroll, text=info_transp, justify="left", font=ctk.CTkFont(size=12)).pack(anchor="w", padx=10, pady=(0, 12))

        # Veículo e Placa
        lbl_t2 = ctk.CTkLabel(scroll, text="🚛 VEÍCULO & PLACA DO CAMINHÃO", font=ctk.CTkFont(size=14, weight="bold"), text_color="#38bdf8")
        lbl_t2.pack(anchor="w", pady=(6, 2))
        placa = self.nota.get('placa_veiculo') or 'Não informada'
        placa_uf = self.nota.get('placa_uf') or '-'
        rntrc = self.nota.get('rntrc') or 'N/A'
        info_veic = (
            f"Placa do Caminhão: {placa}   |   UF da Placa: {placa_uf}\n"
            f"RNTRC: {rntrc}"
        )
        ctk.CTkLabel(scroll, text=info_veic, justify="left", font=ctk.CTkFont(size=13, weight="bold" if placa != 'Não informada' else "normal")).pack(anchor="w", padx=10, pady=(0, 12))

        # Volumes e Pesos
        lbl_t3 = ctk.CTkLabel(scroll, text="📦 VOLUMES E PESOS", font=ctk.CTkFont(size=14, weight="bold"), text_color="#38bdf8")
        lbl_t3.pack(anchor="w", pady=(6, 2))
        info_vol = (
            f"Quantidade de Volumes: {self.nota.get('quantidade_volumes')}   |   Espécie: {self.nota.get('especie_volumes') or 'N/A'}\n"
            f"Marca: {self.nota.get('marca_volumes') or 'N/A'}\n"
            f"Peso Líquido: {self.nota.get('peso_liquido_fmt')} kg\n"
            f"Peso Bruto: {self.nota.get('peso_bruto_fmt')} kg"
        )
        ctk.CTkLabel(scroll, text=info_vol, justify="left", font=ctk.CTkFont(size=12)).pack(anchor="w", padx=10, pady=(0, 6))

    def _build_tab_trib(self, parent):
        scroll = ctk.CTkScrollableFrame(parent)
        scroll.pack(fill="both", expand=True, padx=8, pady=8)

        # Totais da Nota
        lbl_tot = ctk.CTkLabel(scroll, text="💰 VALORES TOTAIS DA NF-e", font=ctk.CTkFont(size=14, weight="bold"), text_color="#10b981")
        lbl_tot.pack(anchor="w", pady=(4, 2))
        info_tot = (
            f"Valor dos Produtos: R$ {format_brazilian_number(self.nota.get('total_produtos'))}\n"
            f"Valor do Frete: R$ {format_brazilian_number(self.nota.get('total_frete'))}   |   Seguro: R$ {format_brazilian_number(self.nota.get('total_seguro'))}\n"
            f"Desconto: R$ {format_brazilian_number(self.nota.get('total_desconto'))}   |   Outras Despesas: R$ {format_brazilian_number(self.nota.get('outras_despesas'))}\n"
            f"VALOR TOTAL DA NOTA: R$ {self.nota.get('total_nota_fmt')}\n"
            f"Quantidade Total de Itens: {self.nota.get('quantidade_total_itens_fmt')}"
        )
        ctk.CTkLabel(scroll, text=info_tot, justify="left", font=ctk.CTkFont(size=12)).pack(anchor="w", padx=10, pady=(0, 12))

        # Detalhamento de Tributos
        lbl_trib = ctk.CTkLabel(scroll, text="🏛️ DISCRIMINAÇÃO DE TRIBUTOS", font=ctk.CTkFont(size=14, weight="bold"), text_color="#f59e0b")
        lbl_trib.pack(anchor="w", pady=(6, 2))
        info_trib = (
            f"Base de Cálculo ICMS: R$ {format_brazilian_number(self.nota.get('bc_icms'))}   ->   Valor ICMS: R$ {format_brazilian_number(self.nota.get('valor_icms'))}\n"
            f"Base de Cálculo ICMS ST: R$ {format_brazilian_number(self.nota.get('bc_icms_st'))}   ->   Valor ICMS ST: R$ {format_brazilian_number(self.nota.get('valor_icms_st'))}\n"
            f"Valor IPI: R$ {format_brazilian_number(self.nota.get('valor_ipi'))}\n"
            f"Valor PIS: R$ {format_brazilian_number(self.nota.get('valor_pis'))}   |   Valor COFINS: R$ {format_brazilian_number(self.nota.get('valor_cofins'))}\n"
            f"TOTAL DE IMPOSTOS CALCULADOS: R$ {self.nota.get('total_impostos_fmt')}\n"
            f"Tributos Aproximados (Lei da Transparência / IBPT): R$ {format_brazilian_number(self.nota.get('total_tributos_aproximado'))}"
        )
        ctk.CTkLabel(scroll, text=info_trib, justify="left", font=ctk.CTkFont(size=12)).pack(anchor="w", padx=10, pady=(0, 6))

    def _build_tab_adic(self, parent):
        scroll = ctk.CTkScrollableFrame(parent)
        scroll.pack(fill="both", expand=True, padx=8, pady=8)

        lbl_cpl = ctk.CTkLabel(scroll, text="📝 INFORMAÇÕES COMPLEMENTARES (infCpl)", font=ctk.CTkFont(size=14, weight="bold"), text_color="#38bdf8")
        lbl_cpl.pack(anchor="w", pady=(4, 2))
        cpl_text = self.nota.get('informacoes_complementares') or "(Nenhuma informação complementar informada)"
        
        box_cpl = ctk.CTkTextbox(scroll, height=140, wrap="word", font=ctk.CTkFont(size=12))
        box_cpl.pack(fill="x", padx=6, pady=(0, 14))
        box_cpl.insert("1.0", cpl_text)
        box_cpl.configure(state="disabled")

        lbl_fisco = ctk.CTkLabel(scroll, text="⚖️ INFORMAÇÕES ADICIONAIS DO FISCO (infAdFisco)", font=ctk.CTkFont(size=14, weight="bold"), text_color="#38bdf8")
        lbl_fisco.pack(anchor="w", pady=(6, 2))
        fisco_text = self.nota.get('informacoes_adicionais_fisco') or "(Nenhuma informação do fisco informada)"

        box_fisco = ctk.CTkTextbox(scroll, height=100, wrap="word", font=ctk.CTkFont(size=12))
        box_fisco.pack(fill="x", padx=6, pady=(0, 6))
        box_fisco.insert("1.0", fisco_text)
        box_fisco.configure(state="disabled")

    def _build_tab_itens(self, parent):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=4, pady=4)

        itens = self.nota.get('itens', [])
        cols = ("n_item", "c_prod", "x_prod", "ncm", "cfop", "u_com", "q_com", "v_un", "v_prod")

        tree = ttk.Treeview(frame, columns=cols, show="headings", height=12)
        tree.heading("n_item", text="Item")
        tree.heading("c_prod", text="Código")
        tree.heading("x_prod", text="Descrição do Produto")
        tree.heading("ncm", text="NCM")
        tree.heading("cfop", text="CFOP")
        tree.heading("u_com", text="Un.")
        tree.heading("q_com", text="Qtd (BR)")
        tree.heading("v_un", text="Vl. Unit")
        tree.heading("v_prod", text="Vl. Total (R$)")

        tree.column("n_item", width=45, anchor="center")
        tree.column("c_prod", width=100, anchor="w")
        tree.column("x_prod", width=260, anchor="w")
        tree.column("ncm", width=80, anchor="center")
        tree.column("cfop", width=60, anchor="center")
        tree.column("u_com", width=50, anchor="center")
        tree.column("q_com", width=90, anchor="e")
        tree.column("v_un", width=90, anchor="e")
        tree.column("v_prod", width=100, anchor="e")

        for item in itens:
            tree.insert("", "end", values=(
                item.get('n_item'),
                item.get('c_prod'),
                item.get('x_prod'),
                item.get('ncm'),
                item.get('cfop'),
                item.get('u_com'),
                item.get('q_com_fmt'),
                item.get('v_un_com_fmt'),
                item.get('v_prod_fmt')
            ))

        scroll_y = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scroll_y.set)
        tree.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")


class AppNFe(ctk.CTk):
    """Janela Principal do Aplicativo."""

    def __init__(self):
        super().__init__()

        self.title("Leitor de XML NF-e Pro - LDC")
        self.geometry("1280x760")
        self.minsize(1050, 620)

        # Estado da aplicação
        self.notas_carregadas: List[Dict[str, Any]] = []
        self.ultimo_excel_salvo: Optional[str] = None
        self.is_loading: bool = False

        self._configurar_estilo_tabela()
        self._criar_layout()

    def _configurar_estilo_tabela(self):
        """Ajusta estilo do Treeview para se integrar com CustomTkinter."""
        style = ttk.Style()
        style.theme_use("clam")
        
        # Configuração para tema escuro por padrão
        style.configure(
            "Treeview",
            background="#1e293b",
            foreground="#f8fafc",
            fieldbackground="#1e293b",
            rowheight=30,
            font=("Segoe UI", 10),
            borderwidth=0
        )
        style.configure(
            "Treeview.Heading",
            background="#0f172a",
            foreground="#93c5fd",
            font=("Segoe UI", 10, "bold"),
            borderwidth=0,
            relief="flat",
            padding=(6, 6)
        )
        style.map(
            "Treeview",
            background=[("selected", "#2563eb")],
            foreground=[("selected", "#ffffff")]
        )
        style.map(
            "Treeview.Heading",
            background=[("active", "#1e3a8a")]
        )

    def _criar_layout(self):
        # 1. BARRA SUPERIOR (HEADER)
        self.header_frame = ctk.CTkFrame(self, height=70, corner_radius=0, fg_color=("gray90", "#0f172a"))
        self.header_frame.pack(fill="x", side="top")

        # Logo e Título
        title_box = ctk.CTkFrame(self.header_frame, fg_color="transparent")
        title_box.pack(side="left", padx=20, pady=12)

        lbl_app = ctk.CTkLabel(
            title_box,
            text="📑 Leitor de XML NF-e Pro",
            font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"),
            text_color=("gray10", "#38bdf8")
        )
        lbl_app.pack(anchor="w")

        lbl_sub = ctk.CTkLabel(
            title_box,
            text="Extração Completa de Dados Fiscais, Logística, Tributos e Exportação Excel",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=("gray40", "gray60")
        )
        lbl_sub.pack(anchor="w")

        # Controles Direita: Tema
        theme_box = ctk.CTkFrame(self.header_frame, fg_color="transparent")
        theme_box.pack(side="right", padx=20, pady=12)

        ctk.CTkLabel(theme_box, text="Tema:", font=ctk.CTkFont(size=12)).pack(side="left", padx=(0, 6))
        self.theme_option = ctk.CTkOptionMenu(
            theme_box,
            values=["Dark", "Light", "System"],
            width=90,
            height=28,
            command=self._alterar_tema
        )
        self.theme_option.set("Dark")
        self.theme_option.pack(side="left")

        # 2. CARDS DE RESUMO (MÉTRICAS RÁPIDAS)
        self.cards_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.cards_frame.pack(fill="x", padx=20, pady=(14, 8))
        self.cards_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self.card_notas = self._criar_card_metrica(self.cards_frame, 0, "Notas Carregadas", "0", "📄", "#38bdf8")
        self.card_qtd = self._criar_card_metrica(self.cards_frame, 1, "Qtd Total Itens", "0,00", "📦", "#818cf8")
        self.card_total = self._criar_card_metrica(self.cards_frame, 2, "Valor Total (R$)", "R$ 0,00", "💰", "#34d399")
        self.card_impostos = self._criar_card_metrica(self.cards_frame, 3, "Total Tributos (R$)", "R$ 0,00", "🏛️", "#f59e0b")

        # 3. BARRA DE FERRAMENTAS / AÇÕES
        self.toolbar_frame = ctk.CTkFrame(self, corner_radius=10, fg_color=("gray95", "#1e293b"))
        self.toolbar_frame.pack(fill="x", padx=20, pady=6)

        # Botões de Carregamento
        btn_box = ctk.CTkFrame(self.toolbar_frame, fg_color="transparent")
        btn_box.pack(side="left", padx=12, pady=10)

        self.btn_carregar_arquivos = ctk.CTkButton(
            btn_box,
            text="📁 Selecionar Arquivo(s)",
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            height=34,
            command=self._carregar_arquivos_dialog
        )
        self.btn_carregar_arquivos.pack(side="left", padx=4)

        self.btn_carregar_pasta = ctk.CTkButton(
            btn_box,
            text="📂 Selecionar Pasta com XMLs",
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#0d9488",
            hover_color="#0f766e",
            height=34,
            command=self._carregar_pasta_dialog
        )
        self.btn_carregar_pasta.pack(side="left", padx=4)

        self.btn_limpar = ctk.CTkButton(
            btn_box,
            text="🗑️ Limpar",
            font=ctk.CTkFont(size=12),
            fg_color=("gray75", "#334155"),
            hover_color=("gray65", "#475569"),
            width=80,
            height=34,
            command=self._limpar_lista
        )
        self.btn_limpar.pack(side="left", padx=4)

        # Barra de Pesquisa / Filtro
        search_box = ctk.CTkFrame(self.toolbar_frame, fg_color="transparent")
        search_box.pack(side="right", padx=12, pady=10)

        self.search_entry = ctk.CTkEntry(
            search_box,
            placeholder_text="🔍 Filtrar por Nº, Razão Social, Placa...",
            width=280,
            height=34
        )
        self.search_entry.pack(side="left", padx=4)
        self.search_entry.bind("<KeyRelease>", lambda e: self._filtrar_tabela())

        # 4. TABELA PRINCIPAL
        self.table_container = ctk.CTkFrame(self, corner_radius=10, fg_color=("gray95", "#1e293b"))
        self.table_container.pack(fill="both", expand=True, padx=20, pady=6)

        colunas = (
            "num_nf", "serie", "data_emi", "emitente", "destinatario", 
            "transp", "placa", "qtd_total", "valor_total", "tributos"
        )
        self.tree = ttk.Treeview(self.table_container, columns=colunas, show="headings", selectmode="browse")
        
        self.tree.heading("num_nf", text="Nº NF")
        self.tree.heading("serie", text="Série")
        self.tree.heading("data_emi", text="Emissão")
        self.tree.heading("emitente", text="Emitente")
        self.tree.heading("destinatario", text="Destinatário")
        self.tree.heading("transp", text="Transportadora")
        self.tree.heading("placa", text="Placa")
        self.tree.heading("qtd_total", text="Qtd Total (BR)")
        self.tree.heading("valor_total", text="Valor Total (R$)")
        self.tree.heading("tributos", text="Tributos (R$)")

        self.tree.column("num_nf", width=75, anchor="center")
        self.tree.column("serie", width=50, anchor="center")
        self.tree.column("data_emi", width=130, anchor="center")
        self.tree.column("emitente", width=210, anchor="w")
        self.tree.column("destinatario", width=210, anchor="w")
        self.tree.column("transp", width=180, anchor="w")
        self.tree.column("placa", width=85, anchor="center")
        self.tree.column("qtd_total", width=95, anchor="e")
        self.tree.column("valor_total", width=115, anchor="e")
        self.tree.column("tributos", width=105, anchor="e")

        self.tree.bind("<Double-1>", self._on_double_click_row)

        scroll_y = ttk.Scrollbar(self.table_container, orient="vertical", command=self.tree.yview)
        scroll_x = ttk.Scrollbar(self.table_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self.tree.grid(row=0, column=0, sticky="nsew", padx=(10, 0), pady=(10, 0))
        scroll_y.grid(row=0, column=1, sticky="ns", padx=(0, 10), pady=(10, 0))
        scroll_x.grid(row=1, column=0, sticky="ew", padx=(10, 0), pady=(0, 10))

        self.table_container.grid_rowconfigure(0, weight=1)
        self.table_container.grid_columnconfigure(0, weight=1)

        # Dica abaixo da tabela
        lbl_hint = ctk.CTkLabel(
            self,
            text="💡 Dica: Dê um duplo-clique em qualquer linha para ver todos os detalhes da nota (itens, tributos, dados adicionais e placa).",
            font=ctk.CTkFont(size=11),
            text_color=("gray40", "gray50")
        )
        lbl_hint.pack(anchor="w", padx=24, pady=(2, 2))

        # 5. BARRA INFERIOR / RODAPÉ COM EXPORTAÇÃO
        self.footer_frame = ctk.CTkFrame(self, height=65, corner_radius=10, fg_color=("gray95", "#0f172a"))
        self.footer_frame.pack(fill="x", padx=20, pady=(4, 14))

        # Status e Progresso à esquerda
        status_box = ctk.CTkFrame(self.footer_frame, fg_color="transparent")
        status_box.pack(side="left", padx=16, pady=10)

        self.lbl_status = ctk.CTkLabel(
            status_box,
            text="Pronto para carregar arquivos XML.",
            font=ctk.CTkFont(size=12),
            text_color=("gray30", "gray70")
        )
        self.lbl_status.pack(anchor="w")

        self.progresso = ctk.CTkProgressBar(status_box, width=280, height=8)
        self.progresso.set(0)
        self.progresso.pack(anchor="w", pady=(4, 0))

        # Botões de Exportação à direita
        actions_box = ctk.CTkFrame(self.footer_frame, fg_color="transparent")
        actions_box.pack(side="right", padx=16, pady=10)

        # Botão Abrir Excel (inicia desabilitado e ganha destaque após exportação)
        self.btn_abrir_excel = ctk.CTkButton(
            actions_box,
            text="🚀 Abrir Arquivo Excel",
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#059669",
            hover_color="#047857",
            height=38,
            state="disabled",
            command=self._abrir_excel_salvo
        )
        self.btn_abrir_excel.pack(side="right", padx=(8, 0))

        # Botão Exportar
        self.btn_exportar = ctk.CTkButton(
            actions_box,
            text="📊 Exportar para Excel (.xlsx)",
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#10b981",
            hover_color="#059669",
            height=38,
            command=self._exportar_excel_dialog
        )
        self.btn_exportar.pack(side="right", padx=4)

    def _criar_card_metrica(self, parent, col, titulo, valor_inicial, icone, cor_icone):
        card = ctk.CTkFrame(parent, corner_radius=10, fg_color=("gray95", "#1e293b"))
        card.grid(row=0, column=col, padx=6, pady=2, sticky="nsew")

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(padx=14, pady=10, fill="both")

        top_row = ctk.CTkFrame(inner, fg_color="transparent")
        top_row.pack(fill="x")

        lbl_tit = ctk.CTkLabel(top_row, text=titulo, font=ctk.CTkFont(size=11, weight="bold"), text_color=("gray40", "gray60"))
        lbl_tit.pack(side="left")

        lbl_ico = ctk.CTkLabel(top_row, text=icone, font=ctk.CTkFont(size=14))
        lbl_ico.pack(side="right")

        lbl_val = ctk.CTkLabel(inner, text=valor_inicial, font=ctk.CTkFont(size=18, weight="bold"), text_color=cor_icone)
        lbl_val.pack(anchor="w", pady=(4, 0))

        return lbl_val

    def _alterar_tema(self, novo_tema: str):
        ctk.set_appearance_mode(novo_tema)
        # Ajusta cores do Treeview
        style = ttk.Style()
        if novo_tema.lower() == "light":
            style.configure("Treeview", background="#ffffff", foreground="#0f172a", fieldbackground="#ffffff")
            style.configure("Treeview.Heading", background="#e2e8f0", foreground="#1e3a8a")
            style.map("Treeview", background=[("selected", "#3b82f6")], foreground=[("selected", "#ffffff")])
        else:
            style.configure("Treeview", background="#1e293b", foreground="#f8fafc", fieldbackground="#1e293b")
            style.configure("Treeview.Heading", background="#0f172a", foreground="#93c5fd")
            style.map("Treeview", background=[("selected", "#2563eb")], foreground=[("selected", "#ffffff")])

    def _carregar_arquivos_dialog(self):
        if self.is_loading:
            return
        arquivos = filedialog.askopenfilenames(
            title="Selecione os Arquivos XML de NF-e",
            filetypes=[("Arquivos XML", "*.xml"), ("Todos os Arquivos", "*.*")]
        )
        if arquivos:
            self._iniciar_leitura_thread(list(arquivos))

    def _carregar_pasta_dialog(self):
        if self.is_loading:
            return
        pasta = filedialog.askdirectory(title="Selecione a Pasta com Arquivos XML")
        if pasta:
            arquivos = []
            for root_dir, _, files in os.walk(pasta):
                for f in files:
                    if f.lower().endswith('.xml'):
                        arquivos.append(os.path.join(root_dir, f))
            if not arquivos:
                messagebox.showwarning("Aviso", f"Nenhum arquivo .xml encontrado na pasta selecionada.")
                return
            self._iniciar_leitura_thread(arquivos)

    def _iniciar_leitura_thread(self, lista_arquivos: List[str]):
        self.is_loading = True
        self.btn_carregar_arquivos.configure(state="disabled")
        self.btn_carregar_pasta.configure(state="disabled")
        self.btn_limpar.configure(state="disabled")
        self.progresso.set(0)
        self.lbl_status.configure(text=f"Processando {len(lista_arquivos)} arquivo(s)...")

        thread = threading.Thread(target=self._processar_arquivos_background, args=(lista_arquivos,), daemon=True)
        thread.start()

    def _processar_arquivos_background(self, lista_arquivos: List[str]):
        total = len(lista_arquivos)
        erros = 0
        sucessos = 0

        for idx, file_path in enumerate(lista_arquivos, start=1):
            dados = NFeParser.parse_file(file_path)
            if dados.get('sucesso'):
                # Evitar duplicar a mesma chave se já carregada
                chave = dados.get('chave')
                ja_existe = any(n.get('chave') == chave and chave != '' for n in self.notas_carregadas)
                if not ja_existe:
                    self.notas_carregadas.append(dados)
                    sucessos += 1
            else:
                erros += 1

            # Atualizar progresso na UI
            prog = idx / total
            self.after(0, self._atualizar_progresso_ui, prog, idx, total)

        self.after(0, self._finalizar_leitura_ui, sucessos, erros)

    def _atualizar_progresso_ui(self, prog: float, idx: int, total: int):
        self.progresso.set(prog)
        self.lbl_status.configure(text=f"Processando: {idx}/{total} XMLs lidos...")

    def _finalizar_leitura_ui(self, sucessos: int, erros: int):
        self.is_loading = False
        self.btn_carregar_arquivos.configure(state="normal")
        self.btn_carregar_pasta.configure(state="normal")
        self.btn_limpar.configure(state="normal")
        self.progresso.set(1.0)

        msg = f"Concluído! {sucessos} nova(s) nota(s) adicionada(s)."
        if erros > 0:
            msg += f" ({erros} com erro ou não-NFe)"
        self.lbl_status.configure(text=msg)

        self._repopular_tabela()
        self._atualizar_metricas()

    def _repopular_tabela(self):
        self.tree.delete(*self.tree.get_children())
        for idx, nf in enumerate(self.notas_carregadas):
            self.tree.insert(
                "",
                "end",
                iid=str(idx),
                values=(
                    nf.get('numero_nota', ''),
                    nf.get('serie', ''),
                    nf.get('data_emissao', ''),
                    nf.get('emit_nome', ''),
                    nf.get('dest_nome', ''),
                    nf.get('transp_nome', '') or 'Não informada',
                    nf.get('placa_veiculo', '') or '-',
                    nf.get('quantidade_total_itens_fmt', '0,00'),
                    f"R$ {nf.get('total_nota_fmt', '0,00')}",
                    f"R$ {nf.get('total_impostos_fmt', '0,00')}"
                )
            )

    def _filtrar_tabela(self):
        termo = self.search_entry.get().strip().lower()
        self.tree.delete(*self.tree.get_children())

        for idx, nf in enumerate(self.notas_carregadas):
            if not termo:
                match = True
            else:
                texto_busca = " ".join([
                    str(nf.get('numero_nota', '')),
                    str(nf.get('chave', '')),
                    str(nf.get('emit_nome', '')),
                    str(nf.get('emit_cnpj', '')),
                    str(nf.get('dest_nome', '')),
                    str(nf.get('dest_cnpj', '')),
                    str(nf.get('transp_nome', '')),
                    str(nf.get('placa_veiculo', ''))
                ]).lower()
                match = termo in texto_busca

            if match:
                self.tree.insert(
                    "",
                    "end",
                    iid=str(idx),
                    values=(
                        nf.get('numero_nota', ''),
                        nf.get('serie', ''),
                        nf.get('data_emissao', ''),
                        nf.get('emit_nome', ''),
                        nf.get('dest_nome', ''),
                        nf.get('transp_nome', '') or 'Não informada',
                        nf.get('placa_veiculo', '') or '-',
                        nf.get('quantidade_total_itens_fmt', '0,00'),
                        f"R$ {nf.get('total_nota_fmt', '0,00')}",
                        f"R$ {nf.get('total_impostos_fmt', '0,00')}"
                    )
                )

    def _atualizar_metricas(self):
        total_notas = len(self.notas_carregadas)
        soma_qtd = sum(n.get('quantidade_total_itens', 0.0) for n in self.notas_carregadas)
        soma_valor = sum(n.get('total_nota', 0.0) for n in self.notas_carregadas)
        soma_impostos = sum(n.get('total_impostos', 0.0) for n in self.notas_carregadas)

        self.card_notas.configure(text=str(total_notas))
        self.card_qtd.configure(text=format_brazilian_number(soma_qtd, 2))
        self.card_total.configure(text=f"R$ {format_brazilian_number(soma_valor, 2)}")
        self.card_impostos.configure(text=f"R$ {format_brazilian_number(soma_impostos, 2)}")

    def _limpar_lista(self):
        if not self.notas_carregadas:
            return
        if messagebox.askyesno("Limpar", "Deseja remover todas as notas da lista?"):
            self.notas_carregadas.clear()
            self._repopular_tabela()
            self._atualizar_metricas()
            self.lbl_status.configure(text="Lista limpa.")
            self.progresso.set(0)
            self.btn_abrir_excel.configure(state="disabled", text="🚀 Abrir Arquivo Excel")

    def _on_double_click_row(self, event):
        item_id = self.tree.focus()
        if not item_id:
            return
        try:
            idx = int(item_id)
            if 0 <= idx < len(self.notas_carregadas):
                nota = self.notas_carregadas[idx]
                DetalhesNotaModal(self, nota)
        except (ValueError, IndexError):
            pass

    def _exportar_excel_dialog(self):
        if not self.notas_carregadas:
            messagebox.showwarning("Aviso", "Nenhuma nota carregada para exportar. Selecione arquivos XML primeiro!")
            return

        data_hora_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        sugestao_nome = f"Relatorio_NFe_{data_hora_str}.xlsx"

        caminho_salvar = filedialog.asksaveasfilename(
            title="Salvar Relatório Excel",
            defaultextension=".xlsx",
            initialfile=sugestao_nome,
            filetypes=[("Planilha Excel", "*.xlsx")]
        )

        if not caminho_salvar:
            return

        try:
            self.lbl_status.configure(text="Gerando planilha Excel...")
            self.update_idletasks()

            caminho_gerado = export_nfe_to_excel(self.notas_carregadas, caminho_salvar)
            self.ultimo_excel_salvo = caminho_gerado

            # Habilitar botão de abrir Excel imediatamente
            nome_arquivo = os.path.basename(caminho_gerado)
            self.btn_abrir_excel.configure(
                state="normal",
                text=f"🚀 Abrir Excel ({nome_arquivo})"
            )
            self.lbl_status.configure(text=f"Salvo com sucesso: {nome_arquivo}")

            # Perguntar se já quer abrir agora
            resposta = messagebox.askyesno(
                "Exportação Concluída com Sucesso!",
                f"A planilha foi gerada com sucesso em:\n{caminho_gerado}\n\nDeseja abrir o arquivo no Excel agora?"
            )
            if resposta:
                self._abrir_excel_salvo()

        except Exception as e:
            messagebox.showerror("Erro na Exportação", f"Ocorreu um erro ao salvar a planilha:\n{str(e)}")
            self.lbl_status.configure(text="Erro ao exportar.")

    def _abrir_excel_salvo(self):
        if not self.ultimo_excel_salvo or not os.path.exists(self.ultimo_excel_salvo):
            messagebox.showwarning("Aviso", "O arquivo Excel gerado não foi encontrado.")
            return

        try:
            # Abertura nativa no Windows
            os.startfile(self.ultimo_excel_salvo)
        except Exception as e:
            messagebox.showerror("Erro ao Abrir", f"Não foi possível abrir o arquivo no Excel automaticamente:\n{str(e)}")


def iniciar_aplicacao():
    app = AppNFe()
    app.mainloop()


if __name__ == "__main__":
    iniciar_aplicacao()
