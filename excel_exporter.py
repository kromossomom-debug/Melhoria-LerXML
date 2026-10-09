# -*- coding: utf-8 -*-
"""
Módulo de exportação de dados de NF-e para planilha Excel (.xlsx) com OpenPyXL.
Inclui formatação profissional corporativa, abas 'Notas Fiscais' e 'Itens das Notas',
zebrado, auto-filtro, formatos numéricos brasileiros e ajuste inteligente de colunas.
"""

from typing import List, Dict, Any, Optional, Set
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# Estilos Visuais Profissionais
COLOR_HEADER_BG = "1E3A8A"       # Azul marinho elegante
COLOR_HEADER_TEXT = "FFFFFF"     # Texto branco
COLOR_ZEBRA_ROW = "F1F5F9"       # Cinza ardósia claro para linhas alternadas
COLOR_BORDER = "CBD5E1"          # Borda cinza suave
FONT_NAME = "Segoe UI"

FONT_HEADER = Font(name=FONT_NAME, size=11, bold=True, color=COLOR_HEADER_TEXT)
FONT_BODY = Font(name=FONT_NAME, size=10)
FONT_TOTAL = Font(name=FONT_NAME, size=11, bold=True)

FILL_HEADER = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
FILL_ZEBRA = PatternFill(start_color=COLOR_ZEBRA_ROW, end_color=COLOR_ZEBRA_ROW, fill_type="solid")
FILL_WHITE = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

BORDER_THIN = Border(
    left=Side(style='thin', color=COLOR_BORDER),
    right=Side(style='thin', color=COLOR_BORDER),
    top=Side(style='thin', color=COLOR_BORDER),
    bottom=Side(style='thin', color=COLOR_BORDER)
)

ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")
ALIGN_HEADER = Alignment(horizontal="center", vertical="center", wrap_text=True)

FMT_CURRENCY = '"R$ "#,##0.00'
FMT_NUMBER_2DEC = '#,##0.00'
FMT_NUMBER_3DEC = '#,##0.000'
FMT_NUMBER_INT = '#,##0'
FMT_TEXT = '@'


# Chave interna, titulo exibido no Excel, formato, alinhamento e largura.
# Estas listas tambem alimentam a janela de selecao de campos da interface.
NOTA_COLUMNS = [
    ("chave", "Chave de Acesso", FMT_TEXT, ALIGN_CENTER, 46),
    ("numero_nota", "Número NF", FMT_NUMBER_INT, ALIGN_CENTER, 14),
    ("serie", "Série", FMT_TEXT, ALIGN_CENTER, 10),
    ("data_emissao", "Data Emissão", FMT_TEXT, ALIGN_CENTER, 20),
    ("data_saida", "Data Saída/Entrada", FMT_TEXT, ALIGN_CENTER, 20),
    ("tipo_operacao", "Tipo Operação", FMT_TEXT, ALIGN_CENTER, 16),
    ("natureza_operacao", "Natureza Operação", FMT_TEXT, ALIGN_LEFT, 28),
    ("emit_nome", "Emitente (Razão Social)", FMT_TEXT, ALIGN_LEFT, 32),
    ("emit_cnpj_formatado", "Emitente CNPJ/CPF", FMT_TEXT, ALIGN_CENTER, 20),
    ("emit_ie", "Emitente IE", FMT_TEXT, ALIGN_CENTER, 16),
    ("emit_municipio", "Emitente Município", FMT_TEXT, ALIGN_LEFT, 20),
    ("emit_uf", "Emitente UF", FMT_TEXT, ALIGN_CENTER, 10),
    ("dest_nome", "Destinatário (Razão Social)", FMT_TEXT, ALIGN_LEFT, 32),
    ("dest_cnpj_formatado", "Destinatário CNPJ/CPF", FMT_TEXT, ALIGN_CENTER, 20),
    ("dest_ie", "Destinatário IE", FMT_TEXT, ALIGN_CENTER, 16),
    ("dest_municipio", "Destinatário Município", FMT_TEXT, ALIGN_LEFT, 20),
    ("dest_uf", "Destinatário UF", FMT_TEXT, ALIGN_CENTER, 10),
    ("modalidade_frete", "Modalidade Frete", FMT_TEXT, ALIGN_LEFT, 26),
    ("transp_nome", "Transportadora", FMT_TEXT, ALIGN_LEFT, 30),
    ("transp_cnpj_formatado", "Transportadora CNPJ/CPF", FMT_TEXT, ALIGN_CENTER, 20),
    ("transp_ie", "Transportadora IE", FMT_TEXT, ALIGN_CENTER, 16),
    ("placa_veiculo", "Placa do Caminhão", FMT_TEXT, ALIGN_CENTER, 18),
    ("placa_uf", "UF Placa", FMT_TEXT, ALIGN_CENTER, 10),
    ("rntrc", "RNTRC", FMT_TEXT, ALIGN_CENTER, 14),
    ("quantidade_volumes", "Qtd Volumes", FMT_NUMBER_INT, ALIGN_RIGHT, 14),
    ("especie_volumes", "Espécie Volumes", FMT_TEXT, ALIGN_LEFT, 16),
    ("peso_liquido", "Peso Líquido (kg)", FMT_NUMBER_3DEC, ALIGN_RIGHT, 18),
    ("peso_bruto", "Peso Bruto (kg)", FMT_NUMBER_3DEC, ALIGN_RIGHT, 18),
    ("quantidade_total_itens", "Qtd Total Itens", FMT_NUMBER_2DEC, ALIGN_RIGHT, 18),
    ("total_produtos", "Valor Produtos", FMT_CURRENCY, ALIGN_RIGHT, 18),
    ("total_frete", "Valor Frete", FMT_CURRENCY, ALIGN_RIGHT, 16),
    ("total_seguro", "Valor Seguro", FMT_CURRENCY, ALIGN_RIGHT, 16),
    ("total_desconto", "Valor Desconto", FMT_CURRENCY, ALIGN_RIGHT, 16),
    ("outras_despesas", "Outras Despesas", FMT_CURRENCY, ALIGN_RIGHT, 16),
    ("total_nota", "Valor Total da NF", FMT_CURRENCY, ALIGN_RIGHT, 20),
    ("bc_icms", "Base ICMS", FMT_CURRENCY, ALIGN_RIGHT, 16),
    ("valor_icms", "Valor ICMS", FMT_CURRENCY, ALIGN_RIGHT, 16),
    ("bc_icms_st", "Base ICMS ST", FMT_CURRENCY, ALIGN_RIGHT, 16),
    ("valor_icms_st", "Valor ICMS ST", FMT_CURRENCY, ALIGN_RIGHT, 16),
    ("valor_ipi", "Valor IPI", FMT_CURRENCY, ALIGN_RIGHT, 16),
    ("valor_pis", "Valor PIS", FMT_CURRENCY, ALIGN_RIGHT, 16),
    ("valor_cofins", "Valor COFINS", FMT_CURRENCY, ALIGN_RIGHT, 16),
    ("total_impostos", "Total Tributos (Soma)", FMT_CURRENCY, ALIGN_RIGHT, 20),
    ("total_tributos_aproximado", "Trib. Aproximado (IBPT)", FMT_CURRENCY, ALIGN_RIGHT, 20),
    ("informacoes_complementares", "Informações Complementares", FMT_TEXT, ALIGN_LEFT, 45),
    ("informacoes_adicionais_fisco", "Informações Adicionais Fisco", FMT_TEXT, ALIGN_LEFT, 35),
    ("arquivo_origem", "Arquivo Origem", FMT_TEXT, ALIGN_LEFT, 35),
]

ITEM_COLUMNS = [
    ("chave", "Chave da NF-e", FMT_TEXT, ALIGN_CENTER, 46),
    ("numero_nota", "Número NF", FMT_NUMBER_INT, ALIGN_CENTER, 14),
    ("serie", "Série", FMT_TEXT, ALIGN_CENTER, 10),
    ("n_item", "Item Nº", FMT_NUMBER_INT, ALIGN_CENTER, 10),
    ("c_prod", "Código Produto", FMT_TEXT, ALIGN_LEFT, 18),
    ("c_ean", "Código EAN", FMT_TEXT, ALIGN_CENTER, 16),
    ("x_prod", "Descrição do Produto", FMT_TEXT, ALIGN_LEFT, 36),
    ("ncm", "NCM", FMT_TEXT, ALIGN_CENTER, 12),
    ("cfop", "CFOP", FMT_TEXT, ALIGN_CENTER, 10),
    ("u_com", "Unidade", FMT_TEXT, ALIGN_CENTER, 10),
    ("q_com", "Quantidade", FMT_NUMBER_2DEC, ALIGN_RIGHT, 16),
    ("v_un_com", "Valor Unitário", FMT_CURRENCY, ALIGN_RIGHT, 16),
    ("v_prod", "Valor Total Produto", FMT_CURRENCY, ALIGN_RIGHT, 18),
    ("v_desc", "Desconto Item", FMT_CURRENCY, ALIGN_RIGHT, 16),
    ("v_frete", "Frete Item", FMT_CURRENCY, ALIGN_RIGHT, 16),
    ("orig_icms", "Origem ICMS", FMT_TEXT, ALIGN_CENTER, 12),
    ("cst_icms", "CST/CSOSN ICMS", FMT_TEXT, ALIGN_CENTER, 14),
    ("v_bc_icms", "Base ICMS", FMT_CURRENCY, ALIGN_RIGHT, 16),
    ("p_icms", "Alíquota ICMS (%)", FMT_NUMBER_2DEC, ALIGN_RIGHT, 16),
    ("v_icms", "Valor ICMS", FMT_CURRENCY, ALIGN_RIGHT, 16),
    ("cst_ipi", "CST IPI", FMT_TEXT, ALIGN_CENTER, 12),
    ("p_ipi", "Alíquota IPI (%)", FMT_NUMBER_2DEC, ALIGN_RIGHT, 16),
    ("v_ipi", "Valor IPI", FMT_CURRENCY, ALIGN_RIGHT, 16),
    ("cst_pis", "CST PIS", FMT_TEXT, ALIGN_CENTER, 12),
    ("p_pis", "Alíquota PIS (%)", FMT_NUMBER_2DEC, ALIGN_RIGHT, 16),
    ("v_pis", "Valor PIS", FMT_CURRENCY, ALIGN_RIGHT, 16),
    ("cst_cofins", "CST COFINS", FMT_TEXT, ALIGN_CENTER, 12),
    ("p_cofins", "Alíquota COFINS (%)", FMT_NUMBER_2DEC, ALIGN_RIGHT, 16),
    ("v_cofins", "Valor COFINS", FMT_CURRENCY, ALIGN_RIGHT, 16),
    ("v_tot_trib", "Trib. Aprox. Item", FMT_CURRENCY, ALIGN_RIGHT, 18),
]


def _selected_columns(columns, selected: Optional[Set[str]]):
    return columns if selected is None else [column for column in columns if column[0] in selected]


def export_nfe_to_excel(
    notas: List[Dict[str, Any]],
    output_path: str,
    selected_fields: Optional[Dict[str, List[str]]] = None,
) -> str:
    """
    Gera uma planilha Excel estilizada com duas abas:
    1. 'Notas Fiscais': resumo consolidado de cada NF-e
    2. 'Itens das Notas': detalhamento de cada item/produto

    ``selected_fields`` aceita as chaves ``notas`` e ``itens`` com as listas de
    campos desejados. Quando omitido, preserva o comportamento anterior e exporta
    todas as colunas.
    """
    selected_fields = selected_fields or {}
    selected_notas = set(selected_fields["notas"]) if "notas" in selected_fields else None
    selected_itens = set(selected_fields["itens"]) if "itens" in selected_fields else None
    colunas_notas = _selected_columns(NOTA_COLUMNS, selected_notas)
    colunas_itens = _selected_columns(ITEM_COLUMNS, selected_itens)
    if not colunas_notas or not colunas_itens:
        raise ValueError("Selecione ao menos um campo para cada aba do Excel.")

    wb = openpyxl.Workbook()

    # ----------------------------------------------------
    # ABA 1: NOTAS FISCAIS
    # ----------------------------------------------------
    ws_notas = wb.active
    ws_notas.title = "Notas Fiscais"
    ws_notas.views.sheetView[0].showGridLines = True
    ws_notas.freeze_panes = "A2"

    _legacy_colunas_notas = [
        ("Chave de Acesso", FMT_TEXT, ALIGN_CENTER, 46),
        ("Número NF", FMT_NUMBER_INT, ALIGN_CENTER, 14),
        ("Série", FMT_TEXT, ALIGN_CENTER, 10),
        ("Data Emissão", FMT_TEXT, ALIGN_CENTER, 20),
        ("Data Saída/Entrada", FMT_TEXT, ALIGN_CENTER, 20),
        ("Tipo Operação", FMT_TEXT, ALIGN_CENTER, 16),
        ("Natureza Operação", FMT_TEXT, ALIGN_LEFT, 28),
        
        # Emitente
        ("Emitente (Razão Social)", FMT_TEXT, ALIGN_LEFT, 32),
        ("Emitente CNPJ/CPF", FMT_TEXT, ALIGN_CENTER, 20),
        ("Emitente IE", FMT_TEXT, ALIGN_CENTER, 16),
        ("Emitente Município", FMT_TEXT, ALIGN_LEFT, 20),
        ("Emitente UF", FMT_TEXT, ALIGN_CENTER, 10),
        
        # Destinatário
        ("Destinatário (Razão Social)", FMT_TEXT, ALIGN_LEFT, 32),
        ("Destinatário CNPJ/CPF", FMT_TEXT, ALIGN_CENTER, 20),
        ("Destinatário IE", FMT_TEXT, ALIGN_CENTER, 16),
        ("Destinatário Município", FMT_TEXT, ALIGN_LEFT, 20),
        ("Destinatário UF", FMT_TEXT, ALIGN_CENTER, 10),
        
        # Logística / Transporte
        ("Modalidade Frete", FMT_TEXT, ALIGN_LEFT, 26),
        ("Transportadora", FMT_TEXT, ALIGN_LEFT, 30),
        ("Transportadora CNPJ/CPF", FMT_TEXT, ALIGN_CENTER, 20),
        ("Transportadora IE", FMT_TEXT, ALIGN_CENTER, 16),
        ("Placa do Caminhão", FMT_TEXT, ALIGN_CENTER, 18),
        ("UF Placa", FMT_TEXT, ALIGN_CENTER, 10),
        ("RNTRC", FMT_TEXT, ALIGN_CENTER, 14),
        ("Qtd Volumes", FMT_NUMBER_INT, ALIGN_RIGHT, 14),
        ("Espécie Volumes", FMT_TEXT, ALIGN_LEFT, 16),
        ("Peso Líquido (kg)", FMT_NUMBER_3DEC, ALIGN_RIGHT, 18),
        ("Peso Bruto (kg)", FMT_NUMBER_3DEC, ALIGN_RIGHT, 18),
        
        # Quantidades e Valores
        ("Qtd Total Itens", FMT_NUMBER_2DEC, ALIGN_RIGHT, 18),
        ("Valor Produtos", FMT_CURRENCY, ALIGN_RIGHT, 18),
        ("Valor Frete", FMT_CURRENCY, ALIGN_RIGHT, 16),
        ("Valor Seguro", FMT_CURRENCY, ALIGN_RIGHT, 16),
        ("Valor Desconto", FMT_CURRENCY, ALIGN_RIGHT, 16),
        ("Outras Despesas", FMT_CURRENCY, ALIGN_RIGHT, 16),
        ("Valor Total da NF", FMT_CURRENCY, ALIGN_RIGHT, 20),
        
        # Todos os Tributos
        ("Base ICMS", FMT_CURRENCY, ALIGN_RIGHT, 16),
        ("Valor ICMS", FMT_CURRENCY, ALIGN_RIGHT, 16),
        ("Base ICMS ST", FMT_CURRENCY, ALIGN_RIGHT, 16),
        ("Valor ICMS ST", FMT_CURRENCY, ALIGN_RIGHT, 16),
        ("Valor IPI", FMT_CURRENCY, ALIGN_RIGHT, 16),
        ("Valor PIS", FMT_CURRENCY, ALIGN_RIGHT, 16),
        ("Valor COFINS", FMT_CURRENCY, ALIGN_RIGHT, 16),
        ("Total Tributos (Soma)", FMT_CURRENCY, ALIGN_RIGHT, 20),
        ("Trib. Aproximado (IBPT)", FMT_CURRENCY, ALIGN_RIGHT, 20),
        
        # Dados Adicionais
        ("Informações Complementares", FMT_TEXT, ALIGN_LEFT, 45),
        ("Informações Adicionais Fisco", FMT_TEXT, ALIGN_LEFT, 35),
        ("Arquivo Origem", FMT_TEXT, ALIGN_LEFT, 35)
    ]

    # Escrever Cabeçalho da Aba 1
    ws_notas.row_dimensions[1].height = 28
    for col_idx, (_, col_name, _, _, _) in enumerate(colunas_notas, start=1):
        cell = ws_notas.cell(row=1, column=col_idx, value=col_name)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = ALIGN_HEADER
        cell.border = BORDER_THIN

    # Escrever Linhas da Aba 1
    for row_idx, nf in enumerate(notas, start=2):
        ws_notas.row_dimensions[row_idx].height = 20
        fill_current = FILL_ZEBRA if row_idx % 2 == 1 else FILL_WHITE

        valores_completos = [
            nf.get('chave', ''),
            to_int_safe(nf.get('numero_nota', 0)),
            nf.get('serie', ''),
            nf.get('data_emissao', ''),
            nf.get('data_saida', ''),
            nf.get('tipo_operacao', ''),
            nf.get('natureza_operacao', ''),
            
            # Emitente
            nf.get('emit_nome', ''),
            nf.get('emit_cnpj_formatado', ''),
            nf.get('emit_ie', ''),
            nf.get('emit_municipio', ''),
            nf.get('emit_uf', ''),
            
            # Destinatário
            nf.get('dest_nome', ''),
            nf.get('dest_cnpj_formatado', ''),
            nf.get('dest_ie', ''),
            nf.get('dest_municipio', ''),
            nf.get('dest_uf', ''),
            
            # Logística
            nf.get('modalidade_frete', ''),
            nf.get('transp_nome', ''),
            nf.get('transp_cnpj_formatado', ''),
            nf.get('transp_ie', ''),
            nf.get('placa_veiculo', ''),
            nf.get('placa_uf', ''),
            nf.get('rntrc', ''),
            nf.get('quantidade_volumes', 0),
            nf.get('especie_volumes', ''),
            nf.get('peso_liquido', 0.0),
            nf.get('peso_bruto', 0.0),
            
            # Valores
            nf.get('quantidade_total_itens', 0.0),
            nf.get('total_produtos', 0.0),
            nf.get('total_frete', 0.0),
            nf.get('total_seguro', 0.0),
            nf.get('total_desconto', 0.0),
            nf.get('outras_despesas', 0.0),
            nf.get('total_nota', 0.0),
            
            # Tributos
            nf.get('bc_icms', 0.0),
            nf.get('valor_icms', 0.0),
            nf.get('bc_icms_st', 0.0),
            nf.get('valor_icms_st', 0.0),
            nf.get('valor_ipi', 0.0),
            nf.get('valor_pis', 0.0),
            nf.get('valor_cofins', 0.0),
            nf.get('total_impostos', 0.0),
            nf.get('total_tributos_aproximado', 0.0),
            
            # Adicionais
            nf.get('informacoes_complementares', ''),
            nf.get('informacoes_adicionais_fisco', ''),
            nf.get('arquivo_origem', '')
        ]

        valores_por_chave = dict(zip((column[0] for column in NOTA_COLUMNS), valores_completos))
        for col_idx, column in enumerate(colunas_notas, start=1):
            key, _, num_fmt, align, _ = column
            val = valores_por_chave[key]
            cell = ws_notas.cell(row=row_idx, column=col_idx, value=val)
            cell.font = FONT_BODY
            cell.fill = fill_current
            cell.alignment = align
            cell.border = BORDER_THIN
            if num_fmt != FMT_TEXT:
                cell.number_format = num_fmt
            else:
                cell.number_format = '@'

    # Ajustar largura das colunas da Aba 1
    for col_idx, (_, _, _, _, default_width) in enumerate(colunas_notas, start=1):
        col_letter = get_column_letter(col_idx)
        ws_notas.column_dimensions[col_letter].width = default_width

    # Adicionar AutoFiltro na Aba 1
    ultima_col_letra_notas = get_column_letter(len(colunas_notas))
    ultima_linha_notas = max(2, len(notas) + 1)
    ws_notas.auto_filter.ref = f"A1:{ultima_col_letra_notas}{ultima_linha_notas}"


    # ----------------------------------------------------
    # ABA 2: ITENS DAS NOTAS
    # ----------------------------------------------------
    ws_itens = wb.create_sheet(title="Itens das Notas")
    ws_itens.views.sheetView[0].showGridLines = True
    ws_itens.freeze_panes = "A2"

    _legacy_colunas_itens = [
        ("Chave da NF-e", FMT_TEXT, ALIGN_CENTER, 46),
        ("Número NF", FMT_NUMBER_INT, ALIGN_CENTER, 14),
        ("Série", FMT_TEXT, ALIGN_CENTER, 10),
        ("Item Nº", FMT_NUMBER_INT, ALIGN_CENTER, 10),
        ("Código Produto", FMT_TEXT, ALIGN_LEFT, 18),
        ("Código EAN", FMT_TEXT, ALIGN_CENTER, 16),
        ("Descrição do Produto", FMT_TEXT, ALIGN_LEFT, 36),
        ("NCM", FMT_TEXT, ALIGN_CENTER, 12),
        ("CFOP", FMT_TEXT, ALIGN_CENTER, 10),
        ("Unidade", FMT_TEXT, ALIGN_CENTER, 10),
        ("Quantidade", FMT_NUMBER_2DEC, ALIGN_RIGHT, 16),
        ("Valor Unitário", FMT_CURRENCY, ALIGN_RIGHT, 16),
        ("Valor Total Produto", FMT_CURRENCY, ALIGN_RIGHT, 18),
        ("Desconto Item", FMT_CURRENCY, ALIGN_RIGHT, 16),
        ("Frete Item", FMT_CURRENCY, ALIGN_RIGHT, 16),
        ("Origem ICMS", FMT_TEXT, ALIGN_CENTER, 12),
        ("CST/CSOSN ICMS", FMT_TEXT, ALIGN_CENTER, 14),
        ("Base ICMS", FMT_CURRENCY, ALIGN_RIGHT, 16),
        ("Alíquota ICMS (%)", FMT_NUMBER_2DEC, ALIGN_RIGHT, 16),
        ("Valor ICMS", FMT_CURRENCY, ALIGN_RIGHT, 16),
        ("CST IPI", FMT_TEXT, ALIGN_CENTER, 12),
        ("Alíquota IPI (%)", FMT_NUMBER_2DEC, ALIGN_RIGHT, 16),
        ("Valor IPI", FMT_CURRENCY, ALIGN_RIGHT, 16),
        ("CST PIS", FMT_TEXT, ALIGN_CENTER, 12),
        ("Alíquota PIS (%)", FMT_NUMBER_2DEC, ALIGN_RIGHT, 16),
        ("Valor PIS", FMT_CURRENCY, ALIGN_RIGHT, 16),
        ("CST COFINS", FMT_TEXT, ALIGN_CENTER, 12),
        ("Alíquota COFINS (%)", FMT_NUMBER_2DEC, ALIGN_RIGHT, 16),
        ("Valor COFINS", FMT_CURRENCY, ALIGN_RIGHT, 16),
        ("Trib. Aprox. Item", FMT_CURRENCY, ALIGN_RIGHT, 18)
    ]

    # Cabeçalho da Aba 2
    ws_itens.row_dimensions[1].height = 28
    for col_idx, (_, col_name, _, _, _) in enumerate(colunas_itens, start=1):
        cell = ws_itens.cell(row=1, column=col_idx, value=col_name)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = ALIGN_HEADER
        cell.border = BORDER_THIN

    # Coletar e escrever todos os itens
    linha_item_idx = 2
    for nf in notas:
        for item in nf.get('itens', []):
            ws_itens.row_dimensions[linha_item_idx].height = 20
            fill_current = FILL_ZEBRA if linha_item_idx % 2 == 1 else FILL_WHITE

            valores_item_completos = [
                item.get('chave', ''),
                to_int_safe(item.get('numero_nota', 0)),
                item.get('serie', ''),
                to_int_safe(item.get('n_item', 1)),
                item.get('c_prod', ''),
                item.get('c_ean', ''),
                item.get('x_prod', ''),
                item.get('ncm', ''),
                item.get('cfop', ''),
                item.get('u_com', ''),
                item.get('q_com', 0.0),
                item.get('v_un_com', 0.0),
                item.get('v_prod', 0.0),
                item.get('v_desc', 0.0),
                item.get('v_frete', 0.0),
                item.get('orig_icms', ''),
                item.get('cst_icms', ''),
                item.get('v_bc_icms', 0.0),
                item.get('p_icms', 0.0),
                item.get('v_icms', 0.0),
                item.get('cst_ipi', ''),
                item.get('p_ipi', 0.0),
                item.get('v_ipi', 0.0),
                item.get('cst_pis', ''),
                item.get('p_pis', 0.0),
                item.get('v_pis', 0.0),
                item.get('cst_cofins', ''),
                item.get('p_cofins', 0.0),
                item.get('v_cofins', 0.0),
                item.get('v_tot_trib', 0.0)
            ]

            valores_item_por_chave = dict(zip((column[0] for column in ITEM_COLUMNS), valores_item_completos))
            for col_idx, column in enumerate(colunas_itens, start=1):
                key, _, num_fmt, align, _ = column
                val = valores_item_por_chave[key]
                cell = ws_itens.cell(row=linha_item_idx, column=col_idx, value=val)
                cell.font = FONT_BODY
                cell.fill = fill_current
                cell.alignment = align
                cell.border = BORDER_THIN
                if num_fmt != FMT_TEXT:
                    cell.number_format = num_fmt
                else:
                    cell.number_format = '@'

            linha_item_idx += 1

    # Ajustar largura das colunas da Aba 2
    for col_idx, (_, _, _, _, default_width) in enumerate(colunas_itens, start=1):
        col_letter = get_column_letter(col_idx)
        ws_itens.column_dimensions[col_letter].width = default_width

    # Adicionar AutoFiltro na Aba 2
    ultima_col_letra_itens = get_column_letter(len(colunas_itens))
    ws_itens.auto_filter.ref = f"A1:{ultima_col_letra_itens}{max(2, linha_item_idx - 1)}"

    # Salvar arquivo
    wb.save(output_path)
    return output_path


def to_int_safe(val: Any) -> int:
    try:
        return int(val)
    except (ValueError, TypeError):
        return 0
