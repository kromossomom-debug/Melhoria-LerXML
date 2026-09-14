# -*- coding: utf-8 -*-
"""
Módulo de extração e parsing de arquivos XML de Notas Fiscais Eletrônicas (NF-e).
Compatível com NF-e modelo 55 e NFC-e modelo 65, versões 2.00, 3.10 e 4.00.
"""

import xml.etree.ElementTree as ET
from datetime import datetime
import re
from typing import Dict, Any, List, Optional, Tuple


def clean_tag(tag: str) -> str:
    """Remove namespace da tag XML para facilitar buscas."""
    if '}' in tag:
        return tag.split('}', 1)[1]
    return tag


def format_brazilian_number(value: Optional[float], decimals: int = 2) -> str:
    """Formata número no padrão brasileiro: 1.234,56"""
    if value is None:
        return "0,00"
    fmt = f"{{:,.{decimals}f}}"
    formatted = fmt.format(value)
    # Trocar separador americano por brasileiro
    return formatted.replace(",", "X").replace(".", ",").replace("X", ".")


def format_cpf_cnpj(doc: str) -> str:
    """Aplica máscara em CPF ou CNPJ."""
    if not doc:
        return ""
    doc = re.sub(r'\D', '', str(doc))
    if len(doc) == 11:
        return f"{doc[:3]}.{doc[3:6]}.{doc[6:9]}-{doc[9:]}"
    elif len(doc) == 14:
        return f"{doc[:2]}.{doc[2:5]}.{doc[5:8]}/{doc[8:12]}-{doc[12:]}"
    return doc


def format_date_br(date_str: str) -> str:
    """Converte ISO 8601 (ex: 2026-09-14T10:30:00-03:00) para DD/MM/AAAA HH:MM:SS."""
    if not date_str:
        return ""
    try:
        # Remover timezone para parsing simplificado
        cleaned = date_str[:19]
        if 'T' in cleaned:
            dt = datetime.strptime(cleaned, "%Y-%m-%dT%H:%M:%S")
            return dt.strftime("%d/%m/%Y %H:%M:%S")
        elif len(cleaned) == 10:
            dt = datetime.strptime(cleaned, "%Y-%m-%d")
            return dt.strftime("%d/%m/%Y")
    except Exception:
        pass
    return date_str


def to_float(val: Any) -> float:
    """Converte com segurança qualquer valor para float."""
    if val is None:
        return 0.0
    try:
        val_str = str(val).strip().replace(',', '.')
        return float(val_str)
    except (ValueError, TypeError):
        return 0.0


def to_int(val: Any) -> int:
    """Converte com segurança qualquer valor para int."""
    if val is None:
        return 0
    try:
        return int(float(str(val).strip()))
    except (ValueError, TypeError):
        return 0


class NFeParser:
    """Classe responsável pelo parsing de arquivos XML de NF-e."""

    FRETE_MAP = {
        '0': '0 - Contratação pelo Emitente (CIF)',
        '1': '1 - Contratação pelo Destinatário (FOB)',
        '2': '2 - Contratação por Terceiros',
        '3': '3 - Transporte Próprio (Remetente)',
        '4': '4 - Transporte Próprio (Destinatário)',
        '9': '9 - Sem Ocorrência de Transporte'
    }

    TIPO_OPERACAO_MAP = {
        '0': '0 - Entrada',
        '1': '1 - Saída'
    }

    @classmethod
    def _find_text(cls, element: Optional[ET.Element], path: str, default: str = "") -> str:
        """Busca o texto de uma tag ignorando namespaces."""
        if element is None:
            return default
        
        parts = path.split('/')
        curr = [element]
        
        for part in parts:
            next_curr = []
            for node in curr:
                for child in node:
                    if clean_tag(child.tag) == part:
                        next_curr.append(child)
            curr = next_curr
            if not curr:
                return default
        
        if curr and curr[0].text:
            return curr[0].text.strip()
        return default

    @classmethod
    def _find_element(cls, element: Optional[ET.Element], path: str) -> Optional[ET.Element]:
        """Busca um elemento ignorando namespaces."""
        if element is None:
            return None
            
        parts = path.split('/')
        curr = [element]
        
        for part in parts:
            next_curr = []
            for node in curr:
                for child in node:
                    if clean_tag(child.tag) == part:
                        next_curr.append(child)
            curr = next_curr
            if not curr:
                return None
                
        return curr[0] if curr else None

    @classmethod
    def _find_all(cls, element: Optional[ET.Element], tag_name: str) -> List[ET.Element]:
        """Busca todos os filhos diretos com tag ignorando namespaces."""
        if element is None:
            return []
        return [child for child in element if clean_tag(child.tag) == tag_name]

    @classmethod
    def parse_file(cls, file_path: str) -> Dict[str, Any]:
        """Lê e processa um arquivo XML do disco."""
        try:
            tree = ET.parse(file_path)
            root = tree.getroot()
            data = cls.parse_element(root)
            data['arquivo_origem'] = file_path
            return data
        except Exception as e:
            return {
                'sucesso': False,
                'erro': f"Erro ao ler XML {file_path}: {str(e)}",
                'arquivo_origem': file_path
            }

    @classmethod
    def parse_string(cls, xml_content: str) -> Dict[str, Any]:
        """Lê e processa uma string XML."""
        try:
            root = ET.fromstring(xml_content)
            return cls.parse_element(root)
        except Exception as e:
            return {
                'sucesso': False,
                'erro': f"Erro ao processar XML: {str(e)}"
            }

    @classmethod
    def parse_element(cls, root: ET.Element) -> Dict[str, Any]:
        """Extrai todos os dados estruturados da NF-e."""
        # Localiza o nó infNFe
        inf_nfe = None
        if clean_tag(root.tag) == 'infNFe':
            inf_nfe = root
        else:
            # Pode estar dentro de NFe ou nfeProc
            for elem in root.iter():
                if clean_tag(elem.tag) == 'infNFe':
                    inf_nfe = elem
                    break

        if inf_nfe is None:
            return {
                'sucesso': False,
                'erro': "Tag <infNFe> não encontrada no arquivo XML."
            }

        # 1. Chave da Nota
        chave = inf_nfe.attrib.get('Id', '').replace('NFe', '').strip()
        if not chave:
            # Tentar achar protNFe/infProt/chNFe
            for elem in root.iter():
                if clean_tag(elem.tag) == 'chNFe' and elem.text:
                    chave = elem.text.strip()
                    break

        # 2. Identificação da Nota (ide)
        ide = cls._find_element(inf_nfe, 'ide')
        numero_nota = cls._find_text(ide, 'nNF')
        serie_nota = cls._find_text(ide, 'serie')
        modelo = cls._find_text(ide, 'mod')
        nat_op = cls._find_text(ide, 'natOp')
        data_emissao_raw = cls._find_text(ide, 'dhEmi') or cls._find_text(ide, 'dEmi')
        data_emissao_formatada = format_date_br(data_emissao_raw)
        data_saida_raw = cls._find_text(ide, 'dhSaiEnt') or cls._find_text(ide, 'dSaiEnt')
        data_saida_formatada = format_date_br(data_saida_raw)
        tp_nf = cls._find_text(ide, 'tpNF')
        tipo_operacao = cls.TIPO_OPERACAO_MAP.get(tp_nf, tp_nf)

        # 3. Emitente (emit)
        emit = cls._find_element(inf_nfe, 'emit')
        emit_nome = cls._find_text(emit, 'xNome')
        emit_fantasia = cls._find_text(emit, 'xFant')
        emit_cnpj = cls._find_text(emit, 'CNPJ') or cls._find_text(emit, 'CPF')
        emit_cnpj_formatado = format_cpf_cnpj(emit_cnpj)
        emit_ie = cls._find_text(emit, 'IE')
        emit_mun = cls._find_text(emit, 'enderEmit/xMun')
        emit_uf = cls._find_text(emit, 'enderEmit/UF')
        emit_logradouro = cls._find_text(emit, 'enderEmit/xLgr')
        emit_numero = cls._find_text(emit, 'enderEmit/nro')
        emit_bairro = cls._find_text(emit, 'enderEmit/xBairro')
        emit_cep = cls._find_text(emit, 'enderEmit/CEP')

        # 4. Destinatário (dest)
        dest = cls._find_element(inf_nfe, 'dest')
        dest_nome = cls._find_text(dest, 'xNome')
        dest_cnpj = cls._find_text(dest, 'CNPJ') or cls._find_text(dest, 'CPF')
        dest_cnpj_formatado = format_cpf_cnpj(dest_cnpj)
        dest_ie = cls._find_text(dest, 'IE')
        dest_mun = cls._find_text(dest, 'enderDest/xMun')
        dest_uf = cls._find_text(dest, 'enderDest/UF')
        dest_logradouro = cls._find_text(dest, 'enderDest/xLgr')
        dest_numero = cls._find_text(dest, 'enderDest/nro')
        dest_bairro = cls._find_text(dest, 'enderDest/xBairro')
        dest_cep = cls._find_text(dest, 'enderDest/CEP')

        # 5. Transporte (transp)
        transp = cls._find_element(inf_nfe, 'transp')
        mod_frete_code = cls._find_text(transp, 'modFrete')
        modalidade_frete = cls.FRETE_MAP.get(mod_frete_code, mod_frete_code)

        transporta = cls._find_element(transp, 'transporta')
        transp_nome = cls._find_text(transporta, 'xNome')
        transp_cnpj = cls._find_text(transporta, 'CNPJ') or cls._find_text(transporta, 'CPF')
        transp_cnpj_formatado = format_cpf_cnpj(transp_cnpj)
        transp_ie = cls._find_text(transporta, 'IE')
        transp_mun = cls._find_text(transporta, 'xMun')
        transp_uf = cls._find_text(transporta, 'UF')
        transp_ender = cls._find_text(transporta, 'xEnder')

        # Veículo e Placa do Caminhão
        veic_transp = cls._find_element(transp, 'veicTransp')
        placa = cls._find_text(veic_transp, 'placa')
        placa_uf = cls._find_text(veic_transp, 'UF')
        rntrc = cls._find_text(veic_transp, 'RNTRC')

        # Se não achou em veicTransp, procura em reboque
        if not placa:
            reboque = cls._find_element(transp, 'reboque')
            placa = cls._find_text(reboque, 'placa')
            placa_uf = cls._find_text(reboque, 'UF')
            if not rntrc:
                rntrc = cls._find_text(reboque, 'RNTRC')

        # Volumes e Pesos
        vol = cls._find_element(transp, 'vol')
        q_vol = to_int(cls._find_text(vol, 'qVol'))
        esp_vol = cls._find_text(vol, 'esp')
        marca_vol = cls._find_text(vol, 'marca')
        peso_liquido = to_float(cls._find_text(vol, 'pesoL'))
        peso_bruto = to_float(cls._find_text(vol, 'pesoB'))

        # 6. Totais e Tributos da Nota (total/ICMSTot)
        total = cls._find_element(inf_nfe, 'total')
        icms_tot = cls._find_element(total, 'ICMSTot')
        v_prod = to_float(cls._find_text(icms_tot, 'vProd'))
        v_frete = to_float(cls._find_text(icms_tot, 'vFrete'))
        v_seg = to_float(cls._find_text(icms_tot, 'vSeg'))
        v_desc = to_float(cls._find_text(icms_tot, 'vDesc'))
        v_outro = to_float(cls._find_text(icms_tot, 'vOutro'))
        v_nf = to_float(cls._find_text(icms_tot, 'vNF'))

        # Tributos
        v_bc_icms = to_float(cls._find_text(icms_tot, 'vBC'))
        v_icms = to_float(cls._find_text(icms_tot, 'vICMS'))
        v_bc_st = to_float(cls._find_text(icms_tot, 'vBCST'))
        v_st = to_float(cls._find_text(icms_tot, 'vST'))
        v_ipi = to_float(cls._find_text(icms_tot, 'vIPI'))
        v_pis = to_float(cls._find_text(icms_tot, 'vPIS'))
        v_cofins = to_float(cls._find_text(icms_tot, 'vCOFINS'))
        v_tot_trib = to_float(cls._find_text(icms_tot, 'vTotTrib'))
        v_fcp = to_float(cls._find_text(icms_tot, 'vFCP'))
        v_fcp_st = to_float(cls._find_text(icms_tot, 'vFCPST'))

        # Total de impostos calculados
        total_impostos = v_icms + v_st + v_ipi + v_pis + v_cofins

        # 7. Dados Adicionais (infAdic)
        inf_adic = cls._find_element(inf_nfe, 'infAdic')
        inf_cpl = cls._find_text(inf_adic, 'infCpl')
        inf_ad_fisco = cls._find_text(inf_adic, 'infAdFisco')

        # 8. Itens da Nota Fiscal (det)
        itens = []
        quantidade_total_itens = 0.0

        for det in cls._find_all(inf_nfe, 'det'):
            n_item = det.attrib.get('nItem', '')
            prod = cls._find_element(det, 'prod')
            c_prod = cls._find_text(prod, 'cProd')
            c_ean = cls._find_text(prod, 'cEAN')
            x_prod = cls._find_text(prod, 'xProd')
            ncm = cls._find_text(prod, 'NCM')
            cfop = cls._find_text(prod, 'CFOP')
            u_com = cls._find_text(prod, 'uCom')
            q_com = to_float(cls._find_text(prod, 'qCom'))
            v_un_com = to_float(cls._find_text(prod, 'vUnCom'))
            v_prod_item = to_float(cls._find_text(prod, 'vProd'))
            v_desc_item = to_float(cls._find_text(prod, 'vDesc'))
            v_frete_item = to_float(cls._find_text(prod, 'vFrete'))
            v_seg_item = to_float(cls._find_text(prod, 'vSeg'))
            v_outro_item = to_float(cls._find_text(prod, 'vOutro'))

            quantidade_total_itens += q_com

            # Tributos do item
            imposto = cls._find_element(det, 'imposto')
            v_tot_trib_item = to_float(cls._find_text(imposto, 'vTotTrib'))

            # ICMS
            icms_node = cls._find_element(imposto, 'ICMS')
            icms_tag_elem = None
            if icms_node is not None and len(icms_node) > 0:
                icms_tag_elem = icms_node[0]

            orig_icms = cls._find_text(icms_tag_elem, 'orig')
            cst_icms = cls._find_text(icms_tag_elem, 'CST') or cls._find_text(icms_tag_elem, 'CSOSN')
            v_bc_icms_item = to_float(cls._find_text(icms_tag_elem, 'vBC'))
            p_icms_item = to_float(cls._find_text(icms_tag_elem, 'pICMS'))
            v_icms_item = to_float(cls._find_text(icms_tag_elem, 'vICMS'))

            # IPI
            ipi_node = cls._find_element(imposto, 'IPI')
            ipi_trib = cls._find_element(ipi_node, 'IPITrib')
            cst_ipi = cls._find_text(ipi_trib, 'CST') or cls._find_text(ipi_node, 'IPINT/CST')
            v_bc_ipi_item = to_float(cls._find_text(ipi_trib, 'vBC'))
            p_ipi_item = to_float(cls._find_text(ipi_trib, 'pIPI'))
            v_ipi_item = to_float(cls._find_text(ipi_trib, 'vIPI'))

            # PIS
            pis_node = cls._find_element(imposto, 'PIS')
            pis_tag_elem = None
            if pis_node is not None and len(pis_node) > 0:
                pis_tag_elem = pis_node[0]
            cst_pis = cls._find_text(pis_tag_elem, 'CST')
            v_bc_pis_item = to_float(cls._find_text(pis_tag_elem, 'vBC'))
            p_pis_item = to_float(cls._find_text(pis_tag_elem, 'pPIS'))
            v_pis_item = to_float(cls._find_text(pis_tag_elem, 'vPIS'))

            # COFINS
            cofins_node = cls._find_element(imposto, 'COFINS')
            cofins_tag_elem = None
            if cofins_node is not None and len(cofins_node) > 0:
                cofins_tag_elem = cofins_node[0]
            cst_cofins = cls._find_text(cofins_tag_elem, 'CST')
            v_bc_cofins_item = to_float(cls._find_text(cofins_tag_elem, 'vBC'))
            p_cofins_item = to_float(cls._find_text(cofins_tag_elem, 'pCOFINS'))
            v_cofins_item = to_float(cls._find_text(cofins_tag_elem, 'vCOFINS'))

            itens.append({
                'n_item': n_item,
                'chave': chave,
                'numero_nota': numero_nota,
                'serie': serie_nota,
                'c_prod': c_prod,
                'c_ean': c_ean,
                'x_prod': x_prod,
                'ncm': ncm,
                'cfop': cfop,
                'u_com': u_com,
                'q_com': q_com,
                'q_com_fmt': format_brazilian_number(q_com, 4 if q_com % 1 != 0 else 2),
                'v_un_com': v_un_com,
                'v_un_com_fmt': format_brazilian_number(v_un_com, 4 if v_un_com % 1 != 0 else 2),
                'v_prod': v_prod_item,
                'v_prod_fmt': format_brazilian_number(v_prod_item, 2),
                'v_desc': v_desc_item,
                'v_frete': v_frete_item,
                'orig_icms': orig_icms,
                'cst_icms': cst_icms,
                'v_bc_icms': v_bc_icms_item,
                'p_icms': p_icms_item,
                'v_icms': v_icms_item,
                'cst_ipi': cst_ipi,
                'p_ipi': p_ipi_item,
                'v_ipi': v_ipi_item,
                'cst_pis': cst_pis,
                'p_pis': p_pis_item,
                'v_pis': v_pis_item,
                'cst_cofins': cst_cofins,
                'p_cofins': p_cofins_item,
                'v_cofins': v_cofins_item,
                'v_tot_trib': v_tot_trib_item
            })

        return {
            'sucesso': True,
            'chave': chave,
            'numero_nota': numero_nota,
            'serie': serie_nota,
            'modelo': modelo,
            'natureza_operacao': nat_op,
            'tipo_operacao': tipo_operacao,
            'data_emissao': data_emissao_formatada,
            'data_emissao_raw': data_emissao_raw,
            'data_saida': data_saida_formatada,
            
            # Emitente
            'emit_nome': emit_nome,
            'emit_fantasia': emit_fantasia,
            'emit_cnpj': emit_cnpj,
            'emit_cnpj_formatado': emit_cnpj_formatado,
            'emit_ie': emit_ie,
            'emit_municipio': emit_mun,
            'emit_uf': emit_uf,
            'emit_endereco': f"{emit_logradouro}, {emit_numero} - {emit_bairro}".strip(' ,-'),
            'emit_cep': emit_cep,

            # Destinatário
            'dest_nome': dest_nome,
            'dest_cnpj': dest_cnpj,
            'dest_cnpj_formatado': dest_cnpj_formatado,
            'dest_ie': dest_ie,
            'dest_municipio': dest_mun,
            'dest_uf': dest_uf,
            'dest_endereco': f"{dest_logradouro}, {dest_numero} - {dest_bairro}".strip(' ,-'),
            'dest_cep': dest_cep,

            # Transporte e Veículo
            'modalidade_frete': modalidade_frete,
            'transp_nome': transp_nome,
            'transp_cnpj': transp_cnpj,
            'transp_cnpj_formatado': transp_cnpj_formatado,
            'transp_ie': transp_ie,
            'transp_municipio': transp_mun,
            'transp_uf': transp_uf,
            'transp_endereco': transp_ender,
            'placa_veiculo': placa.upper().strip() if placa else "",
            'placa_uf': placa_uf.upper().strip() if placa_uf else "",
            'rntrc': rntrc,
            'quantidade_volumes': q_vol,
            'especie_volumes': esp_vol,
            'marca_volumes': marca_vol,
            'peso_liquido': peso_liquido,
            'peso_bruto': peso_bruto,
            'peso_liquido_fmt': format_brazilian_number(peso_liquido, 3),
            'peso_bruto_fmt': format_brazilian_number(peso_bruto, 3),

            # Quantidades e Valores Totais
            'quantidade_total_itens': quantidade_total_itens,
            'quantidade_total_itens_fmt': format_brazilian_number(quantidade_total_itens, 2),
            'total_produtos': v_prod,
            'total_frete': v_frete,
            'total_seguro': v_seg,
            'total_desconto': v_desc,
            'outras_despesas': v_outro,
            'total_nota': v_nf,
            'total_nota_fmt': format_brazilian_number(v_nf, 2),

            # Tributos
            'bc_icms': v_bc_icms,
            'valor_icms': v_icms,
            'bc_icms_st': v_bc_st,
            'valor_icms_st': v_st,
            'valor_ipi': v_ipi,
            'valor_pis': v_pis,
            'valor_cofins': v_cofins,
            'total_tributos_aproximado': v_tot_trib,
            'valor_fcp': v_fcp,
            'valor_fcp_st': v_fcp_st,
            'total_impostos': total_impostos,
            'total_impostos_fmt': format_brazilian_number(total_impostos, 2),

            # Dados Adicionais
            'informacoes_complementares': inf_cpl,
            'informacoes_adicionais_fisco': inf_ad_fisco,

            # Itens
            'itens': itens,
            'quantidade_itens_distintos': len(itens)
        }
