# -*- coding: utf-8 -*-
"""Descrições amigáveis para campos descobertos dinamicamente em XMLs de NF-e."""

import re


EXACT_LABELS = {
    "infNFe/@Id": "Identificador completo da NF-e",
    "infNFe/@versao": "Versão do leiaute da NF-e",
    "infNFe/det/@nItem": "Número sequencial do item",
    "infNFe/det/prod/xPed": "Número do pedido",
    "infNFe/det/prod/nItemPed": "Número do item no pedido",
    "infNFe/compra/xPed": "Referências de pedidos da compra",
    "infNFe/infAdic/infCpl": "Informações complementares da nota",
    "infNFe/infAdic/infAdFisco": "Informações adicionais de interesse do Fisco",
}


LEAF_LABELS = {
    "cUF": "Código da UF",
    "cNF": "Código numérico da NF-e",
    "natOp": "Natureza da operação",
    "mod": "Modelo do documento fiscal",
    "serie": "Série da nota fiscal",
    "nNF": "Número da nota fiscal",
    "dhEmi": "Data e hora de emissão",
    "dhSaiEnt": "Data e hora de saída ou entrada",
    "tpNF": "Tipo da operação (entrada ou saída)",
    "idDest": "Destino da operação",
    "cMunFG": "Município do fato gerador",
    "tpImp": "Formato de impressão do DANFE",
    "tpEmis": "Forma de emissão da NF-e",
    "cDV": "Dígito verificador da chave",
    "tpAmb": "Ambiente de autorização",
    "finNFe": "Finalidade da NF-e",
    "indFinal": "Indicador de consumidor final",
    "indPres": "Indicador de presença do comprador",
    "indIntermed": "Indicador de intermediador",
    "procEmi": "Processo de emissão",
    "verProc": "Versão do sistema emissor",
    "refNFe": "Chave da NF-e referenciada",
    "CNPJ": "CNPJ",
    "CPF": "CPF",
    "xNome": "Razão social ou nome",
    "xFant": "Nome fantasia",
    "xLgr": "Logradouro",
    "nro": "Número do endereço",
    "xCpl": "Complemento do endereço",
    "xBairro": "Bairro",
    "cMun": "Código do município",
    "xMun": "Município",
    "UF": "UF",
    "CEP": "CEP",
    "cPais": "Código do país",
    "xPais": "País",
    "fone": "Telefone",
    "IE": "Inscrição estadual",
    "IEST": "Inscrição estadual do substituto tributário",
    "IM": "Inscrição municipal",
    "CNAE": "CNAE fiscal",
    "CRT": "Regime tributário",
    "indIEDest": "Situação da inscrição estadual do destinatário",
    "email": "E-mail",
    "cProd": "Código do produto",
    "cEAN": "Código de barras comercial",
    "xProd": "Descrição do produto",
    "NCM": "Classificação fiscal NCM",
    "NVE": "Nomenclatura de valor aduaneiro",
    "CEST": "Código CEST",
    "indEscala": "Indicador de produção em escala relevante",
    "CNPJFab": "CNPJ do fabricante",
    "cBenef": "Código de benefício fiscal",
    "EXTIPI": "Exceção da tabela TIPI",
    "CFOP": "CFOP da operação",
    "uCom": "Unidade comercial",
    "qCom": "Quantidade comercial",
    "vUnCom": "Valor unitário comercial",
    "vProd": "Valor dos produtos",
    "vNF": "Valor total da NF-e",
    "cEANTrib": "Código de barras tributável",
    "uTrib": "Unidade tributável",
    "qTrib": "Quantidade tributável",
    "vUnTrib": "Valor unitário tributável",
    "vFrete": "Valor do frete",
    "vSeg": "Valor do seguro",
    "vDesc": "Valor do desconto",
    "vOutro": "Outras despesas",
    "indTot": "Indicador de composição do total da nota",
    "xPed": "Número do pedido",
    "nItemPed": "Número do item no pedido",
    "nFCI": "Número de controle da FCI",
    "nRECOPI": "Número RECOPI",
    "orig": "Origem da mercadoria",
    "CST": "Código de situação tributária",
    "CSOSN": "Código de situação do Simples Nacional",
    "modBC": "Modalidade da base de cálculo",
    "pRedBC": "Redução da base de cálculo (%)",
    "vBC": "Base de cálculo",
    "pICMS": "Alíquota de ICMS (%)",
    "vICMS": "Valor do ICMS",
    "vICMSOp": "Valor do ICMS da operação",
    "pDif": "Percentual de diferimento",
    "vICMSDif": "Valor do ICMS diferido",
    "vBCST": "Base de cálculo do ICMS ST",
    "pICMSST": "Alíquota do ICMS ST (%)",
    "vICMSST": "Valor do ICMS ST",
    "vST": "Valor do ICMS ST",
    "pFCP": "Alíquota do FCP (%)",
    "vFCP": "Valor do FCP",
    "vFCPST": "Valor do FCP ST",
    "cEnq": "Código de enquadramento do IPI",
    "pIPI": "Alíquota do IPI (%)",
    "vIPI": "Valor do IPI",
    "pPIS": "Alíquota do PIS (%)",
    "vPIS": "Valor do PIS",
    "pCOFINS": "Alíquota da COFINS (%)",
    "vCOFINS": "Valor da COFINS",
    "vTotTrib": "Valor aproximado dos tributos",
    "cClassTrib": "Classificação tributária do IBS/CBS",
    "pIBSUF": "Alíquota do IBS estadual (%)",
    "pIBSMun": "Alíquota do IBS municipal (%)",
    "pCBS": "Alíquota da CBS (%)",
    "pRedAliq": "Redução da alíquota (%)",
    "pAliqEfet": "Alíquota efetiva (%)",
    "vDif": "Valor diferido",
    "vDevTrib": "Valor de devolução do tributo",
    "vIBSUF": "Valor do IBS estadual",
    "vIBSMun": "Valor do IBS municipal",
    "vIBS": "Valor total do IBS",
    "vCBS": "Valor da CBS",
    "vCredPres": "Crédito presumido",
    "vCredPresCondSus": "Crédito presumido em condição suspensiva",
    "vBCIBSCBS": "Base de cálculo total do IBS/CBS",
    "vNFTot": "Valor total da NF-e com IBS/CBS",
    "infAdProd": "Informações adicionais do produto",
    "vItem": "Valor total do item",
    "vICMSDeson": "Valor do ICMS desonerado",
    "vFCPUFDest": "FCP destinado à UF de destino",
    "vICMSUFDest": "ICMS destinado à UF de destino",
    "vICMSUFRemet": "ICMS destinado à UF do remetente",
    "vFCPSTRet": "FCP ST retido anteriormente",
    "vII": "Valor do imposto de importação",
    "vIPIDevol": "Valor do IPI devolvido",
    "modFrete": "Responsável pelo frete",
    "xEnder": "Endereço",
    "placa": "Placa do veículo",
    "RNTC": "Registro nacional do transportador",
    "RNTRC": "Registro nacional do transportador",
    "qVol": "Quantidade de volumes",
    "esp": "Espécie dos volumes",
    "marca": "Marca dos volumes",
    "nVol": "Numeração dos volumes",
    "pesoL": "Peso líquido",
    "pesoB": "Peso bruto",
    "nFat": "Número da fatura",
    "vOrig": "Valor original da fatura",
    "vLiq": "Valor líquido da fatura",
    "nDup": "Número da parcela",
    "dVenc": "Data de vencimento",
    "vDup": "Valor da parcela",
    "indPag": "Forma de pagamento à vista ou a prazo",
    "tPag": "Meio de pagamento",
    "vPag": "Valor do pagamento",
    "xCampo": "Identificação da observação",
    "xTexto": "Conteúdo da observação",
    "xContato": "Nome do responsável técnico",
}


def _humanize_unknown(tag: str) -> str:
    tag = tag.lstrip("@")
    words = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", " ", tag).replace("_", " ")
    return f"Informação: {words}"


def _context_suffix(parts: list[str]) -> str:
    path = "/".join(parts)
    contexts = (
        ("enderEmit", "do endereço do emitente"),
        ("emit", "do emitente"),
        ("enderDest", "do endereço do destinatário"),
        ("dest", "do destinatário"),
        ("transporta", "da transportadora"),
        ("veicTransp", "do veículo de transporte"),
        ("prod", "do produto"),
        ("ICMSTot", "nos totais da nota"),
        ("IBSCBSTot", "nos totais de IBS/CBS"),
        ("ICMS", "do ICMS do item"),
        ("IPI", "do IPI do item"),
        ("PIS", "do PIS do item"),
        ("COFINS", "da COFINS do item"),
        ("IBSCBS", "do IBS/CBS do item"),
        ("transporta", "da transportadora"),
        ("vol", "dos volumes transportados"),
        ("fat", "da fatura"),
        ("dup", "da parcela de cobrança"),
        ("detPag", "do pagamento"),
        ("infRespTec", "do responsável técnico"),
    )
    for marker, suffix in contexts:
        if marker in parts or marker in path:
            return suffix
    return ""


def describe_xml_path(path: str) -> str:
    """Retorna o significado fiscal da tag, usando o contexto do caminho."""
    if path in EXACT_LABELS:
        return EXACT_LABELS[path]
    parts = [part for part in path.split("/") if part]
    leaf = parts[-1] if parts else path
    leaf_key = leaf.lstrip("@")
    base = LEAF_LABELS.get(leaf_key, _humanize_unknown(leaf_key))
    suffix = _context_suffix(parts[:-1])
    base_folded = base.casefold()
    suffix_folded = suffix.casefold()
    for tax_name in ("icms", "ipi", "pis", "cofins", "ibs/cbs"):
        if tax_name in base_folded and tax_name in suffix_folded:
            return base
    if suffix and suffix.casefold() not in base.casefold():
        return f"{base} {suffix}"
    return base


def describe_xml_group(path: str) -> str:
    """Retorna a área funcional da NF-e sem expor o nome técnico da tag."""
    parts = path.split("/")
    groups = (
        ("ide", "Identificação da NF-e"),
        ("emit", "Emitente"),
        ("dest", "Destinatário"),
        ("prod", "Produto / item da nota"),
        ("imposto", "Tributos do item"),
        ("total", "Totais da nota"),
        ("transp", "Transporte e volumes"),
        ("cobr", "Cobrança e parcelas"),
        ("pag", "Pagamento"),
        ("infAdic", "Informações adicionais"),
        ("compra", "Compras e pedidos"),
        ("infRespTec", "Responsável técnico"),
    )
    for marker, label in groups:
        if marker in parts:
            return label
    return "Informação geral da NF-e"
