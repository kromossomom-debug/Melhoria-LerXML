# -*- coding: utf-8 -*-
"""
Script de teste automatizado para validar o parsing de XML de NF-e e a geração de Excel.
"""

import os
from nfe_parser import NFeParser
from excel_exporter import export_nfe_to_excel

# XML representativo de uma NF-e 4.00 com todos os campos solicitados
SAMPLE_NFE_XML = """<?xml version="1.0" encoding="UTF-8"?>
<nfeProc versao="4.00" xmlns="http://www.portalfiscal.inf.br/nfe">
  <NFe>
    <infNFe Id="NFe35260912345678000195550010000123451000123456" versao="4.00">
      <ide>
        <cUF>35</cUF>
        <cNF>00012345</cNF>
        <natOp>VENDA DE MERCADORIA ADQUIRIDA DE TERCEIROS</natOp>
        <mod>55</mod>
        <serie>1</serie>
        <nNF>12345</nNF>
        <dhEmi>2026-09-14T10:15:00-03:00</dhEmi>
        <dhSaiEnt>2026-09-14T14:30:00-03:00</dhSaiEnt>
        <tpNF>1</tpNF>
        <idDest>1</idDest>
        <cMunFG>3550308</cMunFG>
        <tpImp>1</tpImp>
        <tpEmis>1</tpEmis>
      </ide>
      <emit>
        <CNPJ>12345678000195</CNPJ>
        <xNome>DISTRIBUIDORA DE GRAOS E ALIMENTOS BRASIL LTDA</xNome>
        <xFant>GRAOS BRASIL</xFant>
        <enderEmit>
          <xLgr>AV DAS INDUSTRIAS</xLgr>
          <nro>1500</nro>
          <xBairro>DISTRITO INDUSTRIAL</xBairro>
          <cMun>3550308</cMun>
          <xMun>SAO PAULO</xMun>
          <UF>SP</UF>
          <CEP>04578000</CEP>
        </enderEmit>
        <IE>111222333444</IE>
        <CRT>3</CRT>
      </emit>
      <dest>
        <CNPJ>98765432000180</CNPJ>
        <xNome>COMERCIO E EXPORTACAO LDC S.A.</xNome>
        <enderDest>
          <xLgr>RODOVIA DOS BANDEIRANTES</xLgr>
          <nro>KM 75</nro>
          <xBairro>ZONA RURAL</xBairro>
          <cMun>3525904</cMun>
          <xMun>JUNDIAI</xMun>
          <UF>SP</UF>
          <CEP>13200000</CEP>
        </enderDest>
        <indIEDest>1</indIEDest>
        <IE>999888777666</IE>
      </dest>
      <det nItem="1">
        <prod>
          <cProd>SOJA-001</cProd>
          <cEAN>7891234567890</cEAN>
          <xProd>SOJA EM GRAOS SAFRA 2026/2027 A GRANEL</xProd>
          <NCM>12019000</NCM>
          <CFOP>5102</CFOP>
          <uCom>TON</uCom>
          <qCom>37.5000</qCom>
          <vUnCom>135.0000</vUnCom>
          <vProd>5062.50</vProd>
          <cEANTrib>7891234567890</cEANTrib>
          <uTrib>TON</uTrib>
          <qTrib>37.5000</qTrib>
          <vUnTrib>135.0000</vUnTrib>
          <indTot>1</indTot>
        </prod>
        <imposto>
          <vTotTrib>650.00</vTotTrib>
          <ICMS>
            <ICMS00>
              <orig>0</orig>
              <CST>00</CST>
              <modBC>3</modBC>
              <vBC>5062.50</vBC>
              <pICMS>18.00</pICMS>
              <vICMS>911.25</vICMS>
            </ICMS00>
          </ICMS>
          <IPI>
            <cEnq>999</cEnq>
            <IPITrib>
              <CST>50</CST>
              <vBC>5062.50</vBC>
              <pIPI>5.00</pIPI>
              <vIPI>253.13</vIPI>
            </IPITrib>
          </IPI>
          <PIS>
            <PISAliq>
              <CST>01</CST>
              <vBC>5062.50</vBC>
              <pPIS>1.65</pPIS>
              <vPIS>83.53</vPIS>
            </PISAliq>
          </PIS>
          <COFINS>
            <COFINSAliq>
              <CST>01</CST>
              <vBC>5062.50</vBC>
              <pCOFINS>7.60</pCOFINS>
              <vCOFINS>384.75</vCOFINS>
            </COFINSAliq>
          </COFINS>
        </imposto>
      </det>
      <det nItem="2">
        <prod>
          <cProd>MILHO-002</cProd>
          <cEAN>SEM GTIN</cEAN>
          <xProd>MILHO EM GRAOS BENEFICIADO</xProd>
          <NCM>10059010</NCM>
          <CFOP>5102</CFOP>
          <uCom>TON</uCom>
          <qCom>15.0000</qCom>
          <vUnCom>80.0000</vUnCom>
          <vProd>1200.00</vProd>
          <cEANTrib>SEM GTIN</cEANTrib>
          <uTrib>TON</uTrib>
          <qTrib>15.0000</qTrib>
          <vUnTrib>80.0000</vUnTrib>
          <indTot>1</indTot>
        </prod>
        <imposto>
          <vTotTrib>180.00</vTotTrib>
          <ICMS>
            <ICMS00>
              <orig>0</orig>
              <CST>00</CST>
              <modBC>3</modBC>
              <vBC>1200.00</vBC>
              <pICMS>18.00</pICMS>
              <vICMS>216.00</vICMS>
            </ICMS00>
          </ICMS>
          <IPI>
            <cEnq>999</cEnq>
            <IPITrib>
              <CST>50</CST>
              <vBC>1200.00</vBC>
              <pIPI>0.00</pIPI>
              <vIPI>0.00</vIPI>
            </IPITrib>
          </IPI>
          <PIS>
            <PISAliq>
              <CST>01</CST>
              <vBC>1200.00</vBC>
              <pPIS>1.65</pPIS>
              <vPIS>19.80</vPIS>
            </PISAliq>
          </PIS>
          <COFINS>
            <COFINSAliq>
              <CST>01</CST>
              <vBC>1200.00</vBC>
              <pCOFINS>7.60</pCOFINS>
              <vCOFINS>91.20</vCOFINS>
            </COFINSAliq>
          </COFINS>
        </imposto>
      </det>
      <total>
        <ICMSTot>
          <vBC>6262.50</vBC>
          <vICMS>1127.25</vICMS>
          <vICMSDeson>0.00</vICMSDeson>
          <vFCP>0.00</vFCP>
          <vBCST>0.00</vBCST>
          <vST>0.00</vST>
          <vFCPST>0.00</vFCPST>
          <vFCPSTRet>0.00</vFCPSTRet>
          <vProd>6262.50</vProd>
          <vFrete>250.00</vFrete>
          <vSeg>50.00</vSeg>
          <vDesc>0.00</vDesc>
          <vII>0.00</vII>
          <vIPI>253.13</vIPI>
          <vIPIDevol>0.00</vIPIDevol>
          <vPIS>103.33</vPIS>
          <vCOFINS>475.95</vCOFINS>
          <vOutro>0.00</vOutro>
          <vNF>6815.63</vNF>
          <vTotTrib>830.00</vTotTrib>
        </ICMSTot>
      </total>
      <transp>
        <modFrete>0</modFrete>
        <transporta>
          <CNPJ>55443322000199</CNPJ>
          <xNome>RODOVIARIO CARGAS PESADAS EXPRESS LTDA</xNome>
          <IE>333444555666</IE>
          <xEnder>RODOVIA ANHANGUERA KM 120</xEnder>
          <xMun>CAMPINAS</xMun>
          <UF>SP</UF>
        </transporta>
        <veicTransp>
          <placa>ABC1D23</placa>
          <UF>SP</UF>
          <RNTRC>12345678</RNTRC>
        </veicTransp>
        <vol>
          <qVol>2</qVol>
          <esp>CARGA A GRANEL</esp>
          <marca>LDC</marca>
          <pesoL>52500.000</pesoL>
          <pesoB>53000.000</pesoB>
        </vol>
      </transp>
      <infAdic>
        <infAdFisco>ICMS RECOLHIDO CONFORME ARTIGO 123 DO RICMS/SP.</infAdFisco>
        <infCpl>PEDIDO DE COMPRA NRO 984512 - ENTREGA NO TERMINAL PORTUARIO SANTOS. MOTORISTA: CARLOS SILVA - CNH 12345678900. PESAGEM BALANCA 02.</infCpl>
      </infAdic>
    </infNFe>
  </NFe>
  <protNFe versao="4.00">
    <infProt>
      <tpAmb>1</tpAmb>
      <verAplic>SP_NFE_PL_009</verAplic>
      <chNFe>35260912345678000195550010000123451000123456</chNFe>
      <dhRecbto>2026-09-14T10:15:30-03:00</dhRecbto>
      <nProt>135260000123456</nProt>
      <cStat>100</cStat>
      <xMotivo>Autorizado o uso da NF-e</xMotivo>
    </infProt>
  </protNFe>
</nfeProc>
"""

def main():
    os.makedirs("exemplos_xml", exist_ok=True)
    xml_path = os.path.join("exemplos_xml", "exemplo_nfe.xml")
    with open(xml_path, "w", encoding="utf-8") as f:
        f.write(SAMPLE_NFE_XML)

    print(f"[1] XML de teste salvo em: {xml_path}")
    
    # Testar parsing
    dados = NFeParser.parse_file(xml_path)
    assert dados['sucesso'] is True, f"Falha no parsing: {dados.get('erro')}"
    
    print("\n--- TESTE DE EXTRAÇÃO ---")
    print(f"Chave da Nota: {dados['chave']}")
    print(f"Número da Nota: {dados['numero_nota']} | Série: {dados['serie']}")
    print(f"Data Emissão: {dados['data_emissao']}")
    print(f"Emitente: {dados['emit_nome']} ({dados['emit_cnpj_formatado']})")
    print(f"Destinatário: {dados['dest_nome']} ({dados['dest_cnpj_formatado']})")
    print(f"Transportadora: {dados['transp_nome']} ({dados['transp_cnpj_formatado']})")
    print(f"Placa do Caminhão: {dados['placa_veiculo']} - UF: {dados['placa_uf']}")
    print(f"Quantidade Total: {dados['quantidade_total_itens_fmt']}")
    print(f"Valor Total Nota: R$ {dados['total_nota_fmt']}")
    print(f"Total de Impostos: R$ {dados['total_impostos_fmt']}")
    print(f"Total Tributos Aproximado: R$ {dados['total_tributos_aproximado']}")
    print(f"ICMS: R$ {dados['valor_icms']} | IPI: R$ {dados['valor_ipi']} | PIS: R$ {dados['valor_pis']} | COFINS: R$ {dados['valor_cofins']}")
    print(f"Peso Líquido: {dados['peso_liquido_fmt']} kg | Peso Bruto: {dados['peso_bruto_fmt']} kg")
    print(f"Dados Adicionais: {dados['informacoes_complementares'][:60]}...")
    print(f"Total de Itens: {len(dados['itens'])}")
    
    # Testar exportação para Excel
    excel_path = os.path.join("exemplos_xml", "teste_relatorio_nfe.xlsx")
    export_nfe_to_excel([dados], excel_path)
    print(f"\n[2] Planilha Excel gerada com sucesso em: {excel_path}")
    assert os.path.exists(excel_path), "Planilha não foi criada!"
    print("[OK] Teste concluído com 100% de sucesso!")

if __name__ == "__main__":
    main()
