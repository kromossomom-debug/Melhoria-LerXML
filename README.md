# 📑 Leitor de XML NF-e Pro

Aplicação desktop moderna em Python com visual corporativo e intuitivo para leitura em lote de arquivos XML de Notas Fiscais Eletrônicas (**NF-e modelo 55** e **NFC-e modelo 65**) e exportação estruturada para **Microsoft Excel (.xlsx)**.

---

## ✨ Principais Recursos

- **Layout Moderno e Intuitivo**:
  - Interface desktop nativa usando `CustomTkinter`.
  - Alternância de tema (**Escuro**, **Claro** ou **Sistema**).
  - Cards com métricas em tempo real: Quantidade de notas, Soma de itens, Valor Total e Total de Tributos.
  - Tabela com colunas organizadas, busca instantânea e duplo-clique para ver detalhes completos.

- **Campos Extraídos**:
  - **Identificação**: Chave de acesso (44 dígitos), Número da Nota, Série, Data e Hora de Emissão, Data de Saída/Entrada, Natureza da Operação e Tipo (Entrada/Saída).
  - **Emitente**: Razão Social, Nome Fantasia, CNPJ/CPF com máscara, Inscrição Estadual, Endereço e Município/UF.
  - **Destinatário**: Razão Social, CNPJ/CPF com máscara, Inscrição Estadual, Endereço e Município/UF.
  - **Logística & Transporte**: Modalidade de frete (CIF/FOB), Transportadora (Nome, CNPJ, IE, Cidade/UF), **Placa do Caminhão** com UF do veículo, RNTRC, Quantidade de Volumes, Espécie e Pesos (Líquido e Bruto no padrão brasileiro).
  - **Quantidades e Valores**: Quantidade total de itens no padrão brasileiro (`1.234,50`), Valor dos Produtos, Frete, Seguro, Desconto, Outras Despesas e **Valor Total da Nota**.
  - **Todos os Tributos**: Base e Valor de ICMS, Base e Valor de ICMS ST, IPI, PIS, COFINS, Soma dos Impostos e Tributos Aproximados (Lei da Transparência/IBPT).
  - **Dados Adicionais**: Informações Complementares de Interesse do Contribuinte (`infCpl`) e Informações Adicionais do Fisco (`infAdFisco`).
  - **Detalhamento de Itens/Produtos**: Código, Descrição, NCM, CFOP, Unidade, Quantidade, Valor Unitário, Valor Total e tributos individuais por item.

- **Exportação para Excel (.xlsx)**:
  - Planilha formatada profissionalmente com **duas abas**:
    1. `Notas Fiscais`: Uma linha consolidada por NF-e com todas as colunas organizadas.
    2. `Itens das Notas`: Detalhamento completo item a item vinculado à nota fiscal.
  - Estilização com cabeçalho azul marinho, fontes corporativas, linhas zebradas, autofiltro ativo e formatação numérica nativa de moeda (`R$ #,##0.00`) e quantidades.
  - **Botão "🚀 Abrir Arquivo Excel"**: Permite abrir o arquivo gerado diretamente no Excel com 1 clique.

---

## 🚀 Como Executar

### 1. Requisitos
- Python 3.10 ou superior instalado.
- Dependências instaladas:
  ```bash
  pip install -r requirements.txt
  ```

### 2. Inicialização Rápida
- Dê um duplo-clique no arquivo **`iniciar_programa.bat`**  
  *OU*
- Execute no terminal:
  ```bash
  python main.py
  ```

---

## 📁 Estrutura do Projeto

```
Melhoria LerXML/
├── main.py                     # Ponto de entrada principal
├── app_gui.py                  # Interface gráfica moderna (CustomTkinter)
├── nfe_parser.py               # Motor de leitura e parsing de XMLs
├── excel_exporter.py           # Gerador de planilhas Excel formatadas (OpenPyXL)
├── requirements.txt            # Lista de dependências Python
├── iniciar_programa.bat        # Inicializador rápido para Windows
├── test_parser_and_excel.py    # Teste automatizado de validação
└── exemplos_xml/               # Pasta com exemplos de XML para teste
    └── exemplo_nfe.xml
```
