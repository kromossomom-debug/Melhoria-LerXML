# 📑 Leitor de XML NF-e Pro

Aplicação desktop moderna em Python com visual corporativo e intuitivo para leitura em lote de arquivos XML de Notas Fiscais Eletrônicas (**NF-e modelo 55** e **NFC-e modelo 65**) e exportação estruturada para **Microsoft Excel (.xlsx)**.

---

## ✨ Principais Recursos

- **Layout Moderno e Intuitivo**:
  - Interface desktop nativa usando `CustomTkinter`.
  - Fluxo guiado em três etapas: carregar XMLs, escolher colunas e exportar.
  - Cards com quantidade de notas, itens, valor total e tags XML descobertas.
  - Tabela pesquisável e duplo-clique para conferir resumo, itens e tags da nota.
  - Seleção de campos integrada à tela principal, com os campos disponíveis de um lado
    e a ordem final das colunas do Excel do outro.
  - Janela maximizada, cartões de seleção maiores em duas colunas e paginação para
    manter a navegação fluida mesmo com centenas de campos ou milhares de notas.

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
  - A área **Campos para exportar** permite marcar exatamente quais informações serão
    incluídas, sem abrir janelas adicionais.
  - Os filtros `Dados da nota`, `Itens e produtos` e `Tags encontradas no XML` servem
    apenas para organizar a busca; nenhuma categoria exige uma seleção mínima.
  - Os campos aparecem em ordem alfabética e podem ser localizados pelo nome ou pela
    tag XML, como `nNF`, `xProd` e `vNF`; a busca ignora maiúsculas e acentos.
  - A interface apresenta o significado fiscal em português, como `Número do pedido`,
    `Quantidade comercial do produto` e `Município do destinatário`; o caminho técnico
    continua disponível internamente para garantir a extração correta.
  - Após carregar os arquivos, a categoria **Tags encontradas no XML** é montada automaticamente
    com todas as tags-folha e atributos existentes nos XMLs, usando o caminho completo
    para diferenciar tags repetidas em pontos distintos da NF-e.
  - A planilha gerada possui uma única aba, `Dados Selecionados`. A primeira coluna é
    sempre `Chave da NF-e`; as colunas seguintes contêm somente os campos marcados.
    Quando um caminho se repete, os valores são reunidos na mesma célula por `|`.
  - A lista de tags encontradas é aberta automaticamente. Tags técnicas conhecidas também
    recebem um nome amigável; por exemplo, pesquisar `pedido` encontra `xPed` e
    `nItemPed`, preservando o caminho XML completo para diferenciar cada ocorrência.
  - Cada caixa corresponde a um único caminho XML e exporta somente o valor dessa tag.
    No exemplo, `Número do pedido` representa `infNFe/det/prod/xPed` e gera apenas o
    valor `4503929479`, sem acrescentar caminho ou outros campos à célula.
  - Estilização com cabeçalho azul marinho, fontes corporativas, linhas zebradas, autofiltro ativo e formatação numérica nativa de moeda (`R$ #,##0.00`) e quantidades.
  - Leitura e exportação são executadas fora da interface. A busca possui debounce,
    as tags descobertas usam cache e tabelas grandes são paginadas.
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
├── app_gui_v2.py               # Interface principal redesenhada (CustomTkinter)
├── nfe_parser.py               # Motor de leitura e parsing de XMLs
├── excel_exporter.py           # Gerador de planilhas Excel formatadas (OpenPyXL)
├── xml_field_labels.py         # Descrições amigáveis dos campos XML
├── requirements.txt            # Lista de dependências Python
├── iniciar_programa.bat        # Inicializador rápido para Windows
└── README.md                    # Documentação do aplicativo
```
