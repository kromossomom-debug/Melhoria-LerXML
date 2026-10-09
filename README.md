# Extrator NF-e — XML para Excel

Aplicativo desktop para leitura em lote de arquivos XML de NF-e e exportação das informações selecionadas para uma planilha Excel.

O usuário carrega os XMLs, escolhe visualmente os campos desejados e gera um relatório com uma linha por nota fiscal. A chave da NF-e é sempre a primeira coluna.

## Funcionalidades

- Carregamento de vários XMLs ou de uma pasta completa.
- Leitura de NF-e modelo 55 e NFC-e modelo 65 que utilizem a estrutura `infNFe`.
- Detecção automática de tags e atributos existentes nos arquivos carregados.
- Campos apresentados por descrições fiscais em português, em vez de nomes técnicos.
- Pesquisa por descrição, nome da tag ou caminho XML.
- Seleção independente de campos da nota, produtos e tags descobertas.
- Visualização prévia das colunas que serão geradas.
- Consulta dos dados, produtos e tags de cada nota com duplo-clique.
- Remoção automática de notas duplicadas pela chave de acesso.
- Exportação para `.xlsx` com cabeçalho, filtro, linhas alternadas e formatação numérica.
- Paginação e processamento em segundo plano para manter a interface responsiva.

## Como o Excel é montado

A planilha gerada contém a aba `Dados Selecionados`.

- A coluna A sempre contém a `Chave da NF-e`.
- As demais colunas correspondem somente aos campos marcados.
- Cada campo exporta apenas o valor da tag correspondente.
- Quando uma tag aparece em vários itens, seus valores ficam na mesma célula, separados por `|`.
- O caminho técnico da tag é usado internamente, mas o cabeçalho utiliza uma descrição amigável.

Exemplo:

| Chave da NF-e | Número do pedido | Razão social do destinatário |
|---|---|---|
| 00000000000000000000000000000000000000000000 | 1234567890 | EMPRESA DESTINATÁRIA EXEMPLO LTDA |

Nesse caso, `Número do pedido` corresponde internamente ao caminho `infNFe/det/prod/xPed`, mas a célula recebe somente `1234567890`.

## Requisitos

- Windows 10 ou superior.
- Python 3.10 ou superior.
- Acesso de gravação à pasta escolhida para o Excel.

Dependências Python:

- `customtkinter`
- `openpyxl`

## Instalação

Abra o PowerShell na pasta do projeto e execute:

```powershell
python -m pip install -r requirements.txt
```

## Execução

No Windows, dê dois cliques em:

```text
iniciar_programa.bat
```

Também é possível iniciar pelo terminal:

```powershell
python main.py
```

## Fluxo de utilização

1. Clique em `Carregar arquivos XML` ou `Carregar uma pasta`.
2. Aguarde o processamento dos documentos.
3. Acesse `Campos para exportar`.
4. Use a busca e os filtros para localizar as informações desejadas.
5. Marque os campos que devem aparecer no relatório.
6. Confira a prévia em `Colunas do Excel`.
7. Clique em `Exportar planilha` e escolha o destino.
8. Use `Abrir Excel` para abrir o último arquivo gerado.

## Categorias de campos

### Dados da nota

Informações consolidadas, como número da nota, datas, emitente, destinatário, transporte, valores e tributos.

### Itens e produtos

Informações individuais dos produtos, como código, descrição, NCM, CFOP, quantidade e valores. Se uma nota possuir vários itens, os valores são reunidos com `|`.

### Tags encontradas no XML

Campos descobertos diretamente nos documentos carregados. O caminho completo é mantido internamente para diferenciar tags iguais em áreas distintas, como os nomes do emitente e do destinatário.

## Desempenho

- A leitura dos XMLs ocorre fora da thread visual.
- O progresso é atualizado em intervalos controlados para grandes lotes.
- A tabela apresenta até 250 notas por página.
- Os campos são exibidos em páginas de 30 cartões.
- A busca possui um pequeno atraso controlado para evitar reconstruções a cada tecla.
- Os caminhos XML descobertos são mantidos em cache durante a sessão.
- A exportação também ocorre em segundo plano.

## Estrutura do projeto

```text
Melhoria LerXML/
|-- main.py                 Ponto de entrada
|-- app_gui_v2.py           Interface gráfica e fluxo da aplicação
|-- nfe_parser.py           Leitura e interpretação dos XMLs
|-- excel_exporter.py       Geração das planilhas Excel
|-- xml_field_labels.py     Descrições amigáveis das tags XML
|-- requirements.txt        Dependências Python
|-- iniciar_programa.bat    Inicializador para Windows
`-- README.md               Documentação
```

## Arquivos principais

### `main.py`

Inicializa a interface gráfica.

### `app_gui_v2.py`

Controla a janela, carregamento dos arquivos, busca, seleção de campos, paginação e exportação em segundo plano.

### `nfe_parser.py`

Interpreta a estrutura da NF-e, remove namespaces e extrai tanto os campos fiscais conhecidos quanto todas as tags-folha e atributos encontrados em `infNFe`.

### `excel_exporter.py`

Monta a aba `Dados Selecionados`, aplica os formatos visuais e salva o arquivo `.xlsx`.

### `xml_field_labels.py`

Converte tags técnicas em descrições compreensíveis para o usuário, considerando a área da NF-e em que cada informação aparece.

## Solução de problemas

### O programa não abre

Confirme a instalação do Python:

```powershell
python --version
```

Depois reinstale as dependências:

```powershell
python -m pip install -r requirements.txt
```

### Um XML não foi carregado

O documento precisa ser um XML válido e conter o elemento `infNFe`. Arquivos inválidos ou documentos fiscais de outra estrutura são contabilizados como erro no rodapé.

### Uma informação não aparece na lista

Carregue primeiro um XML que contenha essa informação. A categoria `Tags encontradas no XML` é formada dinamicamente a partir dos arquivos da sessão.

### O mesmo campo aparece com significados diferentes

O aplicativo utiliza o caminho completo da tag para separar contextos. Por exemplo, campos com o mesmo nome dentro de `emit`, `dest`, `prod` ou `total` são tratados individualmente.

### A planilha não abre automaticamente

O arquivo continua salvo no local escolhido. Verifique se existe um programa associado ao formato `.xlsx` ou abra o documento manualmente.

## Observações

- O aplicativo não altera os XMLs de origem.
- Nenhum dado é enviado para serviços externos.
- As seleções permanecem ativas enquanto a sessão estiver aberta.
- O botão `Limpar sessão` remove da memória as notas e as seleções atuais.
