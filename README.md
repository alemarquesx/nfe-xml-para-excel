# 📄 NF-e XML para Excel

## 🧾 Descrição do Projeto

Projeto Python para extrair dados de notas fiscais eletrônicas (NF-e) em XML e organizá-los em planilhas Excel.

A proposta é facilitar a consulta e a análise das informações das notas e dos produtos, reduzindo o trabalho manual de organização dos dados.

### 🔧 Funcionalidades Implementadas

- **Leitura de XML**: Utiliza `xml.etree.ElementTree` para acessar os dados da NF-e.
- **Compatibilidade de estrutura**: Lê documentos com raiz `<nfeProc>` ou `<NFe>`, utilizando o namespace da NF-e.
- **Extração dos dados da nota**: Identifica emitente, destinatário, valores e informações da operação.
- **Detalhamento dos produtos**: Extrai os itens de cada nota e os relaciona pela chave de acesso.
- **Conversão numérica**: Converte quantidades e valores para números.
- **Formatação das abas**: Utiliza `openpyxl` para aplicar cabeçalhos, filtros, congelamento da primeira linha e ajuste da largura das colunas.

### 🏢 Dados da Nota Fiscal

- Chave de acesso
- Número e série
- Data de emissão
- Natureza da operação
- CNPJ, nome e UF do emitente
- CNPJ ou CPF, nome e UF do destinatário
- Valor dos produtos
- Frete e desconto
- Valores de ICMS e IPI
- Valor total da nota
- Nome do arquivo de origem

### 📦 Dados dos Produtos

- Chave de acesso e número da nota
- Número do item
- Código e descrição do produto
- NCM e CFOP
- Unidade comercial
- Quantidade
- Valor unitário
- Valor total do item

### 🛠️ Tecnologias

- **Python**
- **openpyxl**: criação e formatação de planilhas Excel.
- **xml.etree.ElementTree**: leitura dos arquivos XML.
- **pathlib**: manipulação de caminhos e nomes de arquivos.

As bibliotecas `xml.etree.ElementTree` e `pathlib` fazem parte da biblioteca padrão do Python.

### 📦 Instalação

Instale a dependência externa:

```bash
python -m pip install openpyxl
```

### 🚧 Status do Projeto

Em desenvolvimento.

A versão atual contém as funções de leitura da NF-e e preenchimento de uma aba do Excel. Para executar a conversão completa, ainda é necessário implementar:

- Busca dos arquivos XML na pasta de entrada.
- Processamento e consolidação das notas e dos itens.
- Criação do arquivo Excel e das abas.
- Salvamento da planilha.
- Tratamento de erros por arquivo.

### 📌 Observações

- O projeto trabalha com arquivos XML de NF-e, não com PDFs ou imagens de DANFE.
- A leitura dos campos não verifica a assinatura digital nem confirma a autorização da nota na SEFAZ.
- Campos de texto ausentes retornam vazios; valores numéricos ausentes ou inválidos retornam `None`.
- A formatação numérica atual aplica duas casas decimais às colunas de valores.
- Utilize arquivos fictícios ou anonimizados como exemplos no repositório, pois XMLs de notas podem conter dados de empresas e pessoas.
