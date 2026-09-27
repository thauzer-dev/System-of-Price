# 💰 Price — Análise de Rentabilidade & Sugestão de Preços

Aplicação desktop para análise de preços de peças, construída em **Python**, com interface **Tkinter**, processamento de dados com **Pandas e NumPy** e exportação de resultados para **Excel e CSV**.

O sistema transforma uma base de custos e preços em uma visão operacional de rentabilidade, identificando itens abaixo da margem mínima e calculando sugestões de reajuste conforme os parâmetros de cada categoria.

## 🎯 Sobre o Projeto

O **Price** foi desenvolvido para apoiar decisões de precificação de peças em operações com múltiplas lojas. A aplicação permite configurar margens mínimas, definir quais categorias devem receber sugestões de ajuste e acompanhar o resultado em uma interface centralizada.

A rentabilidade atual é calculada com o custo de cada registro. Para os itens que precisam de reajuste, o sistema utiliza o **maior custo médio do mesmo item entre as lojas presentes na base**, identificado por `COD_ITEM`, como referência para a sugestão de preço.

O projeto demonstra a aplicação de tratamento de dados, regras de negócio e desenvolvimento de interfaces desktop em uma rotina de inteligência comercial.

## ✨ Principais Recursos

### 📊 Análise de Rentabilidade

- Cálculo da margem sobre o preço líquido de imposto.
- Classificação dos registros conforme a margem mínima da categoria.
- Consolidação do maior custo médio por item entre lojas.
- Sugestão de preço para registros válidos, abaixo da margem e com ajuste habilitado.
- Preservação do preço atual para itens dentro da margem ou configurados para permanecer sem ajuste.
- Proteção para que a sugestão calculada não seja inferior ao preço atual.

### 🎛️ Parametrização por Categoria

- Definição de margens mínimas pela interface.
- Opção de habilitar o ajuste ou manter os preços de uma categoria.
- Importação de parâmetros a partir de CSV.
- Reprocessamento da base carregada após alterações nos parâmetros.
- Histórico de alterações com data, horário, categoria, margem e ação.

### 🧹 Importação e Tratamento de Dados

- Detecção automática do separador do CSV.
- Tentativas de leitura com `UTF-8 BOM`, `CP1252` e `Latin-1`.
- Conversão de valores monetários em formatos como `1.234,56`, `1234.56` e `R$ 1.234,56`.
- Normalização dos cabeçalhos da base e dos nomes de categorias.
- Identificação de códigos ou categorias ausentes, custos negativos e preços inválidos.
- Validação de margens, categorias duplicadas e opções de ajuste na importação de parâmetros.

### 🖥️ Interface Operacional

- Tema escuro com destaques visuais por status.
- Indicadores de total de registros, itens dentro e fora da margem e dados inválidos.
- Tabelas de parâmetros, histórico e resultados.
- Formatação de moeda em reais e margens em percentual.
- Prévia limitada aos primeiros **2.000 registros**, com exportação do resultado completo.

### 💾 Configurações e Exportação

- Salvamento manual de parâmetros e histórico em JSON.
- Restauração de parâmetros e histórico para continuar o trabalho.
- Exportação para Excel com as abas `Resultado`, `Parametros` e `Historico`.
- Exportação do resultado para CSV com separador `;` e codificação `UTF-8 BOM`.
- Registro de exceções no console por meio de `logging`.

## 🧮 Regras de Precificação

A versão atual utiliza a constante `TAX_RATE = 0.0925`, equivalente a **9,25%**, como premissa de cálculo do projeto.

```text
Preço líquido = Preço atual × (1 − Imposto)

Rentabilidade atual = (Preço líquido − Custo médio do registro) / Preço líquido

Preço sugerido = Maior custo médio do item / (1 − Margem mínima) / (1 − Imposto)
```

A sugestão é aplicada somente quando os dados são válidos, a categoria possui margem configurada, o ajuste está habilitado e a rentabilidade atual está abaixo da meta. O resultado final considera o maior valor entre o preço calculado e o preço atual.

| Status | Significado |
| --- | --- |
| `DENTRO DA MARGEM` | Rentabilidade igual ou superior à margem mínima, com ajuste habilitado. |
| `FORA DA MARGEM` | Rentabilidade abaixo da margem mínima, com ajuste habilitado. |
| `MANTER - SEM AJUSTE` | Categoria parametrizada para preservar o preço atual. |
| `SEM PARÂMETRO` | Categoria sem margem mínima correspondente. |
| `DADO INVÁLIDO` | Registro com ausência ou inconsistência nos campos validados. |

## 🛠️ Arquitetura e Engenharia de Código

A implementação está concentrada em `price.py`, com funções de leitura e cálculo separadas da classe responsável pela interface.

| Componente | Responsabilidade |
| --- | --- |
| `read_csv_flexible()` | Leitura de CSV com detecção de separador e alternativas de codificação. |
| `parse_number()` | Conversão e normalização de valores numéricos. |
| `normalize_category()` | Padronização dos nomes de categorias. |
| `normalize_params()` | Normalização dos parâmetros recuperados de configurações. |
| `load_parameters()` | Importação e validação de parâmetros em CSV. |
| `calculate()` | Cálculo de rentabilidade, classificação e sugestão de preços. |
| `PricingApp` | Interface, eventos, indicadores, histórico, configurações e exportação. |

O processamento utiliza operações vetorizadas, mapeamento de parâmetros e `groupby().transform("max")` para calcular o maior custo de cada item sem reduzir o número de registros da base.

## 🚀 Como Executar o Projeto

### 1. Prepare o ambiente

Tenha **Python 3.10 ou superior** e suporte ao **Tkinter** em um ambiente com interface gráfica. Na pasta que contém `price.py`, crie um ambiente virtual:

```bash
python -m venv .venv
```

Ative no Windows, pelo PowerShell:

```powershell
.venv\Scripts\Activate.ps1
```

Ou no Linux/macOS:

```bash
source .venv/bin/activate
```

### 2. Instale as dependências

```bash
python -m pip install pandas numpy openpyxl
```

O `Tkinter` faz parte da biblioteca padrão, mas sua disponibilidade depende da instalação do Python. O `openpyxl` é utilizado na exportação para Excel.

### 3. Configure as categorias

**Atenção à versão disponibilizada:** a constante `CATEGORIES` contém apenas o texto provisório `"Dados sobre as Categorias"`. Substitua-o pelas categorias reais da base, em letras maiúsculas, antes de utilizar a parametrização.

Exemplo de configuração:

```python
CATEGORIES = ["FILTROS", "LUBRIFICANTES", "ACESSÓRIOS"]
```

Essa padronização é necessária porque os nomes importados são convertidos para maiúsculas e comparados com as categorias cadastradas no código.

### 4. Inicie a aplicação

```bash
python price.py
```

### 5. Execute uma análise

1. Carregue a base de peças em CSV.
2. Selecione uma categoria e informe sua margem mínima, por exemplo `30` para 30%.
3. Defina se a categoria deve receber sugestões de ajuste.
4. Aplique o parâmetro e repita para as demais categorias.
5. Execute a análise e confira os indicadores e a tabela de resultados.
6. Exporte o resultado e, se desejar, salve a configuração em JSON.

## 📂 Estrutura dos Dados

### Base de peças

| Coluna | Obrigatória | Conteúdo |
| --- | --- | --- |
| `COD_ITEM` | Sim | Identificador do item, utilizado para relacioná-lo entre lojas. |
| `DES_CATEGORIA` | Sim | Categoria associada aos parâmetros de rentabilidade. |
| `CUSTO_MEDIO` | Sim | Custo médio do registro. |
| `PRECO_OFICINA` | Sim | Preço atual do registro. |
| `DES_ITEM` | Não | Descrição exibida na tabela. |
| `NOME_FANTASIA` | Não | Nome da loja exibido na tabela. |

Exemplo fictício, considerando as categorias configuradas acima:

```csv
COD_ITEM;DES_CATEGORIA;CUSTO_MEDIO;PRECO_OFICINA;DES_ITEM;NOME_FANTASIA
1001;FILTROS;50,00;70,00;Filtro de óleo;Loja A
1001;FILTROS;55,00;75,00;Filtro de óleo;Loja B
1002;LUBRIFICANTES;30,00;60,00;Óleo lubrificante;Loja A
```

### Parâmetros opcionais em CSV

```csv
DES_CATEGORIA;MARGEM_MINIMA;AJUSTAR
FILTROS;30;SIM
LUBRIFICANTES;25;NAO
ACESSÓRIOS;35;SIM
```

As margens devem ser maiores ou iguais a zero e menores que 100%. A coluna `AJUSTAR` é opcional; quando omitida, os parâmetros importados habilitam o ajuste.

## 📐 Stack Utilizada

- **Python:** regras de negócio e controle da aplicação.
- **Tkinter / ttk:** interface gráfica desktop e componentes visuais.
- **Pandas:** leitura, transformação, agrupamento e exportação de dados.
- **NumPy:** operações numéricas e cálculos condicionais.
- **Openpyxl:** mecanismo de escrita dos arquivos Excel.
- **JSON, pathlib, datetime e logging:** configurações, caminhos, histórico e diagnóstico de erros.

## 📌 Escopo da Versão Atual

- A aplicação gera sugestões e arquivos de análise; não atualiza preços em um ERP.
- O histórico registra alterações de parâmetros, sem armazenar versões completas de cada resultado.
- A persistência em JSON depende do salvamento manual. Sua restauração não recarrega a base CSV nem altera a constante de imposto.
- Novas análises utilizam a base carregada; as sugestões anteriores não substituem automaticamente os preços de entrada.
- O processamento ocorre na thread da interface, podendo causar espera em bases grandes.

## 📄 Licença

O arquivo fornecido não declara uma licença. A licença de distribuição deverá ser definida e incluída no repositório pelo responsável pelo projeto.
