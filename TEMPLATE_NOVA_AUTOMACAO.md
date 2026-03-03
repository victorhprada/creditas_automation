# Template: Nova Automação Excel com Streamlit

Guia para replicar a estrutura da automação de faturamento para um novo cliente.
A ideia base é a mesma: juntar uma **planilha do parceiro** com uma **planilha BASE**,
processar regras de negócio específicas e gerar um arquivo Excel atualizado para download.

---

## Visão Geral da Arquitetura

```
app.py
├── Seção 1: Funções auxiliares e regras de negócio
│   ├── Utilitários genéricos (estilo, headers, última linha...)
│   ├── Funções de processamento (copiar dados, aplicar fórmulas...)
│   └── Funções de orquestração (atualizar abas, validar estrutura...)
└── Seção 2: Interface Streamlit
    ├── Upload dos dois arquivos Excel
    ├── Inputs de configuração (mês alvo, datas do ciclo)
    └── Pipeline de processamento + download do resultado
```

---

## Stack de Tecnologia

| Pacote | Uso |
|---|---|
| `streamlit` | Interface web (upload, formulário, download) |
| `openpyxl` | Leitura e escrita de arquivos `.xlsx`/`.xlsm` |
| `pandas` | Manipulação de dados tabulares (filtros, DataFrames) |
| `python-dateutil` | Cálculo de datas relativas (mês anterior, etc.) |

### requirements.txt mínimo

```txt
streamlit
openpyxl
pandas
python-dateutil
```

---

## Estrutura de Arquivos do Projeto

```
nome_do_projeto/
├── app.py                ← Código principal
├── requirements.txt
├── .gitignore
└── README.md
```

---

## Checklist de Adaptação para Novo Cliente

### 1. Mapeamento das Planilhas

Antes de escrever código, responda:

- [ ] Qual é o nome da planilha do **parceiro** (arquivo externo)?
- [ ] Quais **abas** da planilha do parceiro serão usadas?
- [ ] Qual é o nome da planilha **BASE** (arquivo interno/master)?
- [ ] Quais **abas** da planilha BASE serão lidas/escritas?
- [ ] Existe uma aba **template** (como `JAN.26`) para clonar a cada mês?
- [ ] Existe uma aba de **resumo** que precisa ser atualizada?
- [ ] Existe uma aba de **inadimplentes** ou equivalente?

### 2. Mapeamento de Colunas

Para cada aba relevante, documente:

| Coluna (letra) | Índice | Nome do Header | Conteúdo |
|---|---|---|---|
| A | 1 | (ex: CCB / ID) | Identificador único do contrato |
| B | 2 | ... | ... |

> **Dica:** Use a função `encontrar_coluna_por_header(ws, nome)` para localizar colunas pelo header
> em vez de hardcodar índices sempre que possível.

### 3. Validações de Entrada

Adapte a função `validar_abas_necessarias()` com os nomes reais das abas do novo cliente:

```python
def validar_abas_necessarias(parceiro_wb, base_wb):
    abas_parceiro_necessarias = ['Nome Aba 1', 'Nome Aba 2']  # ← alterar
    abas_base_necessarias     = ['BASE', 'RESUMO', 'TEMPLATE_MES']  # ← alterar
    for aba in abas_parceiro_necessarias:
        if aba not in parceiro_wb.sheetnames:
            return False, f"Aba '{aba}' não encontrada no arquivo PARCEIRO"
    for aba in abas_base_necessarias:
        if aba not in base_wb.sheetnames:
            return False, f"Aba '{aba}' não encontrada no arquivo BASE"
    return True, "OK"
```

### 4. Função de Cópia de Dados

A função `inserir_dados_colunas_especificas()` copia colunas 1–13 da aba do parceiro
para a aba do mês. Ajuste `col_fim` conforme o número real de colunas do novo cliente:

```python
inserir_dados_colunas_especificas(
    ws_origem=parceiro_wb['Nome Da Aba'],
    ws_destino=ws_mes,
    col_inicio=1,
    col_fim=N,   # ← número de colunas a copiar
    linha_destino_inicio=2
)
```

### 5. Regras de Negócio por Colunas

Substitua `aplicar_regras_colunas_n_x()` pela lógica específica do novo cliente.
Estrutura base a manter:

```python
def aplicar_regras_colunas_cliente(ws, target_month, linha_inicio=2):
    ultima_linha = ...  # encontrar última linha preenchida
    for row in range(linha_inicio, ultima_linha + 1):
        # Escreva aqui as fórmulas/valores específicos do cliente
        ws.cell(row=row, column=N, value="valor ou =FORMULA")
```

### 6. Atualização da Aba BASE

Se o cliente tiver uma aba BASE para receber dados de produção/novos contratos,
adapte `copiar_producao_para_base()`:

- Verifique quais colunas da aba de origem mapeiam para quais colunas da BASE
- Ajuste o range de colunas no loop (`for col in range(1, N)`)
- Adapte fórmulas fixas inseridas na coluna 8 (ou equivalente)

### 7. Aba de Resumo (se existir)

Se o cliente tiver uma aba de resumo com blocos de métricas por mês, recrie as funções:

- `atualizar_resumo_mes_faturamento()` → insere nova coluna com fórmulas de faturamento
- `atualizar_resumo_ciclo_pmt()` → insere bloco de COUNTIFS/SUMIFS por período
- `atualizar_resumo_bloco_final()` → preenche totalizadores finais

> **Atenção a células mescladas:** use `ws.unmerge_cells()` antes de escrever em
> células que possam estar dentro de um intervalo mesclado.

### 8. Processamento de Inadimplentes / Exceções

Se houver uma lógica de identificar registros "problemáticos" (inadimplentes, pendentes, etc.),
adapte `processar_inadimplentes()`:

- Defina qual coluna é o identificador único (CCB, CPF, ID...)
- Defina a coluna de comparação na aba de destino
- Mapeie as colunas de origem para as colunas de destino

---

## Funções Genéricas Reutilizáveis (Copiar sem Alteração)

Estas funções são 100% independentes de regras de negócio e podem ser usadas em qualquer projeto:

```python
def copiar_estilo(celula_origem, celula_destino)
    # Copia font, border, fill, number_format, alignment

def encontrar_coluna_por_header(ws, nome_header) -> int | None
    # Retorna o índice da coluna cujo header (linha 1) é igual a nome_header

def encontrar_ultima_linha(ws) -> int
    # Retorna o número da última linha com algum valor preenchido

def limpar_dados_worksheet(ws, manter_linha_1=True)
    # Apaga todos os valores da planilha (opcionalmente preserva o header)

def calcular_mes_anterior(mes_str) -> str
    # Converte "FEV.26" → "jan/26" (mês anterior no formato textual)
```

---

## Template da Interface Streamlit

Cole este bloco no final do `app.py` e adapte os labels e a lógica do pipeline:

```python
import streamlit as st
import openpyxl
import pandas as pd
from io import BytesIO
import gc

st.set_page_config(page_title="Nome da Automação", page_icon="📊", layout="centered")
st.title("📊 Nome da Automação - Cliente X")

with st.form("form_processamento"):
    arquivo_parceiro = st.file_uploader("1️⃣ Arquivo PARCEIRO (.xlsx)", type=["xlsx"])
    arquivo_base     = st.file_uploader("2️⃣ Arquivo BASE (.xlsx, .xlsm)", type=["xlsx", "xlsm"])

    col1, col2, col3 = st.columns(3)
    with col1:
        target_month = st.text_input("Mês Alvo", value="MAR.26")  # ← ajustar default
    with col2:
        dt_inicio = st.date_input("Início do Ciclo")
    with col3:
        dt_fim = st.date_input("Fim do Ciclo")

    submit = st.form_submit_button("Iniciar Processamento", type="primary")

if submit:
    if not arquivo_parceiro or not arquivo_base:
        st.error("⚠️ Envie as duas planilhas antes de processar.")
    else:
        with st.status("🚀 Processando...", expanded=True) as status:
            try:
                st.write("📥 Lendo arquivos...")
                parceiro_wb = openpyxl.load_workbook(arquivo_parceiro, data_only=True)
                base_wb     = openpyxl.load_workbook(arquivo_base, data_only=False)

                st.write("⚙️ Validando estrutura...")
                valido, msg = validar_abas_necessarias(parceiro_wb, base_wb)
                if not valido:
                    raise ValueError(msg)

                st.write("🔄 Processando dados...")
                # ← CHAMAR AQUI AS FUNÇÕES DE NEGÓCIO DO NOVO CLIENTE

                st.write("💾 Gerando arquivo final...")
                output = BytesIO()
                base_wb.save(output)
                output.seek(0)

                del parceiro_wb, base_wb
                gc.collect()

                status.update(label="✅ Concluído!", state="complete", expanded=False)

                st.download_button(
                    label="📥 Baixar Excel Processado",
                    data=output,
                    file_name=f"Processado_{target_month}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

            except Exception as e:
                status.update(label="❌ Erro", state="error")
                st.error(f"Erro: {str(e)}")
```

---

## Como Rodar Localmente

```bash
# Criar e ativar ambiente virtual
python -m venv venv
venv\Scripts\activate       # Windows
# source venv/bin/activate  # Linux/Mac

# Instalar dependências
pip install -r requirements.txt

# Rodar a aplicação
streamlit run app.py
# Acesse em: http://localhost:8501
```

---

## Pontos de Atenção

| Situação | Como tratar |
|---|---|
| Células mescladas | Usar `ws.unmerge_cells(str(range))` antes de escrever |
| Fórmulas no Excel | Abrir com `data_only=False` para preservar fórmulas |
| Datas com timestamp | Usar `pd.to_datetime(..., format='mixed').dt.date` |
| Fórmulas com `;` vs `,` | Normalizar com `.replace(";", ",")` antes de inserir |
| Memória após processar | Sempre `del wb` + `gc.collect()` no final |
| Arquivo `.xlsm` (com macros) | `openpyxl` lê, mas **não preserva macros** ao salvar |

---

## Referência das Funções do Projeto Original

Para consultar a implementação completa de cada função, veja:
`excel_automation/app.py` — Cliente: Zen / Faturamento
