# 📊 Automação Creditas - Processador de Benefícios e Comissionamento

Automação desenvolvida em **Streamlit** para processar e consolidar planilhas de benefícios e comissionamento de parceiros. A aplicação filtra, copia e organiza dados de múltiplas abas entre arquivos Excel, aplicando fórmulas e estilos automaticamente.

---

## 🚀 Como Executar

### Pré-requisitos
- Python 3.8+
- Dependências listadas em `requirements.txt`

### Instalação
```bash
pip install -r requirements.txt
```

### Rodando a aplicação
```bash
streamlit run app.py
```

A interface será aberta no navegador. Envie os arquivos e clique em **Iniciar** para processar.

---

## 📋 Entradas Necessárias

| Campo | Descrição | Exemplo |
|---|---|---|
| **Arquivo do Parceiro** | Planilha `.xlsx` enviada pelo parceiro | `Benefits_Comissionamento_Sênior.xlsx` |
| **Arquivo BASE** | Planilha mestra `.xlsx` ou `.xlsm` | `Acompanhamento creditas base.xlsx` |
| **Mês de Referência** | Mês/ano para filtro do histórico | `01-2026` |
| **Mês de Faturamento** | Mês usado como referência para comissionamento | `Janeiro` |

---

## 🏗️ Estrutura de Abas Obrigatórias

### No Arquivo do Parceiro
- `Apoio | Originação e Repasse`
- `Histórico de relatórios de comi`

### No Arquivo BASE
- `CREDITAS BASE`
- `Parcelas pagas`

> ⚠️ Se alguma aba não existir, o processo é interrompido com erro.

---

## ⚙️ Como Funciona — Fluxo de Processamento

A automação executa **duas etapas principais**:

### Etapa 1 — Originação para BASE
1. Calcula o **mês anterior** ao Mês de Faturamento informado
2. Filtra a aba `Apoio | Originação e Repasse` pela **coluna M** (mês)
3. Copia as colunas **A a Q** das linhas aprovadas para a aba `CREDITAS BASE`
4. Aplica fórmulas e estilos nas colunas **R a V** com base na linha modelo

### Etapa 2 — Histórico Filtrado
1. Filtra a aba `Histórico de relatórios de comi` pela **coluna Q** (mês de referência)
2. Copia os dados para **dois destinos simultaneamente**:
   - Aba `Parcelas pagas` (com fórmulas extras)
   - Nova aba mensal (ex: `Jan.26`)
3. Aplica fórmulas de comissionamento em cada destino

---

## 📐 Regras de Negócio

### Regra 1 — Filtro por Mês Anterior (Etapa 1)
A coluna **M** da origem é filtrada pelo **mês anterior** ao mês de faturamento:

| Mês de Faturamento | Mês Alvo (Coluna M) |
|---|---|
| Janeiro | Dezembro (12) |
| Fevereiro | Janeiro (1) |
| Março | Fevereiro (2) |
| ... | ... |

### Regra 2 — Formatos Aceitos na Coluna M
A coluna M aceita:
- Datas reais do Excel
- Nomes de mês por extenso (`Janeiro`, `Fevereiro`...)
- Abreviações (`Jan`, `Fev`, `Mar`...)
- Formatos com barras (`dd/mm/aaaa`)

### Regra 3 — Descarte de Linhas Vazias
Linhas totalmente vazias são **ignoradas** automaticamente em ambas as etapas.

### Regra 4 — Linha de Destino
A cópia sempre inicia na **primeira linha vazia real**, identificada pela coluna A. Células apenas formatadas (sem valor) são desconsideradas.

### Regra 5 — Preservação de Estilos
- `number_format` é preservado em cada célula (datas, moeda, CNPJ, etc.)
- Estilos completos (cor de fundo, bordas, fontes) são copiados nas fórmulas arrastadas

### Regra 6 — Normalização do Mês (Coluna P)
Na Etapa 2, a coluna **P** (16ª coluna) é convertida para nome do mês por extenso:
- `1`, `01` ou data com mês 1 → `Janeiro`
- Funciona para todos os 12 meses

### Regra 7 — Nomenclatura da Aba Mensal
| Entrada | Saída |
|---|---|
| `01-2026` | `Jan.26` |
| `10-2026` | `Out.26` |
| `12-2025` | `Dez.25` |

Se a aba mensal **já existe**, os dados são adicionados a partir da última linha preenchida. Se **não existe**, ela é criada com o cabeçalho copiado da origem.

---

## 📊 Fórmulas Aplicadas

### Na aba `CREDITAS BASE` (Etapa 1)
| Coluna | Lógica |
|---|---|
| **R a V** | Arrastadas da linha modelo anterior, com referências atualizadas |

### Na aba `Parcelas pagas` (Etapa 2)
| Coluna | Fórmula | Formato |
|---|---|---|
| **R** (18) | `=N{linha}/M{linha}` | Percentual (`0.00%`) |
| **S** (19) | Mês de Faturamento (texto) | — |

### Na aba Mensal (ex: `Jan.26`) (Etapa 2)
| Coluna | Fórmula | Formato |
|---|---|---|
| **R** (18) | `=N{linha}/M{linha}` | Percentual (`0.00%`) |
| **S** (19) | `=M{linha}*3.5%` | Monetário (`"R$" #,##0.00`) |

---

## 📤 Saída Gerada

Ao concluir sem erro, a aplicação disponibiliza o download de:

**`CREDITAS_BASE_ATUALIZADA.xlsx`**

Este arquivo contém todas as alterações das duas etapas no workbook BASE.

---

## 🔍 Auditoria

Para rastreabilidade, recomenda-se registrar por execução:

1. Data/hora da execução
2. Nome dos arquivos de entrada
3. Mês de referência e mês de faturamento informados
4. Quantidade de linhas copiadas na Etapa 1
5. Quantidade de linhas copiadas na Etapa 2
6. Nome da aba mensal utilizada/criada
7. Status final (sucesso/erro) e mensagem de erro

---

## 🚧 Limitações Conhecidas

- O filtro de mês depende da **padronização correta** da coluna Q
- Entradas fora do formato `MM-AAAA` podem gerar baixa efetividade do filtro
- A qualidade da saída depende da **consistência estrutural** das abas de origem

---

## 🛠️ Tecnologias Utilizadas

| Tecnologia | Versão | Função |
|---|---|---|
| [Streamlit](https://streamlit.io/) | 1.32.0 | Interface web interativa |
| [openpyxl](https://openpyxl.readthedocs.io/) | 3.1.2 | Leitura e escrita de arquivos Excel |

---

## 📄 Licença

Projeto interno Creditas — uso restrito.
