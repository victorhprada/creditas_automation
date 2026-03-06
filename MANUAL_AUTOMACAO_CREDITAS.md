# Manual de Uso e Auditoria - Automacao Creditas

## 1) Objetivo

Este documento descreve o funcionamento da automacao implementada em `app.py`, incluindo:

- entradas exigidas;
- abas e colunas utilizadas;
- regras de copia, filtro e transformacao;
- logica de criacao da nova aba mensal;
- evidencias para auditoria.

A aplicacao foi desenvolvida em Streamlit e processa arquivos Excel via `openpyxl`.

---

## 2) Escopo do Processo

A automacao executa **duas etapas principais** sobre o arquivo BASE:

1. **Etapa 1**  
   Copia dados da aba `Apoio | Originacao e Repasse` (arquivo Parceiro) para a aba `CREDITAS BASE` (arquivo BASE), e arrasta formulas/estilos nas colunas `R` a `V`.

2. **Etapa 2**  
   Filtra dados da aba `Historico de relatorios de comi` (arquivo Parceiro) por mes de referencia e copia:
   - para a aba `Parcelas pagas` (arquivo BASE);
   - para uma **nova aba mensal** (ex.: `Jan.26`) no arquivo BASE.

---

## 3) Entradas da Interface

Na tela da aplicacao, o usuario informa:

1. **Arquivo do Parceiro** (`.xlsx`)  
2. **Arquivo BASE** (`.xlsx` ou `.xlsm`)  
3. **Mes de Referencia** (formato `MM-AAAA`, ex.: `01-2026`)  
4. **Mes de Faturamento** (texto livre, ex.: `Janeiro`)  

Acao: botao **Iniciar**.

---

## 4) Abas Obrigatorias

### 4.1 No Arquivo do Parceiro
- `Apoio | Originacao e Repasse`
- `Historico de relatorios de comi`

### 4.2 No Arquivo BASE
- `CREDITAS BASE`
- `Parcelas pagas`

Se alguma aba obrigatoria nao existir, o processo e interrompido com erro.

---

## 5) Regras de Colunas e Transformacoes

## 5.1 Etapa 1 - Originação para BASE

### Origem
- Aba: `Apoio | Originacao e Repasse`
- Colunas lidas: **A a Q** (`1` a `17`)
- Linhas: a partir da linha `2` (ignora cabecalho)

### Destino
- Aba: `CREDITAS BASE`
- Insercao: primeira linha vazia real, identificada pela coluna `A`

### Filtro pre-copia (coluna M)
- Antes da copia de `A:Q`, a automacao calcula o **mes anterior** ao valor digitado em **Mes de Faturamento**;
  - Ex.: Faturamento `Fevereiro` -> mes anterior = `Janeiro` (mes 1)
  - Ex.: Faturamento `Janeiro` -> mes anterior = `Dezembro` (mes 12)
- A automacao acessa a coluna `M` (13a coluna) da aba `Apoio | Originacao e Repasse` e filtra somente as linhas cujo mes seja igual ao mes anterior calculado;
- A coluna `M` aceita datas reais do Excel, nomes de mes por extenso ou abreviados, e formatos com barras (`dd/mm/aaaa`).

### Regras aplicadas
- Somente as linhas aprovadas no filtro da coluna `M` sao copiadas (colunas `A:Q`);
- Linhas totalmente vazias sao descartadas;
- Preserva `number_format` por celula (datas, moeda, CNPJ etc.);
- A quantidade de linhas copiadas reportada no log representa apenas as linhas que passaram no filtro.

### Pos-processamento na BASE
Se houve linhas copiadas, as colunas **R a V** (`18` a `22`) sao preenchidas para as novas linhas com base na linha anterior existente:
- copia estilo completo da linha modelo;
- se a celula de origem for formula, a referencia e traduzida para a nova linha;
- se for texto/valor fixo, replica o mesmo valor.

---

## 5.2 Etapa 2 - Historico filtrado

### Origem
- Aba: `Historico de relatorios de comi`
- Colunas lidas: **A a Q** (`1` a `17`)
- Linhas: a partir da linha `2`

### Filtro principal
- Campo de filtro: coluna `Q` (17a coluna)
- O input `MM-AAAA` e convertido para `01/MM/AAAA`
  - Ex.: `01-2026` -> `01/01/2026`
- A linha e copiada apenas se valor padronizado da coluna `Q` for igual ao mes alvo.

### Destinos simultaneos
Para cada linha aprovada no filtro, copia para:
1. `Parcelas pagas`
2. Aba mensal (nova ou ja existente, ex.: `Jan.26`)

### Transformacao de coluna P
- Coluna `P` (16a coluna) e convertida para nome do mes por extenso quando possivel
  (ex.: `1`/`01`/data com mes 1 -> `Janeiro`).

---

## 5.3 Formulas adicionadas apos copia da Etapa 2

### Na aba `Parcelas pagas`
- Coluna `R` (18): `=N{linha}/M{linha}` com formato percentual (`0.00%`)
- Coluna `S` (19): recebe o texto informado em **Mes de Faturamento**

### Na aba mensal (`Jan.26`, `Fev.26`, etc.)
- Coluna `R` (18): `=N{linha}/M{linha}` com formato percentual (`0.00%`)
- Coluna `S` (19): `=M{linha}*3.5%` com formato monetario (`"R$" #,##0.00`)

---

## 6) Logica de Criacao da Nova Aba Mensal

Nome da aba mensal:
- Entrada: `MM-AAAA`
- Saida: `AbrevMes.AA`
  - Ex.: `01-2026` -> `Jan.26`
  - Ex.: `10-2026` -> `Out.26`

Comportamento:
- Se a aba mensal **nao existe** no arquivo BASE:
  - cria a aba;
  - copia o cabecalho da linha `1` (colunas `A:Q`) da aba `Historico de relatorios de comi`;
  - preserva estilo do cabecalho.
- Se a aba mensal **ja existe**:
  - reutiliza a aba existente;
  - adiciona novos dados a partir da ultima linha preenchida (coluna `A`).

---

## 7) Evidencias de Execucao (Logs)

Durante a execucao, a interface mostra mensagens de status, por exemplo:

- leitura dos arquivos na memoria;
- validacao das abas;
- quantidade de linhas copiadas na Etapa 1;
- aplicacao de formulas R:V;
- criacao/reuso da aba mensal;
- quantidade de linhas historicas copiadas;
- geracao do arquivo final para download;
- mensagens de erro em caso de falha.

---

## 8) Saida Gerada

Ao concluir sem erro, a aplicacao disponibiliza download de:

- `CREDITAS_BASE_ATUALIZADA.xlsx`

Este arquivo contem as alteracoes das duas etapas no workbook BASE.

---

## 9) Controles de Auditoria Recomendados

Para uso de auditoria, recomenda-se registrar por execucao:

1. data/hora da execucao;
2. nome dos arquivos de entrada;
3. mes de referencia e mes de faturamento informados;
4. quantidade de linhas copiadas na Etapa 1;
5. quantidade de linhas copiadas na Etapa 2;
6. nome da aba mensal utilizada/criada;
7. status final (sucesso/erro) e mensagem de erro (se houver).

---

## 10) Principais Regras de Negocio Resumidas

- Na Etapa 1, a coluna `M` da origem e filtrada pelo mes anterior ao mes de faturamento informado;
- Copia sempre inicia da primeira linha vazia real (coluna `A`);
- Linhas totalmente vazias sao descartadas;
- Filtro mensal considera igualdade com data no formato `DD/MM/AAAA`;
- Coluna `P` pode ser normalizada para mes por extenso;
- Formulas de comissionamento sao aplicadas automaticamente nas colunas `R` e `S` conforme a aba de destino.

---

## 11) Limitacoes Conhecidas

- O filtro de mes depende da padronizacao correta da coluna `Q`;
- Entradas fora do formato esperado (`MM-AAAA`) podem gerar baixa efetividade do filtro;
- A qualidade da saida depende da consistencia estrutural das abas de origem.

---