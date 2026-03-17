✍️ TW: DOCUMENTAÇÃO — Automação Creditas (app.py)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# Automação de Processamento de Benefícios e Comissionamento
> Versão: 1.1.0 | Data: 03/2026 | Autor: main | Status: ✅ Entregue

---

## 🎯 Por que isso existe?

O time precisava tirar da mão o trabalho manual de consolidar relatórios de parceiros em planilhas grandes do Excel.
Especificamente:

- Processar originação / repasse copiando linhas relevantes da aba de apoio do parceiro para a aba `CREDITAS BASE`, mantendo fórmulas e formatações.
- Filtrar o histórico de comissionamento por mês de referência (coluna Q) e gerar, em um único passo:
  - atualizações na aba `Parcelas pagas` (com fórmulas adicionais), e
  - uma **aba mensal nova** (ex: `Fev.26`) com um recorte pronto para análise.
- Copiar o **histórico de antecipo** da aba `Histórico Antecipo` do parceiro (filtro na coluna G pelo mês anterior ao de referência) para a aba `ANTECIPO` da base, preenchendo colunas K–M (valor fixo 2,75, mês numérico e nome do mês por extenso).

Regras de negócio importantes:

- Para originação, o filtro na coluna M usa o **mês anterior ao faturamento**.
- Para histórico, o filtro na coluna Q usa o **mês anterior ao mês de referência digitado**, mas o **nome da aba** permanece o do mês de referência (ex: aba `Fev.26` contendo dados de janeiro).
- Para antecipo, o filtro na coluna G considera **mês e ano** do mês anterior ao de referência; as fórmulas em L e M usam ano dinâmico (extraído do filtro).

O objetivo é reduzir erros manuais, padronizar o processo e permitir que qualquer pessoa do time consiga gerar planilhas atualizadas a partir de dois arquivos de entrada.

---

## 🧠 Linha de raciocínio

### Problema identificado

- Planilhas grandes (10MB+) eram atualizadas manualmente, com risco alto de erro humano.
- Regras de datas (mês de faturamento, mês de referência, mês anterior) eram aplicadas “de cabeça”.
- A criação de abas mensais e dos campos derivados (percentuais, comissões) era repetitiva e demorada.

O script precisava:

- Entender múltiplos formatos de data vindos do Excel.
- Aplicar corretamente regras de “mês anterior”.
- Copiar dados preservando estilos e fórmulas.
- Automatizar a criação de nova aba mensal e o cálculo de campos derivados.

### Abordagens consideradas

**Opção A — Tudo via Excel (Power Query / fórmulas complexas)** — *descartada*  
- Como funcionaria: queries e fórmulas na própria planilha para filtrar por mês, gerar abas e cálculos automáticos.  
- Por que foi descartada:
  - Dificuldade de manutenção por quem não domina Power Query.
  - Mais frágil quando o layout da planilha do parceiro muda.
  - Difícil de versionar e revisar em Git.

**Opção B — Script “flat” só copiando células por coordenada** — *descartada*  
- Como funcionaria: loops simples de linha/coluna, com regras de negócio espalhadas e acopladas à manipulação de células.  
- Por que foi descartada:
  - Baixa legibilidade.
  - Alto risco de regressão ao ajustar detalhes de mês, coluna ou formato de data.
  - Mistura forte entre regra de negócio e detalhes de Excel.

**Opção C — Camada de funções de regra + camada de orquestração Streamlit** — ✅ *escolhida*  
- Como funciona:
  - Funções puras para normalizar textos, interpretar datas e calcular “mês anterior”.
  - Funções específicas para:
    - copiar originação (`copiar_originacao_para_base`),
    - preencher fórmulas (`preencher_formulas_colunas_r_v`),
    - copiar histórico para duas abas (`copiar_historico_filtrado`),
    - copiar antecipo e preencher K–M (`copiar_antecipo_para_base`),
    - gerar nome de aba mensal (`gerar_nome_aba_mes`).
  - A camada Streamlit apenas orquestra: lê arquivos, chama funções de negócio, exibe logs e oferece o download.
- Por que foi escolhida:
  - Mais fácil de evoluir quando mudar a regra de mês ou o layout das planilhas.
  - Separação clara entre UI (Streamlit) e regras de Excel.
  - Melhor capacidade de teste das funções isoladas.
- Trade‑offs conscientes:
  - Toda a lógica ainda está em um único arquivo (`app.py`).
  - As colunas são referenciadas por índice (ex: `row_values[12]` → coluna M), o que exige cuidado em alterações de layout.

---

## 🏗️ Arquitetura da solução

### Fluxo de dados

```text
[Usuário] 
   ↓ (upload de 2 arquivos + meses)
[Streamlit UI (form)] 
   ↓
[openpyxl carrega workbooks em memória]
   ↓
[Funções de regra de negócio]
   - normalização e parsing de datas/textos
   - filtro pela coluna M (originação)
   - filtro pela coluna Q (histórico)
   - criação de nova aba mensal
   - aplicação de fórmulas R–V e cálculos adicionais
   - filtro pela coluna G (antecipo) e preenchimento de K–M
   ↓
[Workbook BASE atualizado em memória]
   ↓
[Streamlit disponibiliza arquivo para download]
```

### Componentes criados / modificados

#### `app.py` *(existente, expandido)*

**Responsabilidade:**  
Orquestrar upload de arquivos, aplicar regras de negócio sobre as planilhas do parceiro e da base e disponibilizar um arquivo consolidado para download.

Principais áreas:

- Funções de normalização e interpretação de datas/textos.
- Regras de cópia e filtro para originação (coluna M).
- Regras de cópia e filtro para histórico (coluna Q), incluindo criação de aba mensal.
- Regras de cópia e filtro para antecipo (coluna G) e preenchimento das colunas K–M na aba `ANTECIPO`.
- Camada Streamlit de interface com o usuário.

---

## 🧩 Funções principais (regra de negócio)

### `normalizar_texto(texto)`  
**Responsabilidade:** Padronizar textos (como nomes de meses) removendo acentos, espaços extras e deixando tudo minúsculo.  
**Por que:** Permite comparar meses escritos de formas diferentes (`"fevereiro"`, `"Fev"`, `"Fevereiro"`) sem depender de acentuação ou capitalização.

---

### `obter_mes_numero_por_nome(nome_mes)`  
**Responsabilidade:** Converter nomes/abreviações de meses em número (1–12).  
**Por que:** Facilita a comparação de meses quando o Excel traz o mês como texto. Usada por funções que dependem do “mês anterior”.

---

### `obter_mes_anterior_numero(mes_faturamento)`  
**Responsabilidade:** Dado um mês de faturamento textual (`"Janeiro"`, `"Fev"`, etc.), retornar o número do **mês anterior**.  
**Por que:** A regra de originação (coluna M) é baseada no mês anterior ao faturamento, não no próprio mês.

---

### `extrair_mes_coluna_m(valor)`  
**Responsabilidade:** Interpretar o valor da coluna M (que pode ser `date`, `datetime` ou string em vários formatos) e devolver o número do mês.  
**Por que:** O Excel pode armazenar datas de maneiras diferentes; encapsular esse parsing deixa o filtro em `copiar_originacao_para_base` mais limpo e tolerante a variações.

---

### `calcular_mes_anterior(mes_referencia)`  
**Responsabilidade:** Receber `MM-AAAA` (ex: `02-2026`) e retornar o mês anterior no mesmo formato (`01-2026`).  
**Por que:**  
- O usuário pensa em termos de mês de referência (ex: `Fev/26`).  
- A regra de histórico quer os dados do **mês anterior** (ex: `Jan/26`).  
- A aba gerada deve manter o nome do mês de referência (ex: aba `Fev.26` com dados de janeiro).

---

### `encontrar_ultima_linha(ws, coluna_referencia=1)`  
**Responsabilidade:** Encontrar a última linha realmente preenchida considerando uma coluna de referência (default: A).  
**Por que:** Garante que novos dados sejam inseridos logo após o último registro real, sem ser enganado por formatação vazia no final da planilha.

---

### `copiar_originacao_para_base(ws_parceiro, ws_base, mes_faturamento)`  
**Responsabilidade:**  
- Filtrar a aba do parceiro pela **coluna M**, considerando o mês anterior ao faturamento.  
- Copiar as colunas A–Q para `CREDITAS BASE`, preservando formatação numérica.

**Por que:** Automatiza a carga de originação/repasse respeitando a regra de mês anterior ao faturamento.

**Saída:**  
Retorna `(linha_inicio, linha_fim, registros_copiados)` para que outras funções saibam onde aplicar fórmulas.

---

### `preencher_formulas_colunas_r_v(ws_base, linha_inicio, linha_fim)`  
**Responsabilidade:**  
- Usar a linha anterior às novas (`linha_inicio - 1`) como “molde”.  
- Replicar estilo completo e fórmulas das colunas R–V nas linhas recém inseridas.

**Por que:**  
Ao adicionar novas linhas na base, as fórmulas de R–V precisam ser “arrastadas” como se o usuário tivesse puxado a alça no Excel.

---

### `copiar_historico_filtrado(ws_origem, ws_destino, ws_nova_aba, mes_filtro, mes_faturamento)`  
**Responsabilidade:**  
Aplicar o filtro de mês na aba de histórico do parceiro (coluna Q) e copiar dados para:

1. `ws_destino` (aba `Parcelas pagas`), com:
   - colunas A–Q;
   - coluna R = `N/M` (percentual);
   - coluna S = `mes_faturamento`.
2. `ws_nova_aba` (aba mensal), com:
   - colunas A–Q;
   - coluna R = `N/M` (percentual);
   - coluna S = `M * 3.5%` (com formato moeda).

**Por que:**  
Centraliza a lógica de “copiar para duas abas diferentes com regras diferentes de fórmula”, evitando duplicação na camada de orquestração.

---

### `copiar_antecipo_para_base(ws_hist_antecipo, ws_base_antecipo, mes_referencia)`  
**Responsabilidade:**  
- Filtrar a aba `Histórico Antecipo` do parceiro pela **coluna G** (datas no formato DD/MM/YYYY), considerando **mês e ano** do mês anterior ao `mes_referencia`.  
- Copiar as colunas A–J para a aba `ANTECIPO` da base na primeira linha vazia.  
- Preencher coluna K com o valor **2,75** usando o formato de moeda já existente na coluna (copiado da última linha preenchida).  
- Preencher coluna L com a fórmula `=MONTH(G{linha})`.  
- Preencher coluna M com a fórmula `=TEXT(DATE({ano},L{linha},1),"mmmm")`, onde o ano é dinâmico (extraído do mês anterior ao de referência).

**Por que:** Automatiza a carga de antecipo e garante que as colunas K–M fiquem preenchidas com valor fixo e fórmulas corretas. O ano na fórmula da coluna M acompanha o período filtrado.

**Detalhe técnico:** As fórmulas são escritas em sintaxe inglesa (vírgulas); o openpyxl exige isso para que o Excel as reconheça ao abrir o arquivo (o Excel converte para ponto e vírgula conforme o locale).

---

### `gerar_nome_aba_mes(mes_referencia)`  
**Responsabilidade:** Converter `MM-AAAA` para o formato de aba `AbrevMes.AA` (ex: `02-2026` → `Fev.26`).  
**Por que:** Padronizar o nome das abas mensais, facilitando navegação e evitando variações de nomenclatura.

---

## 🖥️ Interface Streamlit

### Formulário principal

Campos:

- `arquivo_parceiro`: upload do Excel do parceiro (`Benefits_Comissionamento_Sênior.xlsx`).
- `arquivo_base`: upload do Excel base (`Acompanhamento creditas base.xlsx`).
- `mes_referencia`: texto `MM-AAAA` para o mês de referência do comissionamento (usado na coluna Q, no nome da aba nova e no filtro da coluna G do antecipo).
- `mes_faturamento`: nome do mês de faturamento (usado na coluna M e como rótulo em `Parcelas pagas`).

Botão:

- `Iniciar`: dispara o pipeline completo.

### Fluxo do processamento

1. Valida se os dois arquivos foram enviados.
2. Carrega workbooks:
   - parceiro: `data_only=True` (ignora fórmulas pesadas);
   - base: `data_only=False` (preserva fórmulas).
3. Valida e processa **Etapa 1 — Originação**:
   - Usa `Apoio | Originação e Repasse` e `CREDITAS BASE`.
   - Chama `copiar_originacao_para_base`.
   - Se houver linhas copiadas, aplica `preencher_formulas_colunas_r_v`.
4. Valida e processa **Etapa 2 — Histórico / Parcelas**:
   - Usa `Histórico de relatórios de comi` e `Parcelas pagas`.
   - Gera `nome_nova_aba` com `gerar_nome_aba_mes(mes_referencia)` (ex: `Fev.26`).
   - Cria ou reutiliza a aba mensal (`ws_nova`) e replica cabeçalho.
   - Calcula `mes_filtro = calcular_mes_anterior(mes_referencia)` (ex: `02-2026` → `01-2026`).
   - Chama `copiar_historico_filtrado` para preencher `Parcelas pagas` e a aba mensal.
5. Valida e processa **Etapa 3 — Antecipo**:
   - Usa `Histórico Antecipo` (parceiro) e `ANTECIPO` (base).
   - Chama `copiar_antecipo_para_base`: filtra pela coluna G (mês e ano do mês anterior ao de referência), copia A–J e preenche K (2,75), L (`=MONTH(G...)`) e M (`=TEXT(DATE(ano,L,1),"mmmm")`).
6. Gera o arquivo final em memória e exibe o botão de download.

---

## ⚠️ Pontos de atenção para o futuro

- **Dependência forte do layout das planilhas**  
  Índices de coluna (ex: `row_values[12]` para M e `row[16]` para Q) estão hard‑coded.  
  Se o parceiro mover colunas, a automação pode quebrar silenciosamente ou gerar dados incorretos.  
  → Sugestão: centralizar mapeamentos em um dicionário configurável, talvez por parceiro/versão de layout.

- **Parsing de datas**  
  Já há suporte a vários formatos, mas ainda pode falhar em casos mais exóticos de data/hora.  
  → Se surgirem erros de filtragem, revisar `extrair_mes_coluna_m` e a normalização de `valor_mes_celula` na coluna Q.

- **Arquivo único grande (`app.py`)**  
  A lógica de negócio e a UI estão no mesmo módulo.  
  → Futuro natural: quebrar em módulos (`regras_excel.py`, `ui_streamlit.py`) e adicionar testes unitários.

- **Taxa fixa de 3,5%**  
  A fórmula da coluna S na aba mensal usa `*3.5%` fixo.  
  → Se esse percentual variar por parceiro/campanha, transformar em configuração (entrada de usuário ou tabela de parâmetros).

- **Fórmulas em português vs. openpyxl**  
  O openpyxl espera fórmulas em sintaxe inglesa (vírgulas). Fórmulas escritas com ponto e vírgula podem não ser reconhecidas e deixar células vazias.  
  → Manter vírgulas no código; o Excel converte para o locale ao abrir.

---

## 🧪 Como testar manualmente

### Pré-condições

- Arquivo do parceiro com:
  - Aba `Apoio | Originação e Repasse`, coluna M com datas/meses válidos.
  - Aba `Histórico de relatórios de comi`, coluna Q com datas no formato compatível (ex: `DD/MM/AAAA`).
  - Aba `Histórico Antecipo`, coluna G com datas no formato DD/MM/YYYY.
- Arquivo base com:
  - Aba `CREDITAS BASE` contendo pelo menos uma linha com fórmulas em R–V.
  - Aba `Parcelas pagas` criada.
  - Aba `ANTECIPO` criada (com pelo menos uma linha com formatação na coluna K, se houver dados).

### Cenário principal

1. Abrir o app Streamlit.
2. Fazer upload do arquivo do parceiro e do arquivo base.
3. Preencher:
   - `mes_referencia` = `02-2026`.
   - `mes_faturamento` = `Fevereiro`.
4. Clicar em **Iniciar**.

**Resultados esperados:**

- Etapa 1:
  - Log indica filtro “pelo mês anterior a 'Fevereiro'”.
  - Apenas linhas com coluna M representando janeiro são copiadas para `CREDITAS BASE`.
  - Novas linhas têm fórmulas R–V corretamente replicadas.

- Etapa 2:
  - Criada (se não existir) a aba `Fev.26`.
  - Log indica filtro por `01-2026` (mês anterior a `02-2026`).
  - Apenas linhas da aba de histórico cuja coluna Q corresponde a `01/01/2026` são copiadas.
  - Em `Parcelas pagas`:
    - Coluna R = `N/M` (percentual).
    - Coluna S = `Fevereiro` (mês de faturamento).
  - Na aba `Fev.26`:
    - Coluna R = `N/M` (percentual).
    - Coluna S = `M * 3.5%` com formato moeda.

- Etapa 3:
  - Log indica filtro na coluna G pelo mês anterior (ex: `01-2026`).
  - Apenas linhas da aba `Histórico Antecipo` cuja coluna G tem mês/ano do mês anterior são copiadas para `ANTECIPO`.
  - Novas linhas têm coluna K = 2,75 (moeda), L = `=MONTH(G{linha})`, M = `=TEXT(DATE(2026,L{linha},1),"mmmm")` (ano dinâmico).

### Edge cases importantes

- `mes_referencia` fora do padrão `MM-AAAA`  
  → Pode quebrar o cálculo de `calcular_mes_anterior` e da `data_alvo_str`.  
  → Melhorar mensagens de erro caso isso se torne frequente.

- `mes_faturamento` inválido (não reconhecido por `obter_mes_numero_por_nome`)  
  → Dispara `ValueError` explicando o problema.  
  → Usuário deve ser orientado a usar nomes/abreviações suportados.

---

## ✅ Checklist do Tech Writer

- [x] Documentação explica o **porquê**, não só o **quê**.
- [x] Linha de raciocínio clara para alguém que não participou da implementação.
- [x] Pontos de atenção e cenários de teste descritos para o “eu do futuro”.

