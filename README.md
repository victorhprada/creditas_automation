# 📊 Automação Creditas — Processador de Benefícios e Comissionamento

> **Um único fluxo:** dois arquivos Excel + mês de referência → planilha base consolidada, com originação, histórico de parcelas e antecipo — sem trabalho manual repetitivo.

[![Python 3.11](https://img.shields.io/badge/python-3.11-blue.svg)](https://www.python.org/downloads/)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.32-FF4B4B?logo=streamlit)](https://streamlit.io/)

---

## 🎯 O que este projeto resolve

Consolidação **automática** de relatórios de parceiros na planilha base da Creditas: originação/repasse, comissionamento (parcelas pagas + aba mensal) e histórico de antecipo. Tudo filtrado por **mês de referência** e **mês anterior**, com fórmulas e formatações preservadas.

---

## 😤 Dores (antes da automação)

| Dor | Impacto |
|-----|--------|
| **Planilhas gigantes (10MB+)** atualizadas à mão | Horas perdidas, risco alto de erro humano |
| **Regras de data “de cabeça”** (mês de faturamento vs. mês anterior vs. mês de referência) | Filtros errados, dados no mês equivocado |
| **Cópia manual** de originação, histórico e antecipo para várias abas | Processo repetitivo, cansativo e propenso a esquecimentos |
| **Fórmulas e formatações** que precisavam ser “arrastadas” linha a linha | Inconsistência (CNPJ em notação científica, percentuais, moeda) |
| **Abas mensais** (ex: Fev.26) criadas e preenchidas manualmente | Atrasos e divergência entre relatório do parceiro e base interna |

O time precisava de **um único ponto de entrada**: enviar o arquivo do parceiro + a planilha base + o mês desejado e receber de volta a base atualizada, padronizada e pronta para análise.

---

## ⚙️ Processos (o que a automação faz)

O fluxo é **linear e em 3 etapas**, sempre na mesma ordem:

```
┌─────────────────────────────────────────────────────────────────────────┐
│  ENTRADA                                                                  │
│  • Arquivo do Parceiro (Benefits_Comissionamento_Sênior.xlsx)              │
│  • Arquivo BASE (Acompanhamento creditas base.xlsx)                       │
│  • Mês de Referência Comissionamento (ex: 02-2026)                         │
│  • Mês de Faturamento (ex: Fevereiro)                                     │
└─────────────────────────────────────────────────────────────────────────┘
                                        │
                                        ▼
┌─────────────────────────────────────────────────────────────────────────┐
│  ETAPA 1 — Originação e Repasse                                          │
│  • Aba: Apoio \| Originação e Repasse → CREDITAS BASE                     │
│  • Filtro: coluna M = mês ANTERIOR ao faturamento                         │
│  • Cópia: colunas A–Q + replicação de fórmulas/estilos R–V               │
└─────────────────────────────────────────────────────────────────────────┘
                                        │
                                        ▼
┌─────────────────────────────────────────────────────────────────────────┐
│  ETAPA 2 — Histórico de Comissionamento                                   │
│  • Aba: Histórico de relatórios de comi → Parcelas pagas + nova aba       │
│  • Filtro: coluna Q = mês ANTERIOR ao de referência (ex: 02-2026 → 01)    │
│  • Nome da nova aba: mês de referência (ex: Fev.26), conteúdo = jan/26    │
│  • Fórmulas: R = N/M, S = mês faturamento ou M*3.5% na aba mensal         │
└─────────────────────────────────────────────────────────────────────────┘
                                        │
                                        ▼
┌─────────────────────────────────────────────────────────────────────────┐
│  ETAPA 3 — Antecipo                                                       │
│  • Aba: Histórico Antecipo → ANTECIPO                                     │
│  • Filtro: coluna G (data DD/MM/YYYY) = mês e ano ANTERIOR ao referência  │
│  • Cópia: colunas A–J + preenchimento K = 2,75 (moeda), L = MONTH(G),     │
│           M = nome do mês por extenso (ano dinâmico)                      │
└─────────────────────────────────────────────────────────────────────────┘
                                        │
                                        ▼
┌─────────────────────────────────────────────────────────────────────────┐
│  SAÍDA                                                                    │
│  • CREDITAS_BASE_ATUALIZADA.xlsx para download                            │
└─────────────────────────────────────────────────────────────────────────┘
```

Resumo das **regras de mês**:

- **Originação:** coluna M do parceiro = mês **anterior** ao “Mês de Faturamento”.
- **Histórico (Q):** coluna Q = mês **anterior** ao “Mês de Referência”; nome da aba = mês de referência (ex.: Fev.26 com dados de jan/26).
- **Antecipo (G):** coluna G = **mês e ano** do mês anterior ao de referência; ano usado nas fórmulas da coluna M é dinâmico.

---

## ✅ Soluções (resultados)

| Antes | Depois |
|-------|--------|
| Atualização manual de várias abas e colunas | Um clique: upload + mês → download da base atualizada |
| Risco de filtrar pelo mês errado | Regras de “mês anterior” aplicadas de forma consistente no código |
| Fórmulas e formatos copiados à mão | Replicação automática (R–V na base, K–M no antecipo) |
| Abas mensais e antecipo feitos em planilha | Geração automática de aba mensal (ex: Fev.26) e preenchimento de ANTECIPO |
| Processo difícil de auditar ou repetir | Código versionado (Git), documentado e executável local ou em Streamlit Cloud |

A solução foi implementada com **funções de regra de negócio** (normalização de datas, filtros por coluna, cópia e fórmulas) e uma **interface Streamlit** que só orquestra: valida abas, chama as funções e entrega o arquivo para download. Assim o processo fica previsível, documentável e fácil de evoluir (ex.: novo parceiro ou nova coluna).

---

## 🚀 Desenvolvimento e uso end-to-end

### Pré-requisitos

- **Python 3.11** (recomendado; o projeto usa `runtime.txt` para Streamlit Cloud)
- Dependências: `streamlit`, `openpyxl`

### Instalação e execução local

```bash
# Clone o repositório
git clone https://github.com/victorhprada/creditas_automation.git
cd creditas_automation

# Ambiente virtual (recomendado)
python -m venv venv
# Windows:
venv\Scripts\activate
# Linux/macOS:
# source venv/bin/activate

# Instalar dependências
pip install -r requirements.txt

# Rodar o app
streamlit run app.py
```

O app abre no navegador (em geral em `http://localhost:8501`). Basta fazer upload dos dois arquivos, informar o mês de referência e o mês de faturamento e clicar em **Iniciar**; ao final, baixe `CREDITAS_BASE_ATUALIZADA.xlsx`.

### Deploy (Streamlit Cloud)

- Conecte o repositório ao [Streamlit Community Cloud](https://share.streamlit.io/).
- O projeto já inclui `runtime.txt` (Python 3.11) e `requirements.txt`.
- Após o deploy, acesse a URL gerada e use o mesmo fluxo: upload + meses → download.

### Estrutura do projeto

```text
creditas_automation/
├── app.py                    # Aplicação Streamlit + regras de negócio (originação, histórico, antecipo)
├── requirements.txt          # streamlit, openpyxl
├── runtime.txt              # Python 3.11 (Streamlit Cloud)
├── README.md                # Este arquivo (dores, processos, soluções, uso)
├── docs/
│   └── features/
│       └── automacao-creditas-app.md   # Documentação técnica detalhada (porquê, arquitetura, funções, testes)
└── TEMPLATE_NOVA_AUTOMACAO.md         # Template para novas automações (se aplicável)
```

O desenvolvimento foi feito de forma **end-to-end**: desde a definição das dores e das regras de filtro (mês anterior, colunas M/Q/G), passando pela implementação em Python/Streamlit/openpyxl, até o deploy em nuvem e a documentação técnica em `docs/features/automacao-creditas-app.md`.

---

## 📖 Documentação técnica

Para **porquê** de cada regra, **arquitetura**, **funções** (ex.: `copiar_antecipo_para_base`, `calcular_mes_anterior`), **pontos de atenção** e **como testar** (cenários e edge cases), use:

- **[docs/features/automacao-creditas-app.md](docs/features/automacao-creditas-app.md)** — documentação completa da feature (versão 1.1.0).

---

## 📄 Licença e repositório

- Repositório: [github.com/victorhprada/creditas_automation](https://github.com/victorhprada/creditas_automation)
- Uso interno Creditas; ajuste de licença conforme política do time.
