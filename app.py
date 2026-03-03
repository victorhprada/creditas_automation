import streamlit as st
import openpyxl
from io import BytesIO
import gc

# ---------------------------------------------------------
# SEÇÃO 1: FUNÇÕES DE APOIO E REGRAS DE NEGÓCIO
# ---------------------------------------------------------

def encontrar_ultima_linha(ws, coluna_referencia=1):
    """
    Encontra a verdadeira última linha preenchida baseada na coluna A.
    Previne colar dados no final da planilha por causa de células formatadas vazias.
    """
    for row in range(ws.max_row, 0, -1):
        if ws.cell(row=row, column=coluna_referencia).value is not None:
            return row
    return 1 # Retorna 1 se estiver totalmente vazia (apenas header)

def copiar_originacao_para_base(ws_parceiro, ws_base):
    """
    Copia os dados da coluna A (1) até Q (17) da aba do Parceiro
    para a primeira linha vazia da aba Base.
    """
    linha_destino = encontrar_ultima_linha(ws_base) + 1
    registros_copiados = 0
    
    # iter_rows com values_only=True é crucial para performance com arquivos de 10MB+
    # min_row=2 pula o cabeçalho do parceiro. max_col=16 pega de A até P.
    for row_values in ws_parceiro.iter_rows(min_row=2, max_col=17):
        
        # Validação de segurança: ignora linhas onde todas as células estão vazias
        # (usamos uma compreensão de lista para checar os valores)
        if not any(cell.value is not None and str(cell.value).strip() != "" for cell in row_values):
            continue
            
        for col_idx, cell_origem in enumerate(row_values, start=1):
            cell_destino = ws_base.cell(row=linha_destino, column=col_idx)

            cell_destino.value = cell_origem.value

            # Copia o formato da célula (Isso é o que salva o CNPJ da notação científica,
            # além de preservar formatação de datas e moedas)
            if cell_origem.has_style:
                cell_destino.number_format = cell_origem.number_format
            
        linha_destino += 1
        registros_copiados += 1
        
    return registros_copiados

# ---------------------------------------------------------
# SEÇÃO 2: INTERFACE STREAMLIT
# ---------------------------------------------------------

st.set_page_config(page_title="Automação Creditas", page_icon="📊", layout="centered")
st.title("📊 Processador de Benefícios e Comissionamento")

with st.form("form_processamento"):
    arquivo_parceiro = st.file_uploader("1️⃣ Arquivo do Parceiro (Benefits_Comissionamento_Sênior.xlsx)", type=["xlsx"])
    arquivo_base = st.file_uploader("2️⃣ Arquivo BASE (Acompanhamento creditas base.xlsx)", type=["xlsx", "xlsm"])

    submit = st.form_submit_button("Iniciar", type="primary")

if submit:
    if not arquivo_parceiro or not arquivo_base:
        st.error("⚠️ Por favor, envie as duas planilhas antes de processar.")
    else:
        with st.status("🚀 Processando...", expanded=True) as status:
            try:
                st.write("📥 Lendo arquivos na memória (isso pode levar alguns segundos)...")
                # Parceiro carrega apenas valores (data_only=True) para ignorar fórmulas pesadas
                parceiro_wb = openpyxl.load_workbook(arquivo_parceiro, data_only=True)
                # Base carrega com data_only=False para preservar as fórmulas existentes lá dentro
                base_wb = openpyxl.load_workbook(arquivo_base, data_only=False)

                aba_parceiro_nome = "Apoio | Originação e Repasse"
                aba_base_nome = "CREDITAS BASE"

                st.write("⚙️ Validando abas...")
                if aba_parceiro_nome not in parceiro_wb.sheetnames:
                    raise ValueError(f"Aba '{aba_parceiro_nome}' não encontrada no arquivo do PARCEIRO.")
                if aba_base_nome not in base_wb.sheetnames:
                    raise ValueError(f"Aba '{aba_base_nome}' não encontrada no arquivo BASE.")

                ws_parceiro = parceiro_wb[aba_parceiro_nome]
                ws_base = base_wb[aba_base_nome]

                st.write(f"🔄 Copiando dados de '{aba_parceiro_nome}' para '{aba_base_nome}'...")
                qtd_copiada = copiar_originacao_para_base(ws_parceiro, ws_base)
                st.write(f"✅ {qtd_copiada} linhas copiadas com sucesso!")

                st.write("💾 Gerando arquivo atualizado para download...")
                output = BytesIO()
                base_wb.save(output)
                output.seek(0)

                # Limpeza severa de memória (obrigatório para nosso cenário de 10MB+)
                del parceiro_wb, base_wb, ws_parceiro, ws_base
                gc.collect()

                status.update(label="✅ Processamento Concluído!", state="complete", expanded=False)

                st.download_button(
                    label="📥 Baixar BASE Atualizada",
                    data=output,
                    file_name="CREDITAS_BASE_ATUALIZADA.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

            except Exception as e:
                status.update(label="❌ Erro no Processamento", state="error")
                st.error(f"Ocorreu um erro: {str(e)}")