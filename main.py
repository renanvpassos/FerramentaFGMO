import streamlit as st
import pandas as pd
from io import BytesIO
import os
import signal
import threading
import time
from streamlit.runtime import get_instance


# --- FUNÇÃO DE AUTO-ENCERRAMENTO INTELIGENTE ---
def monitorar_conexoes():
    """
    Verifica se há sessões ativas.
    SÓ encerra o processo se estiver rodando localmente como EXECUTÁVEL.
    Se detectar o ambiente do Streamlit Cloud (GitHub), deixa o servidor em paz.
    """
    # Detecta se está na nuvem do Streamlit (Linux/Adminuser)
    is_cloud = os.environ.get("STREAMLIT_SERVER_HEADLESS") == "true" or os.path.exists("/home/adminuser")

    if is_cloud:
        return  # Aborta o monitoramento para não derrubar o deploy online

    # Se for o executável local do Windows, mantém o monitoramento
    while True:
        time.sleep(10)  # Checa a cada 10 segundos
        try:
            runtime = get_instance()
            sessions = runtime._session_mgr.list_active_sessions()
            if not sessions:
                os.kill(os.getpid(), signal.SIGTERM)
        except:
            pass


# Inicia o monitoramento em segundo plano (Thread Daemon)
t = threading.Thread(target=monitorar_conexoes, daemon=True)
t.start()


# --- LÓGICA DE MANIPULAÇÃO DA PLANILHA ---
def converter_letra_para_indice(letra):
    """Converte letras de colunas do Excel (A, B, Y...) para índice base-0 do Pandas."""
    letra = letra.upper()
    resultado = 0
    for char in letra:
        resultado = resultado * 26 + (ord(char) - ord('A') + 1)
    return resultado - 1


def tratar_planilha(df_origem, num_fatura):
    df_final = pd.DataFrame()

    # Filtra apenas linhas onde a coluna A original (índice 0) tenha conteúdo válido
    mask = df_origem.iloc[:, 0].notna()
    df_dados = df_origem[mask].copy()

    # 1 e 2. Coluna A: FATURA (Preenchida com o número fornecido pelo usuário)
    df_final['FATURA'] = [num_fatura] * len(df_dados)

    # 3. Coluna B: PARTNUMBER (Origem Coluna A + " - " + Origem Coluna Y)
    idx_a = converter_letra_para_indice('A')
    idx_y = converter_letra_para_indice('Y')
    df_final['PARTNUMBER'] = (
            df_dados.iloc[:, idx_a].astype(str) + " - " + df_dados.iloc[:, idx_y].astype(str)
    )

    # 4. Coluna C: QUANTIDADE (Origem Coluna F)
    df_final['QUANTIDADE'] = df_dados.iloc[:, converter_letra_para_indice('F')]

    # 5. Coluna D: PESOUNITARIO (Origem Coluna G)
    df_final['PESOUNITARIO'] = df_dados.iloc[:, converter_letra_para_indice('G')]

    # 6. Coluna E: PRECOUNITARIO (Origem Coluna K)
    df_final['PRECOUNITARIO'] = df_dados.iloc[:, converter_letra_para_indice('K')]

    # 7. Coluna F: MOEDA (Origem Coluna U)
    df_final['MOEDA'] = df_dados.iloc[:, converter_letra_para_indice('U')]

    # 8. Coluna G: INCOTERM (Preenchido com o texto fixo 'CIP')
    df_final['INCOTERM'] = 'CIP'

    # 9. Coluna H: UNIDADE (Origem Coluna D)
    df_final['UNIDADE'] = df_dados.iloc[:, converter_letra_para_indice('D')]

    return df_final


# --- INTERFACE DO USUÁRIO (STREAMLIT) ---
st.set_page_config(page_title="Tratador de Faturas", layout="wide")
st.title("📂 Tratamento planilha FGMO")
st.info("Insira a planilha de PO (itens) do jeito que ela vem abaixo.")
st.markdown("---")

col1, col2 = st.columns(2)
with col1:
    uploaded_file = st.file_uploader("1. Anexe a planilha original (.xlsx)", type=["xlsx"])
with col2:
    num_fatura = st.text_input("2. Digite o número da FATURA:")

if uploaded_file and num_fatura:
    if st.button("🚀 Processar e Gerar Download"):
        try:
            # skiprows=1 pula a primeira linha (Cabeçalho da planilha original)
            # header=None força o Pandas a indexar por números (0, 1, 2...)
            df_input = pd.read_excel(uploaded_file, skiprows=1, header=None)

            if df_input.empty:
                st.error("A planilha anexada não possui dados válidos após a linha 1.")
            else:
                df_resultado = tratar_planilha(df_input, num_fatura)

                st.success("Planilha tratada com sucesso!")
                st.subheader("Pré-visualização dos Dados:")
                st.dataframe(df_resultado.head(10))

                # Preparar o arquivo final em memória para o download
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_resultado.to_excel(writer, index=False, sheet_name='Tratado')

                st.download_button(
                    label="📥 Baixar Planilha Tratada",
                    data=output.getvalue(),
                    file_name=f"Fatura_{num_fatura}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Erro ao processar o arquivo: {e}")
