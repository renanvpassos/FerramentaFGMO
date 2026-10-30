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
    is_cloud = os.environ.get("STREAMLIT_SERVER_HEADLESS") == "true" or os.path.exists("/home/adminuser")

    if is_cloud:
        return

    while True:
        time.sleep(10)
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


# --- INTERFACE DO USUÁRIO (STREAMLIT) ---
# Mudado de "wide" para "centered" para centralizar o layout
st.set_page_config(page_title="Tratador de Faturas", layout="centered")

# --- CSS PERSONALIZADO PARA CENTRALIZAR TÍTULOS E COMPONENTES ---
st.markdown("""
    <style>
        .main .block-container {
            max-width: 800px;
            padding-top: 2rem;
            padding-right: 2rem;
            padding-left: 2rem;
            padding-bottom: 2rem;
        }
        h1, p {
            text-align: center;
        }
        .stAlert {
            text-align: center;
        }
        .row-widget.stButton {
            display: flex;
            justify-content: center;
        }
    </style>
""", unsafe_allow_html=True)

st.title("📂 Tratamento Planilha FGMO")
st.info("Insira a planilha de PO (itens) do jeito que ela vem abaixo.")
st.markdown("---")

uploaded_file = st.file_uploader("1. Anexe a planilha original (.xlsx)", type=["xlsx"])


# --- LÓGICA DE MANIPULAÇÃO DA PLANILHA ---
def converter_letra_para_indice(letra):
    """Converte letras de colunas do Excel (A, B, Y...) para índice base-0 do Pandas."""
    letra = letra.upper()
    resultado = 0
    for char in letra:
        resultado = resultado * 26 + (ord(char) - ord('A') + 1)
    return resultado - 1


def tratar_planilha(df_origem):
    df_final = pd.DataFrame()

    # Filtra apenas linhas onde a coluna A original (índice 0) tenha conteúdo válido
    mask = df_origem.iloc[:, 0].notna()
    df_dados = df_origem[mask].copy()

    # 1 e 2. Coluna FATURA: Busca pela coluna nomeada 'FATURA' no arquivo original
    if 'FATURA' in df_dados.columns:
        df_final['FATURA'] = df_dados['FATURA']
    else:
        st.warning("Coluna 'FATURA' não encontrada pelo nome exato. Verifique se o nome do cabeçalho na planilha de origem está correto.")
        df_final['FATURA'] = df_dados.iloc[:, 0]

    # 3. Coluna B: PARTNUMBER (Origem Coluna A + " - " + Origem Coluna Y)
    idx_a = converter_letra_para_indice('A')
    idx_y = converter_letra_para_indice('Y')
    df_final['PARTNUMBER'] = (
        df_dados.iloc[:, idx_a].astype(str) + " - " + df_dados.iloc[:, idx_y].astype(str)
    )

    # 4. Coluna C: QUANTIDADE (Origem Coluna F)
    df_final['QUANTIDADE'] = df_dados.iloc[:, converter_let_para_indice('F')]

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


if uploaded_file:
    # Centralizando o botão de processamento com colunas do Streamlit
    col_vazia1, col_botao, col_vazia2 = st.columns([1, 2, 1])
    with col_botao:
        processar = st.button("🚀 Processar e Gerar Download", use_container_width=True)

    if processar:
        try:
            # Lê mantendo a primeira linha como cabeçalho para identificar a coluna "FATURA"
            df_input = pd.read_excel(uploaded_file)

            if df_input.empty:
                st.error("A planilha anexada não possui dados válidos.")
            else:
                df_resultado = tratar_planilha(df_input)

                st.success("Planilha tratada com sucesso!")
                st.subheader("Pré-visualização dos Dados:")
                st.dataframe(df_resultado.head(10), use_container_width=True)

                # Preparar o arquivo final em memória para o download
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_resultado.to_excel(writer, index=False, sheet_name='Tratado')

                # Centralizando o botão de download também
                col_down1, col_down2, col_down3 = st.columns([1, 2, 1])
                with col_down2:
                    st.download_button(
                        label="📥 Baixar Planilha Tratada",
                        data=output.getvalue(),
                        file_name="Planilha_Tratada.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
        except Exception as e:
            st.error(f"Erro ao processar o arquivo: {e}")
