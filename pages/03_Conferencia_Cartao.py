import streamlit as st
import pandas as pd

st.set_page_config(page_title="Conferência Cartão", layout="wide")

if 'autenticado' not in st.session_state or not st.session_state.autenticado:
    st.error("🔒 Por favor, faça login na página inicial.")
    st.stop()

st.title("💳 Conferência Cartão (Cielo)")
st.divider()

# Upload dos arquivos
u_receb = st.file_uploader("1. Planilha de Recebíveis Cielo (Excel)", type=['xlsx'])
u_titul = st.file_uploader("2. Relatório de Títulos (Excel)", type=['xlsx'])

if u_receb and u_titul:
    df_receb = pd.read_excel(u_receb)
    df_titul = pd.read_excel(u_titul)
    
    # Filtro automático solicitado
    # Ajustaremos os nomes das colunas 'Tipo' conforme sua resposta
    credito = df_receb[df_receb.iloc[:, 1].str.contains('PC', na=False)]
    debito = df_receb[df_receb.iloc[:, 1].str.contains('PD', na=False)]
    
    col1, col2 = st.columns(2)
    col1.metric("Total Crédito (PC)", f"R$ {credito.iloc[:, -1].sum():,.2f}")
    col2.metric("Total Débito (PD)", f"R$ {debito.iloc[:, -1].sum():,.2f}")
    
    st.warning("⚠️ Lógica de cruzamento por valor bruto aguardando confirmação das colunas.")
