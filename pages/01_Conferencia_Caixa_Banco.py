import streamlit as st
import pandas as pd

# Configuração da página
st.set_page_config(page_title="Conferência Caixa/Banco", layout="wide")

# Verificação de Login
if 'autenticado' not in st.session_state or not st.session_state.autenticado:
    st.error("🔒 Por favor, faça login na página inicial (app.py) antes de acessar este módulo.")
    st.stop()

st.title("🔍 Conferência Caixa / Banco")
st.subheader("Identificação de divergências entre Relatório do Sistema e Extrato Bancário.")
st.divider()

# Upload de arquivos
col1, col2 = st.columns(2)
with col1:
    u_sis = st.file_uploader("Relatório Sistema (PDF/Excel)", type=['pdf', 'xlsx'], key="b1")
with col2:
    u_ext = st.file_uploader("Extrato Bancário (PDF/Excel)", type=['pdf', 'xlsx'], key="b2")

if u_sis and u_ext:
    st.info("⚙️ Processando arquivos... A regra de comparação automática será executada aqui.")
    # Aqui entrará a lógica de cruzamento de dados que estamos desenvolvendo
