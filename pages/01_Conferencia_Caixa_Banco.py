import streamlit as st
import pandas as pd

if not st.session_state.get('autenticado'):
    st.stop()

st.title("🔍 Conferência Caixa / Banco")
st.info("Identificação de divergências entre Relatório do Sistema e Extrato Bancário.")

col1, col2 = st.columns(2)
u_sis = col1.file_uploader("Relatório Sistema (PDF/Excel)", type=['pdf', 'xlsx'], key="b_sis")
u_ext = col2.file_uploader("Extrato Bancário (PDF/Excel)", type=['pdf', 'xlsx'], key="b_ext")

if u_sis and u_ext:
    st.subheader("⚠️ Divergências Identificadas")
    # Aqui o código fará o cruzamento (merge) dos dados
    st.warning("Analisando lançamentos... Verificando datas e valores correspondentes.")
