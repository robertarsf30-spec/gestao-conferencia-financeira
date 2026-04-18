import streamlit as st

if not st.session_state.get('autenticado'):
    st.error("Faça login no início.")
    st.stop()

st.title("🔍 Conferência Caixa / Banco")
st.file_uploader("Relatórios do Banco", type=['pdf'], key="banco_up")
