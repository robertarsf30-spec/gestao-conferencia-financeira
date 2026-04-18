import streamlit as st

if not st.session_state.get('autenticado'):
    st.error("Faça login no início.")
    st.stop()

st.title("💵 Conferência de Caixa")
st.file_uploader("Planilha Cristalina", type=['xlsx'], key="caixa_up")
