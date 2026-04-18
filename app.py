import streamlit as st

st.set_page_config(page_title="Gestão Conferência Financeira", layout="wide")

if 'autenticado' not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    st.title("🔐 Painel de Gestão")
    with st.form("login"):
        senha = st.text_input("Senha do Sistema:", type="password")
        if st.form_submit_button("Entrar"):
            if senha == "1234":
                st.session_state.autenticado = True
                st.rerun()
            else:
                st.error("Senha incorreta.")
    st.stop()

st.title("🚀 Bem-vindo ao Sistema de Gestão Financeira")
st.markdown("---")
st.write("Selecione um dos módulos especializados no menu lateral para iniciar a conferência.")

if st.sidebar.button("🚪 Sair"):
    st.session_state.autenticado = False
    st.rerun()
