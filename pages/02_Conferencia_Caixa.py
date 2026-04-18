import streamlit as st

st.set_page_config(page_title="Conferência de Caixa", layout="wide")

if 'autenticado' not in st.session_state or not st.session_state.autenticado:
    st.error("🔒 Por favor, faça login na página inicial.")
    st.stop()

st.title("💵 Conferência de Caixa")
st.info("Regra: O saldo inicial deve ser igual ao saldo final do dia anterior.")

u_pdf = st.file_uploader("Upload Boletim de Caixa (PDF)", type=['pdf'], key="caixa_pdf")
u_exc = st.file_uploader("Upload Planilha de Conferência (Excel)", type=['xlsx'], key="caixa_xlsx")

if u_pdf and u_exc:
    st.success("Arquivos carregados. Iniciando validação de saldos...")
