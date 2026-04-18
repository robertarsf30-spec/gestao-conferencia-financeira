import streamlit as st

st.set_page_config(page_title="Conferência Cartão", layout="wide")

if 'autenticado' not in st.session_state or not st.session_state.autenticado:
    st.stop()

st.title("💳 Conferência Cartão (Cielo)")
st.write("⚙️ Filtro aplicado: **PC** (Crédito) e **PD** (Débito)")

u_receb = st.file_uploader("Planilha de Recebíveis Cielo", type=['xlsx'], key="cielo_rec")
u_titul = st.file_uploader("Relatório de Títulos Cielo", type=['xlsx'], key="cielo_tit")

if u_receb and u_titul:
    st.info("Cruzando dados por Valor Bruto e Data (margem de 1 dia).")
