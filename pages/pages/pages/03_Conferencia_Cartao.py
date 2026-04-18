import streamlit as st

if not st.session_state.get('autenticado'):
    st.stop()

st.title("💳 Conferência de Cartão (Cielo)")
st.info("Base de conferência: PC (Crédito) e PD (Débito).")

u_receb = st.file_uploader("Planilha de Recebíveis Cielo", type=['xlsx'])
u_titul = st.file_uploader("Relatório de Títulos Cielo", type=['xlsx'])

if u_receb and u_titul:
    st.write("✅ Cruzando títulos por Valor Bruto e Data de Lançamento...")
    # Lógica: Se não encontrar a venda na data exata (ou d+1), deixa espaço em branco
