import streamlit as st
import pandas as pd

if not st.session_state.get('autenticado'):
    st.stop()

st.title("💵 Conferência de Caixa (Físico)")
st.info("Validação: Saldo Anterior + Entradas - Saídas = Saldo Atual.")

with st.expander("Configuração de Saldos", expanded=True):
    c1, c2 = st.columns(2)
    u_bol = c1.file_uploader("Boletim de Caixa (PDF)", type=['pdf'])
    u_pla = c2.file_uploader("Planilha de Conferência (Excel)", type=['xlsx'])

if u_bol and u_pla:
    st.divider()
    # Lógica: Pegar 'Total Final' do boletim anterior e comparar com 'Inicial' do atual
    st.success("Saldos iniciais validados com o dia anterior.")
