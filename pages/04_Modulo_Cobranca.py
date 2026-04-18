import streamlit as st

st.set_page_config(page_title="Módulo Cobrança", layout="wide")

if 'autenticado' not in st.session_state or not st.session_state.autenticado:
    st.stop()

st.title("📑 Gestão de Cobrança")
st.info("Filtro: Títulos vencidos há mais de 30 dias.")

u_pend = st.file_uploader("Relatório de Títulos Pendentes", type=['xlsx'], key="cobr_pend")

if u_pend:
    st.subheader("⚠️ Clientes com Pendências")
    # Estrutura para exibir os blocos de clientes em atraso
    with st.container(border=True):
        st.write("**Cliente Exemplo** | Total: R$ 0,00")
        st.selectbox("Situação:", ["Sem contato", "Promessa de Pagto", "RJ"], key="sit_ex")
