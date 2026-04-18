import streamlit as st

if not st.session_state.get('autenticado'):
    st.stop()

st.title("📑 Módulo de Cobrança")
st.info("Filtro automático: Apenas dívidas vencidas há mais de 30 dias.")

u_pend = st.file_uploader("Relatório de Títulos Pendentes", type=['xlsx'])

if u_pend:
    st.subheader("Clientes com Débitos Pendentes")
    # Exemplo de como aparecerá no sistema:
    with st.container(border=True):
        st.write("**Cliente: João da Silva** | Total: R$ 1.500,00 (com juros)")
        st.selectbox("Situação:", ["Sem Retorno", "Programou Pagamento", "RJ"], key="joao1")
        st.date_input("Data Programada (se houver):")
