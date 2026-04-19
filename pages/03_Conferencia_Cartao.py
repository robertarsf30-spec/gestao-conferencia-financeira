import streamlit as st
import pandas as pd
import pdfplumber # Biblioteca para ler o PDF do sistema

st.set_page_config(page_title="Conferência Cartão Cielo", layout="wide")

if 'autenticado' not in st.session_state or not st.session_state.autenticado:
    st.error("🔒 Por favor, faça login na página inicial.")
    st.stop()

st.title("💳 Conferência Cartão (Cielo)")
st.info("Regra: Comparação de Valor Bruto e Data (PC = Crédito / PD = Débito)")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Dados da Cielo")
    u_excel = st.file_uploader("Upload Planilha Recebíveis (Excel)", type=['xlsx'])

with col2:
    st.subheader("2. Dados do Sistema")
    u_pdf = st.file_uploader("Upload Relatório de Títulos (PDF)", type=['pdf'])

if u_excel and u_pdf:
    # --- PROCESSANDO EXCEL CIELO ---
    df_cielo = pd.read_excel(u_excel)
    # Filtramos por PC (Crédito) e PD (Débito)
    # Ajustaremos as colunas exatas assim que você testar o primeiro arquivo
    cielo_filtrado = df_cielo[df_cielo.iloc[:, 1].str.contains('PC|PD', na=False, case=False)]
    
    # --- PROCESSANDO PDF DO SISTEMA ---
    dados_pdf = []
    with pdfplumber.open(u_pdf) as pdf:
        for pagina in pdf.pages:
            tabela = pagina.extract_table()
            if tabela:
                dados_pdf.extend(tabela)
    
    df_sistema = pd.DataFrame(dados_pdf)
    
    st.success("Arquivos lidos com sucesso! Iniciando cruzamento...")
    
    # Exibição básica dos totais encontrados
    st.divider()
    c1, c2 = st.columns(2)
    c1.metric("Total na Cielo", f"R$ {cielo_filtrado.iloc[:, -1].sum():,.2f}")
    st.dataframe(cielo_filtrado.head())
