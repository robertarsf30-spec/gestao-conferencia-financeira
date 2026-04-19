import streamlit as st
import pandas as pd
import pdfplumber

st.set_page_config(page_title="Conferência Cartão Cielo", layout="wide")

if 'autenticado' not in st.session_state or not st.session_state.autenticado:
    st.error("🔒 Por favor, faça login na página inicial.")
    st.stop()

st.title("💳 Conferência Cartão (Cielo)")

u_excel = st.file_uploader("1. Planilha Cielo (Excel)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    # 1. Processar Excel Cielo
    df_cielo = pd.read_excel(u_excel)
    # Filtra por PC ou PD e limpa valores
    df_cielo['Valor Bruto'] = pd.to_numeric(df_cielo['Valor Bruto'], errors='coerce')
    
    # 2. Processar PDF Sistema
    dados_pdf = []
    with pdfplumber.open(u_pdf) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                dados_pdf.extend(table[1:]) # Pula cabeçalho do PDF
    
    df_sis = pd.DataFrame(dados_pdf)
    # Aqui precisamos ajustar as colunas do PDF conforme sua estrutura
    # Exemplo: df_sis.columns = ['Tipo', 'Data Lcto', 'Data Vcto', 'Valor']

    # 3. Lógica de Cruzamento (Merge com margem de 0.02)
    def encontrar_venda(row_cielo, df_sistema):
        # Filtra sistema por data e tipo (PC/PD)
        possiveis = df_sistema[
            (df_sistema['Data Venda'] == row_cielo['Data Venda']) & 
            (df_sistema['Tipo'] == row_cielo['Tipo'])
        ]
        # Busca valor com margem de 0.02
        for _, s in possiveis.iterrows():
            if abs(s['Valor'] - row_cielo['Valor Bruto']) <= 0.02:
                return "CONFERIDO"
        return "NÃO ENCONTRADO (Preencher Manual)"

    st.success("Arquivos carregados. Clique abaixo para gerar o resultado.")
    
    if st.button("Gerar Planilha de Conferência"):
        # O sistema executará a comparação aqui e exibirá a tabela final
        st.write("Resultado da Conferência:")
        st.dataframe(df_cielo) # Exibe a base para preenchimento
