import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO

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
    # Garante que as colunas essenciais existam (ajuste os nomes conforme sua planilha)
    # Esperado: 'Tipo', 'Data Venda', 'Valor Bruto'
    df_cielo['Valor Bruto'] = pd.to_numeric(df_cielo['Valor Bruto'], errors='coerce')
    df_cielo['Data Venda'] = pd.to_datetime(df_cielo['Data Venda']).dt.date
    
    # Adiciona coluna de Descrição se não existir
    if 'Descrição' not in df_cielo.columns:
        df_cielo['Descrição'] = ""

    # 2. Processar PDF Sistema
    dados_pdf = []
    with pdfplumber.open(u_pdf) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                # Assume que os dados úteis começam após o cabeçalho
                for row in table[1:]:
                    if row[0]: # Filtra linhas vazias
                        dados_pdf.append({
                            'Tipo': row[0], # PC ou PD
                            'Data Lcto': pd.to_datetime(row[1], dayfirst=True).date() if row[1] else None,
                            'Valor': float(row[-1].replace('.', '').replace(',', '.')) if row[-1] else 0.0
                        })
    
    df_sis = pd.DataFrame(dados_pdf)

    # 3. Lógica de Cruzamento e Preenchimento da Descrição
    def conferir_venda(row):
        # Busca no sistema com margem de 1 dia e R$ 0,02
        match = df_sis[
            (df_sis['Tipo'].str.contains(row['Tipo'][:2], case=False, na=False)) &
            (abs((df_sis['Data Lcto'] - row['Data Venda']).days) <= 1) &
            (abs(df_sis['Valor'] - row['Valor Bruto']) <= 0.02)
        ]
        
        if not match.empty:
            return "CONFERIDO"
        else:
            return "VENDA NÃO ENCONTRADA NO SISTEMA"

    if st.button("🚀 Processar Conferência"):
        df_cielo['Descrição'] = df_cielo.apply(conferir_venda, axis=1)
        
        st.success("Conferência concluída!")
        st.dataframe(df_cielo)

        # 4. Gerar arquivo Excel para Download
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_cielo.to_excel(writer, index=False, sheet_name='Conferencia')
        
        processed_data = output.getvalue()
        
        st.download_button(
            label="📥 Baixar Planilha de Conferência (Excel)",
            data=processed_data,
            file_name="conferencia_cielo_final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
