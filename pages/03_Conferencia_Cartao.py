import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO

st.set_page_config(page_title="Conferência Cartão Cielo", layout="wide")

if 'autenticado' not in st.session_state or not st.session_state.autenticado:
    st.error("🔒 Por favor, faça login na página inicial.")
    st.stop()

st.title("💳 Conferência Cartão (Cielo)")
st.divider()

u_excel = st.file_uploader("1. Planilha Cielo (Excel)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    try:
        # 1. PROCESSAR EXCEL CIELO
        # Lendo a partir da linha 14 (cabeçalho real)
        df_cielo = pd.read_excel(u_excel, header=13) 
        
        # Nomes exatos das colunas do seu arquivo
        col_tipo = 'Forma de pagamento'
        col_data = 'Data da venda'
        col_valor = 'Valor bruto'

        # Limpeza: remove linhas que não tenham data de venda
        df_cielo = df_cielo.dropna(subset=[col_data])

        # Conversão de dados
        df_cielo['Valor_Num'] = pd.to_numeric(df_cielo[col_valor], errors='coerce')
        df_cielo['Data_Venda'] = pd.to_datetime(df_cielo[col_data], dayfirst=True).dt.date
        
        if 'Descrição' not in df_cielo.columns:
            df_cielo['Descrição'] = ""

        # 2. PROCESSAR PDF DO SISTEMA
        dados_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    for row in table:
                        try:
                            # Tenta converter o último campo (valor) e a data (terceira coluna)
                            v_limpo = float(str(row[-1]).replace('.', '').replace(',', '.'))
                            d_limpa = pd.to_datetime(row[2], dayfirst=True).date()
                            dados_pdf.append({'Tipo': str(row[0]), 'Data': d_limpa, 'Valor': v_limpo})
                        except:
                            continue
        
        df_sis = pd.DataFrame(dados_pdf)

        # 3. BOTÃO DE PROCESSAMENTO
        if st.button("🚀 Iniciar Conferência"):
            def conferir(row):
                if df_sis.empty: return "ERRO: PDF VAZIO"
                
                # Regra: PC (Crédito) / PD (Débito)
                # Margem de 1 dia e R$ 0,02
                match = df_sis[
                    (df_sis['Tipo'].str.contains('PC|PD', case=False, na=False)) &
                    (abs((df_sis['Data'] - row['Data_Venda']).days) <= 1) &
                    (abs(df_sis['Valor'] - row['Valor_Num']) <= 0.02)
                ]
                return "CONFERIDO" if not match.empty else "NÃO ENCONTRADO"

            df_cielo['Descrição'] = df_cielo.apply(conferir, axis=1)
            
            # Remove colunas auxiliares
            df_final = df_cielo.drop(columns=['Valor_Num', 'Data_Venda'])
            
            st.success("✅ Conferência concluída!")
            st.dataframe(df_final)

            # 4. DOWNLOAD (Usando engine openpyxl corrigida)
            output = BytesIO()
            df_final.to_excel(output, index=False, engine='openpyxl')
            
            st.download_button(
                label="📥 Baixar Planilha de Conferência",
                data=output.getvalue(),
                file_name="conferencia_cielo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"Erro técnico: {e}")
