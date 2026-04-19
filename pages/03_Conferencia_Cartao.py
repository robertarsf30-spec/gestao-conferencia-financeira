import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl

st.set_page_config(page_title="Conferência Cartão Cielo", layout="wide")

if 'autenticado' not in st.session_state or not st.session_state.autenticado:
    st.error("🔒 Por favor, faça login na página inicial.")
    st.stop()

st.title("💳 Conferência Cartão (Cielo)")
st.info("O modelo original da planilha será mantido, preenchendo apenas a coluna 'Descrição'.")

u_excel = st.file_uploader("1. Planilha Cielo Original (Excel)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    try:
        # 1. LER DADOS PARA COMPARAR (Pula as 14 linhas para a lógica)
        df_dados = pd.read_excel(u_excel, header=14)
        
        # Mapeamento por posição conforme seu arquivo
        # Col 1: Data Venda | Col 4: Valor Bruto
        df_dados['Data_Ref'] = pd.to_datetime(df_dados.iloc[:, 1], dayfirst=True, errors='coerce').dt.date
        df_dados['Valor_Ref'] = pd.to_numeric(df_dados.iloc[:, 4], errors='coerce')

        # 2. PROCESSAR PDF DO SISTEMA
        dados_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    for line in text.split('\n'):
                        parts = line.split()
                        if len(parts) >= 8:
                            try:
                                tipo_pdf = str(parts[0])
                                data_pdf = pd.to_datetime(parts[1], dayfirst=True).date()
                                val_str = parts[-3].replace('.', '').replace(',', '.')
                                valor_pdf = float(val_str)
                                dados_pdf.append({'Tipo': tipo_pdf, 'Data': data_pdf, 'Valor': valor_pdf})
                            except:
                                continue
        
        df_sis = pd.DataFrame(dados_pdf)

        if st.button("🚀 Iniciar Conferência"):
            # Carregar o Excel usando openpyxl para manter a "plantilha" original
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active # Pega a planilha ativa
            
            # A coluna 'Descrição' no seu modelo é a coluna H (8)
            # Os dados começam na linha 16 (abaixo do cabeçalho de 15 linhas)
            
            conferidos = 0
            for i, row in df_dados.iterrows():
                if pd.isna(row['Data_Ref']):
                    continue
                
                # Regra de cruzamento (Data e Valor +- 0.02)
                mask = (df_sis['Data'] == row['Data_Ref']) & \
                       ((df_sis['Valor'] - row['Valor_Ref']).abs() <= 0.02)
                
                match = df_sis[mask]
                
                # Linha no Excel = índice do pandas + 16 (15 do header + 1 do Excel)
                linha_excel = i + 16
                
                if not match.empty:
                    ws.cell(row=linha_excel, column=8).value = f"CONFERIDO ({match.iloc[0]['Tipo']})"
                    conferidos += 1
                else:
                    ws.cell(row=linha_excel, column=8).value = "NÃO ENCONTRADO"

            st.success(f"✅ Processamento concluído! {conferidos} vendas encontradas.")

            # 3. GERAR DOWNLOAD DO MODELO ORIGINAL MANTIDO
            output = BytesIO()
            wb.save(output)
            
            st.download_button(
                label="📥 Baixar Planilha Preenchida (Modelo Original)",
                data=output.getvalue(),
                file_name="Cielo_Conferida_Original.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"Erro ao processar: {e}")
