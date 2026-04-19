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

u_excel = st.file_uploader("1. Planilha Cielo (Excel)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    try:
        # 1. LER DADOS DO EXCEL PARA PROCESSAMENTO
        # Pula as 14 linhas de cabeçalho da Cielo
        df_dados = pd.read_excel(u_excel, header=14)
        
        # Mapeamento por posição (Col 1: Data Venda | Col 4: Valor Bruto)
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
                        # Procura linhas com o formato do seu sistema (PC... data valor)
                        if len(parts) >= 8:
                            try:
                                tipo_pdf = str(parts[0])
                                data_pdf = pd.to_datetime(parts[1], dayfirst=True).date()
                                # Valor na antepenúltima posição (Ex: 156,00)
                                val_str = parts[-3].replace('.', '').replace(',', '.')
                                valor_pdf = float(val_str)
                                dados_pdf.append({'Tipo': tipo_pdf, 'Data': data_pdf, 'Valor': valor_pdf})
                            except:
                                continue
        
        df_sis = pd.DataFrame(dados_pdf)

        if st.button("🚀 Iniciar Conferência"):
            # Abre o Excel original para preencher a coluna Descrição (H)
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active 
            
            sucessos = 0
            # Os dados começam na linha 16 do Excel (15 do header + 1)
            for i, row in df_dados.iterrows():
                if pd.isna(row['Data_Ref']):
                    continue
                
                # Regra: Mesma Data e Valor com margem de 0.02
                valor_cielo = row['Valor_Ref']
                data_cielo = row['Data_Ref']
                
                # Filtra o PDF
                achou = False
                for _, p in df_sis.iterrows():
                    if p['Data'] == data_cielo and abs(p['Valor'] - valor_cielo) <= 0.02:
                        ws.cell(row=i+16, column=8).value = f"CONFERIDO ({p['Tipo']})"
                        sucessos += 1
                        achou = True
                        break
                
                if not achou:
                    ws.cell(row=i+16, column=8).value = "NÃO ENCONTRADO"

            st.success(f"✅ Concluído! {sucessos} vendas conferidas.")

            # 3. GERAR DOWNLOAD DO ARQUIVO MANTENDO O MODELO
            output = BytesIO()
            wb.save(output)
            
            st.download_button(
                label="📥 Baixar Planilha Original Preenchida",
                data=output.getvalue(),
                file_name="Cielo_Conferida.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"Erro técnico: {e}")
