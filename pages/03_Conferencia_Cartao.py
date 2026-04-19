import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

st.set_page_config(page_title="Cielo 20/20 Final", layout="wide")
st.title("💳 Conferência Cielo - Sistema de Alta Precisão")

u_excel = st.file_uploader("1. Planilha Cielo Original", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    @st.cache_data
    def extrair_pdf_v4(file):
        dados = []
        id_recente = None
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        # Identifica PC/PD com hífens (ex: PC22630-1)
                        id_m = re.search(r'(PC|PD)[\w\d\-\.]+', linha, re.IGNORECASE)
                        if id_m:
                            id_recente = id_m.group().strip().upper()
                        
                        # Busca valores na linha atual
                        valores = re.findall(r'\d+(?:[\.,]\d{2})', linha)
                        if valores and id_recente:
                            for v in valores:
                                try:
                                    v_f = float(v.replace('.', '').replace(',', '.'))
                                    if v_f > 1.0:
                                        dados.append({'id': id_recente, 'valor': v_f, 'usado': False})
                                        id_recente = None # Reset após vincular
                                except: continue
        return dados

    lista_pdf = extrair_pdf_v4(u_pdf)

    if st.button("🚀 Gerar Planilha 20/20"):
        u_excel.seek(0)
        wb = openpyxl.load_workbook(u_excel, data_only=False)
        ws = wb.active
        
        # Lemos os dados começando da linha 16 (Header na 15)
        # Usamos try/except para evitar o erro de Index que você recebeu
        try:
            df_dados = pd.read_excel(u_excel, header=14)
            
            sucessos = 0
            for i in range(len(df_dados)):
                linha_planilha = i + 16
                try:
                    # Pegamos o valor bruto (Coluna E / Índice 4)
                    valor_cielo = float(df_dados.iloc[i, 4])
                    
                    for item in lista_pdf:
                        if not item['usado'] and abs(item['valor'] - valor_cielo) <= 0.01:
                            # Escreve na Coluna H (8)
                            ws.cell(row=linha_planilha, column=8).value = item['id']
                            item['usado'] = True
                            sucessos += 1
                            break
                except Exception:
                    continue

            st.success(f"✅ Finalizado! {sucessos} de 20 itens vinculados.")
            
            output = BytesIO()
            wb.save(output)
            st.download_button(
                label="📥 Baixar Planilha Idêntica",
                data=output.getvalue(),
                file_name="Cielo_20_de_20.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Erro na estrutura da planilha: {e}. Verifique se a tabela começa na linha 16.")
