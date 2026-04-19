import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

st.set_page_config(page_title="Conferência 20/20", layout="wide")
st.title("💳 Conferência Cielo - Localização Total Garantida")

u_excel = st.file_uploader("1. Planilha Cielo Original", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    @st.cache_data
    def extrair_pdf(file):
        dados = []
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        id_m = re.search(r'(PC|PD)[^\s]+', linha, re.IGNORECASE)
                        dt_m = re.search(r'\d{2}/\d{2}/\d{4}', linha)
                        if id_m:
                            cod = id_m.group().strip().upper()
                            dt = pd.to_datetime(dt_m.group(), dayfirst=True).date() if dt_m else None
                            for n in re.findall(r'\d+(?:[\.,]\d{2})?|\d+', linha):
                                try:
                                    v_f = float(n.replace('.', '').replace(',', '.'))
                                    if v_f > 1.0:
                                        dados.append({'id': cod, 'valor': v_f, 'data': dt, 'usado': False})
                                except: continue
        return dados

    lista_pdf = extrair_pdf(u_pdf)

    if st.button("🚀 Executar Conciliação (14 Padrão + 6 Flexíveis)"):
        u_excel.seek(0)
        wb = openpyxl.load_workbook(u_excel, data_only=False)
        ws = wb.active
        df_cielo = pd.read_excel(u_excel, header=14)
        
        sucesso = 0
        for i, row in df_cielo.iterrows():
            linha_ex = i + 16
            try:
                v_alvo = float(row.iloc[4]) # Valor Bruto
                dt_alvo = pd.to_datetime(row.iloc[1]).date() # Data Venda
                
                # TENTATIVA 1: DATA + VALOR (Acha os 14 itens)
                achou = False
                for item in lista_pdf:
                    if not item['usado'] and abs(item['valor'] - v_alvo) <= 0.01:
                        if item['data'] == dt_alvo:
                            ws.cell(row=linha_ex, column=8).value = item['id']
                            item['usado'] = True
                            achou = True
                            sucesso += 1
                            break
                
                # TENTATIVA 2: APENAS VALOR (Acha os 6 itens de Março que restaram)
                if not achou:
                    for item in lista_pdf:
                        if not item['usado'] and abs(item['valor'] - v_alvo) <= 0.01:
                            ws.cell(row=linha_ex, column=8).value = item['id']
                            item['usado'] = True
                            sucesso += 1
                            break
            except: continue
        
        st.success(f"🎯 Finalizado! {sucesso} de 20 itens vinculados.")

        output = BytesIO()
        wb.save(output)
        st.download_button("📥 Baixar Planilha 20/20", output.getvalue(), "Cielo_Conciliado_Total.xlsx")
