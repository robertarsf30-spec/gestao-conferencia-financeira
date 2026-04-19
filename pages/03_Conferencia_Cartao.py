import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

st.set_page_config(page_title="Conferência 100% Cielo", layout="wide")

st.title("💳 Conferência Cielo - Localização de Precisão")
st.info("Este código ignora datas para garantir que vendas de meses anteriores sejam localizadas pelo valor.")

u_excel = st.file_uploader("1. Planilha Cielo", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    # --- ETAPA 1: MINERAÇÃO DE DADOS DO PDF ---
    base_pdf = []
    with pdfplumber.open(u_pdf) as pdf:
        for page in pdf.pages:
            texto = page.extract_text()
            if texto:
                for linha in texto.split('\n'):
                    # Captura qualquer ID que comece com PC ou PD
                    ids = re.findall(r'(?:PC|PD)[\w\d-]+', linha, re.IGNORECASE)
                    # Captura valores (ex: 156 ou 156,00)
                    valores = re.findall(r'\d+(?:[\.,]\d{2})?', linha)
                    
                    if ids:
                        for v_str in valores:
                            try:
                                v_f = float(v_str.replace('.', '').replace(',', '.'))
                                if v_f > 1.0: # Evita números pequenos que não são valores
                                    base_pdf.append({'id': ids[0].upper(), 'valor': v_f, 'usado': False})
                            except: continue

    if st.button("🚀 Forçar Preenchimento dos 20 Itens"):
        u_excel.seek(0)
        wb = openpyxl.load_workbook(u_excel)
        ws = wb.active
        
        # Lê a Cielo (Coluna E é o Valor Bruto)
        df_cielo = pd.read_excel(u_excel, header=14)
        
        sucessos = 0
        for i, row in df_cielo.iterrows():
            linha_ex = i + 16
            try:
                val_cielo = float(row.iloc[4]) # Valor Bruto
                
                # Busca no PDF apenas pelo VALOR
                # Se o valor bater, o ID daquela linha é o correto
                for item in base_pdf:
                    if not item['usado'] and abs(item['valor'] - val_cielo) <= 0.02:
                        ws.cell(row=linha_ex, column=8).value = item['id']
                        item['usado'] = True
                        sucessos += 1
                        break
            except: continue

        st.success(f"🎯 Resultado: {sucessos} de 20 títulos localizados!")
        
        buffer = BytesIO()
        wb.save(buffer)
        st.download_button(
            label="📥 Baixar Planilha Corrigida",
            data=buffer.getvalue(),
            file_name="Cielo_Conferencia_Final.xlsx"
        )
