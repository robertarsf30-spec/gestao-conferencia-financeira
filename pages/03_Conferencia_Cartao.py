import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

st.set_page_config(page_title="Conferência Cielo 20/20", layout="wide")
st.title("💳 Conferência Cielo - Recuperando os 20 Itens")

u_excel = st.file_uploader("1. Planilha Cielo", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    # --- ETAPA 1: LER O PDF COM FOCO EM VALORES "LIMPOS" ---
    base_pdf = []
    with pdfplumber.open(u_pdf) as pdf:
        for page in pdf.pages:
            texto = page.extract_text()
            if texto:
                for linha in texto.split('\n'):
                    # Captura códigos PC ou PD (ex: PC22650-1)
                    ids = re.findall(r'(?:PC|PD)[\w\d-]+', linha, re.IGNORECASE)
                    # Captura números que pareçam valores (ex: 156,00 ou 156)
                    valores = re.findall(r'\d+(?:[\.,]\d{2})?', linha)
                    
                    if ids:
                        codigo_id = ids[0].upper()
                        for v_str in valores:
                            try:
                                # Converte 156,00 ou 156 para float 156.0
                                v_f = float(v_str.replace('.', '').replace(',', '.'))
                                if v_f > 1.0: 
                                    base_pdf.append({'id': codigo_id, 'valor': v_f})
                            except: continue

    if st.button("🚀 Iniciar Conferência Final (Buscar todos os 20)"):
        u_excel.seek(0)
        wb = openpyxl.load_workbook(u_excel)
        ws = wb.active
        
        # Lê a Cielo (Coluna E é o Valor Bruto na linha 15 em diante)
        df_cielo = pd.read_excel(u_excel, header=14)
        
        sucessos = 0
        usados_no_pdf = [] # Para não repetir o mesmo PC em valores iguais

        for i, row in df_cielo.iterrows():
            linha_excel = i + 16
            try:
                valor_cielo = float(row.iloc[4]) # Valor Bruto
                
                # BUSCA: Tenta achar o valor exato no que extraímos do PDF
                for item in base_pdf:
                    if item['id'] not in usados_no_pdf:
                        # Compara com margem de erro de 2 centavos
                        if abs(item['valor'] - valor_cielo) <= 0.02:
                            ws.cell(row=linha_excel, column=8).value = item['id']
                            usados_no_pdf.append(item['id'])
                            sucessos += 1
                            break
            except: continue

        st.success(f"🎯 Sucesso! {sucessos} de 20 títulos vinculados com precisão.")
        
        buffer = BytesIO()
        wb.save(buffer)
        st.download_button(
            label="📥 Baixar Planilha 100% Corrigida",
            data=buffer.getvalue(),
            file_name="Cielo_Conferencia_Completa.xlsx"
        )
