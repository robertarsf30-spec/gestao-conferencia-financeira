import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
from datetime import timedelta
import re

st.set_page_config(page_title="Conferência Cielo V3", layout="wide")

st.title("💳 Conferência Cielo - Sistema de Dupla Checagem")

u_excel = st.file_uploader("1. Planilha Cielo (Original)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    # --- PROCESSAMENTO DO PDF ---
    base_pdf = []
    with pdfplumber.open(u_pdf) as pdf:
        for page in pdf.pages:
            texto = page.extract_text()
            if texto:
                for linha in texto.split('\n'):
                    # Captura ID (PC/PD) com qualquer símbolo/caractere especial
                    match_id = re.search(r'(PC|PD)[^\s]+', linha, re.IGNORECASE)
                    match_data = re.search(r'\d{2}/\d{2}/\d{4}', linha)
                    
                    if match_id:
                        cod_id = match_id.group().strip().upper()
                        dt_pdf = pd.to_datetime(match_data.group(), dayfirst=True).date() if match_data else None
                        
                        # Captura valores (ex: 156,00 ou 156)
                        numeros = re.findall(r'\d+(?:[\.,]\d{2})?|\d+', linha)
                        for n in numeros:
                            try:
                                v_f = float(n.replace('.', '').replace(',', '.'))
                                if v_f > 1.0:
                                    base_pdf.append({'id': cod_id, 'valor': v_f, 'data': dt_pdf, 'usado': False})
                            except: continue

    col1, col2 = st.columns(2)

    with col1:
        btn_padrao = st.button("🚀 1. Busca Padrão (14/20)")
    
    with col2:
        btn_avancado = st.button("🔍 2. Busca Avançada (Recuperar 6/20)")

    if btn_padrao or btn_avancado:
        u_excel.seek(0)
        wb = openpyxl.load_workbook(u_excel)
        ws = wb.active
        df_cielo = pd.read_excel(u_excel, header=14)
        
        sucessos = 0
        
        for i, row in df_cielo.iterrows():
            linha_ex = i + 16
            status_atual = ws.cell(row=linha_ex, column=8).value
            
            # Se for busca avançada, só processa o que não foi achado na primeira
            if btn_avancado and status_atual not in [None, "NÃO ENCONTRADO"]:
                continue
                
            try:
                val_cielo = float(row.iloc[4]) # Valor Bruto
                dt_venda = pd.to_datetime(row.iloc[1]).date() # Data Venda
                
                achou = False
                # CRITÉRIOS DE BUSCA
                for item in base_pdf:
                    if not item['usado']:
                        # Define margens baseadas no botão clicado
                        margem_valor = 0.03 if btn_avancado else 0.01
                        margem_data = 3 if btn_avancado else 1
                        
                        v_bate = abs(item['valor'] - val_cielo) <= margem_valor
                        
                        # Validação de data (se houver data no PDF)
                        d_bate = True
                        if item['data']:
                            diff = abs((item['data'] - dt_venda).days)
                            d_bate = diff <= margem_data
                        
                        if v_bate and d_bate:
                            ws.cell(row=linha_ex, column=8).value = item['id']
                            item['usado'] = True
                            achou = True
                            sucessos += 1
                            break
                
                if not achou and not btn_avancado:
                    ws.cell(row=linha_ex, column=8).value = "NÃO ENCONTRADO"
            except: continue

        st.success(f"🎯 Processamento concluído! Localizados: {sucessos}")
        
        buffer = BytesIO()
        wb.save(buffer)
        st.download_button("📥 Baixar Planilha Atualizada", buffer.getvalue(), "Cielo_Conferencia.xlsx")
