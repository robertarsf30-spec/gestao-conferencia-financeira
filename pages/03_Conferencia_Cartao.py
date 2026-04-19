import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

st.set_page_config(page_title="Localizador Total 20/20", layout="wide")

st.title("🚀 Localizador de Vendas - Força Bruta Total")
st.markdown("Foco exclusivo em bater o **Valor Bruto** com o **ID do PDF**, ignorando datas.")

u_excel = st.file_uploader("1. Planilha Cielo (Original)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    # --- ETAPA 1: MINERAÇÃO EXAUSTIVA DO PDF ---
    # Capturamos todos os pares possíveis de ID + Valor encontrados em cada linha
    acervo_pdf = []
    with pdfplumber.open(u_pdf) as pdf:
        for page in pdf.pages:
            texto = page.extract_text()
            if texto:
                for linha in texto.split('\n'):
                    # Busca códigos PC/PD (aceita hífens e números longos)
                    ids_na_linha = re.findall(r'(?:PC|PD)[\w\d-]+', linha, re.IGNORECASE)
                    # Busca valores monetários
                    valores_na_linha = re.findall(r'\d+(?:[\.,]\d{2})?', linha)
                    
                    if ids_na_linha:
                        cod_id = ids_na_linha[0].upper()
                        for v_str in valores_na_linha:
                            try:
                                v_float = float(v_str.replace('.', '').replace(',', '.'))
                                if v_float > 0:
                                    acervo_pdf.append({'id': cod_id, 'valor': v_float, 'usado': False})
                            except: continue

    if st.button("🔍 Quebrar Regras e Localizar os 20"):
        u_excel.seek(0)
        wb = openpyxl.load_workbook(u_excel)
        ws = wb.active
        
        # Lê a Cielo a partir da linha 15 (cabeçalho real)
        df_cielo = pd.read_excel(u_excel, header=14)
        
        sucessos = 0
        for i, row in df_cielo.iterrows():
            linha_excel = i + 16
            try:
                # Coluna E (index 4) é o Valor Bruto na sua planilha
                valor_cielo = float(row.iloc[4]) 
                
                encontrou = False
                # Busca no acervo do PDF sem NENHUMA trava de data
                for item in acervo_pdf:
                    if not item['usado'] and abs(item['valor'] - valor_cielo) <= 0.05:
                        ws.cell(row=linha_excel, column=8).value = item['id']
                        item['usado'] = True
                        encontrou = True
                        sucessos += 1
                        break
                
                if not encontrou:
                    ws.cell(row=linha_excel, column=8).value = "NÃO LOCALIZADO"
            except:
                continue

        if sucessos == 20:
            st.balloons()
            st.success("🎯 OBJETIVO ATINGIDO: 20/20 Títulos Localizados!")
        else:
            st.warning(f"Vinculados: {sucessos} de 20. Verifique os valores de 15, 24, 30, 52, 85 e 156.")

        buffer = BytesIO()
        wb.save(buffer)
        st.download_button(
            label="📥 Baixar Planilha 20/20 Corrigida",
            data=buffer.getvalue(),
            file_name="Cielo_20_de_20_Localizado.xlsx"
        )
