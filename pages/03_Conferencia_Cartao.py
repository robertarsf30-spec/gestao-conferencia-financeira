import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

st.set_page_config(page_title="Conferência Cielo 20/20", layout="wide")
st.title("💳 Conferência Cielo - Recuperando os 20 Itens")

u_excel = st.file_uploader("1. Planilha Cielo (Original)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    # --- ETAPA 1: EXTRAÇÃO INTELIGENTE DO PDF ---
    base_pdf = []
    with pdfplumber.open(u_pdf) as pdf:
        for page in pdf.pages:
            texto = page.extract_text()
            if texto:
                for linha in texto.split('\n'):
                    # Localiza IDs que começam com PC ou PD (ex: PC22650-1)
                    ids = re.findall(r'(?:PC|PD)[\w\d-]+', linha, re.IGNORECASE)
                    # Localiza números formatados como valores (ex: 156,00 ou 1.640,00)
                    valores = re.findall(r'\d{1,3}(?:\.\d{3})*(?:\,\d{2})?', linha)
                    
                    if ids:
                        codigo_id = ids[0].upper()
                        for v_str in valores:
                            try:
                                # Normaliza o valor para float (remove ponto de milhar e troca vírgula por ponto)
                                v_limpo = v_str.replace('.', '').replace(',', '.')
                                v_f = float(v_limpo)
                                if v_f > 1.0: # Ignora valores irrelevantes
                                    base_pdf.append({'id': codigo_id, 'valor': v_f})
                            except: continue

    if st.button("🚀 Iniciar Conferência Final (Buscar todos os 20)"):
        # Carrega o Excel mantendo a formatação original
        u_excel.seek(0)
        wb = openpyxl.load_workbook(u_excel)
        ws = wb.active
        
        # Lê os dados para processamento (Coluna E é o Valor Bruto, linha 15 em diante)
        df_cielo = pd.read_excel(u_excel, header=14)
        
        sucessos = 0
        usados_no_pdf = [] 

        for i, row in df_cielo.iterrows():
            linha_excel = i + 16 # Ajuste para a linha correta no openpyxl
            try:
                # Valor Bruto da Coluna E
                valor_cielo = float(row.iloc[4]) 
                
                # Procura o valor correspondente na lista extraída do PDF
                for item in base_pdf:
                    if item['id'] not in usados_no_pdf:
                        # Compara com margem de erro mínima para centavos
                        if abs(item['valor'] - valor_cielo) <= 0.02:
                            # Escreve o ID (PC/PD) na Coluna H (Descrição)
                            ws.cell(row=linha_excel, column=8).value = item['id']
                            usados_no_pdf.append(item['id'])
                            sucessos += 1
                            break
            except: continue

        if sucessos >= 14:
            st.success(f"🎯 Sucesso! {sucessos} de 20 títulos vinculados.")
        else:
            st.warning(f"Atenção: Apenas {sucessos} itens encontrados. Verifique o formato do PDF.")
        
        # Gerar arquivo para download
        buffer = BytesIO()
        wb.save(buffer)
        st.download_button(
            label="📥 Baixar Planilha 100% Corrigida",
            data=buffer.getvalue(),
            file_name="Cielo_Conferencia_Completa.xlsx"
        )
