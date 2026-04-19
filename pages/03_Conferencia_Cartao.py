import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

st.set_page_config(page_title="Localizador Total Cielo", layout="wide")

st.title("🚀 Localizador de Vendas - Força Bruta (Alvo: 20/20)")

u_excel = st.file_uploader("Planilha Cielo", type=['xlsx'])
u_pdf = st.file_uploader("Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    # 1. VARREDURA AGRESSIVA NO PDF
    # Vamos capturar tudo que pareça um ID e tudo que pareça um Valor
    dados_extraidos = []
    with pdfplumber.open(u_pdf) as pdf:
        for page in pdf.pages:
            linhas = page.extract_text().split('\n')
            for linha in linhas:
                # Busca IDs (PC/PD) e valores na mesma linha
                id_match = re.search(r'(PC|PD)[\w\d-]+', linha, re.IGNORECASE)
                valor_matches = re.findall(r'\d+(?:[\.,]\d{2})?', linha)
                
                if id_match:
                    for v_str in valor_matches:
                        try:
                            v_float = float(v_str.replace('.', '').replace(',', '.'))
                            if v_float > 0:
                                dados_extraidos.append({
                                    'id': id_match.group().upper(),
                                    'valor': v_float,
                                    'usado': False
                                })
                        except: continue

    if st.button("🔍 Forçar Localização das 20 Vendas"):
        u_excel.seek(0)
        wb = openpyxl.load_workbook(u_excel)
        ws = wb.active
        
        # Dados da Cielo (Coluna E é o Valor Bruto)
        df_cielo = pd.read_excel(u_excel, header=14)
        
        encontrados = 0
        for i, row in df_cielo.iterrows():
            linha_excel = i + 16
            valor_cielo = float(row.iloc[4]) # Valor Bruto
            
            # BUSCA SEM REGRAS: Apenas valor e disponibilidade
            match_perfeito = False
            for doc in dados_extraidos:
                # Tolerância de 2 centavos para arredondamentos do sistema
                if not doc['usado'] and abs(doc['valor'] - valor_cielo) <= 0.02:
                    ws.cell(row=linha_excel, column=8).value = doc['id']
                    doc['usado'] = True
                    match_perfeito = True
                    encontrados += 1
                    break
            
            if not match_perfeito:
                ws.cell(row=linha_excel, column=8).value = "REVISAR MANUAL"

        st.success(f"✅ Sucesso! {encontrados} de 20 itens localizados.")
        
        # Download do resultado
        buffer = BytesIO()
        wb.save(buffer)
        st.download_button(
            label="📥 Baixar Planilha Localizada",
            data=buffer.getvalue(),
            file_name="Cielo_Localizacao_Forcada.xlsx"
        )
