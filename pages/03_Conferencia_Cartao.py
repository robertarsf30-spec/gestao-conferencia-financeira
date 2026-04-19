import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

st.set_page_config(page_title="Conferência Cielo - Estável", layout="wide")
st.title("💳 Conferência Cielo - Versão Estável")

u_excel = st.file_uploader("1. Planilha Cielo (Original)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    # --- EXTRAÇÃO DO PDF ---
    @st.cache_data
    def extrair_pdf(file):
        base = []
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        # Captura IDs (PC/PD) e Datas
                        id_m = re.search(r'(PC|PD)[^\s]+', linha, re.IGNORECASE)
                        dt_m = re.search(r'\d{2}/\d{2}/\d{4}', linha)
                        if id_m:
                            cod = id_m.group().strip().upper()
                            dt = pd.to_datetime(dt_m.group(), dayfirst=True).date() if dt_m else None
                            # Captura valores monetários
                            for n in re.findall(r'\d+(?:[\.,]\d{2})?', linha):
                                try:
                                    v_f = float(n.replace('.', '').replace(',', '.'))
                                    if v_f > 1.0:
                                        base.append({'id': cod, 'valor': v_f, 'data': dt, 'usado': False})
                                except: continue
        return base

    lista_pdf = extrair_pdf(u_pdf)

    if st.button("🚀 Iniciar Conciliação"):
        # Abrir o Excel original de forma segura
        u_excel.seek(0)
        wb = openpyxl.load_workbook(u_excel, data_only=False)
        ws = wb.active
        
        # Leitura dos dados para comparação (Coluna E = Valor Bruto)
        df_cielo = pd.read_excel(u_excel, header=14)
        
        sucesso = 0
        for i, row in df_cielo.iterrows():
            linha_ex = i + 16 # Ajuste para começar na linha correta do Excel
            v_alvo = float(row.iloc[4])
            dt_alvo = pd.to_datetime(row.iloc[1]).date()
            
            # Busca Rigorosa (Data + Valor) - O que garantiu os 14 itens
            for item in lista_pdf:
                if not item['usado'] and abs(item['valor'] - v_alvo) <= 0.01:
                    if item['data'] == dt_alvo:
                        ws.cell(row=linha_ex, column=8).value = item['id']
                        item['usado'] = True
                        sucesso += 1
                        break
        
        st.success(f"Finalizado! {sucesso} de 20 itens vinculados com precisão.")

        # Gerar o arquivo para download
        output = BytesIO()
        wb.save(output)
        st.download_button(
            label="📥 Baixar Planilha Preenchida",
            data=output.getvalue(),
            file_name="Cielo_Conciliado_Estavel.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
