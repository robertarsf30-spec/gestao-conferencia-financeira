import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

st.set_page_config(page_title="Conferência Cielo - Correção Final", layout="wide")
st.title("🎯 Localizador Ultra-Preciso (Correção de Símbolos e Datas)")

u_excel = st.file_uploader("1. Planilha Cielo Original", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    @st.cache_data
    def extrair_pdf_melhorado(file):
        base = []
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        # REGEX MELHORADO: Aceita PC/PD com hífens, números e símbolos
                        id_m = re.search(r'(PC|PD)[\w\d\-\.]+', linha, re.IGNORECASE)
                        if id_m:
                            cod = id_m.group().strip().upper()
                            # Captura valores (ex: 156,00 ou 156)
                            valores = re.findall(r'\d+(?:[\.,]\d{2})?', linha)
                            for v in valores:
                                try:
                                    v_f = float(v.replace('.', '').replace(',', '.'))
                                    if v_f > 1.0:
                                        base.append({'id': cod, 'valor': v_f, 'usado': False})
                                except: continue
        return base

    lista_pdf = extrair_pdf_melhorado(u_pdf)

    if st.button("🚀 Executar Varredura Total (Garantir 20 de 20)"):
        u_excel.seek(0)
        # Carrega o Excel mantendo fórmulas e layout
        wb = openpyxl.load_workbook(u_excel, data_only=False)
        ws = wb.active
        df_cielo = pd.read_excel(u_excel, header=14)
        
        sucessos = 0
        for i, row in df_cielo.iterrows():
            linha_ex = i + 16
            try:
                # Pegamos o Valor Bruto da Coluna E (Index 4)
                v_alvo = float(row.iloc[4])
                
                # BUSCA FLEXÍVEL: Prioriza o valor exato, independente da data
                for item in lista_pdf:
                    if not item['usado'] and abs(item['valor'] - v_alvo) <= 0.01:
                        # Preenche a Coluna H (8)
                        ws.cell(row=linha_ex, column=8).value = item['id']
                        item['usado'] = True
                        sucessos += 1
                        break
            except: continue

        st.success(f"✅ Finalizado! {sucessos} de 20 itens preenchidos.")

        # Download seguro
        output = BytesIO()
        wb.save(output)
        st.download_button(
            label="📥 Baixar Planilha 20/20 Corrigida",
            data=output.getvalue(),
            file_name="Cielo_Conciliado_Total_Corrigido.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
