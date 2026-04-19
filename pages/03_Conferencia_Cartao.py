import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

st.set_page_config(page_title="Cielo 20/20 Final", layout="wide")
st.title("🎯 Localizador Total: Correção de Quebra de Linha e IDs")

u_excel = st.file_uploader("1. Planilha Cielo Original", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    @st.cache_data
    def extrair_pdf_v3(file):
        dados = []
        id_pendente = None
        
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        # 1. Tenta achar um PC ou PD na linha
                        id_m = re.search(r'(PC|PD)[\w\d\-\.]+', linha, re.IGNORECASE)
                        if id_m:
                            id_pendente = id_m.group().strip().upper()
                        
                        # 2. Busca valores na linha atual (mesmo que o PC esteja na linha de cima)
                        valores = re.findall(r'\d+(?:[\.,]\d{2})', linha)
                        if valores and id_pendente:
                            for v in valores:
                                try:
                                    v_f = float(v.replace('.', '').replace(',', '.'))
                                    if v_f > 1.0:
                                        dados.append({'id': id_pendente, 'valor': v_f, 'usado': False})
                                        # Após vincular o valor ao ID pendente, resetamos para não repetir errado
                                        id_pendente = None 
                                except: continue
        return dados

    lista_pdf = extrair_pdf_v3(u_pdf)

    if st.button("🚀 Executar Conciliação Final (Bater os 20 itens)"):
        u_excel.seek(0)
        # Carrega o Excel mantendo o layout original (Cores, Logos, etc)
        wb = openpyxl.load_workbook(u_excel, data_only=False)
        ws = wb.active
        df_cielo = pd.read_excel(u_excel, header=14)
        
        sucessos = 0
        for i, row in df_cielo.iterrows():
            linha_excel = i + 16 
            try:
                # Valor Bruto da Coluna E
                valor_cielo = float(row.iloc[4])
                
                # Busca flexível por valor (independente de data)
                for item in lista_pdf:
                    if not item['usado'] and abs(item['valor'] - valor_cielo) <= 0.01:
                        ws.cell(row=linha_excel, column=8).value = item['id']
                        item['usado'] = True
                        sucessos += 1
                        break
            except: continue

        if sucessos >= 20:
            st.success(f"✅ EXCELENTE! {sucessos} de 20 títulos preenchidos com sucesso.")
        else:
            st.warning(f"Atenção: {sucessos} encontrados. Verifique se o PDF contém todos os títulos.")

        output = BytesIO()
        wb.save(output)
        st.download_button(
            label="📥 Baixar Planilha 20/20 Idêntica",
            data=output.getvalue(),
            file_name="Cielo_Final_Identica.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
