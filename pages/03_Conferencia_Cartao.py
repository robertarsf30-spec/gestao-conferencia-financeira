import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

st.set_page_config(page_title="Conferência Cielo Identica", layout="wide")

st.title("💳 Conferência Cielo - Preservação de Layout")

u_excel = st.file_uploader("1. Planilha Cielo Original", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    # 1. Extração Inteligente do PDF
    @st.cache_data
    def extrair_pdf(file):
        dados = []
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        id_m = re.search(r'(PC|PD)[^\s]+', linha, re.IGNORECASE)
                        if id_m:
                            cod = id_m.group().strip().upper()
                            # Captura todos os números na linha para achar o valor
                            for v in re.findall(r'\d+(?:[\.,]\d{2})?|\d+', linha):
                                try:
                                    v_f = float(v.replace('.', '').replace(',', '.'))
                                    if v_f > 1.0:
                                        dados.append({'id': cod, 'valor': v_f, 'usado': False})
                                except: continue
        return dados

    base_pdf = extrair_pdf(u_pdf)

    if st.button("🚀 Localizar Vendas e Manter Planilha Idêntica"):
        # Abrir o arquivo original para edição
        u_excel.seek(0)
        # data_only=False preserva fórmulas; keep_vba=False remove macros que travam o salvamento
        wb = openpyxl.load_workbook(u_excel, data_only=False, keep_vba=False)
        ws = wb.active
        
        # Lemos os dados para processamento (Coluna E = Valor Bruto)
        df_dados = pd.read_excel(u_excel, header=14)
        
        sucessos = 0
        for i, row in df_dados.iterrows():
            linha_excel = i + 16 # Linha real no Excel
            try:
                valor_cielo = float(row.iloc[4]) # Valor Bruto (Coluna E)
                
                # Busca por valor (Ignorando data para pegar os 6 itens de março/abril)
                for item in base_pdf:
                    if not item['usado'] and abs(item['valor'] - valor_cielo) <= 0.02:
                        # Preenche apenas a Coluna H (8), mantendo o resto intacto
                        ws.cell(row=linha_excel, column=8).value = item['id']
                        item['usado'] = True
                        sucessos += 1
                        break
            except: continue

        st.success(f"🎯 Concluído! {sucessos} de 20 itens preenchidos na sua planilha original.")

        # Gerar o arquivo para download mantendo o buffer aberto
        output = BytesIO()
        wb.save(output)
        st.download_button(
            label="📥 Baixar Planilha Original Preenchida",
            data=output.getvalue(),
            file_name="Cielo_Identica_Preenchida.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
