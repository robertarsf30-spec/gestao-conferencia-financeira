import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

st.set_page_config(page_title="Conferência Cielo 20/20", layout="wide")
st.title("💳 Conferência Cielo - Sistema de Preservação Total")

u_excel = st.file_uploader("1. Planilha Cielo Original", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
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
                            for v in re.findall(r'\d+(?:[\.,]\d{2})?', linha):
                                try:
                                    v_f = float(v.replace('.', '').replace(',', '.'))
                                    if v_f > 1.0:
                                        dados.append({'id': cod, 'valor': v_f, 'usado': False})
                                except: continue
        return dados

    base_pdf = extrair_pdf(u_pdf)

    if st.button("🚀 Iniciar Conciliação (Preservar Layout)"):
        u_excel.seek(0)
        # Carrega o workbook garantindo que não vamos quebrar as células mescladas
        wb = openpyxl.load_workbook(u_excel, data_only=False)
        ws = wb.active
        
        # Lemos os dados via Pandas apenas para saber os valores e linhas (Header na 15)
        df_dados = pd.read_excel(u_excel, header=14)
        
        sucessos = 0
        # A varredura começa na linha 16 do Excel (onde começam os dados reais)
        for i, row in df_dados.iterrows():
            linha_excel = i + 16 
            try:
                valor_cielo = float(row.iloc[4]) # Coluna E (Valor Bruto)
                
                for item in base_pdf:
                    if not item['usado'] and abs(item['valor'] - valor_cielo) <= 0.02:
                        # TRATAMENTO PARA MERGED CELL:
                        cell = ws.cell(row=linha_excel, column=8) # Coluna H
                        # Se a célula for parte de uma mesclagem, o openpyxl precisa tratar diferente
                        if isinstance(cell, openpyxl.cell.cell.MergedCell):
                            # Encontra a célula mestre da mesclagem para atribuir o valor
                            for range_ in ws.merged_cells.ranges:
                                if cell.coordinate in range_:
                                    ws.cell(range_.min_row, range_.min_col).value = item['id']
                                    break
                        else:
                            cell.value = item['id']
                            
                        item['usado'] = True
                        sucessos += 1
                        break
            except: continue

        st.success(f"🎯 Concluído! {sucessos} de 20 itens localizados.")

        # Gerar o arquivo final
        output = BytesIO()
        wb.save(output)
        st.download_button(
            label="📥 Baixar Planilha Idêntica Preenchida",
            data=output.getvalue(),
            file_name="Cielo_Identica_Final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
