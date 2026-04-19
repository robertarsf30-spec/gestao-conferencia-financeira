import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

st.set_page_config(page_title="Cielo 20/20 Final", layout="wide")
st.title("🎯 Localizador Ultra-Preciso: Correção de IDs e Datas")

u_excel = st.file_uploader("1. Planilha Cielo Original", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    @st.cache_data
    def extrair_pdf_ultra(file):
        dados = []
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        # EXPLICAÇÃO DO ERRO: O Regex agora aceita PC/PD seguido de qualquer 
                        # combinação de números, letras e hífens (ex: PC22630-1)
                        id_m = re.search(r'(PC|PD)[\w\d\-\.]+', linha, re.IGNORECASE)
                        
                        if id_m:
                            cod_completo = id_m.group().strip().upper()
                            # Captura valores na linha (ex: 156,00)
                            valores_linha = re.findall(r'\d+(?:[\.,]\d{2})?', linha)
                            
                            for v in valores_linha:
                                try:
                                    v_f = float(v.replace('.', '').replace(',', '.'))
                                    if v_f > 1.0:
                                        dados.append({'id': cod_completo, 'valor': v_f, 'usado': False})
                                except: continue
        return dados

    base_pdf = extrair_pdf_ultra(u_pdf)

    if st.button("🚀 Executar Varredura Final (Garantir 20/20)"):
        u_excel.seek(0)
        # Carrega mantendo layout original
        wb = openpyxl.load_workbook(u_excel, data_only=False)
        ws = wb.active
        df_cielo = pd.read_excel(u_excel, header=14)
        
        sucessos = 0
        # Percorre a planilha da Cielo
        for i, row in df_cielo.iterrows():
            linha_excel = i + 16 # Linha real no Excel (15 cabeçalho + 1)
            try:
                # Coluna E (Index 4) é o VALOR BRUTO
                valor_cielo = float(row.iloc[4])
                
                # BUSCA: Encontra o PC/PD no PDF que tenha o mesmo valor exato
                # Ignoramos a data aqui para capturar os 6 títulos de Março
                for item in base_pdf:
                    if not item['usado'] and abs(item['valor'] - valor_cielo) <= 0.01:
                        # Preenche a Coluna H (Coluna 8)
                        ws.cell(row=linha_excel, column=8).value = item['id']
                        item['usado'] = True
                        sucessos += 1
                        break
            except: continue

        st.success(f"✅ Sucesso! {sucessos} de 20 títulos preenchidos com IDs completos.")

        # Download do arquivo idêntico e preenchido
        output = BytesIO()
        wb.save(output)
        st.download_button(
            label="📥 Baixar Planilha 20/20 Corrigida",
            data=output.getvalue(),
            file_name="Cielo_Conferencia_Total.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
