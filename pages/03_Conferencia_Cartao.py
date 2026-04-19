import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

st.set_page_config(page_title="Conferência Cartão Cielo", layout="wide")

st.title("💳 Conferência Cartão Cielo - Versão Definitiva")

u_excel = st.file_uploader("1. Planilha Cielo (Original)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    try:
        # 1. MAPEAMENTO DO PDF - BUSCA POR CÓDIGO E VALOR NA MESMA LINHA
        base_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        # Busca códigos PC ou PD (incluindo símbolos)
                        match_id = re.search(r'(PC|PD)[0-9\-/\\*#]+', linha, re.IGNORECASE)
                        
                        if match_id:
                            cod_id = match_id.group().strip().upper()
                            
                            # Extrai todos os números que podem ser valores (15,00 ou 15.00 ou 15)
                            numeros_linha = re.findall(r'\d+(?:[.,]\d{2})?', linha)
                            
                            for n in numeros_linha:
                                # Normaliza para float
                                n_limpo = n.replace('.', '').replace(',', '.')
                                try:
                                    base_pdf.append({
                                        'id': cod_id,
                                        'valor': float(n_limpo),
                                        'usado': False
                                    })
                                except: continue

        if st.button("🚀 Iniciar Conferência (Localizar PC22648 e outros)"):
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            # Lê o Excel a partir da linha 11
            df_cielo = pd.read_excel(u_excel, header=10)
            
            sucessos = 0
            for i, row in df_cielo.iterrows():
                lin_ex = i + 12 
                try:
                    # Tenta converter o valor bruto da planilha Cielo
                    val_cielo = float(row['Valor bruto'])
                    
                    achou = False
                    for item in base_pdf:
                        # REGRA: Valor com margem de 0.02
                        if not item['usado'] and abs(item['valor'] - val_cielo) <= 0.02:
                            # Preenche a coluna H (Descrição)
                            ws.cell(row=lin_ex, column=8).value = item['id']
                            item['usado'] = True
                            achou = True
                            sucessos += 1
                            break
                    
                    if not achou:
                        ws.cell(row=lin_ex, column=8).value = "NÃO ENCONTRADO"
                except: continue

            st.success(f"🎯 Finalizado! {sucessos} itens encontrados. Confira o PC22648!")
            
            buffer = BytesIO()
            wb.save(buffer)
            st.download_button(
                label="📥 Baixar Planilha Original Preenchida",
                data=buffer.getvalue(),
                file_name="Cielo_Conferida_Completa.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
