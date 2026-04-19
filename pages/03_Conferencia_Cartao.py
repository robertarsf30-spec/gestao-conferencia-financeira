import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

st.set_page_config(page_title="Conferência Cartão Cielo", layout="wide")

st.title("💳 Conferência Cartão (Cielo) - Foco em Valor e Símbolos")

u_excel = st.file_uploader("1. Planilha Cielo (Original)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    try:
        # 1. MAPEAMENTO DO PDF (BUSCA ULTRA FLEXÍVEL)
        base_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        # Busca códigos PC/PD com símbolos: / - * #
                        match_id = re.search(r'(PC|PD)[0-9\-/ \*\#]+', linha, re.IGNORECASE)
                        
                        if match_id:
                            cod_id = match_id.group().strip()
                            # Extrai todos os números da linha que parecem valores (0.00)
                            numeros = re.findall(r'\d+[.,]\d{2}', linha)
                            
                            for n in numeros:
                                valor_f = float(n.replace('.', '').replace(',', '.'))
                                base_pdf.append({
                                    'id': cod_id,
                                    'valor': valor_f,
                                    'usado': False
                                })

        if st.button("🚀 Iniciar Conferência (Foco em Valor)"):
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            # header=10 para ler a partir da linha 11
            df_cielo = pd.read_excel(u_excel, header=10)
            
            sucessos = 0
            for i, row in df_cielo.iterrows():
                lin_ex = i + 12 
                try:
                    # Tenta ler o valor bruto da coluna E
                    val_ex = float(row['Valor bruto'])
                    
                    achou = False
                    for t in base_pdf:
                        if not t['usado']:
                            # REGRA 1: Valor com margem de 0.02
                            if abs(t['valor'] - val_ex) <= 0.02:
                                ws.cell(row=lin_ex, column=8).value = t['id']
                                t['usado'] = True
                                achou = True
                                sucessos += 1
                                break
                    
                    if not achou:
                        ws.cell(row=lin_ex, column=8).value = "NÃO ENCONTRADO"
                except:
                    continue

            st.success(f"🎯 Finalizado! {sucessos} itens encontrados pelo valor.")
            
            buffer = BytesIO()
            wb.save(buffer)
            st.download_button(
                label="📥 Baixar Planilha Preenchida",
                data=buffer.getvalue(),
                file_name="Cielo_Conferencia_Símbolos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
