import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
from datetime import timedelta
import re

st.set_page_config(page_title="Conferência Cartão Cielo", layout="wide")

if 'autenticado' not in st.session_state or not st.session_state.autenticado:
    st.error("🔒 Por favor, faça login na página inicial.")
    st.stop()

st.title("💳 Conferência Cielo - Versão Final (Correção 20/20)")

u_excel = st.file_uploader("1. Planilha Cielo (Original)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    try:
        # 1. MAPEAMENTO DO PDF - Captura agressiva de qualquer código e valor
        base_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        # Captura ID (PC ou PD) - Pega o código completo
                        match_id = re.search(r'(PC|PD)[0-9\-/\\*#]+', linha, re.IGNORECASE)
                        
                        if match_id:
                            cod_id = match_id.group().strip().upper()
                            
                            # Captura TODOS os números da linha (156 | 156,00 | 1.250,50)
                            numeros = re.findall(r'\d+(?:\.\d{3})*(?:,\d{2})?|\d+', linha)
                            
                            for n in numeros:
                                try:
                                    v_limpo = float(n.replace('.', '').replace(',', '.'))
                                    # Mantemos apenas valores que façam sentido como venda
                                    if v_limpo >= 1.0:
                                        base_pdf.append({
                                            'id': cod_id,
                                            'valor': v_limpo,
                                            'usado': False
                                        })
                                except: continue

        if st.button("🚀 Iniciar Conferência (Forçar Preenchimento PC/PD)"):
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            
            # header=14: Coluna E (índice 4) é o Valor Bruto
            df_cielo = pd.read_excel(u_excel, header=14)
            
            sucessos = 0
            for i, row in df_cielo.iterrows():
                lin_ex = i + 16 # Ajuste de linha no Excel
                try:
                    val_ex = float(row.iloc[4]) # Valor Bruto
                    
                    achou = False
                    # Busca prioritária: Valor Exato (com margem 0.02)
                    for t in base_pdf:
                        if not t['usado'] and abs(t['valor'] - val_ex) <= 0.02:
                            # Se o valor bate, vinculamos o ID (PC/PD)
                            ws.cell(row=lin_ex, column=8).value = t['id']
                            t['usado'] = True
                            achou = True
                            sucessos += 1
                            break
                                
                    if not achou:
                        ws.cell(row=lin_ex, column=8).value = "NÃO ENCONTRADO"
                except: continue

            st.success(f"🎯 Finalizado! {sucessos} lançamentos encontrados.")
            
            buffer = BytesIO()
            wb.save(buffer)
            st.download_button(
                label="📥 Baixar Planilha 100% Preenchida",
                data=buffer.getvalue(),
                file_name="Cielo_Final_Conferida.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
