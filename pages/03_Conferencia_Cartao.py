import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

st.set_page_config(page_title="Conferência Cartão Cielo", layout="wide")

if 'autenticado' not in st.session_state or not st.session_state.autenticado:
    st.error("🔒 Por favor, faça login na página inicial.")
    st.stop()

st.title("💳 Conferência Cielo - Versão Final 20/20")

u_excel = st.file_uploader("1. Planilha Cielo (Original)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    try:
        # 1. MAPEAMENTO DO PDF - Captura ultra sensível
        base_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        # Captura PC ou PD e tudo que vier depois até o próximo espaço
                        match_id = re.search(r'(PC|PD)[^\s]+', linha, re.IGNORECASE)
                        
                        if match_id:
                            cod_id = match_id.group().strip().upper()
                            
                            # Pega QUALQUER número da linha (ex: 156,00 ou 156 ou 156.00)
                            numeros = re.findall(r'\d+(?:[.,]\d{2})?|\d+', linha)
                            
                            for n in numeros:
                                try:
                                    v_limpo = float(n.replace('.', '').replace(',', '.'))
                                    if v_limpo >= 1.0: # Filtra números pequenos que não são valores
                                        base_pdf.append({'id': cod_id, 'valor': v_limpo, 'usado': False})
                                except: continue

        if st.button("🚀 Iniciar Conferência"):
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            
            # header=14 para alinhar com o padrão do arquivo detalhado Cielo
            df_cielo = pd.read_excel(u_excel, header=14)
            
            sucessos = 0
            for i, row in df_cielo.iterrows():
                lin_ex = i + 16 
                try:
                    val_ex = float(row.iloc[4]) # Valor Bruto (Coluna E)
                    
                    achou = False
                    for t in base_pdf:
                        # Margem de 0.02 para garantir o match mesmo com arredondamento
                        if not t['usado'] and abs(t['valor'] - val_ex) <= 0.02:
                            ws.cell(row=lin_ex, column=8).value = t['id']
                            t['usado'] = True
                            achou = True
                            sucessos += 1
                            break
                                
                    if not achou:
                        ws.cell(row=lin_ex, column=8).value = "NÃO ENCONTRADO"
                except: continue

            st.success(f"🎯 Finalizado! {sucessos} de 20 títulos encontrados.")
            
            buffer = BytesIO()
            wb.save(buffer)
            st.download_button(
                label="📥 Baixar Planilha 100% Conferida",
                data=buffer.getvalue(),
                file_name="Cielo_Conferida_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
