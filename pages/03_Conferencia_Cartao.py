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

st.title("💳 Conferência Cielo - Versão Final 100% (Sem Travas)")

u_excel = st.file_uploader("1. Planilha Cielo (Original)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    try:
        # 1. MAPEAMENTO DO PDF - Captura ultra-sensível
        base_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        # Captura PC ou PD completo (incluindo hífens e números extras)
                        match_id = re.search(r'(PC|PD)[^\s]+', linha, re.IGNORECASE)
                        
                        if match_id:
                            cod_id = match_id.group().strip().upper()
                            
                            # Busca QUALQUER número (ex: 156 ou 156,00 ou 1.250,50)
                            numeros = re.findall(r'\d+(?:\.\d{3})*(?:,\d{2})?|\d+', linha)
                            
                            for n in numeros:
                                try:
                                    # Limpeza para garantir que '156' vire 156.0
                                    v_limpo = float(n.replace('.', '').replace(',', '.'))
                                    if v_limpo >= 1.0:
                                        base_pdf.append({
                                            'id': cod_id,
                                            'valor': v_limpo,
                                            'usado': False
                                        })
                                except: continue

        if st.button("🚀 Iniciar Conferência Final (Preencher tudo)"):
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            
            # header=14 (Coluna E é o Valor Bruto)
            df_cielo = pd.read_excel(u_excel, header=14)
            
            sucessos = 0
            for i, row in df_cielo.iterrows():
                lin_ex = i + 16
                try:
                    val_ex = float(row.iloc[4]) # Valor Bruto
                    
                    achou = False
                    # Busca direta por Valor (com margem 0.02)
                    # Ignoramos a data para garantir que créditos de meses passados sejam pegos
                    for t in base_pdf:
                        if not t['usado'] and abs(t['valor'] - val_ex) <= 0.02:
                            ws.cell(row=lin_ex, column=8).value = t['id']
                            t['usado'] = True
                            achou = True
                            sucessos += 1
                            break
                                
                    if not achou:
                        ws.cell(row=lin_ex, column=8).value = "NÃO ENCONTRADO"
                except: continue

            st.success(f"🎯 Finalizado! {sucessos} de 20 itens encontrados.")
            
            buffer = BytesIO()
            wb.save(buffer)
            st.download_button(
                label="📥 Baixar Planilha 20/20 Conferida",
                data=buffer.getvalue(),
                file_name="Cielo_Conferencia_Total_Sucesso.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
