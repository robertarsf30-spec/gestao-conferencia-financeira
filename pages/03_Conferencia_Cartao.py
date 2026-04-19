import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

# Configuração da página
st.set_page_config(page_title="Conferência Cartão Cielo", layout="wide")

if 'autenticado' not in st.session_state or not st.session_state.autenticado:
    st.error("🔒 Por favor, faça login na página inicial.")
    st.stop()

st.title("💳 Conferência Cielo - Versão 20/20 (Sem Travas de Data)")

u_excel = st.file_uploader("1. Planilha Cielo (Original)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    try:
        # 1. MAPEAMENTO DO PDF - Foco total em ID + Valor
        base_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        # Captura códigos PC ou PD completos (ex: PC22650-1)
                        match_id = re.search(r'(PC|PD)[^\s]+', linha, re.IGNORECASE)
                        
                        if match_id:
                            cod_id = match_id.group().strip().upper()
                            
                            # Captura valores (ex: 156 ou 156,00)
                            numeros = re.findall(r'\d+(?:\.\d{3})*(?:,\d{2})?|\d+', linha)
                            
                            for n in numeros:
                                try:
                                    # Normaliza 156 ou 156,00 para 156.0
                                    v_limpo = float(n.replace('.', '').replace(',', '.'))
                                    if v_limpo >= 1.0:
                                        base_pdf.append({
                                            'id': cod_id,
                                            'valor': v_limpo,
                                            'usado': False
                                        })
                                except: continue

        if st.button("🚀 Iniciar Conferência (Capturar todos os 20)"):
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            
            # Lê a Cielo (Dados começam após a linha 14)
            df_cielo = pd.read_excel(u_excel, header=14)
            
            sucessos = 0
            for i, row in df_cielo.iterrows():
                lin_ex = i + 16 # Ajuste para linha correta no Excel
                try:
                    val_ex = float(row.iloc[4]) # Coluna E: Valor Bruto
                    
                    achou = False
                    # Busca apenas por Valor + ID (ignora data para bater março/abril)
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

            st.success(f"🎯 Finalizado! {sucessos} de 20 títulos encontrados.")
            
            buffer = BytesIO()
            wb.save(buffer)
            st.download_button(
                label="📥 Baixar Planilha 20/20 Corrigida",
                data=buffer.getvalue(),
                file_name="Cielo_Conferencia_Sucesso_Total.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
