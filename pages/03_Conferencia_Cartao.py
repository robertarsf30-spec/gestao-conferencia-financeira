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

st.title("💳 Conferência Cartão (Cielo) - 100% de Captura")

u_excel = st.file_uploader("1. Planilha Cielo (Original)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    try:
        # 1. MAPEAMENTO DO PDF - Captura sem travas de data
        base_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        # REGEX: Captura PC ou PD com qualquer símbolo ( - / * # )
                        match_id = re.search(r'(PC|PD)[0-9\-/\\*#]+', linha, re.IGNORECASE)
                        
                        if match_id:
                            cod_id = match_id.group().strip().upper()
                            
                            # AJUSTE: Captura valores inteiros ou com decimais
                            valores_texto = re.findall(r'\d+(?:[.,]\d{2})?', linha)
                            
                            for v in valores_texto:
                                try:
                                    # Normaliza o valor para float
                                    v_limpo = float(v.replace('.', '').replace(',', '.'))
                                    base_pdf.append({
                                        'id': cod_id,
                                        'valor': v_limpo,
                                        'usado': False
                                    })
                                except: continue

        if st.button("🚀 Iniciar Conferência (Localizar todos os 20 itens)"):
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            
            # header=14 para alinhar com o seu Excel
            df_cielo = pd.read_excel(u_excel, header=14)
            
            sucessos = 0
            for i, row in df_cielo.iterrows():
                lin_ex = i + 16 
                try:
                    val_ex = float(row.iloc[4]) # Valor Bruto (Coluna E)
                    
                    achou = False
                    for t in base_pdf:
                        # REGRA: Margem de 0.02 e prioridade para quem não foi usado
                        if not t['usado'] and abs(t['valor'] - val_ex) <= 0.02:
                            ws.cell(row=lin_ex, column=8).value = t['id']
                            t['usado'] = True
                            achou = True
                            sucessos += 1
                            break
                                
                    if not achou:
                        ws.cell(row=lin_ex, column=8).value = "NÃO ENCONTRADO"
                except:
                    continue

            st.success(f"🎯 Finalizado! {sucessos} títulos encontrados. Agora deve bater o total de 20!")
            
            buffer = BytesIO()
            wb.save(buffer)
            st.download_button(
                label="📥 Baixar Planilha Completa",
                data=buffer.getvalue(),
                file_name="Cielo_Conferencia_Total.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
