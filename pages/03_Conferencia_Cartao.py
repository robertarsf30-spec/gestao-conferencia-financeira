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

st.title("💳 Conferência Cartão (Cielo) - Versão Final PC/PD")

u_excel = st.file_uploader("1. Planilha Cielo (Original)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    try:
        # 1. MAPEAMENTO DO PDF - Focado em capturar códigos complexos e valores
        base_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        # REGEX: Captura PC ou PD com qualquer símbolo ( - / * # )
                        match_id = re.search(r'(PC|PD)[0-9\-/\\*#]+', linha, re.IGNORECASE)
                        
                        if match_id:
                            cod_id = match_id.group().strip()
                            
                            # Extrai valores (considera ponto de milhar e vírgula decimal)
                            valores_texto = re.findall(r'\d+(?:\.\d{3})*(?:,\d{2})', linha)
                            valores_float = [float(v.replace('.', '').replace(',', '.')) for v in valores_texto]
                            
                            # Busca data na linha, mas não torna obrigatória para os PC
                            datas = re.findall(r'\d{2}/\d{2}/\d{4}', linha)
                            dt_lcto = pd.to_datetime(datas[-1], dayfirst=True).date() if datas else None
                            
                            if cod_id and valores_float:
                                base_pdf.append({
                                    'id': cod_id,
                                    'data': dt_lcto,
                                    'valores': valores_float,
                                    'usado': False
                                })

        if st.button("🚀 Iniciar Conferência (Capturar PC e PD)"):
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            
            # header=14 conforme sua estrutura
            df_cielo = pd.read_excel(u_excel, header=14)
            
            sucessos = 0
            for i, row in df_cielo.iterrows():
                lin_ex = i + 16 # Ajuste para a linha correta no Excel
                try:
                    val_ex = float(row.iloc[4]) # Valor Bruto (Coluna E)
                    
                    achou = False
                    for t in base_pdf:
                        # REGRA: Margem de 0.02 centavos
                        if not t['usado'] and any(abs(v_pdf - val_ex) <= 0.02 for v_pdf in t['valores']):
                            
                            # Para os PC, se o valor bate e o código existe, fazemos o vínculo
                            # sem travar rigidamente na data de lançamento
                            ws.cell(row=lin_ex, column=8).value = t['id']
                            t['usado'] = True
                            achou = True
                            sucessos += 1
                            break
                                
                    if not achou:
                        ws.cell(row=lin_ex, column=8).value = "NÃO ENCONTRADO"
                except:
                    continue

            st.success(f"🎯 Finalizado! {sucessos} títulos (PC/PD) encontrados com sucesso.")
            
            buffer = BytesIO()
            wb.save(buffer)
            st.download_button(
                label="📥 Baixar Planilha Preenchida",
                data=buffer.getvalue(),
                file_name="Cielo_Conferida_Completa.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
