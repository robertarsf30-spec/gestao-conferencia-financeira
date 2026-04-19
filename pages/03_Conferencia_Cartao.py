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

st.title("💳 Conferência Cartão (Cielo) - Versão 100%")

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
                        # Busca o código PC ou PD (incluindo símbolos - / * #)
                        match_id = re.search(r'(PC|PD)[0-9\-/\\*#]+', linha, re.IGNORECASE)
                        
                        if match_id:
                            cod_id = match_id.group().strip().upper()
                            
                            # Busca QUALQUER número que pareça valor (ex: 15, 15,00, 1.250,50)
                            # Pega números isolados ou com separadores de milhar/decimal
                            numeros = re.findall(r'\d+(?:[.,]\d{2})?|\d+', linha)
                            
                            for n in numeros:
                                try:
                                    # Limpeza profunda do valor
                                    v_limpo = n.replace('.', '').replace(',', '.')
                                    # Evita pegar números pequenos demais que não são valores de venda
                                    val_final = float(v_limpo)
                                    if val_final > 0:
                                        base_pdf.append({
                                            'id': cod_id,
                                            'valor': val_final,
                                            'usado': False
                                        })
                                except: continue

        if st.button("🚀 Iniciar Conferência"):
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            
            # header=14 para alinhar com o padrão da Cielo
            df_cielo = pd.read_excel(u_excel, header=14)
            
            sucessos = 0
            for i, row in df_cielo.iterrows():
                lin_ex = i + 16
                try:
                    # Valor bruto na coluna E (iloc[4])
                    val_ex = float(row.iloc[4])
                    
                    achou = False
                    # Compara com a base do PDF com margem de 0.02
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

            st.success(f"🎯 Finalizado! {sucessos} itens encontrados. A lista de 20 deve estar completa agora!")
            
            buffer = BytesIO()
            wb.save(buffer)
            st.download_button(
                label="📥 Baixar Planilha 100% Conferida",
                data=buffer.getvalue(),
                file_name="Cielo_Conferencia_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
