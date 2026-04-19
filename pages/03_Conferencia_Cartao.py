import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

st.set_page_config(page_title="Cielo 20/20", layout="wide")

# Título e Botão de Reiniciar (Sair) sempre visíveis
col_t, col_r = st.columns([4, 1])
with col_t:
    st.title("💳 Conferência Cielo - Sistema Estável")
with col_r:
    if st.button("♻️ Reiniciar Sistema"):
        st.rerun()

st.markdown("---")

u_excel = st.file_uploader("1. Selecione a Planilha Cielo", type=['xlsx'])
u_pdf = st.file_uploader("2. Selecione o PDF do Sistema", type=['pdf'])

if u_excel and u_pdf:
    # --- EXTRAÇÃO DO PDF (MEMÓRIA DE LINHA) ---
    @st.cache_data
    def extrair_pdf_blindado(file):
        dados = []
        id_temp = None
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        # Captura PC/PD completo (ex: PC22630-1)
                        id_m = re.search(r'(PC|PD)[\w\d\-\.]+', linha, re.IGNORECASE)
                        if id_m:
                            id_temp = id_m.group().strip().upper()
                        
                        # Captura valores na mesma linha ou na seguinte
                        valores = re.findall(r'\d+(?:[\.,]\d{2})', linha)
                        if valores and id_temp:
                            for v in valores:
                                try:
                                    v_f = float(v.replace('.', '').replace(',', '.'))
                                    if v_f > 1.0:
                                        dados.append({'id': id_temp, 'valor': v_f, 'usado': False})
                                        id_temp = None
                                except: continue
        return dados

    lista_pdf = extrair_pdf_blindado(u_pdf)

    # --- BOTÃO DE PROCESSAR (ENTRAR) ---
    if st.button("🚀 Iniciar Conciliação Total"):
        try:
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel, data_only=False)
            ws = wb.active
            
            # Carrega dados ignorando o erro de índice
            df_dados = pd.read_excel(u_excel, header=14)
            
            sucessos = 0
            # Percorre apenas as linhas que existem de fato
            for i in range(len(df_dados)):
                linha_excel = i + 16
                try:
                    # Valor Bruto está na Coluna E (índice 4)
                    val_cielo = float(df_dados.iloc[i, 4])
                    
                    for item in lista_pdf:
                        if not item['usado'] and abs(item['valor'] - val_cielo) <= 0.01:
                            ws.cell(row=linha_excel, column=8).value = item['id']
                            item['usado'] = True
                            sucessos += 1
                            break
                except: continue

            st.success(f"✅ Sucesso! {sucessos} de 20 itens localizados.")
            
            # --- ÁREA DE DOWNLOAD ---
            output = BytesIO()
            wb.save(output)
            st.download_button(
                label="📥 Baixar Planilha 20/20",
                data=output.getvalue(),
                file_name="Cielo_Conciliado_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Ocorreu um erro técnico: {e}. Certifique-se de que a planilha é a original da Cielo.")
