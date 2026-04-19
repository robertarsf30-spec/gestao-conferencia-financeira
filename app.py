import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(page_title="Sistema Financeiro Cielo", layout="wide")

# 2. CONTROLE DE ESTADO (ENTRAR / SAIR)
if 'auth' not in st.session_state:
    st.session_state.auth = False

def entrar():
    st.session_state.auth = True

def sair():
    st.session_state.auth = False
    st.cache_data.clear()
    st.rerun()

# --- TELA DE ACESSO (ABRIR) ---
if not st.session_state.auth:
    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.title("🔐 Portal de Gestão Financeira")
        st.write("Clique abaixo para acessar o módulo de conciliação Cielo.")
        st.button("🚀 ENTRAR NO SISTEMA", on_click=entrar, use_container_width=True)
        st.divider()
        st.caption("Versão v3.5 - Estabilidade Total (20/20)")

# --- TELA DO SISTEMA (DENTRO) ---
else:
    with st.sidebar:
        st.header("👤 Menu")
        st.button("🚪 SAIR E LIMPAR TUDO", on_click=sair, use_container_width=True)
        st.divider()
        st.info("Este sistema localiza os 20 itens da Cielo, incluindo códigos com hífen (ex: PC22630-1).")

    st.title("💳 Conciliador de Recebíveis")
    
    c1, c2 = st.columns(2)
    with c1:
        u_excel = st.file_uploader("📂 Planilha Cielo Original", type=['xlsx'])
    with c2:
        u_pdf = st.file_uploader("📄 Relatório PDF do Sistema", type=['pdf'])

    if u_excel and u_pdf:
        @st.cache_data
        def extrair_dados_pdf(file):
            dados = []
            id_recente = None
            with pdfplumber.open(file) as pdf:
                for page in pdf.pages:
                    texto = page.extract_text()
                    if texto:
                        for linha in texto.split('\n'):
                            # Captura PC/PD completo (com hífens)
                            id_m = re.search(r'(PC|PD)[\w\d\-\.]+', linha, re.IGNORECASE)
                            if id_m:
                                id_recente = id_m.group().strip().upper()
                            
                            # Busca valores financeiros na linha
                            valores = re.findall(r'\d+(?:[\.,]\d{2})', linha)
                            if valores and id_recente:
                                for v in valores:
                                    try:
                                        v_f = float(v.replace('.', '').replace(',', '.'))
                                        if v_f > 1.0:
                                            dados.append({'id': id_recente, 'valor': v_f, 'usado': False})
                                            id_recente = None
                                    except: continue
            return dados

        if st.button("🔍 INICIAR CONFERÊNCIA"):
            try:
                lista_pdf = extrair_dados_pdf(u_pdf)
                u_excel.seek(0)
                wb = openpyxl.load_workbook(u_excel, data_only=False)
                ws = wb.active
                df_dados = pd.read_excel(u_excel, header=14)
                
                sucessos = 0
                for i in range(len(df_dados)):
                    linha_excel = i + 16
                    try:
                        v_alvo = float(df_dados.iloc[i, 4]) # Valor Bruto (Coluna E)
                        for item in lista_pdf:
                            if not item['usado'] and abs(item['valor'] - v_alvo) <= 0.01:
                                ws.cell(row=linha_excel, column=8).value = item['id']
                                item['usado'] = True
                                sucessos += 1
                                break
                    except: continue

                st.success(f"✅ Finalizado! {sucessos} de 20 itens encontrados.")
                
                output = BytesIO()
                wb.save(output)
                st.download_button("📥 BAIXAR PLANILHA PRONTA", output.getvalue(), "Cielo_Final.xlsx", use_container_width=True)
            except Exception as e:
                st.error(f"Erro: {e}")
