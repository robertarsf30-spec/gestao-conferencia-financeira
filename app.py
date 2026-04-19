import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(page_title="Sistema Gestão 20/20", layout="wide")

# --- DEFINIÇÃO DA SENHA (Mude aqui se quiser) ---
SENHA_CORRETA = "1234" 

# 2. INICIALIZAÇÃO DO ESTADO DE ACESSO
if 'autenticado' not in st.session_state:
    st.session_state.autenticado = False

# Funções de Acesso
def realizar_login(senha_digitada):
    if senha_digitada == SENHA_CORRETA:
        st.session_state.autenticado = True
        st.success("Acesso liberado!")
    else:
        st.error("Senha incorreta. Acesso Negado.")

def realizar_logout():
    st.session_state.autenticado = False
    st.cache_data.clear()
    st.rerun()

# --- FLUXO DE TELAS ---

if not st.session_state.autenticado:
    # --- TELA DE LOGIN (ACESSO BLOQUEADO) ---
    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.5, 1])
    
    with col2:
        st.title("🔐 Acesso Restrito")
        st.write("Digite sua senha para desbloquear os módulos de conferência.")
        
        senha_input = st.text_input("Senha do Sistema", type="password")
        
        if st.button("🔓 ENTRAR E DESBLOQUEAR", use_container_width=True):
            realizar_login(senha_input)
            if st.session_state.autenticado:
                st.rerun()

else:
    # --- TELA DO SISTEMA DESBLOQUEADA ---
    
    # Barra lateral com botão de saída
    with st.sidebar:
        st.header("👤 Gestor Ativo")
        st.write("Módulo: **Cielo 20/20**")
        st.divider()
        if st.button("🚪 SAIR E BLOQUEAR TUDO", use_container_width=True):
            realizar_logout()
        st.divider()
        st.caption("Ao sair, os módulos são bloqueados e os dados limpos.")

    # Área de Trabalho
    st.title("💳 Conciliador de Recebíveis Cielo")
    st.info("Todos os módulos estão desbloqueados. Carregue os arquivos abaixo.")

    c1, c2 = st.columns(2)
    with c1:
        u_excel = st.file_uploader("📂 Planilha Cielo Original", type=['xlsx'])
    with c2:
        u_pdf = st.file_uploader("📄 PDF do Sistema", type=['pdf'])

    if u_excel and u_pdf:
        # Lógica de extração que resolve os IDs complexos (ex: PC22630-1)
        @st.cache_data
        def extrair_pdf_ninja(file):
            dados = []
            id_temp = None
            with pdfplumber.open(file) as pdf:
                for page in pdf.pages:
                    texto = page.extract_text()
                    if texto:
                        for linha in texto.split('\n'):
                            # Captura IDs com símbolos e hífens
                            id_m = re.search(r'(PC|PD)[\w\d\-\.]+', linha, re.IGNORECASE)
                            if id_m:
                                id_temp = id_m.group().strip().upper()
                            
                            # Captura valores mesmo com quebra de linha no PDF
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

        if st.button("🚀 PROCESSAR 20 ITENS AGORA"):
            try:
                lista_pdf = extrair_pdf_ninja(u_pdf)
                u_excel.seek(0)
                wb = openpyxl.load_workbook(u_excel, data_only=False)
                ws = wb.active
                df_dados = pd.read_excel(u_excel, header=14)
                
                contador = 0
                for i in range(len(df_dados)):
                    linha_ws = i + 16
                    try:
                        v_alvo = float(df_dados.iloc[i, 4]) # Coluna E
                        for item in lista_pdf:
                            if not item['usado'] and abs(item['valor'] - v_alvo) <= 0.01:
                                ws.
