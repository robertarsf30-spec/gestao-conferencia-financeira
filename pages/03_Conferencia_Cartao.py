import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

# Configuração da Página
st.set_page_config(page_title="Gestão Financeira - Cielo", layout="wide")

# --- BARRA LATERAL (CONTROLE DE ACESSO) ---
with st.sidebar:
    st.header("🎮 Painel de Controle")
    btn_entrar = st.button("🚀 Entrar e Processar", use_container_width=True)
    if st.button("🚪 Sair e Limpar Tudo", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    st.markdown("---")
    st.info("Esta versão preserva o layout original da planilha e busca IDs com hífens (ex: PC22630-1).")

# --- ÁREA PRINCIPAL ---
st.title("💳 Conciliador de Recebíveis Cielo")
st.markdown("Faça o upload dos arquivos para iniciar a conferência total (20/20).")

col1, col2 = st.columns(2)
with col1:
    u_excel = st.file_uploader("📂 Planilha Cielo Original", type=['xlsx'])
with col2:
    u_pdf = st.file_uploader("📄 Relatório PDF do Sistema", type=['pdf'])

if u_excel and u_pdf:
    # EXTRAÇÃO DO PDF COM MEMÓRIA DE CONTEXTO
    @st.cache_data
    def extrair_pdf_total(file):
        dados = []
        id_pendente = None
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        # Captura PC/PD completo, incluindo símbolos e hífens
                        id_m = re.search(r'(PC|PD)[\w\d\-\.]+', linha, re.IGNORECASE)
                        if id_m:
                            id_pendente = id_m.group().strip().upper()
                        
                        # Busca valores financeiros (mesmo que em linhas diferentes do ID)
                        valores = re.findall(r'\d+(?:[\.,]\d{2})', linha)
                        if valores and id_pendente:
                            for v in valores:
                                try:
                                    v_f = float(v.replace('.', '').replace(',', '.'))
                                    if v_f > 1.0:
                                        dados.append({'id': id_pendente, 'valor': v_f, 'usado': False})
                                        id_pendente = None # Reset para próxima busca
                                except: continue
        return dados

    lista_pdf = extrair_pdf_total(u_pdf)

    # PROCESSAMENTO DISPARADO PELO BOTÃO "ENTRAR"
    if btn_entrar:
        try:
            u_excel.seek(0)
            # Abre preservando formatação, fórmulas e imagens
            wb = openpyxl.load_workbook(u_excel, data_only=False)
            ws = wb.active
            
            # Lê dados para comparação (Header na linha 15)
            df_dados = pd.read_excel(u_excel, header=14)
            
            contagem = 0
            # Varredura segura linha por linha
            for i in range(len(df_dados)):
                linha_planilha = i + 16 # Linha 15
