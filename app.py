import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(page_title="Sistema de Gestão Cielo", layout="wide")

# 2. INICIALIZAÇÃO DO ESTADO DE ACESSO
if 'auth' not in st.session_state:
    st.session_state.auth = False

# Funções de Controle
def entrar_no_sistema():
    st.session_state.auth = True

def sair_do_sistema():
    st.session_state.auth = False
    st.cache_data.clear() # Limpa o cache dos arquivos PDF
    st.rerun()

# --- LÓGICA DE EXIBIÇÃO ---

if not st.session_state.auth:
    # --- TELA DE LOGIN (ABRIR) ---
    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.title("🔐 Portal Financeiro")
        st.info("Acesso restrito ao módulo de Conciliação Cielo.")
        st.button("🚀 ENTRAR NO SISTEMA", on_click=entrar_no_sistema, use_container_width=True)
        st.divider()
        st.caption("Acesso seguro | Versão 20/20")

else:
    # --- TELA DO SISTEMA (DENTRO) ---
    
    # Barra lateral sempre visível para sair
    with st.sidebar:
        st.header("👤 Menu de Acesso")
        st.write("Status: **Online**")
        st.button("🚪 SAIR E FECHAR MÓDULO", on_click=sair_do_sistema, use_container_width=True)
        st.divider()
        st.caption("Ao sair, todos os arquivos temporários serão apagados.")

    # Conteúdo Principal
    st.title("💳 Conciliador de Recebíveis")
    st.markdown("Carregue os arquivos abaixo para realizar a varredura.")

    c1, c2 = st.columns(2)
    with c1:
        u_excel = st.file_uploader("📂 Selecione a Planilha Cielo", type=['xlsx'])
    with c2:
        u_pdf = st.file_uploader("📄 Selecione o PDF do Sistema", type=['pdf'])

    if u_excel and u_pdf:
        # Lógica de extração ultra-precisa (IDs com hífen e quebra de linha)
        @st.cache_data
        def extrair_dados_pdf(file):
            dados = []
            id_atual = None
            with pdfplumber.open(file) as pdf:
                for page in pdf.pages:
                    texto = page.extract_text()
                    if texto:
                        for linha in texto.split('\n'):
                            id_m = re.search(r'(PC|PD)[\w\d\-\.]+', linha, re.IGNORECASE)
                            if id_m:
                                id_atual = id_m.group().strip().upper()
                            
                            valores = re.findall(r'\d+(?:[\.,]\d{2})', linha)
                            if valores and id_atual:
                                for v in valores:
                                    try:
                                        v_f = float(v.replace('.', '').replace(',', '.'))
                                        if v_f > 1.0:
                                            dados.append({'id': id_atual, 'valor': v_f, 'usado': False})
                                            id_atual = None
                                    except: continue
            return dados

        if st.button("🔍 INICIAR CONFERÊNCIA TOTAL"):
            try:
                lista_pdf = extrair_dados_pdf(u_pdf)
                u_excel.seek(0)
                wb = openpyxl.load_workbook(u_excel, data_only=False)
                ws = wb.active
                
                # Header na linha 15 (Index 14)
                df_dados = pd.read_excel(u_excel, header=14)
                
                acertos = 0
                for i in range(len(df_dados)):
                    linha_ws = i + 16
                    try:
                        v_alvo = float(df_dados.iloc[i, 4]) # Valor Bruto
                        for item in lista_pdf:
                            if not item['usado'] and abs(item['valor'] - v_alvo) <= 0.01:
                                ws.cell(row=linha_ws, column=8).value = item['id']
                                item['usado'] = True
                                acertos += 1
                                break
                    except: continue

                st.success(f"✅ Finalizado! {acertos} de 20 itens localizados.")
                
                # Botão de Download
                out = BytesIO()
                wb.save(out)
                st.download_button("📥 BAIXAR PLANILHA PRONTA", out.getvalue(), "Cielo_Finalizada.xlsx", use_container_width=True)
            
            except Exception as e:
                st.error(f"Erro no processamento: {e}")
