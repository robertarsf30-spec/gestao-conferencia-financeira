import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

st.set_page_config(page_title="Conferência Cielo 20/20", layout="wide")

st.title("💳 Conferência Cielo - Sistema Unificado")
st.markdown("---")

u_excel = st.file_uploader("1. Planilha Cielo (Original)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    # --- PROCESSAMENTO DO PDF ---
    @st.cache_data
    def extrair_pdf(file):
        base = []
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        # Regex para capturar IDs com símbolos (ex: PC22650-1)
                        match_id = re.search(r'(PC|PD)[^\s]+', linha, re.IGNORECASE)
                        match_data = re.search(r'\d{2}/\d{2}/\d{4}', linha)
                        if match_id:
                            cod = match_id.group().strip().upper()
                            dt = pd.to_datetime(match_data.group(), dayfirst=True).date() if match_data else None
                            # Captura valores na linha
                            nums = re.findall(r'\d+(?:[\.,]\d{2})?', linha)
                            for n in nums:
                                try:
                                    v_f = float(n.replace('.', '').replace(',', '.'))
                                    if v_f > 1.0:
                                        base.append({'id': cod, 'valor': v_f, 'data': dt, 'usado': False})
                                except: continue
        return base

    lista_pdf = extrair_pdf(u_pdf)

    # Inicialização segura da planilha no estado da sessão
    if 'wb' not in st.session_state:
        u_excel.seek(0)
        # Carrega sem imagens para evitar o erro "I/O operation on closed file"
        st.session_state.wb = openpyxl.load_workbook(u_excel, data_only=False)
        st.session_state.wb.vba_archive = None 

    wb = st.session_state.wb
    ws = wb.active
    df_cielo = pd.read_excel(u_excel, header=14)

    col1, col2 = st.columns(2)

    # --- BOTÃO 1: BUSCA PADRÃO ---
    if col1.button("🚀 1. Busca Padrão"):
        achados_p = 0
        for i, row in df_cielo.iterrows():
            linha_ex = i + 16
            v_alvo = float(row.iloc[4])
            dt_alvo = pd.to_datetime(row.iloc[1]).date()
            
            for item in lista_pdf:
                if not item['usado'] and abs(item['valor'] - v_alvo) <= 0.01:
                    if item['data'] and item['data'] == dt_alvo:
                        ws.cell(row=linha_ex, column=8).value = item['id']
                        item['usado'] = True
                        achados_p += 1
                        break
        st.success(f"Busca Padrão: {achados_p} itens conciliados.")

    # --- BOTÃO 2: BUSCA AVANÇADA (SÓ O QUE FALTOU) ---
    if col2.button("🔍 2. Busca Avançada (Recuperar Faltantes)"):
        achados_a = 0
        for i, row in df_cielo.iterrows():
            linha_ex = i + 16
            # Só tenta se a célula estiver vazia
            if ws.cell(row=linha_ex, column=8).value in [None, "NÃO ENCONTRADO"]:
                v_alvo = float(row.iloc[4])
                dt_alvo = pd.to_datetime(row.iloc[1]).date()
                
                for item in lista_pdf:
                    if not item['usado']:
                        # Critério flexível: valor 0.03 e data até 3 dias
                        v_ok = abs(item['valor'] - v_alvo) <= 0.03
                        d_ok = True
                        if item['data']:
                            d_ok = abs((item['data'] - dt_alvo).days) <= 3
                        
                        if v_ok and d_ok:
                            ws.cell(row=linha_ex, column=8).value = item['id']
                            item['usado'] = True
                            achados_a += 1
                            break
        st.success(f"Busca Avançada: {achados_a} itens recuperados.")

    # --- EXIBIÇÃO E DOWNLOAD ---
    st.markdown("---")
    output = BytesIO()
    try:
        wb.save(output)
        st.download_button(
            label="📥 Baixar Planilha Consolidada (20/20)",
            data=output.getvalue(),
            file_name="Conferencia_Cielo_Final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Erro ao salvar: {e}. Certifique-se de que a planilha original não tenha imagens.")
