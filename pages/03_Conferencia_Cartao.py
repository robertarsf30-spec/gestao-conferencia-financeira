import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

st.set_page_config(page_title="Conferência Cielo Total", layout="wide")

st.title("💳 Conferência Cielo - Sistema Consolidado 20/20")
st.info("Dica: Use a 'Busca Padrão' primeiro e a 'Busca Avançada' para os itens restantes.")

u_excel = st.file_uploader("1. Planilha Cielo", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    # --- EXTRAÇÃO DO PDF (CACHE PARA VELOCIDADE) ---
    @st.cache_data
    def processar_pdf(pdf_file):
        dados = []
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        id_match = re.search(r'(PC|PD)[^\s]+', linha, re.IGNORECASE)
                        data_match = re.search(r'\d{2}/\d{2}/\d{4}', linha)
                        valores = re.findall(r'\d+(?:[\.,]\d{2})?', linha)
                        
                        if id_match:
                            cod = id_match.group().strip().upper()
                            dt = pd.to_datetime(data_match.group(), dayfirst=True).date() if data_match else None
                            for v in valores:
                                try:
                                    v_f = float(v.replace('.', '').replace(',', '.'))
                                    if v_f > 1.0:
                                        dados.append({'id': cod, 'valor': v_f, 'data': dt, 'usado': False})
                                except: continue
        return dados

    base_pdf = processar_pdf(u_pdf)

    # Inicializa a planilha no estado da sessão (evita ValueError de imagens)
    if 'wb' not in st.session_state:
        u_excel.seek(0)
        # data_only=False mantém fórmulas, mas ignoramos imagens para evitar o erro da imagem 15B
        st.session_state.wb = openpyxl.load_workbook(u_excel, data_only=False)
        # Limpa o erro de imagens que causa falha no salvamento
        st.session_state.wb.vba_archive = None 

    wb = st.session_state.wb
    ws = wb.active
    df_cielo = pd.read_excel(u_excel, header=14)

    col1, col2 = st.columns(2)

    # --- BOTÃO 1: BUSCA PADRÃO ---
    if col1.button("🚀 1. Busca Padrão"):
        cont = 0
        for i, row in df_cielo.iterrows():
            linha_ex = i + 16
            v_alvo = float(row.iloc[4])
            dt_alvo = pd.to_datetime(row.iloc[1]).date()
            
            for item in base_pdf:
                if not item['usado'] and abs(item['valor'] - v_alvo) <= 0.01:
                    if item['data'] and item['data'] == dt_alvo:
                        ws.cell(row=linha_ex, column=8).value = item['id']
                        item['usado'] = True
                        cont += 1
                        break
        st.success(f"Busca padrão finalizada: {cont} itens.")

    # --- BOTÃO 2: BUSCA AVANÇADA ---
    if col2.button("🔍 2. Busca Avançada (Recuperar Restantes)"):
        cont_adv = 0
        for i, row in df_cielo.iterrows():
            linha_ex = i + 16
            # Só processa se a coluna H estiver vazia
            if ws.cell(row=linha_ex, column=8).value is None:
                v_alvo = float(row.iloc[4])
                dt_alvo = pd.to_datetime(row.iloc[1]).date()
                
                for item in base_pdf:
                    if not item['usado']:
                        # Critério: Valor até 0.03 e Data até 3 dias de diferença
                        v_ok = abs(item['valor'] - v_alvo) <= 0.03
                        d_ok = True
                        if item['data']:
                            d_ok = abs((item['data'] - dt_alvo).days) <= 3
                        
                        if v_ok and d_ok:
                            ws.cell(row=linha_ex, column=8).value = item['id']
                            item['usado'] = True
                            cont_adv += 1
                            break
        st.success(f"Busca avançada finalizada: {cont_adv} itens recuperados.")

    # --- BOTÃO DE DOWNLOAD ÚNICO ---
    st.markdown("---")
    buffer = BytesIO()
    # Ajuste técnico para evitar o erro de 'fp.seek(0)' da imagem 15B
    try:
        wb.save(buffer)
        st.download_button(
            label="📥 Baixar Planilha Consolidada (20/20)",
            data=buffer.getvalue(),
            file_name="Conferencia_Cielo_Final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Erro ao gerar arquivo: {e}. Tente remover imagens da sua planilha original.")
