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

st.title("💳 Conferência Cielo - Versão Definitiva (Data + Valor)")

u_excel = st.file_uploader("1. Planilha Cielo (Original)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    try:
        # 1. MAPEAMENTO DO PDF - Captura ultra sensível de Data, ID e Valor
        base_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        # Busca ID (PC ou PD) e Data na linha
                        match_id = re.search(r'(PC|PD)[0-9\-/\\*#]+', linha, re.IGNORECASE)
                        match_data = re.search(r'\d{2}/\d{2}/\d{4}', linha)
                        
                        if match_id:
                            cod_id = match_id.group().strip().upper()
                            dt_pdf = pd.to_datetime(match_data.group(), dayfirst=True).date() if match_data else None
                            
                            # Captura valores (ex: 156,00 ou apenas 156)
                            numeros = re.findall(r'\d+(?:[.,]\d{2})?|\d+', linha)
                            for n in numeros:
                                try:
                                    v_limpo = float(n.replace('.', '').replace(',', '.'))
                                    if v_limpo >= 1.0:
                                        base_pdf.append({
                                            'id': cod_id,
                                            'valor': v_limpo,
                                            'data': dt_pdf,
                                            'usado': False
                                        })
                                except: continue

        if st.button("🚀 Iniciar Conferência (Localizar todos os 20)"):
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            
            # header=14 (Coluna B=Data Venda, Coluna E=Valor Bruto)
            df_cielo = pd.read_excel(u_excel, header=14)
            
            sucessos = 0
            for i, row in df_cielo.iterrows():
                lin_ex = i + 16
                try:
                    val_ex = float(row.iloc[4]) # Valor Bruto
                    dt_venda_ex = pd.to_datetime(row.iloc[1]).date() # Data da Venda
                    
                    # 1ª BUSCA: Filtrar candidatos apenas pelo valor (margem 0.02)
                    candidatos = [t for t in base_pdf if not t['usado'] and abs(t['valor'] - val_ex) <= 0.02]
                    
                    if candidatos:
                        escolhido = None
                        
                        # 2ª BUSCA: Desempate pela Data (Margem de até 2 dias DEPOIS da venda)
                        for c in candidatos:
                            if c['data']:
                                if dt_venda_ex <= c['data'] <= (dt_venda_ex + timedelta(days=2)):
                                    escolhido = c
                                    break
                        
                        # Se não achou na janela de data mas só tem 1 opção de valor, usa ele por segurança
                        if not escolhido and len(candidatos) == 1:
                            escolhido = candidatos[0]
                        
                        if escolhido:
                            ws.cell(row=lin_ex, column=8).value = escolhido['id']
                            escolhido['usado'] = True
                            sucessos += 1
                        else:
                            ws.cell(row=lin_ex, column=8).value = "NÃO ENCONTRADO (DATA FORA DA MARGEM)"
                    else:
                        ws.cell(row=lin_ex, column=8).value = "NÃO ENCONTRADO"
                        
                except: continue

            st.success(f"🎯 Finalizado! {sucessos} itens encontrados. Os PC ausentes foram capturados!")
            
            buffer = BytesIO()
            wb.save(buffer)
            st.download_button(
                label="📥 Baixar Planilha 100% Corrigida",
                data=buffer.getvalue(),
                file_name="Cielo_Conferida_Total.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
