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

st.title("💳 Conferência Cartão (Cielo) - PC e PD")

u_excel = st.file_uploader("1. Planilha Cielo (Original)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    try:
        # 1. MAPEAMENTO DO PDF (BUSCA POR PC E PD)
        base_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        # Filtro amplo: busca qualquer linha que tenha PC ou PD no texto
                        if re.search(r'\b(PC|PD)\d+', linha):
                            partes = linha.split()
                            try:
                                # Localiza o código (ex: PC22650-1 ou PD23052)
                                cod_id = next((x for x in partes if 'PC' in x or 'PD' in x), None)
                                
                                # Localiza a data de lançamento no PDF
                                datas = re.findall(r'\d{2}/\d{2}/\d{4}', linha)
                                dt_lcto = pd.to_datetime(datas[-1], dayfirst=True).date() if datas else None
                                
                                # Extrai todos os valores numéricos da linha
                                valores = []
                                for p in partes:
                                    limpo = p.replace('.', '').replace(',', '.')
                                    # Procura formato de moeda (ex: 156.00)
                                    if re.match(r'^\d+\.\d{2}$', limpo):
                                        valores.append(float(limpo))
                                
                                if cod_id and dt_lcto:
                                    base_pdf.append({
                                        'id': cod_id,
                                        'data': dt_lcto,
                                        'valores': valores,
                                        'usado': False
                                    })
                            except:
                                continue

        if st.button("🚀 Iniciar Conferência (PC + PD)"):
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            df_cielo = pd.read_excel(u_excel, header=14)
            
            sucessos = 0
            for i, row in df_cielo.iterrows():
                lin_ex = i + 16
                try:
                    dt_venda_ex = pd.to_datetime(row.iloc[1], dayfirst=True).date()
                    val_ex = float(row.iloc[4]) # Valor Bruto no Excel
                    
                    achou = False
                    # FUNIL DE BUSCA: 1º Valor -> 2º Data (+2 dias) -> 3º Não Usado
                    for t in base_pdf:
                        if not t['usado']:
                            # Verifica se o valor do Excel está na linha do PDF
                            if any(abs(v_pdf - val_ex) <= 0.05 for v_pdf in t['valores']):
                                # Verifica a janela de 2 dias (Venda <= Lançamento <= Venda + 2)
                                dt_max = dt_venda_ex + timedelta(days=2)
                                if dt_venda_ex <= t['data'] <= dt_max:
                                    ws.cell(row=lin_ex, column=8).value = t['id']
                                    t['usado'] = True
                                    achou = True
                                    sucessos += 1
                                    break
                    
                    if not achou:
                        ws.cell(row=lin_ex, column=8).value = "NÃO ENCONTRADO"
                except:
                    continue

            st.success(f"🎯 Sucesso! {sucessos} vendas (PC/PD) conferidas.")
            
            buffer = BytesIO()
            wb.save(buffer)
            st.download_button(
                label="📥 Baixar Planilha Final",
                data=buffer.getvalue(),
                file_name="Cielo_Conferida_Completa.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
