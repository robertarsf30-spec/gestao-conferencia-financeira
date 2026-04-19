import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

st.set_page_config(page_title="Conferência Cartão Cielo", layout="wide")

st.title("💳 Scanner de Conferência (Foco Total em Valor)")

u_excel = st.file_uploader("1. Planilha Cielo (Original)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    try:
        # 1. MAPEAMENTO RADICAL DO PDF
        base_geral_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                palavras = page.extract_words()
                # Transforma a página em uma lista de strings para busca rápida
                texto_pagina = page.extract_text()
                linhas = texto_pagina.split('\n') if texto_pagina else []
                
                for linha in linhas:
                    # Busca qualquer código que pareça PC... ou PD... com qualquer símbolo
                    codigos = re.findall(r'(?:PC|PD)[0-9\-/\\*#]+', linha, re.IGNORECASE)
                    # Busca qualquer coisa que pareça valor financeiro (ex: 156,00 ou 156.00)
                    valores = re.findall(r'\d+[.,]\d{2}', linha)
                    
                    if codigos and valores:
                        id_encontrado = codigos[0].strip()
                        for v in valores:
                            v_float = float(v.replace('.', '').replace(',', '.'))
                            base_geral_pdf.append({
                                'id': id_encontrado,
                                'valor': v_float,
                                'usado': False
                            })

        if st.button("🚀 Iniciar Scanner de Precisão"):
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            
            # Lê o Excel - Começando da linha 11 (header=10) conforme seu print
            df_cielo = pd.read_excel(u_excel, header=10)
            
            sucessos = 0
            for i, row in df_cielo.iterrows():
                lin_ex = i + 12 
                try:
                    # Pega o valor bruto da coluna E
                    val_alvo = float(row['Valor bruto'])
                    
                    encontrado = False
                    # Busca na base do PDF pelo valor com margem de 0.02
                    for item in base_geral_pdf:
                        if not item['usado']:
                            if abs(item['valor'] - val_alvo) <= 0.02:
                                # Escreve o código encontrado na coluna H (Descrição)
                                ws.cell(row=lin_ex, column=8).value = item['id']
                                item['usado'] = True
                                encontrado = True
                                sucessos += 1
                                break
                    
                    if not encontrado:
                        ws.cell(row=lin_ex, column=8).value = "NÃO ENCONTRADO"
                except:
                    continue

            st.success(f"🎯 Scanner Finalizado! {sucessos} itens vinculados.")
            
            buffer = BytesIO()
            wb.save(buffer)
            st.download_button(
                label="📥 Baixar Planilha Preenchida",
                data=buffer.getvalue(),
                file_name="Cielo_Conferencia_Scanner.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro no Scanner: {e}")
