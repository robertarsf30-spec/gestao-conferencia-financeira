import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

st.set_page_config(page_title="Conferência Cartão Cielo", layout="wide")

st.title("💳 Conferência Cartão Cielo (Versão Final)")

u_excel = st.file_uploader("1. Planilha Cielo (Original)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    try:
        # 1. MAPEAMENTO DO PDF - Extração de alta sensibilidade
        base_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        # Busca códigos PC/PD que podem conter / - * #
                        # Ex: PC22650-1 ou PD/23041
                        match_id = re.search(r'(PC|PD)[0-9\-/\\*#]+', linha, re.IGNORECASE)
                        
                        if match_id:
                            cod_id = match_id.group().strip()
                            
                            # Busca todos os números com padrão de moeda na linha (ex: 156,00)
                            valores_encontrados = re.findall(r'\d+(?:\.\d{3})*(?:,\d{2})', linha)
                            
                            for v in valores_encontrados:
                                # Converte para float removendo pontos de milhar e ajustando a vírgula
                                v_limpo = float(v.replace('.', '').replace(',', '.'))
                                base_pdf.append({
                                    'id': cod_id,
                                    'valor': v_limpo,
                                    'usado': False
                                })

        if st.button("🚀 Iniciar Conferência"):
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            
            # Lê o Excel a partir da linha 11
            df_cielo = pd.read_excel(u_excel, header=10)
            
            sucessos = 0
            for i, row in df_cielo.iterrows():
                lin_ex = i + 12 
                try:
                    # Valor bruto da coluna E
                    val_alvo = float(row['Valor bruto'])
                    
                    achou = False
                    for t in base_pdf:
                        if not t['usado']:
                            # Critério de diferença de até 0,02
                            if abs(t['valor'] - val_alvo) <= 0.02:
                                # Preenche a coluna H (8)
                                ws.cell(row=lin_ex, column=8).value = t['id']
                                t['usado'] = True
                                achou = True
                                sucessos += 1
                                break
                    
                    if not achou:
                        ws.cell(row=lin_ex, column=8).value = "NÃO ENCONTRADO"
                except:
                    continue

            st.success(f"🎯 Finalizado! {sucessos} itens encontrados e vinculados.")
            
            buffer = BytesIO()
            wb.save(buffer)
            st.download_button(
                label="📥 Baixar Planilha Preenchida",
                data=buffer.getvalue(),
                file_name="Conferencia_Cielo_Finalizada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
