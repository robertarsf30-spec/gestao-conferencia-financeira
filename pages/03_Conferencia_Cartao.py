import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

st.set_page_config(page_title="Conferência Cartão Cielo", layout="wide")

st.title("💳 Conferência Cartão (Cielo) - Final")

u_excel = st.file_uploader("1. Planilha Cielo (Original)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    try:
        # 1. MAPEAMENTO DO PDF (ACEITANDO / E -)
        base_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        # Regex melhorada: busca PC/PD seguidos de números, podendo ter / ou -
                        # Ex: PC22650-1, PC22650/1, PD23052
                        match_id = re.search(r'\b(PC|PD)[\w/-]+\b', linha)
                        
                        if match_id:
                            cod_id = match_id.group()
                            partes = linha.split()
                            try:
                                # Captura o valor (geralmente o último da linha)
                                val_texto = partes[-1].replace('.', '').replace(',', '.')
                                valor_pdf = float(val_texto)
                                
                                base_pdf.append({
                                    'id': cod_id, 
                                    'valor': valor_pdf, 
                                    'usado': False
                                })
                            except:
                                continue

        if st.button("🚀 Iniciar Conferência (PC com / e -)"):
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            # header=10 para começar a ler da linha 11 do Excel
            df_cielo = pd.read_excel(u_excel, header=10)
            
            sucessos = 0
            for i, row in df_cielo.iterrows():
                # lin_ex ajustada para bater com a célula do Excel (i + header + 2)
                lin_ex = i + 12 
                try:
                    # Verifica se a coluna de valor bruto existe e tem número
                    val_ex = float(row['Valor bruto'])
                    
                    achou = False
                    for t in base_pdf:
                        if not t['usado']:
                            # REGRA: Diferença de até 0,02 centavos
                            if abs(t['valor'] - val_ex) <= 0.02:
                                ws.cell(row=lin_ex, column=8).value = t['id']
                                t['usado'] = True
                                achou = True
                                sucessos += 1
                                break
                    
                    if not achou:
                        # Se não achar nada, deixa em branco ou mantém o que já tinha
                        pass
                except:
                    continue

            st.success(f"🎯 Finalizado! {sucessos} títulos (PC/PD) vinculados com sucesso.")
            
            buffer = BytesIO()
            wb.save(buffer)
            st.download_button(
                label="📥 Baixar Planilha Original Preenchida",
                data=buffer.getvalue(),
                file_name="Cielo_Conferencia_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
