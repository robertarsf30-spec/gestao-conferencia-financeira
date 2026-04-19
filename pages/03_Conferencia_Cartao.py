import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

st.set_page_config(page_title="Conferência Cartão Cielo", layout="wide")

st.title("💳 Conferência Cartão (Cielo) - Versão Definitiva")

u_excel = st.file_uploader("1. Planilha Cielo (Original)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    try:
        # 1. MAPEAMENTO DO PDF (Baseado no que funcionou para os PD)
        base_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        # Regex aprimorada para pegar PC/PD com símbolos: / - * #
                        # Ex: PC22650-1, PD/23041, PC*889, PD#123
                        match_id = re.search(r'(PC|PD)[0-9\-/\\*#]+', linha, re.IGNORECASE)
                        
                        if match_id:
                            cod_id = match_id.group().strip()
                            
                            # Extrai valores da linha (considerando que o valor pode estar em qualquer lugar)
                            # Pega formatos como 1.234,56 ou 156,00
                            valores_encontrados = re.findall(r'\d+(?:\.\d{3})*(?:,\d{2})', linha)
                            
                            for v in valores_encontrados:
                                # Converte para float: remove ponto de milhar e troca vírgula por ponto
                                v_limpo = v.replace('.', '').replace(',', '.')
                                base_pdf.append({
                                    'id': cod_id,
                                    'valor': float(v_limpo),
                                    'usado': False
                                })

        if st.button("🚀 Iniciar Conferência (Aprimorada)"):
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            
            # header=10 para ler a partir da linha 11 do seu Excel
            df_cielo = pd.read_excel(u_excel, header=10)
            
            sucessos = 0
            for i, row in df_cielo.iterrows():
                lin_ex = i + 12 
                try:
                    # Pega o valor bruto da coluna E
                    val_ex = float(row['Valor bruto'])
                    
                    achou = False
                    for t in base_pdf:
                        if not t['usado']:
                            # REGRA: Foco no valor com margem de 0.02
                            if abs(t['valor'] - val_ex) <= 0.02:
                                ws.cell(row=lin_ex, column=8).value = t['id']
                                t['usado'] = True
                                achou = True
                                sucessos += 1
                                break
                    
                    if not achou:
                        ws.cell(row=lin_ex, column=8).value = "NÃO ENCONTRADO"
                except:
                    continue

            st.success(f"✅ Sucesso! {sucessos} títulos vinculados (incluindo PC/PD com símbolos).")
            
            buffer = BytesIO()
            wb.save(buffer)
            st.download_button(
                label="📥 Baixar Planilha Finalizada",
                data=buffer.getvalue(),
                file_name="Cielo_Conferida_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
