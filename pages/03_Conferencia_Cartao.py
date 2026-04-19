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
        # 1. MAPEAMENTO DO PDF - Lógica que funcionou anteriormente
        base_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        # REGEX TURBINADA: Captura PC ou PD seguidos de números e símbolos / - * #
                        match_id = re.search(r'(PC|PD)[0-9\-/\\*#]+', linha, re.IGNORECASE)
                        
                        if match_id:
                            cod_id = match_id.group().strip()
                            
                            # Busca valores financeiros na linha (ex: 1.250,00 ou 85,02)
                            valores_linha = re.findall(r'\d+(?:\.\d{3})*(?:,\d{2})', linha)
                            
                            for v in valores_linha:
                                # Converte o valor para float (limpa ponto e troca vírgula por ponto)
                                valor_limpo = float(v.replace('.', '').replace(',', '.'))
                                base_pdf.append({
                                    'id': cod_id,
                                    'valor': valor_limpo,
                                    'usado': False
                                })

        if st.button("🚀 Iniciar Conferência"):
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            
            # Lê o Excel a partir da linha 11 (header=10)
            df_cielo = pd.read_excel(u_excel, header=10)
            
            sucessos = 0
            for i, row in df_cielo.iterrows():
                # Linha real no Excel (i + 12 devido ao header 10 e index 0)
                lin_ex = i + 12 
                try:
                    # Valor bruto da coluna E
                    val_cielo = float(row['Valor bruto'])
                    
                    achou = False
                    for item in base_pdf:
                        if not item['usado']:
                            # REGRA DE OURO: Diferença de até 0,02
                            if abs(item['valor'] - val_cielo) <= 0.02:
                                # Preenche a coluna H (Descrição)
                                ws.cell(row=lin_ex, column=8).value = item['id']
                                item['usado'] = True
                                achou = True
                                sucessos += 1
                                break
                    
                    if not achou:
                        ws.cell(row=lin_ex, column=8).value = "NÃO ENCONTRADO"
                except:
                    continue

            st.success(f"✅ Finalizado! {sucessos} itens encontrados e vinculados.")
            
            # Gerar download
            buffer = BytesIO()
            wb.save(buffer)
            st.download_button(
                label="📥 Baixar Planilha Original Preenchida",
                data=buffer.getvalue(),
                file_name="Cielo_Conferida_Ajustada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Ocorreu um erro no processamento: {e}")
