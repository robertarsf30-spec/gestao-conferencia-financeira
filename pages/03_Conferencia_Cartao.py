import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

st.set_page_config(page_title="Solução Final 20/20", layout="wide")
st.title("🎯 Localizador Ultra-Preciso (Modo Varredura)")

u_excel = st.file_uploader("1. Planilha Cielo", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    # --- ETAPA 1: MAPEAMENTO GEOMÉTRICO DO PDF ---
    # Em vez de ler linhas, vamos ler a posição de cada palavra
    mapa_financeiro = []
    with pdfplumber.open(u_pdf) as pdf:
        for page in pdf.pages:
            palavras = page.extract_words()
            for p in palavras:
                texto = p['text'].upper()
                # Verifica se é um código PC/PD
                if re.search(r'(PC|PD)[\w\d-]+', texto):
                    mapa_financeiro.append({'tipo': 'ID', 'val': texto, 'y': p['top'], 'x': p['x0']})
                # Verifica se é um valor financeiro
                else:
                    try:
                        v_limpo = float(texto.replace('.', '').replace(',', '.'))
                        if v_limpo > 1.0:
                            mapa_financeiro.append({'tipo': 'VALOR', 'val': v_f, 'y': p['top'], 'x': p['x0']})
                    except: continue

    if st.button("🚀 Forçar Localização Total (20 de 20)"):
        u_excel.seek(0)
        wb = openpyxl.load_workbook(u_excel)
        ws = wb.active
        df_cielo = pd.read_excel(u_excel, header=14)
        
        sucessos = 0
        ids_usados = set()

        for i, row in df_cielo.iterrows():
            linha_excel = i + 16
            valor_alvo = float(row.iloc[4]) # Valor Bruto da Cielo
            
            # BUSCA INTELIGENTE: Acha o valor e pega o ID mais próximo no eixo Y (mesma linha)
            match_id = None
            menor_distancia = 999
            
            # 1. Acha todos os locais no PDF que tem o valor da planilha
            locais_valor = [m for m in mapa_financeiro if m['tipo'] == 'VALOR' and abs(m['val'] - valor_alvo) <= 0.02]
            
            for loc in locais_valor:
                # 2. Para cada valor achado, busca o ID (PC/PD) mais próximo dele na página
                for m in mapa_financeiro:
                    if m['tipo'] == 'ID' and m['val'] not in ids_usados:
                        distancia = abs(m['y'] - loc['y']) # Distância vertical (mesma linha)
                        if distancia < 10 and distancia < menor_distancia:
                            match_id = m['val']
                            menor_distancia = distancia

            if match_id:
                ws.cell(row=linha_excel, column=8).value = match_id
                ids_usados.add(match_id)
                sucessos += 1
            else:
                ws.cell(row=linha_excel, column=8).value = "REVISAR"

        st.success(f"✅ Finalizado! {sucessos} de 20 itens vinculados.")
        
        buffer = BytesIO()
        wb.save(buffer)
        st.download_button("📥 Baixar Planilha 20/20", buffer.getvalue(), "Conferencia_Cielo_Fechada.xlsx")
