import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

st.set_page_config(page_title="Conferência Cartão Cielo", layout="wide")

if 'autenticado' not in st.session_state or not st.session_state.autenticado:
    st.error("🔒 Por favor, faça login na página inicial.")
    st.stop()

st.title("💳 Conferência Cartão (Cielo)")

u_excel = st.file_uploader("1. Planilha Cielo Original (Excel)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    try:
        # 1. EXTRAÇÃO DE DADOS DO PDF
        lista_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        partes = linha.split()
                        # Procura linhas que começam com o código do título (PC ou PD)
                        if len(partes) >= 6 and re.match(r'^(PC|PD)', partes[0]):
                            try:
                                cod_id = partes[0]
                                data_v = pd.to_datetime(partes[1], dayfirst=True).date()
                                # O valor bruto no seu PDF está na penúltima posição
                                v_texto = partes[-3].replace('.', '').replace(',', '.')
                                valor_v = float(v_texto)
                                lista_pdf.append({'id': cod_id, 'data': data_v, 'valor': valor_v})
                            except:
                                continue
        
        df_pdf = pd.DataFrame(lista_pdf)

        if st.button("🚀 Iniciar Conferência"):
            # 2. CARREGAR EXCEL ORIGINAL PARA MANTER O FORMATO
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            
            # DataFrame para leitura lógica (ignora o cabeçalho de 14 linhas)
            df_cielo = pd.read_excel(u_excel, header=14)
            
            sucessos = 0
            # Os dados reais começam na linha 16 do Excel
            for i, row in df_cielo.iterrows():
                linha_atual = i + 16
                try:
                    # B = Data da Venda (pos 1) | E = Valor Bruto (pos 4)
                    dt_excel = pd.to_datetime(row.iloc[1], dayfirst=True).date()
                    vl_excel = float(row.iloc[4])
                    
                    match_encontrado = "NÃO ENCONTRADO"
                    
                    if not df_pdf.empty:
                        for _, p in df_pdf.iterrows():
                            # Comparação: Mesma data e valor aproximado (margem 0.05)
                            if p['data'] == dt_excel and abs(p['valor'] - vl_excel) <= 0.05:
                                match_encontrado = f"CONFERIDO ({p['id']})"
                                sucessos += 1
                                break
                    
                    # Escreve apenas na coluna H (8) do seu modelo original
                    ws.cell(row=linha_atual, column=8).value = match_encontrado
                except:
                    continue

            st.success(f"✅ Concluído! {sucessos} itens conferidos.")
            
            # 3. GERAR DOWNLOAD PRESERVANDO O MODELO
            saida = BytesIO()
            wb.save(saida)
            
            st.download_button(
                label="📥 Baixar Planilha Original Preenchida",
                data=saida.getvalue(),
                file_name="Cielo_Conferida_Original.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
