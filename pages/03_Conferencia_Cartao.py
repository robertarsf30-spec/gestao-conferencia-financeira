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
                    linhas = texto.split('\n')
                    for linha in linhas:
                        partes = linha.split()
                        # Verifica se começa com PC ou PD (ex: PC22650-1)
                        if len(partes) >= 6:
                            inicio = partes[0]
                            if inicio.startswith('PC') or inicio.startswith('PD'):
                                try:
                                    cod_pdf = inicio
                                    # Data na segunda coluna do PDF
                                    dt_pdf = pd.to_datetime(partes[1], dayfirst=True).date()
                                    # Valor bruto é o 3º de trás para frente
                                    v_txt = partes[-3].replace('.', '').replace(',', '.')
                                    v_pdf = float(v_txt)
                                    lista_pdf.append({'id': cod_pdf, 'data': dt_pdf, 'valor': v_pdf})
                                except:
                                    continue
        
        df_pdf = pd.DataFrame(lista_pdf)

        if st.button("🚀 Iniciar Conferência"):
            # 2. CARREGAR MODELO ORIGINAL
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            
            # DataFrame para leitura (pula o cabeçalho de 14 linhas)
            df_cielo = pd.read_excel(u_excel, header=14)
            
            sucessos = 0
            # Os dados reais começam na linha 16 do arquivo Excel
            for i, row in df_cielo.iterrows():
                linha_excel = i + 16
                try:
                    # B (pos 1) = Data | E (pos 4) = Valor Bruto
                    dt_cielo = pd.to_datetime(row.iloc[1], dayfirst=True).date()
                    vl_cielo = float(row.iloc[4])
                    
                    match_final = "NÃO ENCONTRADO"
                    
                    if not df_pdf.empty:
                        for _, p in df_pdf.iterrows():
                            # Margem de 0.05 para pequenas diferenças de arredondamento
                            if p['data'] == dt_cielo and abs(p['valor'] - vl_cielo) <= 0.05:
                                match_final = f"CONFERIDO ({p['id']})"
                                sucessos += 1
                                break
                    
                    # Escreve na Coluna H (8) do arquivo original
                    ws.cell(row=linha_excel, column=8).value = match_final
                except:
                    continue

            st.success(f"✅ Concluído! {sucessos} itens conferidos.")
            
            # 3. EXPORTAÇÃO
            saida = BytesIO()
            wb.save(saida)
            
            st.download_button(
                label="📥 Baixar Planilha Original Preenchida",
                data=saida.getvalue(),
                file_name="Cielo_Conferida_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
