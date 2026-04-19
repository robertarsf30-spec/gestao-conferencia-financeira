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

u_excel = st.file_uploader("1. Planilha Cielo (Excel)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    try:
        # 1. LER DADOS DO PDF
        dados_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        partes = linha.split()
                        # Procura linhas que começam com PC ou PD (seu código de título)
                        if len(partes) >= 6 and re.match(r'^(PC|PD)', partes[0]):
                            try:
                                cod_titulo = partes[0]
                                data_titulo = pd.to_datetime(partes[1], dayfirst=True).date()
                                # O valor no seu PDF está na penúltima posição antes de 'cielo'
                                val_limpo = partes[-3].replace('.', '').replace(',', '.')
                                valor_titulo = float(val_limpo)
                                dados_pdf.append({'id': cod_titulo, 'dt': data_titulo, 'vl': valor_titulo})
                            except:
                                continue
        
        df_pdf = pd.DataFrame(dados_pdf)

        if st.button("🚀 Iniciar Conferência"):
            # 2. CARREGAR EXCEL PARA EDIÇÃO (Mantém modelo original)
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            
            # Dados para lógica (Pula as 14 linhas de cabeçalho)
            df_cielo = pd.read_excel(u_excel, header=14)
            
            conferidos = 0
            # Percorre o Excel (Dados começam na linha 16 do arquivo)
            for i, row in df_cielo.iterrows():
                num_linha = i + 16
                try:
                    # B = Data (pos 1) | E = Valor Bruto (pos 4)
                    dt_cielo = pd.to_datetime(row.iloc[1], dayfirst=True).date()
                    vl_cielo = float(row.iloc[4])
                    
                    encontrado = False
                    for _, p in df_pdf.iterrows():
                        # Bate data exata e valor com margem de 0.05
                        if p['dt'] == dt_cielo and abs(p['vl'] - vl_cielo) <= 0.05:
                            # Escreve na coluna H (8) o código que você destacou em amarelo
                            ws.cell(row=num_linha, column=8).value = f"CONFERIDO ({p['id']})"
                            encontrado = True
                            conferidos += 1
                            break
                    
                    if not encontrado:
                        ws.cell(row=num_linha, column=8).value = "NÃO ENCONTRADO"
                except:
                    continue

            # 3. SUCESSO E DOWNLOAD
            st.success(f"✅ Sucesso! {conferidos} títulos identificados no PDF.")
            
            saida = BytesIO()
            wb.save(saida)
            
            st.download_button(
                label="📥 Baixar Planilha Original Preenchida",
                data=saida.getvalue(),
                file_name="Conferencia_Cielo_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
