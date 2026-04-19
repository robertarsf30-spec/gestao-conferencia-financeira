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
                        # Identifica linhas que começam com PC ou PD (ex: PC22650-1)
                        if len(partes) >= 6 and re.match(r'^(PC|PD)', partes[0]):
                            try:
                                codigo = partes[0]
                                data_v = pd.to_datetime(partes[1], dayfirst=True).date()
                                # No seu PDF, o valor bruto é o 3º elemento de trás para frente
                                val_limpo = partes[-3].replace('.', '').replace(',', '.')
                                valor_v = float(val_limpo)
                                lista_pdf.append({'id': codigo, 'data': data_v, 'valor': valor_v})
                            except:
                                continue
        
        df_pdf = pd.DataFrame(lista_pdf)

        if st.button("🚀 Iniciar Conferência"):
            # 2. CARREGAR MODELO ORIGINAL PARA MANTER LOGOS E FORMATO
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            
            # DataFrame para leitura lógica (pula as 14 linhas de cabeçalho)
            df_cielo = pd.read_excel(u_excel, header=14)
            
            sucessos = 0
            # A leitura começa na linha 16 do Excel (15 do header + 1)
            for i, row in df_cielo.iterrows():
                num_linha = i + 16
                try:
                    # B = Data (pos 1) | E = Valor Bruto (pos 4)
                    dt_cielo = pd.to_datetime(row.iloc[1], dayfirst=True).date()
                    vl_cielo = float(row.iloc[4])
                    
                    match_final = "NÃO ENCONTRADO"
                    
                    if not df_pdf.empty:
                        for _, p in df_pdf.iterrows():
                            # Bate Data e Valor (com margem de 0.05 para centavos)
                            if p['data'] == dt_cielo and abs(p['valor'] - vl_cielo) <= 0.05:
                                match_final = f"CONFERIDO ({p['id']})"
                                sucessos += 1
                                break
                    
                    # Escreve apenas na coluna H (8)
                    ws.cell(row=num_linha, column=8).value = match_final
                except:
                    continue

            st.success(f"✅ Conferência realizada! {sucessos} itens encontrados.")
            
            # 3. DOWNLOAD DO ARQUIVO MODIFICADO
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
