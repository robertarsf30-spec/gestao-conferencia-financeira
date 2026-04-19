import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl

st.set_page_config(page_title="Conferência Cartão Cielo", layout="wide")

if 'autenticado' not in st.session_state or not st.session_state.autenticado:
    st.error("🔒 Por favor, faça login na página inicial.")
    st.stop()

st.title("💳 Conferência Cartão (Cielo)")

u_excel = st.file_uploader("1. Planilha Cielo Original (Excel)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    try:
        # 1. EXTRAÇÃO SIMPLIFICADA DO PDF
        lista_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        p = linha.split()
                        # Verifica se começa com PC ou PD (ex: PC22650-1)
                        if len(p) >= 6 and (p[0].startswith('PC') or p[0].startswith('PD')):
                            try:
                                # Data na 2ª posição | Valor na antepenúltima
                                d_v = pd.to_datetime(p[1], dayfirst=True).date()
                                v_v = float(p[-3].replace('.', '').replace(',', '.'))
                                lista_pdf.append({'id': p[0], 'data': d_v, 'valor': v_v})
                            except:
                                continue
        
        df_pdf = pd.DataFrame(lista_pdf)

        if st.button("🚀 Iniciar Conferência"):
            # 2. CARREGAR EXCEL ORIGINAL PARA MANTER O MODELO
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            
            # Lê dados (pula 14 linhas de cabeçalho)
            df_cielo = pd.read_excel(u_excel, header=14)
            
            conferidos = 0
            # Percorre o Excel (dados começam na linha 16)
            for i, row in df_cielo.iterrows():
                lin = i + 16
                try:
                    dt_c = pd.to_datetime(row.iloc[1], dayfirst=True).date()
                    vl_c = float(row.iloc[4])
                    
                    final = "NÃO ENCONTRADO"
                    
                    if not df_pdf.empty:
                        for _, p_item in df_pdf.iterrows():
                            # Bate Data e Valor (margem 0.05)
                            if p_item['data'] == dt_c and abs(p_item['valor'] - vl_c) <= 0.05:
                                final = f"CONFERIDO ({p_item['id']})"
                                conferidos += 1
                                break
                    
                    # Escreve na Coluna H (8)
                    ws.cell(row=lin, column=8).value = final
                except:
                    continue

            st.success(f"✅ Sucesso! {conferidos} itens marcados.")
            
            # 3. DOWNLOAD
            saida = BytesIO()
            wb.save(saida)
            
            st.download_button(
                label="📥 Baixar Planilha Original Preenchida",
                data=saida.getvalue(),
                file_name="Cielo_Conferida.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro: {e}")
