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

u_excel = st.file_uploader("1. Planilha Cielo (Excel)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    try:
        # 1. EXTRAÇÃO DE DADOS DO PDF
        dados_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        partes = linha.split()
                        if len(partes) >= 8:
                            try:
                                p_tipo = str(partes[0])
                                p_data = pd.to_datetime(partes[1], dayfirst=True).date()
                                # Valor bruto na penúltima posição do PDF
                                v_limpo = partes[-3].replace('.', '').replace(',', '.')
                                p_valor = float(v_limpo)
                                dados_pdf.append({'tipo': p_tipo, 'data': p_data, 'valor': p_valor})
                            except:
                                continue
        df_pdf = pd.DataFrame(dados_pdf)

        if st.button("🚀 Iniciar Conferência"):
            # 2. CARREGAR EXCEL ORIGINAL PARA EDITAR APENAS A COLUNA H
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            
            # Lendo dados para lógica (pula as 14 linhas de cabeçalho)
            df_cielo = pd.read_excel(u_excel, header=14)
            
            sucessos = 0
            # Percorre a planilha (começando da linha 16 no Excel)
            for i, row in df_cielo.iterrows():
                linha_excel = i + 16
                
                try:
                    # Posições: Data Venda (Col 2) | Valor Bruto (Col 5)
                    data_c = pd.to_datetime(row.iloc[1], dayfirst=True).date()
                    valor_c = float(row.iloc[4])
                    
                    encontrado = False
                    if not df_pdf.empty:
                        for _, pdf_row in df_pdf.iterrows():
                            # Regra: Mesma data e diferença de valor <= 0.02
                            if pdf_row['data'] == data_c and abs(pdf_row['valor'] - valor_c) <= 0.02:
                                ws.cell(row=linha_excel, column=8).value = f"CONFERIDO ({pdf_row['tipo']})"
                                encontrado = True
                                sucessos += 1
                                break
                    
                    if not encontrado:
                        ws.cell(row=linha_excel, column=8).value = "NÃO ENCONTRADO"
                except:
                    continue

            # 3. DOWNLOAD
            st.success(f"✅ Conferência concluída! {sucessos} itens encontrados.")
            
            output = BytesIO()
            wb.save(output)
            
            st.download_button(
                label="📥 Baixar Planilha Original Preenchida",
                data=output.getvalue(),
                file_name="Cielo_Conferida_Original.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
