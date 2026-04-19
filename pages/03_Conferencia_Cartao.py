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
                                # Tipo (PC/PD), Data (2ª col) e Valor (antepenúltima)
                                p_tipo = str(partes[0])
                                p_data = pd.to_datetime(partes[1], dayfirst=True).date()
                                p_valor = float(partes[-3].replace('.', '').replace(',', '.'))
                                dados_pdf.append({'tipo': p_tipo, 'data': p_data, 'valor': p_valor})
                            except:
                                continue
        df_pdf = pd.DataFrame(dados_pdf)

        if st.button("🚀 Iniciar Conferência"):
            # 2. CARREGAR EXCEL ORIGINAL PARA EDIÇÃO
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            
            # Lendo dados para lógica (pula as 14 linhas de cabeçalho)
            df_cielo = pd.read_excel(u_excel, header=14)
            
            sucessos = 0
            # Percorre a planilha linha por linha (começando da linha 16 no Excel)
            for i, row in df_cielo.iterrows():
                linha_atual = i + 16
                
                # Extrai Data (Col 2) e Valor (Col 5) por posição
                try:
                    data_c = pd.to_datetime(row.iloc[1], dayfirst=True).date()
                    valor_c = float(row.iloc[4])
                except:
                    continue
                
                # Procura no PDF
                encontrado = False
                if not df_pdf.empty:
                    for _, pdf_row in df_pdf.iterrows():
                        # Critério: Mesma data e diferença de valor menor que 0.02
                        dif_valor = abs(pdf_row['valor'] - valor_c)
                        if pdf_row['data'] == data_c and dif_valor <= 0.02:
                            ws.cell(row=linha_atual, column=8).value = f"CONFERIDO ({pdf_row['tipo']})"
                            encontrado = True
                            sucessos += 1
                            break
                
                if not encontrado:
                    ws.cell(row=linha_atual, column=8).value = "NÃO ENCONTRADO"

            # 3. FINALIZAÇÃO
            st.success(f"✅ Conferência finalizada! {sucessos} itens encontrados.")
            
            # Gerar arquivo para download mantendo o modelo
            output = BytesIO()
            wb.save(output)
            
            st.download_button(
                label="📥 Baixar Planilha Original Preenchida",
                data=output.getvalue(),
                file_name="conferencia_cielo_original.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
