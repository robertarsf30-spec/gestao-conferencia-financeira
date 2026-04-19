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
        # 1. PROCESSAR PDF (Extração de Títulos PC/PD e Valores)
        dados_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        partes = linha.split()
                        # Verifica se a linha começa com PC ou PD
                        if len(partes) >= 6 and re.match(r'^(PC|PD)', partes[0]):
                            try:
                                # ID do Título (Ex: PC22650-1)
                                id_titulo = partes[0]
                                # Data do Título (2ª parte)
                                dt_titulo = pd.to_datetime(partes[1], dayfirst=True).date()
                                # Valor (Penúltima posição antes de textos como 'cielo')
                                v_texto = partes[-3].replace('.', '').replace(',', '.')
                                v_num = float(v_texto)
                                dados_pdf.append({'id': id_titulo, 'data': dt_titulo, 'valor': v_num})
                            except:
                                continue
        
        df_pdf = pd.DataFrame(dados_pdf)

        if st.button("🚀 Iniciar Conferência"):
            # 2. CARREGAR EXCEL USANDO OPENPYXL PARA MANTER O MODELO
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            
            # DataFrame auxiliar para leitura (pula as 14 linhas de cabeçalho)
            df_cielo = pd.read_excel(u_excel, header=14)
            
            conferidos = 0
            # Os dados reais começam na linha 16 do Excel (15 do header + 1)
            for i, row in df_cielo.iterrows():
                idx_excel = i + 16
                try:
                    # B = Data (pos 1) | E = Valor Bruto (pos 4)
                    data_c = pd.to_datetime(row.iloc[1], dayfirst=True).date()
                    valor_c = float(row.iloc[4])
                    
                    status = "NÃO ENCONTRADO"
                    
                    # Comparação direta
                    if not df_pdf.empty:
                        for _, p in df_pdf.iterrows():
                            # Se a data bater e o valor tiver margem de 0.05
                            if p['data'] == data_c and abs(p['valor'] - valor_c) <= 0.05:
                                status = f"CONFERIDO ({p['id']})"
                                conferidos += 1
                                break
                    
                    # Escreve apenas na coluna H (8)
                    ws.cell(row=idx_excel, column=8).value = status
                except:
                    continue

            # 3. FINALIZAÇÃO E DOWNLOAD
            st.success(f"✅ Sucesso! {conferidos} títulos conferidos e marcados.")
            
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
