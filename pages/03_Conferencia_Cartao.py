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
        # 1. EXTRAÇÃO DE DADOS DO PDF
        dados_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        partes = linha.split()
                        # Procuramos o padrão PC ou PD no início da linha
                        if len(partes) >= 6 and re.match(r'^(PC|PD)', partes[0]):
                            try:
                                t_codigo = partes[0] # Ex: PC22650-1
                                t_data = pd.to_datetime(partes[1], dayfirst=True).date()
                                # O valor bruto no seu PDF é o penúltimo ou antepenúltimo
                                # Vamos converter o valor que aparece antes de "cielo" ou similar
                                val_str = partes[-3].replace('.', '').replace(',', '.')
                                t_valor = float(val_str)
                                
                                dados_pdf.append({'codigo': t_codigo, 'data': t_data, 'valor': t_valor})
                            except:
                                continue
        
        df_pdf = pd.DataFrame(dados_pdf)

        if st.button("🚀 Iniciar Conferência"):
            # 2. CARREGAR EXCEL ORIGINAL
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            
            # Lê dados para conferência (pula as 14 linhas de topo)
            df_cielo = pd.read_excel(u_excel, header=14)
            
            conferidos = 0
            # Percorre o Excel (dados começam na linha 16)
            for i, row in df_cielo.iterrows():
                linha_excel = i + 16
                
                try:
                    # Coluna 2 (B) é Data | Coluna 5 (E) é Valor Bruto
                    data_c = pd.to_datetime(row.iloc[1], dayfirst=True).date()
                    valor_c = float(row.iloc[4])
                    
                    achou = False
                    if not df_pdf.empty:
                        # Filtra PDF por data e valor (margem 0.05 para garantir)
                        for _, p in df_pdf.iterrows():
                            if p['data'] == data_c and abs(p['valor'] - valor_c) <= 0.05:
                                ws.cell(row=linha_excel, column=8).value = f"CONFERIDO ({p['codigo']})"
                                achou = True
                                conferidos += 1
                                break
                    
                    if not achou:
                        ws.cell(row=linha_excel, column=8).value = "NÃO ENCONTRADO"
                except:
                    continue

            # 3. RESULTADO E DOWNLOAD
            st.success(f"✅ Finalizado! {conferidos} títulos encontrados e marcados.")
            
            output = BytesIO()
            wb.save(output)
            
            st.download_button(
                label="📥 Baixar Planilha Original Preenchida",
                data=output.getvalue(),
                file_name="Cielo_Conferida_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
