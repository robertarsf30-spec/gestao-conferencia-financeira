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
        # 1. EXTRAÇÃO CIRÚRGICA DO PDF
        dados_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        partes = linha.split()
                        # Identifica linhas que começam com PC ou PD
                        if len(partes) >= 6 and re.match(r'^(PC|PD)', partes[0]):
                            try:
                                t_id = partes[0]
                                t_data = pd.to_datetime(partes[1], dayfirst=True).date()
                                # Pega o valor bruto (antepenúltimo antes de 'cielo')
                                v_limpo = partes[-3].replace('.', '').replace(',', '.')
                                t_valor = float(v_limpo)
                                dados_pdf.append({'id': t_id, 'data': t_data, 'valor': t_valor})
                            except:
                                continue
        
        df_pdf = pd.DataFrame(dados_pdf)

        if st.button("🚀 Iniciar Conferência"):
            # 2. CARREGAR MODELO ORIGINAL (Mantém logos e cabeçalho)
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            
            # DataFrame para processamento (pula as 14 linhas de topo)
            df_cielo = pd.read_excel(u_excel, header=14)
            
            sucessos = 0
            # Dados começam na linha 16 do Excel (15 do header + 1)
            for i, row in df_cielo.iterrows():
                num_linha = i + 16
                try:
                    # B = Data (pos 1) | E = Valor Bruto (pos 4)
                    dt_c = pd.to_datetime(row.iloc[1], dayfirst=True).date()
                    vl_c = float(row.iloc[4])
                    
                    resultado = "NÃO ENCONTRADO"
                    
                    if not df_pdf.empty:
                        for _, p in df_pdf.iterrows():
                            # Bate data e valor com margem de 0.05
                            diferenca = abs(p['valor'] - vl_c)
                            if p['data'] == dt_c and diferenca <= 0.05:
                                resultado = f"CONFERIDO ({p['id']})"
                                sucessos += 1
                                break
                    
                    # Escreve na Coluna H (8) preservando o resto da linha
                    ws.cell(row=num_linha, column=8).value = resultado
                except:
                    continue

            st.success(f"✅ Concluído! {sucessos} títulos conferidos com sucesso.")
            
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
