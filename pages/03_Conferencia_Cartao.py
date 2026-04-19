import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl

# Configuração da página
st.set_page_config(page_title="Conferência Cartão Cielo", layout="wide")

# Verificação de segurança
if 'autenticado' not in st.session_state or not st.session_state.autenticado:
    st.error("🔒 Por favor, faça login na página inicial.")
    st.stop()

st.title("💳 Conferência Cartão (Cielo)")
st.markdown("---")

u_excel = st.file_uploader("1. Envie a Planilha Cielo (Original)", type=['xlsx'])
u_pdf = st.file_uploader("2. Envie o Relatório do Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    try:
        # --- PASSO 1: LER O PDF COM BUSCA FLEXÍVEL ---
        dados_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if not texto:
                    continue
                
                for linha in texto.split('\n'):
                    partes = linha.split()
                    # Identifica se a linha começa com PC ou PD (seus títulos)
                    if len(partes) > 5 and (partes[0].startswith('PC') or partes[0].startswith('PD')):
                        try:
                            # O código é sempre o primeiro item
                            codigo_id = partes[0]
                            # A data é sempre o segundo item
                            data_venda = pd.to_datetime(partes[1], dayfirst=True).date()
                            
                            # BUSCA O VALOR: Varre os itens da linha para achar o valor bruto
                            # Tentamos converter cada parte para número até achar o valor que bate
                            valores_da_linha = []
                            for p in partes[2:]:
                                num_limpo = p.replace('.', '').replace(',', '.')
                                try:
                                    valores_da_linha.append(float(num_limpo))
                                except:
                                    continue
                            
                            # Adicionamos à nossa base de busca
                            dados_pdf.append({
                                'id': codigo_id,
                                'data': data_venda,
                                'valores': valores_da_linha
                            })
                        except:
                            continue

        if st.button("🚀 Iniciar Processamento"):
            # --- PASSO 2: EDITAR O EXCEL ORIGINAL ---
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            
            # Lê para lógica (pula as 14 linhas de cabeçalho da Cielo)
            df_cielo = pd.read_excel(u_excel, header=14)
            
            sucessos = 0
            # Varre o Excel a partir da linha 16
            for i, row in df_cielo.iterrows():
                num_linha_excel = i + 16
                try:
                    # Coluna B (Data) e E (Valor Bruto)
                    dt_excel = pd.to_datetime(row.iloc[1], dayfirst=True).date()
                    vl_excel = float(row.iloc[4])
                    
                    status = "NÃO ENCONTRADO"
                    
                    # Busca no banco de dados que criamos do PDF
                    for item in dados_pdf:
                        if item['data'] == dt_excel:
                            # Verifica se o valor do Excel está presente nos números daquela linha do PDF
                            # Usamos margem de 0.05 para evitar erro de centavos
                            for v_pdf in item['valores']:
                                if abs(v_pdf - vl_excel) <= 0.05:
                                    status = f"CONFERIDO ({item['id']})"
                                    sucessos += 1
                                    break
                            if "CONFERIDO" in status: break
                    
                    # Escreve na coluna H (8)
                    ws.cell(row=num_linha_excel, column=8).value = status
                except:
                    continue

            # --- PASSO 3: FINALIZAR E DOWNLOAD ---
            st.success(f"✅ Conferência Concluída! {sucessos} itens identificados.")
            
            buffer = BytesIO()
            wb.save(buffer)
            
            st.download_button(
                label="📥 Baixar Planilha Original Preenchida",
                data=buffer.getvalue(),
                file_name="Cielo_Conferida_Original.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Ocorreu um erro inesperado: {e}")
