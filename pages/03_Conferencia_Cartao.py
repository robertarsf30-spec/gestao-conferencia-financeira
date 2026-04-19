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

u_excel = st.file_uploader("1. Planilha Cielo (Original)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    try:
        # 1. EXTRAÇÃO DO PDF (Cria uma lista de todos os títulos disponíveis)
        base_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        p = linha.split()
                        # Verifica se a linha tem o padrão de título PC/PD
                        if len(p) >= 6 and (p[0].startswith('PC') or p[0].startswith('PD')):
                            try:
                                cod = p[0]
                                data = pd.to_datetime(p[1], dayfirst=True).date()
                                # Localiza o valor bruto na linha (tentativa em múltiplas posições)
                                v_limpo = p[-3].replace('.', '').replace(',', '.')
                                valor = float(v_limpo)
                                
                                base_pdf.append({
                                    'id': cod,
                                    'data': data,
                                    'valor': valor,
                                    'usado': False # Trava para não repetir
                                })
                            except:
                                continue

        if st.button("🚀 Iniciar Conferência Inteligente"):
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            df_cielo = pd.read_excel(u_excel, header=14)
            
            sucessos = 0
            
            # 2. PROCESSAMENTO COM FILTRO TRIPLO E TRAVA DE REPETIÇÃO
            for i, row in df_cielo.iterrows():
                lin_ex = i + 16
                try:
                    data_ex = pd.to_datetime(row.iloc[1], dayfirst=True).date()
                    val_ex = float(row.iloc[4])
                    
                    achou_nesta_linha = False
                    # Busca na base do PDF
                    for titulo in base_pdf:
                        # CRITÉRIO: Valor igual + Data igual + Ainda não usado
                        if not titulo['usado']:
                            if abs(titulo['valor'] - val_ex) <= 0.05 and titulo['data'] == data_ex:
                                # Preenche e trava o título
                                ws.cell(row=lin_ex, column=8).value = f"CONFERIDO ({titulo['id']})"
                                titulo['usado'] = True 
                                achou_nesta_linha = True
                                sucessos += 1
                                break
                    
                    if not achou_nesta_linha:
                        ws.cell(row=lin_ex, column=8).value = "NÃO ENCONTRADO"
                except:
                    continue

            st.success(f"✅ Finalizado! {sucessos} títulos vinculados sem repetições.")
            
            # 3. DOWNLOAD
            buffer = BytesIO()
            wb.save(buffer)
            st.download_button(
                label="📥 Baixar Planilha Original",
                data=buffer.getvalue(),
                file_name="Cielo_Conferida_Sem_Repeticao.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro técnico: {e}")
