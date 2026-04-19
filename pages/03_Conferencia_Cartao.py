import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
from datetime import timedelta

st.set_page_config(page_title="Conferência Cartão Cielo", layout="wide")

if 'autenticado' not in st.session_state or not st.session_state.autenticado:
    st.error("🔒 Por favor, faça login na página inicial.")
    st.stop()

st.title("💳 Conferência Cartão (Cielo) - Precisão Máxima")

u_excel = st.file_uploader("1. Planilha Cielo (Original)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    try:
        # 1. MAPEAMENTO DO PDF (Lançamentos PC/PD)
        base_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        p = linha.split()
                        # Identifica o título (PC ou PD)
                        if len(p) >= 6 and (p[0].startswith('PC') or p[0].startswith('PD')):
                            try:
                                cod = p[0]
                                data_lcto = pd.to_datetime(p[1], dayfirst=True).date()
                                
                                # Coleta todos os números da linha para achar o valor bruto
                                valores_linha = []
                                for item in p[2:]:
                                    num_limpo = item.replace('.', '').replace(',', '.')
                                    try:
                                        valores_linha.append(float(num_limpo))
                                    except:
                                        continue
                                
                                base_pdf.append({
                                    'id': cod,
                                    'data': data_lcto,
                                    'valores': valores_linha,
                                    'usado': False
                                })
                            except:
                                continue

        if st.button("🚀 Iniciar Conferência Inteligente"):
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            df_cielo = pd.read_excel(u_excel, header=14)
            
            sucessos = 0
            for i, row in df_cielo.iterrows():
                lin_ex = i + 16 # Ajuste para começar na linha 16 do Excel
                try:
                    data_venda_ex = pd.to_datetime(row.iloc[1], dayfirst=True).date()
                    val_ex = float(row.iloc[4])
                    
                    achou = False
                    # Busca na base extraída do PDF
                    for t in base_pdf:
                        if not t['usado']:
                            # LÓGICA: Data do PDF entre (Data Excel) e (Data Excel + 2 dias)
                            data_limite = data_venda_ex + timedelta(days=2)
                            
                            if data_venda_ex <= t['data'] <= data_limite:
                                # Verifica se o valor bate com algum número daquela linha do PDF
                                for v_pdf in t['valores']:
                                    if abs(v_pdf - val_ex) <= 0.05:
                                        ws.cell(row=lin_ex, column=8).value = f"CONFERIDO ({t['id']})"
                                        t['usado'] = True # Marca como usado para não repetir
                                        achou = True
                                        sucessos += 1
                                        break
                        if achou: break
                    
                    if not achou:
                        ws.cell(row=lin_ex, column=8).value = "NÃO ENCONTRADO"
                except:
                    continue

            st.success(f"🎯 Finalizado! {sucessos} itens vinculados com sucesso.")
            
            buffer = BytesIO()
            wb.save(buffer)
            st.download_button(
                label="📥 Baixar Planilha Original Preenchida",
                data=buffer.getvalue(),
                file_name="Conferencia_Cielo_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
