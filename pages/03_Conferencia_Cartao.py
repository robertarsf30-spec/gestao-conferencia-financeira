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
                        if len(p) >= 6 and (p[0].startswith('PC') or p[0].startswith('PD')):
                            try:
                                cod = p[0]
                                data_lcto = pd.to_datetime(p[1], dayfirst=True).date()
                                
                                # Captura todos os números da linha (Valor Bruto pode estar em qualquer posição)
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

        if st.button("🚀 Iniciar Conferência (Janela de 2 dias)"):
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            df_cielo = pd.read_excel(u_excel, header=14)
            
            sucessos = 0
            for i, row in df_cielo.iterrows():
                lin_ex = i + 16
                try:
                    data_venda_ex = pd.to_datetime(row.iloc[1], dayfirst=True).date()
                    val_ex = float(row.iloc[4])
                    
                    achou = False
                    for t in base_pdf:
                        if not t['usado']:
                            # LÓGICA ELÁSTICA: Data Lcto (PDF) pode ser até 2 dias após Data Venda (Excel)
                            data_limite = data_venda_ex + timedelta(days=2)
                            
                            if data_venda_ex <= t['data'] <= data_limite:
                                # Verifica se o valor bate com qualquer número da linha do PDF
                                for v_pdf in t['valores']:
                                    if abs(v_pdf - val_ex) <= 0.05:
                                        ws.cell(row=lin_ex, column=8).value = f"CONFERIDO ({t['id']})"
                                        t['usado'] = True
                                        achou = True
                                        sucessos += 1
                                        break
                        if achou: break
                    
                    if not achou:
                        ws.cell(row=lin_ex, column=8).value = "NÃO ENCONTRADO"
                except:
                    continue

            st.success(f"🎯 Conferência finalizada! {sucessos} itens vinculados.")
            
            buffer = BytesIO()
            wb.save(buffer)
            st.download_button(
                label="📥 Baixar Planilha Preenchida",
                data=buffer.getvalue(),
                file_name="Cielo_Conferencia_Final_99p.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro técnico: {e}")
