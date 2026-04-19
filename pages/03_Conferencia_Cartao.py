import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
from datetime import timedelta
import re

st.set_page_config(page_title="Conferência Cartão Cielo", layout="wide")

# Mantive sua trava de autenticação
if 'autenticado' not in st.session_state or not st.session_state.autenticado:
    st.error("🔒 Por favor, faça login na página inicial.")
    st.stop()

st.title("💳 Conferência Cartão (Cielo) - Versão Final PC/PD")

u_excel = st.file_uploader("1. Planilha Cielo (Original)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    try:
        # 1. MAPEAMENTO COMPLETO DO PDF
        base_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        # AJUSTE 1: Regex aprimorada para aceitar PC/PD com símbolos / - * #
                        match_id = re.search(r'(PC|PD)[0-9\-/\\*#]+', linha, re.IGNORECASE)
                        
                        if match_id:
                            cod_id = match_id.group().strip()
                            # Localiza as datas (DD/MM/AAAA)
                            datas = re.findall(r'\d{2}/\d{2}/\d{4}', linha)
                            dt_lcto = pd.to_datetime(datas[-1], dayfirst=True).date() if datas else None
                            
                            # Extrai valores numéricos (Ex: 156,00 ou 1.250,50)
                            # AJUSTE 2: Captura valores com ponto de milhar opcional e vírgula decimal
                            valores_texto = re.findall(r'\d+(?:\.\d{3})*(?:,\d{2})', linha)
                            valores_float = []
                            for v in valores_texto:
                                v_limpo = v.replace('.', '').replace(',', '.')
                                valores_float.append(float(v_limpo))
                            
                            # Só adiciona se tiver o código. A data agora é opcional para não travar os PC
                            if cod_id:
                                base_pdf.append({
                                    'id': cod_id,
                                    'data': dt_lcto,
                                    'valores': valores_float,
                                    'usado': False
                                })

        if st.button("🚀 Iniciar Conferência (Foco em Valor e Símbolos)"):
            u_excel.seek(0)
            wb = openpyxl.load_workbook(u_excel)
            ws = wb.active
            
            # Ajuste de cabeçalho conforme seu padrão anterior (header=14)
            df_cielo = pd.read_excel(u_excel, header=14)
            
            sucessos = 0
            for i, row in df_cielo.iterrows():
                lin_ex = i + 16
                try:
                    dt_venda_ex = pd.to_datetime(row.iloc[1], dayfirst=True).date()
                    val_ex = float(row.iloc[4]) # Valor Bruto no Excel
                    
                    achou = False
                    for t in base_pdf:
                        # AJUSTE 3: Margem de 0.02 e tolerância maior na data para os PC
                        if not t['usado'] and any(abs(v_pdf - val_ex) <= 0.02 for v_pdf in t['valores']):
                            
                            # Se o PDF tiver data, valida o intervalo. Se não tiver, aceita pelo valor.
                            if t['data']:
                                dt_limite = dt_venda_ex + timedelta(days=5) # Aumentei para 5 dias para os PC
                                if dt_venda_ex <= t['data'] <= dt_limite:
                                    achou = True
                            else:
                                # Se não tem data no PDF mas o valor é exato, assume que é o título
                                achou = True

                            if achou:
                                ws.cell(row=lin_ex, column=8).value = t['id']
                                t['usado'] = True
                                sucessos += 1
                                break
                                
                    if not achou:
                        ws.cell(row=lin_ex, column=8).value = "NÃO ENCONTRADO"
                except:
                    continue

            st.success(f"🎯 Finalizado! {sucessos} títulos conferidos (PD e PC com símbolos).")
            
            buffer = BytesIO()
            wb.save(buffer)
            st.download_button(
                label="📥 Baixar Planilha Original Preenchida",
                data=buffer.getvalue(),
                file_name="Cielo_Conferencia_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
