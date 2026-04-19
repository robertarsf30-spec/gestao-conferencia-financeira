import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO

st.set_page_config(page_title="Conferência Cartão Cielo", layout="wide")

if 'autenticado' not in st.session_state or not st.session_state.autenticado:
    st.error("🔒 Por favor, faça login na página inicial.")
    st.stop()

st.title("💳 Conferência Cartão (Cielo)")

u_excel = st.file_uploader("1. Planilha Cielo (Excel)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    try:
        # 1. PROCESSAR EXCEL CIELO (Pula as 14 linhas de cabeçalho)
        df_cielo = pd.read_excel(u_excel, header=14) 
        
        # Identificação por posição conforme seu exemplo
        # Col 1: Data Venda | Col 2: Forma Pagamento | Col 4: Valor Bruto
        df_cielo['Data_Venda'] = pd.to_datetime(df_cielo.iloc[:, 1], dayfirst=True, errors='coerce').dt.date
        df_cielo['Tipo_Cielo'] = df_cielo.iloc[:, 2].astype(str)
        df_cielo['Valor_Cielo'] = pd.to_numeric(df_cielo.iloc[:, 4], errors='coerce')

        # 2. PROCESSAR PDF DO SISTEMA (Baseado no seu print e exemplo)
        dados_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    for line in text.split('\n'):
                        parts = line.split()
                        if len(parts) >= 8: # Linhas com dados financeiros
                            try:
                                # Pega o código (PC/PD) no início e o valor na penúltima posição
                                tipo_pdf = parts[0] 
                                data_pdf = pd.to_datetime(parts[1], dayfirst=True).date()
                                valor_pdf = float(parts[-3].replace('.', '').replace(',', '.'))
                                dados_pdf.append({'Tipo': tipo_pdf, 'Data': data_pdf, 'Valor': valor_pdf})
                            except:
                                continue
        
        df_sis = pd.DataFrame(dados_pdf)

        if st.button("🚀 Iniciar Conferência"):
            def conferir(row):
                if df_sis.empty: return "PDF NÃO LIDO"
                
                # Regra: Mesma Data, Tipo (PC/PD) e Valor (margem 0.02)
                # PC = Crédito | PD = Débito
                match = df_sis[
                    (df_sis['Tipo'].str.contains('PC|PD', case=False, na=False)) &
                    (df_sis['Data'] == row['Data_Venda']) &
                    (abs(df_sis['Valor'] - row['Valor_Cielo']) <= 0.02)
                ]
                
                if not match.empty:
                    return f"CONFERIDO ({match.iloc[0]['Tipo']})"
                return "NÃO ENCONTRADO NO SISTEMA"

            df_cielo['Descrição'] = df_cielo.apply(conferir, axis=1)
            
            # Limpeza das colunas extras antes de mostrar
            df_final = df_cielo.drop(columns=['Data_Venda', 'Tipo_Cielo', 'Valor_Cielo'])
            
            st.success("✅ Conferência concluída!")
            st.dataframe(df_final)

            # 4. DOWNLOAD
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Baixar Planilha Finalizada",
                data=output.getvalue(),
                file_name="conferencia_finalizada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"Erro ao processar: {e}")
