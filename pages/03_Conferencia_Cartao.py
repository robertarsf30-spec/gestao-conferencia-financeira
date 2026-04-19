import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO

st.set_page_config(page_title="Conferência Cartão Cielo", layout="wide")

if 'autenticado' not in st.session_state or not st.session_state.autenticado:
    st.error("🔒 Por favor, faça login na página inicial.")
    st.stop()

st.title("💳 Conferência Cartão (Cielo)")
st.divider()

u_excel = st.file_uploader("1. Planilha Cielo (Excel)", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório Sistema (PDF)", type=['pdf'])

if u_excel and u_pdf:
    try:
        # 1. PROCESSAR EXCEL CIELO
        # O cabeçalho real da Cielo começa na linha 14 (índice 13)
        df_cielo = pd.read_excel(u_excel, header=13) 
        
        # Mapeamento das colunas do seu arquivo Cielo
        col_tipo = 'Forma de pagamento'
        col_data = 'Data da venda'
        col_valor = 'Valor bruto'

        # Limpeza de dados nulos nas colunas essenciais
        df_cielo = df_cielo.dropna(subset=[col_data, col_valor])
        
        # Conversão para garantir cálculos corretos
        df_cielo['Valor_Num'] = pd.to_numeric(df_cielo[col_valor], errors='coerce')
        df_cielo['Data_Venda'] = pd.to_datetime(df_cielo[col_data], dayfirst=True, errors='coerce').dt.date
        
        if 'Descrição' not in df_cielo.columns:
            df_cielo['Descrição'] = ""

        # 2. PROCESSAR PDF DO SISTEMA
        dados_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    for row in table:
                        try:
                            # Tenta converter valor (última col) e data (terceira col)
                            v_str = str(row[-1]).replace('.', '').replace(',', '.')
                            valor_pdf = float(v_str)
                            data_pdf = pd.to_datetime(row[2], dayfirst=True).date()
                            dados_pdf.append({'Tipo': str(row[0]), 'Data': data_pdf, 'Valor': valor_pdf})
                        except:
                            continue
        
        df_sis = pd.DataFrame(dados_pdf)

        # 3. BOTÃO DE AÇÃO
        if st.button("🚀 Iniciar Conferência"):
            if df_sis.empty:
                st.warning("⚠️ Não foi possível extrair dados do PDF. Verifique o formato.")
            
            def conferir(row):
                # Regra: Mesma Data (ou D+1), Tipo (PC/PD) e Valor (margem 0.02)
                # PC = Crédito | PD = Débito
                condicao_tipo = df_sis['Tipo'].str.contains('PC|PD', case=False, na=False)
                condicao_data = (df_sis['Data'] - row['Data_Venda']).map(lambda x: abs(x.days) <= 1)
                condicao_valor = (df_sis['Valor'] - row['Valor_Num']).abs() <= 0.02
                
                match = df_sis[condicao_tipo & condicao_data & condicao_valor]
                return "CONFERIDO" if not match.empty else "NÃO ENCONTRADO"

            df_cielo['Descrição'] = df_cielo.apply(conferir, axis=1)
            
            # Limpeza final para exibição
            df_final = df_cielo.drop(columns=['Valor_Num', 'Data_Venda'])
            
            st.success("✅ Conferência concluída!")
            st.dataframe(df_final)

            # 4. DOWNLOAD
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Baixar Planilha de Conferência",
                data=output.getvalue(),
                file_name="conferencia_finalizada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"Erro ao processar: {e}")
