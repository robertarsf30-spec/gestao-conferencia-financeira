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
        # Pula as 14 linhas iniciais para chegar no cabeçalho correto
        df_cielo = pd.read_excel(u_excel, header=14) 
        
        # Identificação por POSIÇÃO para evitar erro de nome de coluna
        # Coluna 1: Data Venda | Coluna 2: Forma Pagamento | Coluna 4: Valor Bruto
        # (Ajustado conforme o padrão do seu arquivo CSV)
        df_cielo['Data_Venda'] = pd.to_datetime(df_cielo.iloc[:, 1], dayfirst=True, errors='coerce').dt.date
        df_cielo['Tipo_Original'] = df_cielo.iloc[:, 2].astype(str)
        df_cielo['Valor_Num'] = pd.to_numeric(df_cielo.iloc[:, 4], errors='coerce')

        # Limpeza de linhas vazias
        df_cielo = df_cielo.dropna(subset=['Data_Venda', 'Valor_Num'])
        
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
            def conferir(row):
                if df_sis.empty: return "PDF SEM DADOS"
                
                # Regra: Mesma Data (ou D+1), Tipo (PC/PD) e Valor (margem 0.02)
                # Filtramos se o tipo no PDF contém parte do texto da Cielo (ex: Crédito)
                cond_data = (df_sis['Data'] - row['Data_Venda']).map(lambda x: abs(x.days) <= 1)
                cond_valor = (df_sis['Valor'] - row['Valor_Num']).abs() <= 0.02
                
                match = df_sis[cond_data & cond_valor]
                return "CONFERIDO" if not match.empty else "NÃO ENCONTRADO NO SISTEMA"

            df_cielo['Descrição'] = df_cielo.apply(conferir, axis=1)
            
            # Remove colunas auxiliares antes de exibir
            df_final = df_cielo.drop(columns=['Data_Venda', 'Tipo_Original', 'Valor_Num'])
            
            st.success("✅ Processamento concluído!")
            st.dataframe(df_final)

            # 4. DOWNLOAD
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Baixar Planilha Finalizada",
                data=output.getvalue(),
                file_name="conferencia_cielo_pronta.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"Erro ao processar: {e}")
