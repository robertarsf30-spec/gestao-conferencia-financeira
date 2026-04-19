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
        # 1. PROCESSAR EXCEL CIELO (Header na linha 15 do Excel = índice 14)
        df_cielo = pd.read_excel(u_excel, header=14) 
        
        # Mapeamento por posição (evita erro de nome de coluna)
        df_cielo['Data_Venda'] = pd.to_datetime(df_cielo.iloc[:, 1], dayfirst=True, errors='coerce').dt.date
        df_cielo['Valor_Cielo'] = pd.to_numeric(df_cielo.iloc[:, 4], errors='coerce')

        # 2. PROCESSAR PDF DO SISTEMA (Extrator de texto para colunas coladas)
        dados_pdf = []
        with pdfplumber.open(u_pdf) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    for line in text.split('\n'):
                        parts = line.split()
                        if len(parts) >= 8:
                            try:
                                # Tipo no início, Data na segunda parte, Valor na antepenúltima
                                tipo_pdf = str(parts[0])
                                data_pdf = pd.to_datetime(parts[1], dayfirst=True).date()
                                val_str = parts[-3].replace('.', '').replace(',', '.')
                                valor_pdf = float(val_str)
                                dados_pdf.append({'Tipo': tipo_pdf, 'Data': data_pdf, 'Valor': valor_pdf})
                            except:
                                continue
        
        df_sis = pd.DataFrame(dados_pdf)

        if st.button("🚀 Iniciar Conferência"):
            if df_sis.empty:
                st.warning("⚠️ O sistema não conseguiu ler os dados do PDF. Verifique o arquivo.")
            
            def conferir(row):
                if df_sis.empty: return "ERRO LEITURA PDF"
                
                # Regra: Mesma Data e Valor (margem 0.02)
                # Verifica se o tipo no PDF (PC/PD) bate com o que você espera
                mask = (df_sis['Data'] == row['Data_Venda']) & \
                       ((df_sis['Valor'] - row['Valor_Cielo']).abs() <= 0.02)
                
                match = df_sis[mask]
                
                if not match.empty:
                    return f"CONFERIDO ({match.iloc[0]['Tipo']})"
                return "NÃO ENCONTRADO"

            # Aplica a conferência
            df_cielo['Descrição'] = df_cielo.apply(conferir, axis=1)
            
            # Limpa colunas temporárias para o usuário
            df_final = df_cielo.drop(columns=['Data_Venda', 'Valor_Cielo'])
            
            st.success("✅ Processamento concluído!")
            st.dataframe(df_final)

            # 4. DOWNLOAD EXCEL
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Baixar Planilha Final",
                data=output.getvalue(),
                file_name="conferencia_cielo_ok.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"Erro ao processar arquivos: {e}")
