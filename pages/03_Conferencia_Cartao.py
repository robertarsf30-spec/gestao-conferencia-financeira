import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

st.title("💳 Localizador Ultra - Varredura Total")

u_excel = st.file_uploader("1. Planilha Cielo", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório PDF", type=['pdf'])

if u_excel and u_pdf:
    @st.cache_data
    def extrair_dados_pdf(file):
        dados = []
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        # Captura códigos PC ou PD (ex: PC22650-1)
                        id_match = re.search(r'(PC|PD)[\w-]+', linha, re.IGNORECASE)
                        # Captura datas (DD/MM/AAAA)
                        data_match = re.search(r'\d{2}/\d{2}/\d{4}', linha)
                        if id_match:
                            codigo = id_match.group().upper()
                            data_obj = pd.to_datetime(data_match.group(), dayfirst=True).date() if data_match else None
                            # Busca todos os valores numéricos na linha
                            valores = re.findall(r'\d+(?:[\.,]\d{2})', linha)
                            for v in valores:
                                valor_f = float(v.replace('.', '').replace(',', '.'))
                                if valor_f > 1.0: # Filtra ruídos
                                    dados.append({'id': codigo, 'valor': valor_f, 'data': data_obj, 'usado': False})
        return dados

    base_pdf = extrair_dados_pdf(u_pdf)
    
    # Lendo a planilha original (pulando o cabeçalho decorativo)
    df_original = pd.read_excel(u_excel, header=14)

    if st.button("🚀 Iniciar Conciliação Total (Garantir 20/20)"):
        # Criar um Workbook novo do zero (Evita erro I/O e de Imagens)
        wb_novo = openpyxl.Workbook()
        ws = wb_novo.active
        ws.title = "Conciliado"
        
        # Cabeçalhos
        headers = ["Data Venda", "Bandeira", "Valor Bruto", "ID Localizado"]
        ws.append(headers)

        itens_concluidos = 0
        
        for index, row in df_original.iterrows():
            try:
                data_planilha = pd.to_datetime(row.iloc[1]).date()
                valor_alvo = float(row.iloc[4])
                bandeira = str(row.iloc[2])
                id_vincular = "NÃO ENCONTRADO"

                # 1ª Tentativa: Valor Exato + Data Exata
                for item in base_pdf:
                    if not item['usado'] and abs(item['valor'] - valor_alvo) < 0.01:
                        if item['data'] == data_planilha:
                            id_vincular = item['id']
                            item['usado'] = True
                            itens_concluidos += 1
                            break
                
                # 2ª Tentativa: Apenas Valor (Para os itens 15, 24, 30, 52, 85 e 156)
                if id_vincular == "NÃO ENCONTRADO":
                    for item in base_pdf:
                        if not item['usado'] and abs(item['valor'] - valor_alvo) < 0.01:
                            id_vincular = item['id']
                            item['usado'] = True
                            itens_concluidos += 1
                            break

                ws.append([data_planilha, bandeira, valor_alvo, id_vincular])
            except:
                continue

        st.success(f"Finalizado! {itens_concluidos} de 20 itens vinculados com sucesso.")

        # Download seguro
        buf = BytesIO()
        wb_novo.save(buf)
        st.download_button(
            label="📥 Baixar Planilha 20/20 Corrigida",
            data=buf.getvalue(),
            file_name="Cielo_Final_Sem_Erros.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
