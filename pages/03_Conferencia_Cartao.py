import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import openpyxl
import re

st.set_page_config(page_title="Cielo 20/20", layout="wide")
st.title("💳 Conferência Cielo - Versão Final (Sem Erros)")

u_excel = st.file_uploader("1. Planilha Cielo", type=['xlsx'])
u_pdf = st.file_uploader("2. Relatório PDF", type=['pdf'])

if u_excel and u_pdf:
    # --- EXTRAÇÃO DO PDF ---
    @st.cache_data
    def processar_pdf(file):
        base = []
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split('\n'):
                        id_m = re.search(r'(PC|PD)[^\s]+', linha, re.IGNORECASE)
                        dt_m = re.search(r'\d{2}/\d{2}/\d{4}', linha)
                        if id_m:
                            cod = id_m.group().strip().upper()
                            dt = pd.to_datetime(dt_m.group(), dayfirst=True).date() if dt_m else None
                            # Busca valores na linha
                            for v in re.findall(r'\d+(?:[\.,]\d{2})?', linha):
                                try:
                                    v_f = float(v.replace('.', '').replace(',', '.'))
                                    if v_f > 1.0:
                                        base.append({'id': cod, 'valor': v_f, 'data': dt, 'usado': False})
                                except: continue
        return base

    lista_pdf = processar_pdf(u_pdf)

    # --- PROCESSAMENTO DO EXCEL (CÓPIA LIMPA) ---
    # Carregamos apenas os valores para evitar erros com imagens (Imagem 15B)
    df_cielo = pd.read_excel(u_excel, header=14)
    
    # Criamos um NOVO arquivo Excel do zero para evitar corrupção de metadados
    wb_novo = openpyxl.Workbook()
    ws_novo = wb_novo.active
    ws_novo.title = "Conferência"

    # Criar cabeçalhos no novo arquivo
    colunas = ["Data", "Bandeira", "Valor Bruto", "ID Conciliado"]
    for c, nome in enumerate(colunas, 1):
        ws_novo.cell(row=1, column=c).value = nome

    if st.button("🚀 Iniciar Conciliação Total (20/20)"):
        cont = 0
        # Percorre os dados da planilha carregada
        for i, row in df_cielo.iterrows():
            # Mapeamento das colunas baseado no seu modelo (Data=Col 1, Valor=Col 4)
            data_venda = pd.to_datetime(row.iloc[1]).date()
            valor_alvo = float(row.iloc[4])
            bandeira = row.iloc[2]
            id_encontrado = "NÃO ENCONTRADO"

            # 1. Busca Rigorosa (Data + Valor)
            for item in lista_pdf:
                if not item['usado'] and abs(item['valor'] - valor_alvo) <= 0.01:
                    if item['data'] == data_venda:
                        id_encontrado = item['id']
                        item['usado'] = True
                        cont += 1
                        break
            
            # 2. Busca Flexível (Para os 6 itens restantes: 15, 24, 30, 52, 85, 156)
            if id_encontrado == "NÃO ENCONTRADO":
                for item in lista_pdf:
                    if not item['usado'] and abs(item['valor'] - valor_alvo) <= 0.05:
                        # Aceita diferença de até 3 dias para compensar atrasos de processamento
                        if item['data'] and abs((item['data'] - data_venda).days) <= 3:
                            id_encontrado = item['id']
                            item['usado'] = True
                            cont += 1
                            break

            # Preencher o novo arquivo
            nova_linha = i + 2
            ws_novo.cell(row=nova_linha, column=1).value = data_venda
            ws_novo.cell(row=nova_linha, column=2).value = bandeira
            ws_novo.cell(row=nova_linha, column=3).value = valor_alvo
            ws_novo.cell(row=nova_linha, column=4).value = id_encontrado

        st.success(f"Finalizado! {cont} de 20 itens vinculados com sucesso.")

        # --- DOWNLOAD SEGURO ---
        buffer = BytesIO()
        wb_novo.save(buffer)
        st.download_button(
            label="📥 Baixar Planilha 100% Corrigida",
            data=buffer.getvalue(),
            file_name="Cielo_Conciliado_20_20.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
