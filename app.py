import streamlit as st

# Configuração da página - Deve ser a primeira linha
st.set_page_config(page_title="Gestão Financeira", layout="wide")

# 1. Definir a senha mestra
SENHA_ACESSO = "1006" 

# 2. Inicializar o estado de autenticação
if 'autenticado' not in st.session_state:
    st.session_state.autenticado = False

# --- TELA DE LOGIN ---
if not st.session_state.autenticado:
    # Esconde a barra lateral enquanto não logar
    st.markdown("""
        <style>
            [data-testid="stSidebarNav"] {display: none;}
        </style>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        st.title("🔐 Login do Sistema")
        st.write("Acesso restrito aos módulos financeiros.")
        
        with st.form("login"):
            senha = st.text_input("Senha:", type="password")
            btn_entrar = st.form_submit_button("Acessar Sistema")
            
            if btn_entrar:
                if senha == SENHA_ACESSO:
                    st.session_state.autenticado = True
                    st.success("Acesso liberado!")
                    st.rerun()
                else:
                    st.error("Senha incorreta!")

# --- SISTEMA DESBLOQUEADO ---
else:
    # Mostra a barra lateral e as abas dos módulos
    st.sidebar.success("✅ Conectado")
    if st.sidebar.button("🚪 Sair"):
        st.session_state.autenticado = False
        st.rerun()

    st.title("🎯 Painel de Controle")
    st.info("Utilize o menu lateral para acessar os módulos de conferência.")
    
    # Lista dos seus módulos disponíveis conforme o seu projeto:
    # - 01_Conferencia_Caixa_Banco
    # - 02_Conferencia_Caixa
    # - 03_Conferencia_Cartao
    # - 04_Modulo_Cobranca
