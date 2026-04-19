import streamlit as st

# Configuração da página - DEVE ser a primeira linha
st.set_page_config(page_title="Gestão Financeira", layout="wide")

# 1. Definir a senha (você pode alterar para a que preferir)
SENHA_ACESSO = "1006" 

# 2. Inicializar o estado de autenticação
if 'autenticado' not in st.session_state:
    st.session_state.autenticado = False

# --- CSS PARA BLOQUEAR A BARRA LATERAL ---
# Se não estiver autenticado, escondemos o menu de páginas
if not st.session_state.autenticado:
    st.markdown("""
        <style>
            [data-testid="stSidebarNav"] {display: none !important;}
            [data-testid="stSidebar"] {display: none !important;}
        </style>
    """, unsafe_allow_html=True)

# --- TELA DE LOGIN ---
if not st.session_state.autenticado:
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        st.write("# 🔐 Login do Sistema")
        st.info("Digite a senha para liberar os módulos.")
        
        with st.form("login_sistema"):
            senha = st.text_input("Senha de Acesso", type="password")
            btn_entrar = st.form_submit_button("Desbloquear Módulos")
            
            if btn_entrar:
                if senha == SENHA_ACESSO:
                    st.session_state.autenticado = True
                    st.success("Senha correta! Liberando acesso...")
                    st.rerun()
                else:
                    st.error("Senha incorreta. Tente novamente.")

# --- SISTEMA LIBERADO ---
else:
    with st.sidebar:
        st.success("✅ Acesso Liberado")
        if st.button("🚪 Sair e Bloquear"):
            st.session_state.autenticado = False
            st.rerun()

    st.title("🎯 Painel de Gestão Financeira")
    st.write("Agora você pode acessar as páginas na barra lateral à esquerda.")
    st.markdown("---")
    st.write("Selecione um dos módulos para trabalhar:")
    st.write("- Conferência Caixa/Banco")
    st.write("- Conferência Caixa")
    st.write("- Conferência Cartão")
    st.write("- Módulo Cobrança")
