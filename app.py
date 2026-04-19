import streamlit as st

# 1. Configuração da Página
st.set_page_config(page_title="Gestão Financeira", layout="wide")

# 2. Definição da Senha
# Você pode mudar o "1006" para a senha que preferir
SENHA_MESTRE = "1006" 

# 3. Inicialização do Estado de Acesso
if 'autenticado' not in st.session_state:
    st.session_state.autenticado = False

# --- TELA DE LOGIN ---
if not st.session_state.autenticado:
    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.5, 1])
    
    with col2:
        st.title("🔐 Login do Sistema")
        with st.form("login_form"):
            senha = st.text_input("Digite a Senha de Acesso:", type="password")
            entrar = st.form_submit_button("Entrar no Painel")
            
            if entrar:
                if senha == SENHA_MESTRE:
                    st.session_state.autenticado = True
                    st.success("Acesso Liberado!")
                    st.rerun()
                else:
                    st.error("Senha Incorreta. Tente novamente.")

# --- SE ESTIVER AUTENTICADO, MOSTRA O CONTEÚDO ---
else:
    st.sidebar.success("✅ Sistema Desbloqueado")
    if st.sidebar.button("🚪 Sair"):
        st.session_state.autenticado = False
        st.rerun()

    st.title("🎯 Painel de Controle - Módulos")
    st.write("Selecione um módulo na barra lateral para começar o trabalho.")
    
    # Aqui o Streamlit automaticamente mostrará as páginas da pasta /pages
    # (01_Conferencia_Caixa_Banco, 02_Conferencia_Caixa, etc.)
