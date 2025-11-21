import streamlit as st

st.set_page_config(page_title="Painel ADM", page_icon="⚙️", layout="centered")

st.title("⚙️ Painel Administrativo de Desvios")
HIDE_SIDEBAR_NAVIGATION = """
<style>
    [data-testid="stSidebarNav"] {display: none;}
    [data-testid="collapsedControl"] {display: none;}
</style>
"""

st.markdown(HIDE_SIDEBAR_NAVIGATION, unsafe_allow_html=True)

st.write("Testando painel adm")