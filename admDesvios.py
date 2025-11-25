import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

# ------------------------
# CONFIG & GLOBAL STYLES
# ------------------------
st.set_page_config(page_title="Painel ADM", page_icon="⚙️", layout="centered")

HIDE_SHELL = """
<style>
    /* Oculta navegação da sidebar e botão de collapse */
    [data-testid="stSidebarNav"], [data-testid="collapsedControl"] {display: none !important;}
    /* Remove rodapé padrão do Streamlit */
    footer {visibility: hidden;}
    /* Largura máxima do conteúdo para leitura confortável */
    .main > div {max-width: 1100px; margin: 0 auto;}
    /* Cabeçalho 'sticky' simples */
    .app-header {position: sticky; top: 0; z-index: 100; background: var(--background-color);
                 padding: 0.75rem 0; border-bottom: 1px solid rgba(49,51,63,0.2);}
    .muted {opacity: .7}
    /* Badges de status */
    .badge {display:inline-block; padding:.2rem .55rem; border-radius:999px; font-size:.75rem; font-weight:600}
    .b-aberto {background:#fff3cd; color:#7a5e00; border:1px solid #ffe69c}
    .b-analise {background:#e7f1ff; color:#0b5ed7; border:1px solid #cfe2ff}
    .b-resolvido {background:#e6f4ea; color:#146c43; border:1px solid #c9e7d5}
    .b-vencido {background:#fde7e9; color:#b02a37; border:1px solid #f5c2c7}
    /* Cartões KPI */
    .kpi {border:1px solid rgba(49,51,63,.2); border-radius:16px; padding:1rem 1.25rem}
    .kpi h3 {margin:.25rem 0 0 0; font-size:1.75rem}
    .kpi small {display:block; margin-top:.25rem}
    .actions {display:flex; gap:.5rem; flex-wrap:wrap}
</style>
"""
st.markdown(HIDE_SHELL, unsafe_allow_html=True)

st.write("Testando painel adm")