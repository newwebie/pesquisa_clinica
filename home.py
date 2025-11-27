"""Tela inicial com autenticação via Microsoft e navegação principal."""
from __future__ import annotations

from pathlib import Path
from typing import Any, Dict

import streamlit as st

from auth_microsoft import (
    AuthManager,
    MicrosoftAuth,
    create_login_page,
    create_user_header,
)

st.set_page_config(page_title="Portal Pesquisa Clínica", page_icon="🧪", layout="wide")

PAGE_STYLE = """
<style>
    /* Remove completamente a sidebar e seus controles */
    [data-testid="stSidebar"] {display: none !important;}
    [data-testid="stSidebarNav"] {display: none !important;}
    [data-testid="collapsedControl"] {display: none !important;}

    }
</style>
"""

st.markdown(PAGE_STYLE, unsafe_allow_html=True)

HIDE_SIDEBAR_NAVIGATION = """
<style>
    [data-testid="stSidebarNav"] {display: none;}
    [data-testid="collapsedControl"] {display: none;}
</style>
"""

st.markdown(HIDE_SIDEBAR_NAVIGATION, unsafe_allow_html=True)

def _load_config() -> Dict[str, Any]:
    """Carrega o config.yaml da raiz do projeto."""
    config_path = Path("config.yaml")
    if not config_path.exists():
        return {}

    return yaml_load(config_path)


def yaml_load(path: Path) -> Dict[str, Any]:
    """Wrapper separado para facilitar testes e linting."""
    import yaml
    from yaml.loader import SafeLoader

    with path.open("r", encoding="utf-8") as file:
        return yaml.load(file, Loader=SafeLoader) or {}


def _admin_email_whitelist(config: Dict[str, Any]) -> set[str]:
    """Retorna o conjunto de e-mails autorizados a acessar a área administrativa."""

    def _normalize(entries: Any) -> set[str]:
        if not isinstance(entries, (list, tuple, set)):
            return set()
        return {str(item).strip().lower() for item in entries if str(item).strip()}

    admin_section = config.get("admin") if isinstance(config.get("admin"), dict) else {}
    candidates = [
        admin_section.get("emails"),
        config.get("admin_emails"),
        (config.get("pre-authorized", {}) or {}).get("emails"),
    ]

    for candidate in candidates:
        normalized = _normalize(candidate)
        if normalized:
            return normalized

    return set()


def _switch_page(target: str) -> None:
    """Encapsula o redirecionamento para outras páginas Streamlit."""
    target_file = target if target.endswith(".py") else f"{target}.py"
    if hasattr(st, "switch_page"):
        st.switch_page(target_file)
    else:
        st.info(f"Abra **{target_file}** pelo seletor de páginas do Streamlit.")


auth = MicrosoftAuth()
config_data = _load_config()
admin_whitelist = _admin_email_whitelist(config_data)

logged_in = create_login_page(auth)
if not logged_in:
    st.stop()

# Garantir token válido durante a sessão
AuthManager.check_and_refresh_token(auth)
create_user_header()

user = AuthManager.get_current_user() or {}
user_email = (user.get("mail") or user.get("userPrincipalName") or "").lower()

is_admin = True if not admin_whitelist else user_email in admin_whitelist

st.title("Portal Pesquisa Clínica")
st.caption("Acesse as funcionalidades principais com sua conta Microsoft corporativa.")

with st.container():
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Cadastro")
        st.write("Registre novos desvios e mantenha os dados organizados.")
        if st.button("Ir para Cadastro", use_container_width=True):
            _switch_page("cadastro.py")

    with col2:
        st.subheader("Administração de Desvios")
        st.write("Área dedicada para análises e ajustes administrativos.")
        if st.button("Abrir Administração", use_container_width=True):
            if is_admin:
                _switch_page("admDesvios.py")
            else:
                st.error("Acesso restrito aos administradores cadastrados.")

if not is_admin:
    st.info(
        "Solicite ao administrador a inclusão do seu e-mail na lista de acesso se precisar da área administrativa."
    )
