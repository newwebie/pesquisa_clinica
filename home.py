"""Tela inicial do app com autenticação usando config.yaml."""
from __future__ import annotations

from pathlib import Path
from typing import Dict, Any

import streamlit as st
import streamlit_authenticator as stauth
import yaml
from yaml.loader import SafeLoader

st.set_page_config(page_title="Portal Pesquisa Clínica", page_icon="🧪", layout="centered")

HIDE_SIDEBAR_STYLE = """
    <style>
        [data-testid="stSidebar"] {display: none;}
        [data-testid="stSidebarNav"] {display: none;}
        [data-testid="collapsedControl"] {display: none;}
    </style>
"""

st.markdown(HIDE_SIDEBAR_STYLE, unsafe_allow_html=True)


def _load_config() -> Dict[str, Any]:
    """Carrega o config.yaml da raiz do projeto."""
    config_path = Path("config.yaml")
    if not config_path.exists():
        return {}

    with open(config_path, "r", encoding="utf-8") as file:
        return yaml.load(file, Loader=SafeLoader) or {}


def _normalize_credentials(config: Dict[str, Any]) -> Dict[str, Any]:
    """Garante que todos os usuários tenham senha em formato bcrypt."""
    credentials = config.get("credentials", {})
    usernames = credentials.get("usernames", {})
    normalized: Dict[str, Dict[str, Any]] = {}

    for username, data in usernames.items():
        if not isinstance(data, dict):
            continue

        password = data.get("password")
        if not password:
            continue

        password_str = str(password)
        if not password_str.startswith("$2"):
            password_str = stauth.Hasher([password_str]).generate()[0]

        normalized[username] = {**data, "password": password_str}

    return {"usernames": normalized}


def _cookie_settings(config: Dict[str, Any]) -> Dict[str, Any]:
    cookie = config.get("cookie", {})
    return {
        "name": cookie.get("name") or "pesquisa_clinica_auth",
        "key": cookie.get("key") or "troque-esta-chave",
        "expiry_days": int(cookie.get("expiry_days") or 1),
    }


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


config_data = _load_config()
credentials = _normalize_credentials(config_data)
admin_whitelist = _admin_email_whitelist(config_data)

if not credentials.get("usernames"):
    st.error("Nenhum usuário definido em config.yaml.")
    st.stop()

cookie_cfg = _cookie_settings(config_data)

authenticator = stauth.Authenticate(
    credentials,
    cookie_cfg["name"],
    cookie_cfg["key"],
    cookie_cfg["expiry_days"],
)

st.title("Portal Pesquisa Clínica")
st.write("Acesse usando o usuário e senha definidos no arquivo `config.yaml`.")

name, authentication_status, username = authenticator.login("Login", "main")

if authentication_status:
    st.success(f"Bem-vindo(a), {name}!")
    st.session_state["authentication_status"] = True
    st.session_state["username"] = username
    user_data = credentials.get("usernames", {}).get(username, {})
    user_email = str(user_data.get("email", "")).strip().lower()
    is_admin = bool(user_email and user_email in admin_whitelist)

    st.subheader("Escolha a funcionalidade desejada")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Cadastro", use_container_width=True):
            _switch_page("cadastro.py")
    with col2:
        if st.button("Administração de Desvios", use_container_width=True):
            if is_admin:
                _switch_page("admDesvios.py")
            else:
                st.error("Acesso restrito a e-mails autorizados pela administração.")

    if not is_admin:
        st.caption(
            "Somente usuários com e-mail autorizado podem acessar a área administrativa."
        )

    authenticator.logout("Sair", "main")
elif authentication_status is False:
    st.error("Usuário ou senha incorretos. Verifique os valores em config.yaml.")
else:
    st.info("Informe as credenciais para continuar.")