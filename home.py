# Home.py — Tela de login híbrida (Microsoft como principal + Usuário/Senha via config.yaml)
# Dependências:
#   pip install streamlit msal pyyaml streamlit-authenticator
#
# Configuração necessária (será enviado em seguida):
# - OAuth2 Microsoft (TENANT / CLIENT / SECRET / REDIRECT) em st.secrets["microsoft"]
# - Usuários locais (username/senha/roles/cargo/email) em config.yaml
#
# O que este arquivo faz:
# 1) Mostra o botão principal: "Entrar com Microsoft" (SSO corporativo via MSAL)
# 2) Mostra um botão secundário: "Entrar com usuário e senha" (abre UI do streamlit-authenticator)
# 3) Após autenticar, padroniza st.session_state["author"] com {username, name, email, cargo, perfil, roles, source}
# 4) Redireciona conforme perfil: administrador/editor/visualizador

from __future__ import annotations

import os
import json
from pathlib import Path
from urllib.parse import urlencode
from typing import Optional

import streamlit as st
import yaml
from yaml.loader import SafeLoader
import streamlit_authenticator as stauth

# MSAL só é necessário se o usuário clicar em Microsoft
try:
    from msal import ConfidentialClientApplication
except Exception:  # pragma: no cover
    ConfidentialClientApplication = None

# -----------------------------------------------------
# Config da página
# -----------------------------------------------------
st.set_page_config(page_title="Login", page_icon="🔐", layout="centered")

# -----------------------------------------------------
# Navegação e utilitários
# -----------------------------------------------------
PAGES_ROOT = Path(__file__).parent
DESTINO_POR_PERFIL = {
    "administrador": "pages/01_Administrador.py",
    "editor": "pages/02_Editor.py",
    "visualizador": "Home.py",
}
DEFAULT_PAGE = "Home.py"


def _go(destino: str):
    """Redireciona para uma página do Streamlit."""
    if not destino:
        destino = DEFAULT_PAGE

    if not destino.endswith(".py"):
        if destino.startswith("pages/"):
            destino = f"{destino}.py"
        else:
            destino = f"pages/{destino}.py"

    destino_path = PAGES_ROOT / destino
    if destino == "Home.py":
        destino_path = PAGES_ROOT / "Home.py"

    if not destino_path.exists():
        st.error(f"Página não encontrada: {destino}")
        return

    if hasattr(st, "switch_page"):
        if destino_path.name == "Home.py":
            st.switch_page("Home.py")
        else:
            st.switch_page(destino)
    else:
        st.success(f"Login efetuado. Abra a página **{destino}** no menu à esquerda.")
        st.stop()


def _first_nonempty(*vals):
    for v in vals:
        if v and str(v).strip():
            return str(v).strip()
    return None


# -----------------------------------------------------
# Carregar config.yaml (usuários locais)
# -----------------------------------------------------
@st.cache_data(show_spinner=False)
def load_config_yaml() -> dict:
    cfg_path = Path("config.yaml")
    if not cfg_path.exists():
        return {}
    with open(cfg_path, "r", encoding="utf-8") as f:
        return yaml.load(f, Loader=SafeLoader) or {}


# -----------------------------------------------------
# Microsoft OAuth (MSAL)
# -----------------------------------------------------
@st.cache_resource(show_spinner=False)
def get_msal_app():
    ms = st.secrets.get("microsoft", {})
    tenant = ms.get("tenant_id")
    client_id = ms.get("client_id")
    client_secret = ms.get("client_secret")
    authority = f"https://login.microsoftonline.com/{tenant}"
    if not (tenant and client_id and client_secret):
        raise RuntimeError("Config Microsoft incompleta em st.secrets['microsoft'].")
    if ConfidentialClientApplication is None:
        raise RuntimeError("Pacote 'msal' não encontrado. Instale com: pip install msal")
    return ConfidentialClientApplication(client_id, authority=authority, client_credential=client_secret)


def microsoft_login_url() -> str:
    ms = st.secrets.get("microsoft", {})
    tenant = ms.get("tenant_id")
    client_id = ms.get("client_id")
    redirect_uri = ms.get("redirect_uri")
    scope = ms.get("scope", ["User.Read", "email", "openid", "profile"])  # escopos padrão

    if not (tenant and client_id and redirect_uri):
        raise RuntimeError("Preencha tenant_id/client_id/redirect_uri em st.secrets['microsoft'].")

    params = {
        "client_id": client_id,
        "response_type": "code",
        "redirect_uri": redirect_uri,
        "response_mode": "query",
        "scope": " ".join(scope),
        "state": "msal_login",  # pode trocar para anti-CSRF
    }
    return f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize?{urlencode(params)}"


def handle_microsoft_callback() -> Optional[dict]:
    """Se a URL tiver ?code=..., troca por tokens e devolve claims (id_token)."""
    qs = st.query_params
    if "code" not in qs:
        return None

    code = qs.get("code")
    ms = st.secrets.get("microsoft", {})
    redirect_uri = ms.get("redirect_uri")
    scope = ms.get("scope", ["User.Read", "email", "openid", "profile"])  # escopos padrão

    app = get_msal_app()
    token = app.acquire_token_by_authorization_code(code, scopes=scope, redirect_uri=redirect_uri)

    if "id_token_claims" not in token:
        # Mostra erro detalhado, se houver
        err = token.get("error_description") or token
        st.error(f"Falha na autenticação Microsoft: {err}")
        return None

    claims = token["id_token_claims"]

    # Campos úteis para montar o "author"
    email = (
        claims.get("email")
        or (claims.get("preferred_username") if "@" in str(claims.get("preferred_username", "")) else None)
        or (claims.get("unique_name") if "@" in str(claims.get("unique_name", "")) else None)
    )
    nome = claims.get("name") or email or claims.get("oid")

    # Regra de perfil/cargo: default visualizador; você pode mapear domínio/email se quiser
    perfil = "visualizador"
    cargo = None

    # Exemplo de mapeamento opcional por domínio (se desejar):
    # dom = str(email).split("@")[-1].lower() if email else ""
    # if dom == "synvia.com":
    #     perfil = "editor"

    author = {
        "username": email or claims.get("oid"),
        "name": nome,
        "email": email,
        "cargo": cargo,
        "roles": [],
        "perfil": perfil,
        "source": "oauth2_microsoft",
    }
    return author


# -----------------------------------------------------
# UI
# -----------------------------------------------------
st.title("🔐 Login")

# Esconde sidebar enquanto não autenticado
if not st.session_state.get("authentication_status"):
    st.markdown(
        """
        <style>
            [data-testid="stSidebar"] {display: none;}
            [data-testid="stSidebarNav"] {display: none;}
            [data-testid="collapsedControl"] {display: none;}
        </style>
        """,
        unsafe_allow_html=True,
    )

# Se já autenticado anteriormente (por qualquer método), redireciona
if st.session_state.get("authentication_status") is True and st.session_state.get("author"):
    perfil = st.session_state.get("perfil") or st.session_state["author"].get("perfil", "visualizador")
    destino = "Home.py" if perfil == "visualizador" else DESTINO_POR_PERFIL.get(perfil, DEFAULT_PAGE)
    _go(destino)
    st.stop()

# 1) Fluxo Microsoft como principal
with st.container(border=True):
    st.subheader("Entrar com Microsoft (SSO)")
    col1, col2 = st.columns([1, 2])
    with col1:
        if st.button("🔵 Entrar com Microsoft", use_container_width=True):
            try:
                st.session_state["_ms_login_clicked"] = True
                st.rerun()
            except Exception as e:
                st.error(e)

    # Se clicou, redireciona para o authorize
    if st.session_state.get("_ms_login_clicked"):
        try:
            url = microsoft_login_url()
            st.markdown(f"<meta http-equiv='refresh' content='0; url={url}'>", unsafe_allow_html=True)
        except Exception as e:
            st.error(e)

# Se retornou de Microsoft com code, processa agora
ms_author = handle_microsoft_callback()
if ms_author:
    # Marca como autenticado
    st.session_state["author"] = ms_author
    st.session_state["authentication_status"] = True
    st.session_state["username"] = ms_author.get("username")
    st.session_state["perfil"] = ms_author.get("perfil", "visualizador")
    st.success(f"Bem-vindo(a), {ms_author.get('name')}!")
    destino = "Home.py" if ms_author["perfil"] == "visualizador" else DESTINO_POR_PERFIL.get(ms_author["perfil"], DEFAULT_PAGE)
    _go(destino)
    st.stop()

# 2) Fluxo Usuário/Senha via config.yaml (secundário)
with st.container(border=True):
    st.subheader("Ou entrar com Usuário e Senha")

    cfg = load_config_yaml()
    creds_cfg = cfg.get("credentials") or cfg.get("credentials".upper())

    if not creds_cfg:
        st.info("Arquivo config.yaml não encontrado ou sem a chave 'credentials'.")
    else:
        try:
            authenticator = stauth.Authenticate(
                creds_cfg,
                cookie_name=cfg.get("cookie", {}).get("name", "app_cookie"),
                key=cfg.get("cookie", {}).get("key", "troque-esta-chave"),
                cookie_expiry_days=int(cfg.get("cookie", {}).get("expiry_days", 7)),
            )

            # Renderiza o login da lib
            name, auth_status, username = authenticator.login("Entrar", "main")

            if auth_status:
                urec = creds_cfg["usernames"][username]
                full_name = _first_nonempty(
                    f"{urec.get('first_name','')} {urec.get('last_name','')}".strip(),
                    urec.get("name"),
                    username,
                )
                # Monta o author padronizado
                author = {
                    "username": username,
                    "name": full_name,
                    "email": urec.get("email"),
                    "cargo": urec.get("cargo"),
                    "roles": urec.get("roles", []),
                    "perfil": urec.get("perfil", "visualizador"),
                    "source": "password",
                }
                st.session_state["author"] = author
                st.session_state["authentication_status"] = True
                st.session_state["username"] = username
                st.session_state["perfil"] = author["perfil"]

                st.success(f"Bem-vindo(a), {author['name']}!")
                destino = "Home.py" if author["perfil"] == "visualizador" else DESTINO_POR_PERFIL.get(author["perfil"], DEFAULT_PAGE)
                _go(destino)
                st.stop()
            elif auth_status is False:
                st.error("Usuário ou senha inválidos.")
            else:
                st.caption("Informe suas credenciais para entrar.")
        except Exception as e:
            st.error(f"Erro na autenticação local: {e}")

# Botão de logout (apenas se autenticado; normalmente ficará em páginas internas)
if st.session_state.get("authentication_status"):
    try:
        # Se login foi via streamlit-authenticator, o botão deles gerencia o cookie
        authenticator.logout("Sair", "sidebar")  # pode ignorar se não autenticou via senha
    except Exception:
        # Logout manual para o caso do Microsoft
        if st.sidebar.button("Sair"):
            for k in ("authentication_status", "username", "author", "perfil", "_ms_login_clicked"):
                if k in st.session_state:
                    del st.session_state[k]
            st.rerun()
