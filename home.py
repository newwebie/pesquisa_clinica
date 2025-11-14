# Home.py — Tela de login híbrida (Microsoft como principal + Usuário/Senha via config.yaml)
# Dependências:
#   pip install streamlit msal pyyaml streamlit-authenticator
#
# Configuração necessária:
# - OAuth2 Microsoft em st.secrets["auth"]  (client_id / client_secret / tenant_id / redirect_uri_local / redirect_uri_prod / scope)
#   (seus valores antigos em st.secrets["microsoft"] devem ser migrados para "auth")
# - Usuários locais (username/senha/roles/cargo/email) em config.yaml
#
# O que este arquivo faz:
# 1) Mostra Microsoft SSO estilizado vindo do módulo (create_login_page)
# 2) Mostra fallback Usuário/Senha (streamlit-authenticator)
# 3) Após autenticar, padroniza st.session_state["author"] = {username, name, email, cargo, perfil, roles, source}
# 4) Redireciona por perfil

from __future__ import annotations

# ===== IMPORTA O SEU MÓDULO =====
from auth_microsoft import MicrosoftAuth, AuthManager, create_login_page, create_user_header
import os
from pathlib import Path
from typing import Optional, Dict, Any
import streamlit as st
import yaml
from yaml.loader import SafeLoader
import streamlit_authenticator as stauth

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
# Funções de mapeamento Microsoft → author (PADRÃO DO SEU APP)
# -----------------------------------------------------
def _graph_to_author(user_info: Dict[str, Any]) -> Dict[str, Any]:
    """
    Converte o payload do Microsoft Graph /me no formato padronizado do app:
    {username, name, email, cargo, perfil, roles, source}
    """
    email = user_info.get("mail") or user_info.get("userPrincipalName")
    name = user_info.get("displayName") or email or user_info.get("id") or "Usuário"
    cargo = user_info.get("jobTitle")

    # Regra default (ajuste se quiser mapear por domínio, grupos, etc.)
    perfil = "visualizador"
    # Exemplo opcional: domínio Synvia como editor
    # dom = (email or "").split("@")[-1].lower() if email else ""
    # if dom == "synvia.com":
    #     perfil = "editor"

    return {
        "username": email or user_info.get("id"),
        "name": name,
        "email": email,
        "cargo": cargo,
        "roles": [],  # mapeie grupos do AAD aqui, se necessário
        "perfil": perfil,
        "source": "oauth2_microsoft",
    }

def _redirect_by_perfil():
    perfil = st.session_state.get("perfil") or st.session_state.get("author", {}).get("perfil", "visualizador")
    destino = "Home.py" if perfil == "visualizador" else DESTINO_POR_PERFIL.get(perfil, DEFAULT_PAGE)
    _go(destino)
    st.stop()

# -----------------------------------------------------
# UI
# -----------------------------------------------------


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
    _redirect_by_perfil()

# ===== 1) Fluxo Microsoft como principal — via SEU MÓDULO =====
auth = MicrosoftAuth()

# Renderiza a tela de login estilizada e processa callback ?code=
if create_login_page(auth):
    # Já autenticou pelo módulo. Garante token válido e popula o padrão do app.
    AuthManager.check_and_refresh_token(auth)
    user_info = AuthManager.get_current_user() or {}

    # Normaliza para o seu formato st.session_state["author"]
    author = _graph_to_author(user_info)
    st.session_state["author"] = author
    st.session_state["authentication_status"] = True
    st.session_state["username"] = author.get("username")
    st.session_state["perfil"] = author.get("perfil", "visualizador")

    st.success(f"Bem-vindo(a), {author.get('name')}!")
    _redirect_by_perfil()




# ---------------------------------------------
# Helpers
# ---------------------------------------------
def _first_nonempty(*vals):
    for v in vals:
        if isinstance(v, str) and v.strip():
            return v.strip()
        if v:  # qualquer truthy
            return v
    return None

def load_config_yaml():
    """
    Lê o config.yaml do caminho indicado na env CONFIG_YAML (se existir)
    ou do arquivo local ./config.yaml. Retorna dict (ou {}).
    """
    path = os.getenv("CONFIG_YAML", "config.yaml")
    if not os.path.exists(path):
        return {}

    try:
        with open(path, "r", encoding="utf-8") as f:
            cfg = yaml.safe_load(f) or {}
            return cfg
    except Exception as e:
        st.error(f"Falha ao ler {path}: {e}")
        return {}

def _normalize_creds(creds_cfg: dict) -> dict:
    """
    Garante o formato esperado pelo streamlit-authenticator:
    {
      "usernames": {
         "login": {
            "name": "Nome",
            "email": "email@x.com",
            "password": "<bcrypt>",
            "roles": [...],
            "perfil": "visualizador",
            ...
         }
      }
    }
    Suporta 'password_plain' e converte para 'password' (bcrypt).
    Ignora usuários sem senha válida (plain ou hash).
    """
    if not isinstance(creds_cfg, dict):
        return {}

    # aceitar tanto "usernames" quanto "USERS"/"users" etc.
    usernames = (
        creds_cfg.get("usernames")
        or creds_cfg.get("USERNAMES")
        or creds_cfg.get("users")
        or creds_cfg.get("USERS")
        or {}
    )
    if not isinstance(usernames, dict):
        return {}

    normalized_users = {}
    to_hash_batch = []

    # 1) primeira passada: coletar quem tem password_plain
    for uname, rec in usernames.items():
        rec = rec or {}
        pwd_hash = rec.get("password") or rec.get("PASSWORD")
        pwd_plain = rec.get("password") or rec.get("PASSWORD_PLAIN")

        # se só tem plain, vamos hashear depois
        if (not pwd_hash) and pwd_plain:
            to_hash_batch.append((uname, str(pwd_plain)))

    # 2) gerar hashes (se houver)
    if to_hash_batch:
        # Hasher recebe lista de senhas e devolve lista de hashes na mesma ordem
        hashed_list = stauth.Hasher([p for _, p in to_hash_batch]).generate()
        for (uname, _), h in zip(to_hash_batch, hashed_list):
            # injetar de volta no dict original para unificar caminho
            user_rec = usernames.get(uname) or {}
            user_rec["password"] = h
            usernames[uname] = user_rec

    # 3) segunda passada: montar normalized_users só dos válidos
    for uname, rec in usernames.items():
        rec = rec or {}
        pwd_hash = rec.get("password") or rec.get("PASSWORD")

        # streamlit-authenticator exige hash bcrypt aqui
        if not pwd_hash or not isinstance(pwd_hash, str) or not pwd_hash.startswith("$2"):
            # se não for hash válido, pula o usuário
            continue

        normalized_users[uname] = {
            "name": _first_nonempty(
                f"{(rec.get('first_name') or '').strip()} {(rec.get('last_name') or '').strip()}".strip(),
                rec.get("name"),
                uname,
            ),
            "email": rec.get("email"),
            "password": pwd_hash,
            "cargo": rec.get("cargo"),
            "roles": rec.get("roles", []) or [],
            "perfil": rec.get("perfil", "visualizador"),
        }

    if not normalized_users:
        return {}

    return {"usernames": normalized_users}

def _read_cookie_cfg(cfg: dict) -> dict:
    cookie = cfg.get("cookie") or cfg.get("COOKIE") or {}
    return {
        "name": (cookie.get("name") or cookie.get("NAME") or "app_cookie"),
        "key": (cookie.get("key") or cookie.get("KEY") or "troque-esta-chave"),
        "expiry_days": int(cookie.get("expiry_days") or cookie.get("EXPIRY_DAYS") or 7),
    }

def _redirect_by_perfil():
    """
    Seu roteamento por perfil. Troque pelo seu fluxo.
    Exemplo: st.switch_page("pages/01_Painel.py")
    """
    pass