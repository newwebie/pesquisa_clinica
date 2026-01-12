"""Portal Pesquisa Clínica - Fluxo: Login → Estudos → Desvios"""
from __future__ import annotations
import pandas as pd
import psycopg2
import streamlit as st
from functools import lru_cache
import requests
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import logging

from auth_microsoft import (
    AuthManager,
    MicrosoftAuth,
    create_login_page,
    create_user_header,
)

# -------------------------------------------------
# Configuração básica da página
# -------------------------------------------------
st.set_page_config(
    page_title="Portal Pesquisa Clínica",
    page_icon="🧪",
    layout="wide",
)

# CSS customizado para layout mais limpo
CUSTOM_CSS = """
<style>
    [data-testid="stSidebarNav"] {display: none;}
    [data-testid="collapsedControl"] {display: none;}
    .block-container {padding-top: 1.5rem; padding-bottom: 2rem;}
    .stTabs [data-baseweb="tab-list"] {gap: 4px; margin-bottom: 1rem;}
    .stTabs [data-baseweb="tab"] {
        padding: 8px 16px;
        border-radius: 6px 6px 0 0;
        font-size: 0.9rem;
    }
    div[data-testid="stMetric"] {
        background-color: #f8f9fa;
        padding: 12px;
        border-radius: 6px;
    }
    .stExpander {margin-bottom: 0.5rem;}
    div[data-testid="stExpander"] summary {padding: 0.5rem 1rem;}
    .stButton > button {transition: all 0.2s ease;}
    hr {margin: 1rem 0;}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# -------------------------------------------------
# Conexão com o banco
# -------------------------------------------------
def get_connection():
    db = st.secrets["postgres"]
    return psycopg2.connect(
        host=db["host"],
        port=db["port"],
        dbname=db["database"],
        user=db["user"],
        password=db["password"],
    )


# -------------------------------------------------
# SharePoint Upload via Microsoft Graph API
# -------------------------------------------------
def get_graph_token() -> str | None:
    """Obtém token de acesso para o Microsoft Graph API."""
    try:
        graph = st.secrets["graph"]
        tenant_id = graph["tenant_id_graph"]
        client_id = graph["client_id_graph"]
        client_secret = graph["client_secret_graph"]

        url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        data = {
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default"
        }

        response = requests.post(url, data=data)
        if response.status_code == 200:
            return response.json().get("access_token")
        return None
    except Exception as e:
        st.error(f"Erro ao obter token Graph: {e}")
        return None


def get_sharepoint_site_id(token: str) -> str | None:
    """Obtém o ID do site SharePoint."""
    try:
        graph = st.secrets["graph"]
        hostname = graph["hostname"]
        site_path = graph["site_path"]

        url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:/{site_path}"
        headers = {"Authorization": f"Bearer {token}"}

        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            return response.json().get("id")
        return None
    except Exception:
        return None


def get_drive_id(token: str, site_id: str) -> str | None:
    """Obtém o ID do drive (biblioteca de documentos) do SharePoint."""
    try:
        graph = st.secrets["graph"]
        library_name = graph["library_name"]

        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
        headers = {"Authorization": f"Bearer {token}"}

        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            drives = response.json().get("value", [])
            for drive in drives:
                if drive.get("name") == library_name:
                    return drive.get("id")
        return None
    except Exception:
        return None


def upload_to_sharepoint(file_content: bytes, file_name: str, estudo_codigo: str, desvio_numero: int) -> str | None:
    """
    Faz upload de arquivo para o SharePoint.
    Retorna a URL do arquivo ou None em caso de erro.
    """
    try:
        token = get_graph_token()
        if not token:
            st.error("Não foi possível autenticar no SharePoint.")
            return None

        site_id = get_sharepoint_site_id(token)
        if not site_id:
            st.error("Site SharePoint não encontrado.")
            return None

        drive_id = get_drive_id(token, site_id)
        if not drive_id:
            st.error("Biblioteca de documentos não encontrada.")
            return None

        # Organiza em pasta: Desvios/{estudo_codigo}/
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        folder_path = f"Desvios/{estudo_codigo}"
        safe_filename = f"desvio_{desvio_numero}_{timestamp}_{file_name}"

        # Upload do arquivo
        upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{folder_path}/{safe_filename}:/content"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/octet-stream"
        }

        response = requests.put(upload_url, headers=headers, data=file_content)

        if response.status_code in [200, 201]:
            file_data = response.json()
            return file_data.get("webUrl")
        else:
            st.error(f"Erro no upload: {response.status_code} - {response.text}")
            return None

    except Exception as e:
        st.error(f"Erro ao fazer upload: {e}")
        return None


def pode_acessar_painel_adm(email: str) -> bool:
    """Verifica se o usuário pode acessar o Painel Administrativo (apenas Administrador)."""
    if not email:
        return False
    try:
        conn = get_connection()
        cursor = conn.cursor()
        query = "SELECT 1 FROM usuarios WHERE LOWER(email) = %s AND perfil = 'Administrador' LIMIT 1"
        cursor.execute(query, (email.lower(),))
        exists = cursor.fetchone() is not None
        cursor.close()
        conn.close()
        return exists
    except Exception as e:
        st.error(f"Erro ao verificar permissões: {e}")
        return False


def get_user_perfil(email: str) -> str | None:
    """Retorna o perfil do usuário (Administrador, Editor, Monitor, Visualizador)."""
    if not email:
        return None
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT perfil FROM usuarios WHERE LOWER(email) = %s", (email.lower(),))
        row = cursor.fetchone()
        cursor.close()
        conn.close()
        return row[0] if row else None
    except Exception as e:
        st.error(f"Erro ao buscar perfil: {e}")
        return None


def usuario_existe(email: str) -> bool:
    """Verifica se o usuário já existe na tabela usuarios."""
    if not email:
        return False
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT 1 FROM usuarios WHERE LOWER(email) = %s", (email.lower(),))
        exists = cursor.fetchone() is not None
        cursor.close()
        conn.close()
        return exists
    except Exception:
        return False


def auto_cadastrar_usuario(email: str, nome: str) -> bool:
    """
    Auto-cadastra usuário no primeiro login.
    Apenas emails @synvia.com são permitidos.
    Perfil padrão: Usuário
    """
    if not email or not email.lower().endswith("@synvia.com"):
        return False

    # Se já existe, não faz nada
    if usuario_existe(email):
        return True

    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute(
            """INSERT INTO usuarios (nome, email, cargo, perfil)
               VALUES (%s, %s, %s, %s)""",
            (nome.strip(), email.lower(), "", "Usuário")
        )
        conn.commit()
        cursor.close()
        conn.close()
        return True
    except Exception as e:
        st.error(f"Erro ao auto-cadastrar usuário: {e}")
        return False


# -------------------------------------------------
# Funções de dados: Estudos
# -------------------------------------------------
@st.cache_data(ttl=60)
def load_estudos_do_monitor(email: str) -> pd.DataFrame:
    """Carrega estudos ATIVOS onde o monitor está alocado."""
    try:
        conn = get_connection()
        query = """
            SELECT e.id, e.codigo, e.nome, e.status
            FROM estudos e
            INNER JOIN estudo_monitores em ON e.id = em.estudo_id
            WHERE LOWER(em.monitor_email) = %s
              AND e.status = 'ativo'
            ORDER BY e.nome
        """
        df = pd.read_sql_query(query, conn, params=(email.lower(),))
        conn.close()
        return df
    except Exception as e:
        st.error(f"Erro ao carregar estudos: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=60)
def is_gerente_medico(email: str) -> bool:
    """Verifica se o email pertence a um gerente médico cadastrado."""
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "SELECT 1 FROM gerentes_medicos WHERE LOWER(email) = %s",
            (email.lower(),)
        )
        result = cursor.fetchone()
        cursor.close()
        conn.close()
        return result is not None
    except Exception:
        return False


@st.cache_data(ttl=60)
def load_estudos_do_gerente_medico(email: str) -> pd.DataFrame:
    """Carrega estudos ATIVOS onde o gerente médico está alocado."""
    try:
        conn = get_connection()
        query = """
            SELECT e.id, e.codigo, e.nome, e.status
            FROM estudos e
            INNER JOIN estudo_gerente_medico egm ON e.id = egm.estudo_id
            INNER JOIN gerentes_medicos gm ON gm.id = egm.gerente_medico_id
            WHERE LOWER(gm.email) = %s
              AND e.status = 'ativo'
            ORDER BY e.nome
        """
        df = pd.read_sql_query(query, conn, params=(email.lower(),))
        conn.close()
        return df
    except Exception as e:
        st.error(f"Erro ao carregar estudos do gerente médico: {e}")
        return pd.DataFrame()


def get_estudo_by_id(estudo_id: int) -> dict | None:
    """Busca info do estudo pelo ID, incluindo data de criação."""
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "SELECT id, codigo, nome, status, criado_em FROM estudos WHERE id = %s",
            (estudo_id,)
        )
        row = cursor.fetchone()
        cursor.close()
        conn.close()
        if row:
            return {
                "id": row[0],
                "codigo": row[1],
                "nome": row[2],
                "status": row[3],
                "criado_em": row[4]
            }
        return None
    except Exception as e:
        st.error(f"Erro ao buscar estudo: {e}")
        return None


def get_patrocinador_do_estudo(estudo_id: int) -> str | None:
    """Retorna o patrocinador do estudo (via gerente médico alocado)."""
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute(
            """SELECT gm.patrocinador
               FROM gerentes_medicos gm
               INNER JOIN estudo_gerente_medico egm ON gm.id = egm.gerente_medico_id
               WHERE egm.estudo_id = %s""",
            (estudo_id,)
        )
        row = cursor.fetchone()
        cursor.close()
        conn.close()
        return row[0] if row else None
    except Exception:
        return None


# -------------------------------------------------
# Funções de dados: Desvios
# -------------------------------------------------
@st.cache_data(ttl=30)
def load_desvios_do_estudo(estudo_id: int) -> pd.DataFrame:
    """Carrega desvios de um estudo específico (exclui deletados)."""
    try:
        conn = get_connection()
        query = """
            SELECT
                id, numero_desvio_estudo, status, participante, data_ocorrido, formulario_status,
                identificacao_desvio, centro, visita, descricao_desvio,
                causa_raiz, acao_preventiva, acao_corretiva, importancia,
                data_identificacao_texto, categoria, subcategoria, codigo,
                escopo, avaliacao_gerente_medico, avaliacao_investigador,
                formulario_arquivado, recorrencia, num_ocorrencia_previa,
                prazo_escalonamento, data_escalonamento, atendeu_prazos_report,
                populacao, data_submissao_cep, data_finalizacao,
                criado_por_nome, criado_por_email, atualizado_por, data_atualizacao,
                url_anexo, xmin AS row_version
            FROM desvios
            WHERE estudo_id = %s
              AND deleted_at IS NULL
            ORDER BY numero_desvio_estudo DESC
        """
        df = pd.read_sql_query(query, conn, params=(estudo_id,))
        conn.close()
        return df
    except Exception as e:
        st.error(f"Erro ao carregar desvios: {e}")
        return pd.DataFrame()


def snake_to_title(name: str) -> str:
    parts = name.split("_")
    return " ".join(p.capitalize() for p in parts)


def get_proximo_numero_desvio(estudo_id: int) -> int:
    """
    Retorna o próximo número de desvio para um estudo específico.
    Cada estudo tem sua própria sequência começando em 1.
    """
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "SELECT COALESCE(MAX(numero_desvio_estudo), 0) + 1 FROM desvios WHERE estudo_id = %s",
            (estudo_id,)
        )
        proximo = cursor.fetchone()[0]
        cursor.close()
        conn.close()
        return proximo
    except Exception:
        return 1


# -------------------------------------------------
# Funções de dados: Admin - Estudos
# -------------------------------------------------
@st.cache_data(ttl=30)
def load_todos_estudos() -> pd.DataFrame:
    """Carrega todos os estudos (admin)."""
    try:
        conn = get_connection()
        query = """
            SELECT id, codigo, nome, status, criado_em
            FROM estudos
            ORDER BY status DESC, nome
        """
        df = pd.read_sql_query(query, conn)
        conn.close()
        return df
    except Exception as e:
        st.error(f"Erro ao carregar estudos: {e}")
        return pd.DataFrame()


def criar_estudo(codigo: str, nome: str) -> bool:
    """Cria um novo estudo."""
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO estudos (codigo, nome, status) VALUES (%s, %s, 'ativo')",
            (codigo.strip(), nome.strip())
        )
        conn.commit()
        cursor.close()
        conn.close()
        return True
    except Exception as e:
        st.error(f"Erro ao criar estudo: {e}")
        return False


def atualizar_estudo(estudo_id: int, codigo: str, nome: str, status: str) -> bool:
    """Atualiza um estudo existente."""
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "UPDATE estudos SET codigo = %s, nome = %s, status = %s WHERE id = %s",
            (codigo.strip(), nome.strip(), status, estudo_id)
        )
        conn.commit()
        cursor.close()
        conn.close()
        return True
    except Exception as e:
        st.error(f"Erro ao atualizar estudo: {e}")
        return False


# -------------------------------------------------
# Funções de dados: Admin - Monitores
# -------------------------------------------------
@st.cache_data(ttl=30)
def load_monitores_do_estudo(estudo_id: int) -> pd.DataFrame:
    """Carrega monitores alocados em um estudo (com nome da tabela usuarios)."""
    try:
        conn = get_connection()
        query = """
            SELECT em.id, em.monitor_email, u.nome as monitor_nome, em.alocado_em
            FROM estudo_monitores em
            LEFT JOIN usuarios u ON LOWER(u.email) = LOWER(em.monitor_email)
            WHERE em.estudo_id = %s
            ORDER BY u.nome, em.monitor_email
        """
        df = pd.read_sql_query(query, conn, params=(estudo_id,))
        conn.close()
        return df
    except Exception as e:
        st.error(f"Erro ao carregar monitores: {e}")
        return pd.DataFrame()


def alocar_monitor(estudo_id: int, email: str) -> bool:
    """Aloca um monitor em um estudo."""
    try:
        conn = get_connection()
        cursor = conn.cursor()
        # Verifica se já está alocado
        cursor.execute(
            "SELECT 1 FROM estudo_monitores WHERE estudo_id = %s AND LOWER(monitor_email) = %s",
            (estudo_id, email.lower())
        )
        if cursor.fetchone():
            st.warning("Este monitor já está alocado neste estudo.")
            cursor.close()
            conn.close()
            return False
        cursor.execute(
            "INSERT INTO estudo_monitores (estudo_id, monitor_email) VALUES (%s, %s)",
            (estudo_id, email.lower())
        )
        conn.commit()
        cursor.close()
        conn.close()
        return True
    except Exception as e:
        st.error(f"Erro ao alocar monitor: {e}")
        return False


def remover_monitor(alocacao_id: int) -> bool:
    """Remove um monitor de um estudo."""
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM estudo_monitores WHERE id = %s", (alocacao_id,))
        conn.commit()
        cursor.close()
        conn.close()
        return True
    except Exception as e:
        st.error(f"Erro ao remover monitor: {e}")
        return False


# -------------------------------------------------
# Funções de dados: Admin - Gerentes Médicos
# -------------------------------------------------
@st.cache_data(ttl=30)
def load_gerentes_medicos() -> pd.DataFrame:
    """Carrega lista de todos os gerentes médicos cadastrados."""
    try:
        conn = get_connection()
        query = "SELECT id, nome, email, patrocinador FROM gerentes_medicos ORDER BY nome"
        df = pd.read_sql_query(query, conn)
        conn.close()
        return df
    except Exception as e:
        st.error(f"Erro ao carregar gerentes médicos: {e}")
        return pd.DataFrame()


def criar_gerente_medico(nome: str, email: str, patrocinador: str) -> bool:
    """Cadastra um novo gerente médico."""
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT 1 FROM gerentes_medicos WHERE LOWER(email) = %s", (email.lower(),))
        if cursor.fetchone():
            st.warning("Este email já está cadastrado como gerente médico.")
            cursor.close()
            conn.close()
            return False
        cursor.execute(
            "INSERT INTO gerentes_medicos (nome, email, patrocinador) VALUES (%s, %s, %s)",
            (nome.strip(), email.lower(), patrocinador.strip())
        )
        conn.commit()
        cursor.close()
        conn.close()
        return True
    except Exception as e:
        st.error(f"Erro ao criar gerente médico: {e}")
        return False


def remover_gerente_medico(gerente_id: int) -> bool:
    """Remove um gerente médico."""
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM gerentes_medicos WHERE id = %s", (gerente_id,))
        conn.commit()
        cursor.close()
        conn.close()
        return True
    except Exception as e:
        st.error(f"Erro ao remover gerente médico: {e}")
        return False


def get_gerente_medico_do_estudo(estudo_id: int) -> dict | None:
    """Retorna o gerente médico alocado em um estudo."""
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute(
            """SELECT gm.id, gm.nome, gm.email
               FROM gerentes_medicos gm
               INNER JOIN estudo_gerente_medico egm ON gm.id = egm.gerente_medico_id
               WHERE egm.estudo_id = %s""",
            (estudo_id,)
        )
        row = cursor.fetchone()
        cursor.close()
        conn.close()
        if row:
            return {"id": row[0], "nome": row[1], "email": row[2]}
        return None
    except Exception:
        return None


def alocar_gerente_medico(estudo_id: int, gerente_id: int) -> bool:
    """Aloca um gerente médico em um estudo (substitui o anterior se houver)."""
    try:
        conn = get_connection()
        cursor = conn.cursor()
        # Remove alocação anterior
        cursor.execute("DELETE FROM estudo_gerente_medico WHERE estudo_id = %s", (estudo_id,))
        # Insere nova alocação
        cursor.execute(
            "INSERT INTO estudo_gerente_medico (estudo_id, gerente_medico_id) VALUES (%s, %s)",
            (estudo_id, gerente_id)
        )
        conn.commit()
        cursor.close()
        conn.close()
        return True
    except Exception as e:
        st.error(f"Erro ao alocar gerente médico: {e}")
        return False


def remover_gerente_medico_do_estudo(estudo_id: int) -> bool:
    """Remove o gerente médico de um estudo."""
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM estudo_gerente_medico WHERE estudo_id = %s", (estudo_id,))
        conn.commit()
        cursor.close()
        conn.close()
        return True
    except Exception as e:
        st.error(f"Erro ao remover gerente médico do estudo: {e}")
        return False


def get_estudos_do_gerente_medico_por_id(gerente_id: int) -> list[str]:
    """Retorna lista de códigos dos estudos onde o GM está alocado."""
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute(
            """SELECT e.codigo
               FROM estudos e
               INNER JOIN estudo_gerente_medico egm ON e.id = egm.estudo_id
               WHERE egm.gerente_medico_id = %s
               ORDER BY e.codigo""",
            (gerente_id,)
        )
        rows = cursor.fetchall()
        cursor.close()
        conn.close()
        return [row[0] for row in rows]
    except Exception:
        return []


def contar_desvios_do_estudo(estudo_id: int) -> int:
    """Retorna a quantidade de desvios de um estudo (exclui deletados)."""
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "SELECT COUNT(*) FROM desvios WHERE estudo_id = %s AND deleted_at IS NULL",
            (estudo_id,)
        )
        count = cursor.fetchone()[0]
        cursor.close()
        conn.close()
        return count
    except Exception:
        return 0


def contar_pendencias_do_estudo(estudo_id: int) -> int:
    """Retorna a quantidade de desvios sem avaliação do Gerente Médico (pendências), excluindo deletados."""
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute(
            """SELECT COUNT(*) FROM desvios
               WHERE estudo_id = %s
               AND (avaliacao_gerente_medico IS NULL OR avaliacao_gerente_medico = '')
               AND deleted_at IS NULL""",
            (estudo_id,)
        )
        count = cursor.fetchone()[0]
        cursor.close()
        conn.close()
        return count
    except Exception:
        return 0


def get_nomes_monitores_do_estudo(estudo_id: int) -> list[str]:
    """Retorna lista de nomes dos monitores alocados em um estudo."""
    df = load_monitores_do_estudo(estudo_id)
    if df.empty:
        return []
    nomes = []
    for _, mon in df.iterrows():
        nome = mon['monitor_nome'] if mon['monitor_nome'] else mon['monitor_email']
        nomes.append(nome)
    return nomes


def get_emails_do_estudo(estudo_id: int) -> list[str]:
    """Retorna lista de emails de todos os participantes do estudo (monitores + gerente médico)."""
    emails = set()

    # Monitores
    df_monitores = load_monitores_do_estudo(estudo_id)
    if not df_monitores.empty:
        for _, mon in df_monitores.iterrows():
            if mon['monitor_email']:
                emails.add(mon['monitor_email'].lower())

    # Gerente Médico
    gerente = get_gerente_medico_do_estudo(estudo_id)
    if gerente and gerente.get('email'):
        emails.add(gerente['email'].lower())

    return list(emails)


# Mapeamento de nomes de campos do banco para nomes de exibição
CAMPOS_DISPLAY_NAMES = {
    'participante': 'Participante',
    'data_ocorrido': 'Data do Ocorrido',
    'formulario_status': 'Formulário',
    'identificacao_desvio': 'Identificação do Desvio',
    'centro': 'Centro',
    'visita': 'Visita',
    'descricao_desvio': 'Descrição do Desvio',
    'causa_raiz': 'Causa Raiz',
    'acao_preventiva': 'Ação Preventiva',
    'acao_corretiva': 'Ação Corretiva',
    'importancia': 'Importância',
    'data_identificacao_texto': 'Data de Identificação',
    'categoria': 'Categoria',
    'subcategoria': 'Subcategoria',
    'codigo': 'Código',
    'recorrencia': 'Recorrência',
    'num_ocorrencia_previa': 'N° Desvio Ocorrência Prévia',
    'escopo': 'Escopo',
    'prazo_escalonamento': 'Prazo para Escalonamento',
    'data_escalonamento': 'Data de Escalonamento',
    'atendeu_prazos_report': 'Atendeu os Prazos de Report?',
    'motivo_nao_atendeu_prazo': 'Motivo (se não atendeu prazos)',
    'populacao': 'População',
    'data_submissao_cep': 'Data de Submissão ao CEP',
    'data_finalizacao': 'Data de Finalização',
    'formulario_arquivado': 'Formulário Arquivado (ISF e TFM)?',
    'avaliacao_investigador': 'Avaliação Investigador Principal',
    'avaliacao_gerente_medico': 'Avaliação Gerente Médico',
    'status': 'Status',
    'url_anexo': 'Anexo',
}

# Mapeamento de tradução PT -> EN para campos de selectbox (usado pelo site externo)
TRADUCAO_PT_EN = {
    # Status do desvio
    'Novo': 'New',
    'Modificado': 'Modified',
    'Avaliado': 'Evaluated',
    # Sim/Não (formulário, atendeu prazos, etc.)
    'Sim': 'Yes',
    'Não': 'No',
    # Importância
    'Maior': 'Major',
    'Menor': 'Minor',
    # Recorrência
    'Recorrente': 'Recurring',
    'Não Recorrente': 'Non-Recurring',
    # Escopo
    'Protocolo': 'Protocol',
    'GCP': 'GCP',
    # N/A permanece igual
    'N/A': 'N/A',
    # Prazo de Escalonamento
    'Mensal': 'Monthly',
    'Padrão': 'Default',
    'Imediato': 'Immediate',
    # População
    'Intenção de Tratar (ITT)': 'ITT',
    'Por Protocolo (PP)': 'Per Protocol',
}

# Mapeamento inverso EN -> PT (para quando precisar converter de volta)
TRADUCAO_EN_PT = {v: k for k, v in TRADUCAO_PT_EN.items()}


def traduzir_valor_para_ingles(valor: str) -> str:
    """Traduz um valor de selectbox de português para inglês."""
    if not valor:
        return valor
    return TRADUCAO_PT_EN.get(valor, valor)


def traduzir_valor_para_portugues(valor: str) -> str:
    """Traduz um valor de selectbox de inglês para português."""
    if not valor:
        return valor
    return TRADUCAO_EN_PT.get(valor, valor)


def traduzir_desvio_para_ingles(desvio: dict) -> dict:
    """Traduz os campos de selectbox de um desvio para inglês."""
    campos_traduzir = [
        'status', 'importancia', 'recorrencia', 'escopo',
        'atendeu_prazos_report', 'formulario_arquivado', 'formulario_status'
    ]
    desvio_traduzido = desvio.copy()
    for campo in campos_traduzir:
        if campo in desvio_traduzido and desvio_traduzido[campo]:
            desvio_traduzido[campo] = traduzir_valor_para_ingles(desvio_traduzido[campo])
    return desvio_traduzido


def traduzir_desvios_para_ingles(desvios: list[dict]) -> list[dict]:
    """Traduz uma lista de desvios para inglês."""
    return [traduzir_desvio_para_ingles(d) for d in desvios]


def get_campo_display_name(campo: str) -> str:
    """Retorna o nome de exibição de um campo."""
    return CAMPOS_DISPLAY_NAMES.get(campo, campo.replace('_', ' ').title())


def enviar_email_notificacao_desvio(
    estudo_id: int,
    estudo_codigo: str,
    estudo_nome: str,
    desvio_id: int,
    numero_desvio: int,
    alteracoes: list[dict],  # Lista de {'campo': str, 'valor_antigo': any, 'valor_novo': any}
    alterado_por: str
):
    """Envia email de notificação para todos os participantes do estudo quando um desvio é modificado."""
    try:
        # Configurações de email do secrets.toml
        email_config = st.secrets.get("email", {})
        smtp_server = email_config.get("smtp_server")
        smtp_port = email_config.get("smtp_port", 587)
        sender = email_config.get("sender")
        password = email_config.get("password")

        if not all([smtp_server, sender, password]):
            logging.warning("Configurações de email incompletas no secrets.toml")
            return False

        # Destinatários
        destinatarios = get_emails_do_estudo(estudo_id)
        if not destinatarios:
            logging.info("Nenhum destinatário para notificação de desvio")
            return True

        # Monta o email
        assunto = f"[Desvio Modificado] {estudo_codigo} - Desvio {numero_desvio}"

        # Monta a tabela de alterações
        alteracoes_html = ""
        for alt in alteracoes:
            campo_display = get_campo_display_name(alt['campo'])
            valor_antigo = alt.get('valor_antigo') or '-'
            valor_novo = alt.get('valor_novo') or '-'
            alteracoes_html += f"""
                <tr>
                    <td style="padding: 12px 15px; border-bottom: 1px solid #e0e0e0; font-weight: 500; color: #333;">{campo_display}</td>
                    <td style="padding: 12px 15px; border-bottom: 1px solid #e0e0e0; color: #999; text-decoration: line-through;">{valor_antigo}</td>
                    <td style="padding: 12px 15px; border-bottom: 1px solid #e0e0e0; color: #2e7d32; font-weight: 500;">{valor_novo}</td>
                </tr>
            """

        corpo_html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
        </head>
        <body style="margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f5f5f5;">
            <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #f5f5f5; padding: 30px 0;">
                <tr>
                    <td align="center">
                        <table width="600" cellpadding="0" cellspacing="0" style="background-color: #ffffff; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.1); overflow: hidden;">

                            <!-- Header -->
                            <tr>
                                <td style="background: linear-gradient(135deg, #6BBF47 0%, #52B54B 100%); padding: 30px 40px; text-align: center;">
                                    <h1 style="margin: 0; color: #ffffff; font-size: 24px; font-weight: 600;">
                                        🔔 Notificação de Alteração
                                    </h1>
                                    <p style="margin: 10px 0 0; color: rgba(255,255,255,0.9); font-size: 14px;">
                                        Portal Pesquisa Clínica
                                    </p>
                                </td>
                            </tr>

                            <!-- Content -->
                            <tr>
                                <td style="padding: 40px;">

                                    <!-- Info Cards -->
                                    <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom: 30px;">
                                        <tr>
                                            <td style="padding: 15px; background-color: #f8f9fa; border-radius: 8px; border-left: 4px solid #6BBF47;">
                                                <table width="100%" cellpadding="0" cellspacing="0">
                                                    <tr>
                                                        <td width="50%" style="padding: 8px 0;">
                                                            <span style="color: #666; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Estudo</span><br>
                                                            <span style="color: #333; font-size: 16px; font-weight: 600;">{estudo_codigo}</span><br>
                                                            <span style="color: #666; font-size: 13px;">{estudo_nome}</span>
                                                        </td>
                                                        <td width="50%" style="padding: 8px 0; text-align: right;">
                                                            <span style="color: #666; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Desvio ID</span><br>
                                                            <span style="color: #6BBF47; font-size: 24px; font-weight: 700;">{numero_desvio}</span>
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>
                                        </tr>
                                    </table>

                                    <!-- Meta Info -->
                                    <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom: 25px;">
                                        <tr>
                                            <td style="padding: 10px 0; border-bottom: 1px solid #eee;">
                                                <span style="color: #999; font-size: 13px;">👤 Alterado por:</span>
                                                <span style="color: #333; font-size: 14px; font-weight: 500; margin-left: 10px;">{alterado_por}</span>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td style="padding: 10px 0;">
                                                <span style="color: #999; font-size: 13px;">📅 Data/Hora:</span>
                                                <span style="color: #333; font-size: 14px; font-weight: 500; margin-left: 10px;">{datetime.now().strftime('%d/%m/%Y às %H:%M')}</span>
                                            </td>
                                        </tr>
                                    </table>

                                    <!-- Changes Table -->
                                    <h3 style="margin: 0 0 15px; color: #333; font-size: 16px; font-weight: 600;">
                                        📝 Campos Alterados
                                    </h3>
                                    <table width="100%" cellpadding="0" cellspacing="0" style="border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;">
                                        <tr style="background-color: #f8f9fa;">
                                            <th style="padding: 12px 15px; text-align: left; font-size: 12px; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #e0e0e0;">Campo</th>
                                            <th style="padding: 12px 15px; text-align: left; font-size: 12px; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #e0e0e0;">Antes</th>
                                            <th style="padding: 12px 15px; text-align: left; font-size: 12px; color: #666; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #e0e0e0;">Depois</th>
                                        </tr>
                                        {alteracoes_html}
                                    </table>

                                </td>
                            </tr>

                            <!-- Footer -->
                            <tr>
                                <td style="background-color: #f8f9fa; padding: 25px 40px; text-align: center; border-top: 1px solid #eee;">
                                    <p style="margin: 0; color: #999; font-size: 12px;">
                                        Este é um email automático do sistema Portal Pesquisa Clínica.<br>
                                        Por favor, não responda a este email.
                                    </p>
                                    <p style="margin: 15px 0 0; color: #6BBF47; font-size: 11px; font-weight: 600;">
                                        © {datetime.now().year} Synvia
                                    </p>
                                </td>
                            </tr>

                        </table>
                    </td>
                </tr>
            </table>
        </body>
        </html>
        """

        # Envia para cada destinatário
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender, password)

            for destinatario in destinatarios:
                msg = MIMEMultipart('alternative')
                msg['Subject'] = assunto
                msg['From'] = sender
                msg['To'] = destinatario

                msg.attach(MIMEText(corpo_html, 'html'))

                server.sendmail(sender, destinatario, msg.as_string())

        logging.info(f"Email de notificação enviado para {len(destinatarios)} destinatário(s)")
        return True

    except Exception as e:
        logging.error(f"Erro ao enviar email de notificação: {e}")
        return False


# -------------------------------------------------
# Constantes de Perfis e Permissões
# -------------------------------------------------
PERFIS_DISPONIVEIS = ["Administrador", "Usuário"]

# Campos editáveis por Monitora (26 campos operacionais)
CAMPOS_MONITORA = [
    'participante', 'data_ocorrido', 'formulario_status', 'identificacao_desvio',
    'centro', 'visita', 'causa_raiz', 'acao_preventiva', 'acao_corretiva',
    'importancia', 'data_identificacao_texto', 'categoria', 'subcategoria',
    'codigo', 'recorrencia', 'num_ocorrencia_previa', 'escopo',
    'prazo_escalonamento', 'data_escalonamento', 'atendeu_prazos_report',
    'avaliacao_investigador', 'formulario_arquivado', 'data_submissao_cep',
    'data_finalizacao', 'populacao'
]

# Campos editáveis por Gerente/Gerente de Projetos (campos de análise/avaliação)
CAMPOS_GERENTE = [
    'descricao_desvio',           # Descrição do desvio
    'causa_raiz',                 # Reporte de causa raiz
    'acao_preventiva',            # Reporte de ação preventiva
    'acao_corretiva',             # Reporte de ação corretiva
    'importancia',                # Maior ou Menor
    'avaliacao_gerente_medico',   # Avaliação Gerente Médico
    'avaliacao_investigador',     # Avaliação Investigador
]

# Campos que nunca podem ser editados (controle do sistema)
CAMPOS_SISTEMA = [
    'id', 'numero_desvio_estudo', 'status', 'criado_por_nome',
    'criado_por_email', 'atualizado_por', 'data_atualizacao', 'row_version', 'url_anexo'
]


def get_campos_editaveis_por_perfil(perfil: str) -> list[str]:
    """Retorna lista de campos editáveis baseado no perfil do usuário."""
    if perfil == "Administrador":
        # Administrador pode editar todos os campos (monitora + gerente)
        return list(set(CAMPOS_MONITORA + CAMPOS_GERENTE))
    elif perfil == "Usuário":
        # Usuário pode editar os campos operacionais (como monitora)
        return CAMPOS_MONITORA
    else:
        return []  # Perfil desconhecido não edita nada


def get_campos_nao_editaveis_para_display(perfil: str) -> list[str]:
    """Retorna lista de nomes de colunas (display) que não podem ser editados."""
    campos_editaveis = get_campos_editaveis_por_perfil(perfil)
    rename_map = get_column_rename_map()

    # Todos os campos que NÃO estão na lista de editáveis
    campos_nao_editaveis = []
    for snake, display in rename_map.items():
        if snake not in campos_editaveis:
            campos_nao_editaveis.append(display)

    return campos_nao_editaveis


# -------------------------------------------------
# Funções de dados: Admin - Usuários
# -------------------------------------------------
@st.cache_data(ttl=30)
def load_usuarios() -> pd.DataFrame:
    """Carrega lista de todos os usuários."""
    try:
        conn = get_connection()
        query = """
            SELECT id, nome, email, cargo, perfil
            FROM usuarios
            ORDER BY perfil, nome
        """
        df = pd.read_sql_query(query, conn)
        conn.close()
        return df
    except Exception as e:
        st.error(f"Erro ao carregar usuários: {e}")
        return pd.DataFrame()


def criar_usuario(nome: str, email: str, cargo: str, perfil: str) -> bool:
    """Cadastra um novo usuário."""
    try:
        conn = get_connection()
        cursor = conn.cursor()
        # Verifica se já existe
        cursor.execute("SELECT 1 FROM usuarios WHERE LOWER(email) = %s", (email.lower(),))
        if cursor.fetchone():
            st.warning("Este email já está cadastrado.")
            cursor.close()
            conn.close()
            return False
        cursor.execute(
            """INSERT INTO usuarios (nome, email, cargo, perfil)
               VALUES (%s, %s, %s, %s)""",
            (nome.strip(), email.lower(), cargo, perfil)
        )
        conn.commit()
        cursor.close()
        conn.close()
        return True
    except Exception as e:
        st.error(f"Erro ao criar usuário: {e}")
        return False


def atualizar_usuario(user_id: int, nome: str, cargo: str, perfil: str, current_user_email: str) -> bool:
    """Atualiza um usuário existente."""
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute(
            """UPDATE usuarios SET nome = %s, cargo = %s, perfil = %s
               WHERE id = %s""",
            (nome.strip(), cargo, perfil, user_id)
        )
        conn.commit()
        cursor.close()
        conn.close()
        return True
    except Exception as e:
        st.error(f"Erro ao atualizar usuário: {e}")
        return False


def remover_usuario(user_id: int, current_user_email: str) -> bool:
    """Remove um usuário."""
    try:
        conn = get_connection()
        cursor = conn.cursor()
        # Verifica se está tentando remover a si mesmo
        cursor.execute("SELECT email FROM usuarios WHERE id = %s", (user_id,))
        row = cursor.fetchone()
        if row and row[0].lower() == current_user_email.lower():
            st.error("Você não pode remover seu próprio acesso.")
            cursor.close()
            conn.close()
            return False
        cursor.execute("DELETE FROM usuarios WHERE id = %s", (user_id,))
        conn.commit()
        cursor.close()
        conn.close()
        return True
    except Exception as e:
        st.error(f"Erro ao remover usuário: {e}")
        return False


# -------------------------------------------------
# Tela: Meus Estudos (Home) - Cards em Grid
# -------------------------------------------------
def render_meus_estudos(user_email: str):
    col_title, col_reload = st.columns([6, 1])
    with col_title:
        st.title("🧪 Portal Pesquisa Clínica")
    with col_reload:
        st.write("")
        if st.button("🔄", key="reload_estudos", help="Atualizar lista de estudos"):
            load_estudos_do_monitor.clear()
            load_estudos_do_gerente_medico.clear()
            is_gerente_medico.clear()
            st.rerun()

    with st.spinner("Carregando estudos..."):
        # Verifica se o usuário é um Gerente Médico
        usuario_eh_gm = is_gerente_medico(user_email)

        # Carrega estudos como monitor
        df_monitor = load_estudos_do_monitor(user_email)

        # Se for GM, carrega também os estudos como GM
        if usuario_eh_gm:
            df_gm = load_estudos_do_gerente_medico(user_email)
            # Combina os dataframes removendo duplicatas (caso seja monitor e GM no mesmo estudo)
            df = pd.concat([df_monitor, df_gm]).drop_duplicates(subset=['id']).reset_index(drop=True)
        else:
            df = df_monitor

    if df.empty:
        st.info("Você não está alocado em nenhum estudo ativo no momento.")
        return

    # Adiciona coluna de pendências para cada estudo
    df['pendencias'] = df['id'].apply(contar_pendencias_do_estudo)

    # Ordena: primeiro os com pendências (decrescente), depois os sem
    df = df.sort_values(by='pendencias', ascending=False)

    # Conta total de pendências
    total_pendencias = df['pendencias'].sum()
    if total_pendencias > 0:
        st.caption(f"{len(df)} estudo(s) ativo(s) • 🔴 {total_pendencias} pendência(s) total")
    else:
        st.caption(f"{len(df)} estudo(s) ativo(s)")

    # Grid de cards (3 colunas)
    cols_per_row = 3
    rows = [df.iloc[i:i + cols_per_row] for i in range(0, len(df), cols_per_row)]

    for row_data in rows:
        cols = st.columns(cols_per_row)
        for idx, (_, estudo) in enumerate(row_data.iterrows()):
            with cols[idx]:
                with st.container(border=True):
                    st.markdown(f"**{estudo['codigo']}**")
                    st.caption(estudo['nome'])

                    # Exibe status de pendências
                    pendencias = estudo['pendencias']
                    if pendencias > 0:
                        st.markdown(f"🔴 **{pendencias}** pendência(s)")
                    else:
                        st.markdown("🟢 Sem pendências")

                    if st.button("Entrar", key=f"entrar_{estudo['id']}", use_container_width=True):
                        st.session_state["estudo_ativo_id"] = estudo['id']
                        st.session_state["pagina_estudo"] = "menu"
                        st.rerun()


# -------------------------------------------------
# Tela: Ver Desvios do Estudo
# -------------------------------------------------
def format_date_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Formata colunas de data para DD/MM/YYYY."""
    date_cols = ['data_ocorrido', 'data_escalonamento', 'data_submissao_cep',
                 'data_finalizacao', 'data_atualizacao']
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%d/%m/%Y')
            df[col] = df[col].fillna('')
    return df


def get_column_rename_map() -> dict:
    """Retorna mapeamento de nomes de colunas para exibição."""
    return {
        'numero_desvio_estudo': 'ID',
        'status': 'Status',
        'participante': 'Participante',
        'data_ocorrido': 'Data Ocorrido',
        'formulario_status': 'Formulário',
        'identificacao_desvio': 'Identificação',
        'centro': 'Centro',
        'visita': 'Visita',
        'descricao_desvio': 'Descrição',
        'causa_raiz': 'Causa Raiz',
        'acao_preventiva': 'Ação Preventiva',
        'acao_corretiva': 'Ação Corretiva',
        'importancia': 'Importância',
        'data_identificacao_texto': 'Data Identificação',
        'categoria': 'Categoria',
        'subcategoria': 'Subcategoria',
        'codigo': 'Código',
        'escopo': 'Escopo',
        'avaliacao_gerente_medico': 'Avaliação Gerente Médico',
        'avaliacao_investigador': 'Avaliação Investigador',
        'formulario_arquivado': 'Formulário Arquivado',
        'recorrencia': 'Recorrência',
        'num_ocorrencia_previa': 'Nº Ocorrência Prévia',
        'prazo_escalonamento': 'Prazo Escalonamento',
        'data_escalonamento': 'Data Escalonamento',
        'atendeu_prazos_report': 'Atendeu Prazos Report',
        'populacao': 'População',
        'data_submissao_cep': 'Data Submissão CEP',
        'data_finalizacao': 'Data Finalização',
        'criado_por_nome': 'Criado Por',
        'criado_por_email': 'Email Criador',
        'atualizado_por': 'Atualizado Por',
        'data_atualizacao': 'Data Atualização',
        'url_anexo': 'Anexo',
    }


def render_desvios_estudo(estudo: dict, display_name: str, user_email: str):
    # Obtém perfil do usuário para aplicar permissões
    perfil_usuario = get_user_perfil(user_email) or "Visualizador"
    campos_editaveis = get_campos_editaveis_por_perfil(perfil_usuario)

    # Força reload se solicitado
    cache_key = f"desvios_df_{estudo['id']}"
    cache_key_orig = f"desvios_df_orig_{estudo['id']}"

    # Barra de controles
    col_filtro, col_reload = st.columns([3, 1])
    with col_filtro:
        filtro_status = st.selectbox(
            "Filtrar por status",
            ["Todos", "Novo", "Modificado", "Avaliado"],
            index=0,
            label_visibility="collapsed",
        )
    with col_reload:
        if st.button("🔄 Atualizar", use_container_width=True):
            st.session_state.pop(cache_key, None)
            st.session_state.pop(cache_key_orig, None)
            load_desvios_do_estudo.clear()
            st.rerun()

    # Carrega dados
    if cache_key not in st.session_state:
        with st.spinner("Carregando desvios..."):
            df = load_desvios_do_estudo(estudo['id'])
            st.session_state[cache_key] = df
            st.session_state[cache_key_orig] = df.copy()

    df_full = st.session_state[cache_key]

    if df_full.empty:
        st.info("Nenhum desvio cadastrado neste estudo.")
        return

    # Aplica filtro de status
    if filtro_status != "Todos":
        df_filtrado = df_full[df_full['status'] == filtro_status]
    else:
        df_filtrado = df_full

    if df_filtrado.empty:
        st.info(f"Nenhum desvio com status '{filtro_status}'.")
        return

    # Contador de resultados
    st.caption(f"{len(df_filtrado)} desvio(s) encontrado(s)")

    # Tabela resumida
    colunas_tabela = ["numero_desvio_estudo", "status", "participante", "centro", "visita", "importancia", "descricao_desvio"]
    df_tabela = df_filtrado[colunas_tabela].copy()
    df_tabela.columns = ["ID", "Status", "Participante", "Centro", "Visita", "Importância", "Descrição"]
    df_tabela["Descrição"] = df_tabela["Descrição"].apply(
        lambda x: (x[:60] + "...") if x and len(str(x)) > 60 else x
    )

    st.dataframe(df_tabela, use_container_width=True, hide_index=True)

    # Seção de edição
    st.subheader("Selecione um desvio para editar")

    # Seletor de desvio
    desvios_list = df_filtrado.to_dict('records')
    opcoes_desvio = {str(d['numero_desvio_estudo']): d for d in desvios_list}

    desvio_sel_key = st.selectbox(
        "Selecione o desvio:",
        options=list(opcoes_desvio.keys()),
        index=None,
        placeholder="Selecione o ID do desvio...",
        label_visibility="collapsed",
        key="select_desvio_editar",
    )

    if not desvio_sel_key:
        st.info("Selecione um desvio na lista acima para visualizar os detalhes e editar.")
        return

    desvio = opcoes_desvio[desvio_sel_key]

    st.markdown("")
    st.markdown("")

    # Exibe detalhes do desvio em cards
    exibir_detalhes_desvio(desvio)

    # Formulário de edição
    st.subheader("Editar campos")
    st.caption(f"Perfil: {perfil_usuario} ({len(campos_editaveis)} campos editáveis)")

    with st.form(f"form_editar_desvio_{desvio['id']}"):
        col1, col2 = st.columns(2)

        # Campos editáveis pela Monitora
        with col1:
            participante = st.text_input("Participante", value=desvio['participante'] or "", disabled='participante' not in campos_editaveis)
            centro = st.text_input("Centro", value=desvio['centro'] or "", disabled='centro' not in campos_editaveis)
            visita = st.text_input("Visita", value=desvio['visita'] or "", disabled='visita' not in campos_editaveis)
            identificacao = st.text_input("Identificação", value=desvio['identificacao_desvio'] or "", disabled='identificacao_desvio' not in campos_editaveis)
            formulario_status = st.selectbox("Formulário", ["", "Sim", "Pendente", "N/A"],
                index=["", "Sim", "Pendente", "N/A"].index(desvio['formulario_status']) if desvio['formulario_status'] in ["", "Sim", "Pendente", "N/A"] else 0,
                disabled='formulario_status' not in campos_editaveis)

        with col2:
            importancia = st.selectbox("Importância", ["", "Maior", "Menor"],
                index=["", "Maior", "Menor"].index(desvio['importancia']) if desvio['importancia'] in ["", "Maior", "Menor"] else 0,
                disabled='importancia' not in campos_editaveis)
            categoria = st.selectbox("Categoria", ["", "Avaliações", "Consentimento informado", "Procedimentos", "PSI", "Segurança", "Outros"],
                index=["", "Avaliações", "Consentimento informado", "Procedimentos", "PSI", "Segurança", "Outros"].index(desvio['categoria']) if desvio['categoria'] in ["", "Avaliações", "Consentimento informado", "Procedimentos", "PSI", "Segurança", "Outros"] else 0,
                disabled='categoria' not in campos_editaveis)
            escopo = st.selectbox("Escopo", ["", "Protocolo", "GCP"],
                index=["", "Protocolo", "GCP"].index(desvio['escopo']) if desvio['escopo'] in ["", "Protocolo", "GCP"] else 0,
                disabled='escopo' not in campos_editaveis)
            recorrencia = st.selectbox("Recorrência", ["", "Recorrente", "Não Recorrente", "Isolado"],
                index=["", "Recorrente", "Não Recorrente", "Isolado"].index(desvio['recorrencia']) if desvio['recorrencia'] in ["", "Recorrente", "Não Recorrente", "Isolado"] else 0,
                disabled='recorrencia' not in campos_editaveis)
            arquivado = st.selectbox("Formulário Arquivado?", ["", "Sim", "Não", "N/A"],
                index=["", "Sim", "Não", "N/A"].index(desvio['formulario_arquivado']) if desvio['formulario_arquivado'] in ["", "Sim", "Não", "N/A"] else 0,
                disabled='formulario_arquivado' not in campos_editaveis)

        # Campos de texto maiores
        descricao = st.text_area("Descrição do Desvio", value=desvio['descricao_desvio'] or "", disabled='descricao_desvio' not in campos_editaveis)

        col3, col4 = st.columns(2)
        with col3:
            causa_raiz = st.text_area("Causa Raiz", value=desvio['causa_raiz'] or "", disabled='causa_raiz' not in campos_editaveis)
            acao_preventiva = st.text_area("Ação Preventiva", value=desvio['acao_preventiva'] or "", disabled='acao_preventiva' not in campos_editaveis)
        with col4:
            acao_corretiva = st.text_area("Ação Corretiva", value=desvio['acao_corretiva'] or "", disabled='acao_corretiva' not in campos_editaveis)
            avaliacao_investigador = st.text_area("Avaliação Investigador", value=desvio['avaliacao_investigador'] or "", disabled='avaliacao_investigador' not in campos_editaveis)

        submitted = st.form_submit_button("💾 Salvar alterações", type="primary", use_container_width=True)

    if submitted:
        salvar_edicao_desvio(
            desvio_id=desvio['id'],
            row_version=desvio['row_version'],
            estudo_id=estudo['id'],
            display_name=display_name,
            campos_editaveis=campos_editaveis,
            novos_valores={
                'participante': participante,
                'centro': centro,
                'visita': visita,
                'identificacao_desvio': identificacao,
                'formulario_status': formulario_status,
                'importancia': importancia,
                'categoria': categoria,
                'escopo': escopo,
                'recorrencia': recorrencia,
                'formulario_arquivado': arquivado,
                'descricao_desvio': descricao,
                'causa_raiz': causa_raiz,
                'acao_preventiva': acao_preventiva,
                'acao_corretiva': acao_corretiva,
                'avaliacao_investigador': avaliacao_investigador,
            },
            valores_originais=desvio
        )

    # Botão de excluir desvio (soft delete) - apenas para Administradores
    if perfil_usuario == "Administrador":
        st.markdown("")
        st.markdown("---")
        st.markdown("### ⚠️ Zona de Perigo")

        col_del1, col_del2 = st.columns([3, 1])
        with col_del1:
            st.caption(f"Excluir permanentemente o desvio #{desvio['numero_desvio_estudo']}. Esta ação não pode ser desfeita.")
        with col_del2:
            if st.button("🗑️ Excluir Desvio", key=f"btn_del_desvio_{desvio['id']}", type="secondary", use_container_width=True):
                st.session_state[f"confirmar_exclusao_{desvio['id']}"] = True

        # Modal de confirmação
        if st.session_state.get(f"confirmar_exclusao_{desvio['id']}", False):
            with st.container(border=True):
                st.warning(f"Tem certeza que deseja excluir o desvio #{desvio['numero_desvio_estudo']}?")
                col_conf1, col_conf2 = st.columns(2)
                with col_conf1:
                    if st.button("✅ Sim, excluir", key=f"confirma_del_{desvio['id']}", type="primary", use_container_width=True):
                        if soft_delete_desvio(desvio['id'], estudo['id'], display_name):
                            st.success("Desvio excluído com sucesso!")
                            st.session_state.pop(f"confirmar_exclusao_{desvio['id']}", None)
                            st.session_state.pop(f"desvios_df_{estudo['id']}", None)
                            st.session_state.pop(f"desvios_df_orig_{estudo['id']}", None)
                            st.rerun()
                with col_conf2:
                    if st.button("❌ Cancelar", key=f"cancela_del_{desvio['id']}", use_container_width=True):
                        st.session_state.pop(f"confirmar_exclusao_{desvio['id']}", None)
                        st.rerun()


def exibir_detalhes_desvio(desvio: dict):
    """Exibe os detalhes do desvio em formato de cards somente leitura"""
    # Linha 1: ID, Status, Importância
    col1, col2, col3 = st.columns(3)

    with col1:
        with st.container(border=True):
            st.caption("ID do Desvio")
            st.markdown(f"**#{desvio['numero_desvio_estudo']}**")

    with col2:
        with st.container(border=True):
            st.caption("Status")
            status = desvio['status'] or '-'
            if status == 'Avaliado':
                st.success(status)
            elif status == 'Novo':
                st.info(status)
            else:
                st.markdown(f"**{status}**")

    with col3:
        with st.container(border=True):
            st.caption("Importância")
            importancia = desvio['importancia'] or '-'
            if importancia == 'Maior':
                st.error(importancia)
            else:
                st.markdown(f"**{importancia}**")

    # Linha 2: Participante, Centro, Visita
    col4, col5, col6 = st.columns(3)

    with col4:
        with st.container(border=True):
            st.caption("Participante")
            st.markdown(f"**{desvio['participante'] or '-'}**")

    with col5:
        with st.container(border=True):
            st.caption("Centro")
            st.markdown(f"**{desvio['centro'] or '-'}**")

    with col6:
        with st.container(border=True):
            st.caption("Visita")
            st.markdown(f"**{desvio['visita'] or '-'}**")

    # Linha 3: Identificação e Anexo
    col7, col8 = st.columns(2)

    with col7:
        with st.container(border=True):
            st.caption("Identificação do Desvio")
            st.markdown(f"**{desvio['identificacao_desvio'] or '-'}**")

    with col8:
        with st.container(border=True):
            st.caption("Anexo")
            if desvio.get('url_anexo'):
                st.markdown(f"[📎 Ver anexo]({desvio['url_anexo']})")
            else:
                st.markdown("**-**")


def salvar_edicao_desvio(desvio_id: int, row_version, estudo_id: int, display_name: str,
                         campos_editaveis: list, novos_valores: dict, valores_originais: dict):
    """Salva as alterações de um desvio específico"""
    try:
        conn = get_connection()
        cursor = conn.cursor()

        # Filtra apenas campos que o usuário pode editar e que foram alterados
        campos_alterados = []
        alteracoes_detalhadas = []  # Lista de {'campo', 'valor_antigo', 'valor_novo'}
        valores = []

        # Campos que possuem coluna equivalente em inglês
        campos_com_en = {
            'formulario_status': 'formulario_status_en',
            'importancia': 'importancia_en',
            'recorrencia': 'recorrencia_en',
            'escopo': 'escopo_en',
            'atendeu_prazos_report': 'atendeu_prazos_report_en',
            'formulario_arquivado': 'formulario_arquivado_en',
            'prazo_escalonamento': 'prazo_escalonamento_en',
            'populacao': 'populacao_en',
        }

        for campo, novo_valor in novos_valores.items():
            if campo in campos_editaveis:
                valor_original = valores_originais.get(campo)
                if str(novo_valor or '') != str(valor_original or ''):
                    campos_alterados.append(campo)
                    valores.append(novo_valor if novo_valor else None)
                    # Se o campo tem versão em inglês, adiciona também
                    if campo in campos_com_en:
                        campos_alterados.append(campos_com_en[campo])
                        valores.append(traduzir_valor_para_ingles(novo_valor) if novo_valor else None)
                    # Registra log
                    registrar_log(cursor, desvio_id, estudo_id, display_name, campo, valor_original, novo_valor)
                    # Guarda para o email
                    alteracoes_detalhadas.append({
                        'campo': campo,
                        'valor_antigo': valor_original,
                        'valor_novo': novo_valor
                    })

        if not campos_alterados:
            st.info("Nenhuma alteração detectada.")
            cursor.close()
            conn.close()
            return

        # Monta UPDATE
        set_clause = ", ".join([f"{campo} = %s" for campo in campos_alterados])
        set_clause += ", atualizado_por = %s, data_atualizacao = NOW(), status = 'Modificado', status_en = 'Modified'"
        valores.append(display_name)
        valores.append(desvio_id)
        valores.append(row_version)

        sql = f"UPDATE desvios SET {set_clause} WHERE id = %s AND xmin = %s::xid"

        cursor.execute(sql, valores)

        if cursor.rowcount == 0:
            st.warning("Conflito detectado! Os dados foram alterados por outro usuário. Atualize a página.")
        else:
            conn.commit()
            st.success(f"Desvio atualizado! ({len(campos_alterados)} campo(s) alterado(s))")
            # Limpa cache
            st.session_state.pop(f"desvios_df_{estudo_id}", None)
            st.session_state.pop(f"desvios_df_orig_{estudo_id}", None)
            load_desvios_do_estudo.clear()

            # Envia notificação por email
            try:
                cursor.execute("SELECT numero_desvio_estudo FROM desvios WHERE id = %s", (desvio_id,))
                numero_desvio = cursor.fetchone()[0] or desvio_id
                estudo = get_estudo_by_id(estudo_id)
                if estudo:
                    enviar_email_notificacao_desvio(
                        estudo_id=estudo_id,
                        estudo_codigo=estudo['codigo'],
                        estudo_nome=estudo['nome'],
                        desvio_id=desvio_id,
                        numero_desvio=numero_desvio,
                        alteracoes=alteracoes_detalhadas,
                        alterado_por=display_name
                    )
            except Exception as email_err:
                logging.error(f"Erro ao enviar email de notificação: {email_err}")

        cursor.close()
        conn.close()

    except Exception as e:
        st.error(f"Erro ao salvar: {e}")


def convert_to_str(val):
    """Converte valores para string (campos varchar do banco)."""
    import numpy as np
    if pd.isna(val):
        return None
    if isinstance(val, (np.integer, np.floating, int, float)):
        return str(int(val) if isinstance(val, (np.integer, int)) or (isinstance(val, float) and val.is_integer()) else val)
    if isinstance(val, (np.bool_,)):
        return str(bool(val))
    return val


def convert_to_int(val):
    """Converte valores numpy para int Python nativo."""
    import numpy as np
    if pd.isna(val):
        return None
    if isinstance(val, (np.integer,)):
        return int(val)
    return val


def registrar_log(cursor, desvio_id: int, estudo_id: int, usuario: str, campo: str, valor_antigo, valor_novo):
    """Registra uma alteração na tabela de logs."""
    cursor.execute(
        """INSERT INTO desvios_log (desvio_id, estudo_id, usuario, campo, valor_antigo, valor_novo)
           VALUES (%s, %s, %s, %s, %s, %s)""",
        (desvio_id, estudo_id, usuario, campo, str(valor_antigo) if valor_antigo else None, str(valor_novo) if valor_novo else None)
    )


def soft_delete_desvio(desvio_id: int, estudo_id: int, deleted_by: str) -> bool:
    """Realiza soft delete de um desvio (marca como excluído sem remover do banco)."""
    try:
        conn = get_connection()
        cursor = conn.cursor()

        # Atualiza o desvio com deleted_at e deleted_by
        cursor.execute(
            """UPDATE desvios
               SET deleted_at = NOW(), deleted_by = %s
               WHERE id = %s""",
            (deleted_by, desvio_id)
        )

        # Registra no log
        registrar_log(cursor, desvio_id, estudo_id, deleted_by, "EXCLUSÃO", "Ativo", "Excluído")

        conn.commit()
        cursor.close()
        conn.close()

        # Limpa cache
        load_desvios_do_estudo.clear()

        return True
    except Exception as e:
        st.error(f"Erro ao excluir desvio: {e}")
        return False


def save_desvios_changes(edited_df, original_df, display_name, estudo_id):
    """Salva alterações dos desvios com controle de concorrência e registra logs."""
    if "id" not in edited_df.columns:
        st.error("Coluna 'id' obrigatória.")
        return

    edited_idx = edited_df.set_index("id")
    original_idx = original_df.set_index("id")

    data_cols = [c for c in edited_df.columns if c not in ("id", "row_version")]

    try:
        diffs = edited_idx[data_cols].ne(original_idx[data_cols]).any(axis=1)
    except:
        diffs = pd.Series(True, index=edited_idx.index)

    ids_alterados = edited_idx.index[diffs].tolist()

    if not ids_alterados:
        st.info("Nenhuma alteração detectada.")
        return

    # Colunas editáveis (exclui campos de controle)
    campos_editaveis = [c for c in data_cols if c not in ("criado_por_nome", "criado_por_email", "atualizado_por", "status", "data_atualizacao", "numero_desvio_estudo")]

    # Colunas para UPDATE (campos editáveis + campos de controle)
    update_cols = campos_editaveis.copy()
    update_cols.append("atualizado_por")
    update_cols.append("data_atualizacao")
    update_cols.append("status")

    conflitos, atualizados = [], 0

    try:
        conn = get_connection()
        cursor = conn.cursor()

        for row_id in ids_alterados:
            row_edit = edited_idx.loc[row_id]
            row_orig = original_idx.loc[row_id]

            # Registra logs para cada campo alterado
            for campo in campos_editaveis:
                valor_orig = row_orig.get(campo)
                valor_novo = row_edit.get(campo)
                # Verifica se houve alteração real no campo
                if str(valor_orig) != str(valor_novo):
                    registrar_log(cursor, convert_to_int(row_id), estudo_id, display_name, campo, valor_orig, valor_novo)

            values = [convert_to_str(row_edit[c]) for c in campos_editaveis]
            values.append(display_name)  # atualizado_por
            values.append('NOW()')  # data_atualizacao - será substituído abaixo
            values.append('Modificado')  # status sempre muda para 'Modificado' quando monitor edita
            values.append(convert_to_int(row_id))
            values.append(convert_to_int(row_orig["row_version"]))

            # Ajusta SQL para usar NOW() diretamente
            set_clause_now = ", ".join(
                f"{col} = NOW()" if col == "data_atualizacao" else f"{col} = %s"
                for col in update_cols
            )
            sql_now = f"UPDATE desvios SET {set_clause_now} WHERE id = %s AND xmin = %s::xid"

            # Remove o placeholder de data_atualizacao dos values
            values_sem_data = [v for i, v in enumerate(values) if update_cols[i] != "data_atualizacao"]

            cursor.execute(sql_now, values_sem_data)
            if cursor.rowcount == 0:
                conflitos.append(row_id)
            else:
                atualizados += 1

        conn.commit()
        cursor.close()
        conn.close()

        if atualizados:
            st.success(f"{atualizados} registro(s) atualizado(s)! ✅")

            # Envia notificação por email (consolidada para edições em lote)
            try:
                estudo = get_estudo_by_id(estudo_id)
                if estudo:
                    enviar_email_notificacao_desvio(
                        estudo_id=estudo_id,
                        estudo_codigo=estudo['codigo'],
                        estudo_nome=estudo['nome'],
                        desvio_id=0,
                        numero_desvio=atualizados,  # Número de registros alterados
                        alteracoes=[{
                            'campo': 'Edição em Lote',
                            'valor_antigo': '-',
                            'valor_novo': f'{atualizados} desvio(s) modificado(s) via tabela'
                        }],
                        alterado_por=display_name
                    )
            except Exception as email_err:
                logging.error(f"Erro ao enviar email de notificação: {email_err}")

        if conflitos:
            st.warning(f"{len(conflitos)} registro(s) com conflito. Recarregue os dados.")

        # Limpa cache para recarregar
        st.session_state.pop(f"desvios_df_{estudo_id}", None)
        st.session_state.pop(f"desvios_df_orig_{estudo_id}", None)

    except Exception as e:
        st.error(f"Erro ao salvar: {e}")


# -------------------------------------------------
# Tela: Cadastrar Desvio
# -------------------------------------------------
def render_cadastro_desvio(estudo: dict, user_email: str, display_name: str):
    st.caption(f"Estudo: {estudo['codigo']} - {estudo['nome']}")

    # Primeiro: seleção de importância (fora do form para permitir rerun e mostrar upload)
    importancia = st.selectbox("Importância", ["", "Maior", "Menor"], key="sel_importancia")

    # Upload de imagem - só aparece se importância for "Maior"
    uploaded_file = None
    if importancia == "Maior":
        st.info("📎 Desvios de importância Maior requerem anexo de evidência.")
        uploaded_file = st.file_uploader(
            "Anexar imagem/documento",
            type=["png", "jpg", "jpeg", "pdf", "doc", "docx"],
            key="upload_evidencia"
        )

    # Form para os demais campos (não causa rerun ao preencher)
    with st.form("form_cadastro_desvio"):
        col1, col2 = st.columns(2)

        with col1:
            participante = st.text_input("Participante")
            centro = st.text_input("Centro")
            visita = st.text_input("Visita")
            identificacao = st.text_input("Identificação do Desvio")
            data_ocorrido = st.date_input("Data do Ocorrido", format="DD/MM/YYYY")
            data_identificacao = st.text_input("Data de Identificação", placeholder="Ex.: MOV01-2005 (01/01-07/01)")
            categoria = st.selectbox("Categoria", [
                "", "Avaliações", "Consentimento informado", "Procedimentos",
                "PSI", "Segurança", "Outros"
            ])
            subcategoria = st.selectbox("Subcategoria", [
                "", "Avaliações Perdidas ou não realizadas",
                "Avaliações realizadas fora da janela", "Desvios Recorrentes",
                "Visitas perdidas ou fora da janela", "Amostras Laboratoriais",
                "Critérios de Inclusão / Exclusão (Elegibilidade)", "Outros"
            ])
            codigo = st.selectbox("Código", ["", "A8", "A7", "O4"])
            escopo = st.selectbox("Escopo", ["", "Protocolo", "GCP"])
            prazo_escalonamento = st.selectbox("Prazo para Escalonamento", ["", "Imediata", "Mensal", "Padrão"])

        with col2:
            formulario = st.selectbox("Formulário", ["", "Sim", "Pendente", "N/A"])
            arquivado = st.selectbox("Formulário Arquivado (ISF e TFM)?", ["", "Sim", "Não", "N/A"])
            recorrencia = st.selectbox("Recorrência", ["", "Recorrente", "Não Recorrente", "Isolado"])
            ocorrencia_previa = st.number_input("N° Desvio Ocorrência Prévia", min_value=0, step=1)
            populacao = st.selectbox("População", ["", "Intenção de Tratar (ITT)", "Por Protocolo (PP)"])
            prazo_report = st.selectbox("Atendeu os Prazos de Report?", ["", "Sim", "Não"])
            motivo_nao_atendeu_prazo = st.text_area("Motivo (se não atendeu prazos)", placeholder="Preencha se não atendeu", height=150)
            data_escalonamento = st.date_input("Data de Escalonamento", format="DD/MM/YYYY")
            data_cep = st.date_input("Data de Submissão ao CEP", format="DD/MM/YYYY")
            data_finalizacao = st.date_input("Data de Finalização", format="DD/MM/YYYY")

        # Campos de texto maiores em grid 2x2
        col3, col4 = st.columns(2)
        with col3:
            descricao = st.text_area("Descrição do Desvio")
            causa_raiz = st.text_area("Causa Raiz")
        with col4:
            acao_preventiva = st.text_area("Ação Preventiva")
            acao_corretiva = st.text_area("Ação Corretiva")

        avaliacao_investigador = st.text_area("Avaliação Investigador Principal")

        submitted = st.form_submit_button("💾 Salvar Desvio", type="primary", use_container_width=True)

    if submitted:
        # Validação de campos obrigatórios
        campos_obrigatorios = {
            "Participante": participante,
            "Formulário": formulario,
            "Identificação do Desvio": identificacao,
            "Centro": centro,
            "Visita": visita,
            "Descrição do Desvio": descricao,
            "Causa Raiz": causa_raiz,
            "Ação Preventiva": acao_preventiva,
            "Ação Corretiva": acao_corretiva,
            "Importância": importancia,
            "Data de Identificação": data_identificacao,
            "Categoria": categoria,
            "Subcategoria": subcategoria,
            "Código": codigo,
            "Escopo": escopo,
            "Avaliação Investigador Principal": avaliacao_investigador,
            "Formulário Arquivado": arquivado,
            "Recorrência": recorrencia,
            "Prazo para Escalonamento": prazo_escalonamento,
            "Atendeu os Prazos de Report": prazo_report,
            "População": populacao,
        }

        campos_vazios = [nome for nome, valor in campos_obrigatorios.items() if not valor or valor == ""]

        # Valida anexo obrigatório para desvios "Maior"
        if importancia == "Maior" and not uploaded_file:
            campos_vazios.append("Anexo de evidência (obrigatório para desvios Maior)")

        # Valida motivo obrigatório se não atendeu prazos de report
        if prazo_report == "Não" and not motivo_nao_atendeu_prazo.strip():
            campos_vazios.append("Motivo (obrigatório quando não atendeu os prazos de report)")

        if campos_vazios:
            st.error("Preencha todos os campos obrigatórios:")
            for campo in campos_vazios:
                st.warning(f"⚠️ {campo}")
        else:
            try:
                # Obtém o próximo número de desvio para este estudo
                numero_desvio = get_proximo_numero_desvio(estudo['id'])

                # Upload para SharePoint se houver arquivo
                url_anexo = None
                if uploaded_file and importancia == "Maior":
                    with st.spinner("Enviando arquivo para o SharePoint..."):
                        file_content = uploaded_file.getvalue()
                        url_anexo = upload_to_sharepoint(
                            file_content,
                            uploaded_file.name,
                            estudo['codigo'],
                            numero_desvio
                        )
                        if url_anexo:
                            st.success(f"Arquivo enviado!")
                        else:
                            st.error("Falha no upload do arquivo. O desvio será salvo sem anexo.")

                conn = get_connection()
                cursor = conn.cursor()

                sql = """
                    INSERT INTO desvios (
                        estudo_id, numero_desvio_estudo, status, participante, data_ocorrido, formulario_status,
                        identificacao_desvio, centro, visita, descricao_desvio,
                        causa_raiz, acao_preventiva, acao_corretiva, importancia,
                        data_identificacao_texto, categoria, subcategoria, codigo,
                        escopo, avaliacao_gerente_medico, avaliacao_investigador,
                        formulario_arquivado, recorrencia, num_ocorrencia_previa,
                        prazo_escalonamento, data_escalonamento, atendeu_prazos_report,
                        motivo_nao_atendeu_prazo, populacao, data_submissao_cep, data_finalizacao,
                        criado_por_nome, criado_por_email, url_anexo,
                        status_en, formulario_status_en, importancia_en, recorrencia_en,
                        escopo_en, atendeu_prazos_report_en, formulario_arquivado_en,
                        prazo_escalonamento_en, populacao_en
                    ) VALUES (
                        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                        %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
                    )
                """

                values = (
                    estudo['id'], numero_desvio, 'Novo', participante, data_ocorrido, formulario,
                    identificacao, centro, visita, descricao, causa_raiz,
                    acao_preventiva, acao_corretiva, importancia,
                    data_identificacao, categoria, subcategoria, codigo,
                    escopo, None, avaliacao_investigador,  # avaliacao_gerente_medico preenchido pelo app externo
                    arquivado, recorrencia, int(ocorrencia_previa) if ocorrencia_previa else None,
                    prazo_escalonamento, data_escalonamento, prazo_report,
                    motivo_nao_atendeu_prazo.strip() if motivo_nao_atendeu_prazo else None,
                    populacao, data_cep, data_finalizacao,
                    display_name, user_email, url_anexo,
                    # Colunas em inglês
                    'New',  # status_en
                    traduzir_valor_para_ingles(formulario),  # formulario_status_en
                    traduzir_valor_para_ingles(importancia),  # importancia_en
                    traduzir_valor_para_ingles(recorrencia),  # recorrencia_en
                    traduzir_valor_para_ingles(escopo),  # escopo_en
                    traduzir_valor_para_ingles(prazo_report),  # atendeu_prazos_report_en
                    traduzir_valor_para_ingles(arquivado),  # formulario_arquivado_en
                    traduzir_valor_para_ingles(prazo_escalonamento),  # prazo_escalonamento_en
                    traduzir_valor_para_ingles(populacao),  # populacao_en
                )

                cursor.execute(sql, values)
                conn.commit()
                cursor.close()
                conn.close()

                st.success("Desvio cadastrado com sucesso!")

                # Limpa cache de desvios desse estudo
                st.session_state.pop(f"desvios_df_{estudo['id']}", None)
                st.session_state.pop(f"desvios_df_orig_{estudo['id']}", None)
                load_desvios_do_estudo.clear()

            except Exception as e:
                st.error(f"Erro ao cadastrar: {e}")


# -------------------------------------------------
# Telas placeholder (Admin e Relatórios)
# -------------------------------------------------
def render_painel_adm(pode_acessar: bool, user_email: str):
    st.subheader("🛠 Painel Administrativo")
    if not pode_acessar:
        st.error("Acesso restrito a Administradores.")
        return

    # Tabs para organizar as seções
    tab_estudos, tab_monitores, tab_gerentes, tab_admins = st.tabs([
        "📚 Gerenciar Estudos",
        "👥 Alocar Monitores",
        "🩺 Gerentes Médicos",
        "🔐 Usuários"
    ])

    # ----- TAB: Gerenciar Estudos -----
    with tab_estudos:
        with st.form("form_criar_estudo", clear_on_submit=True):
            col1, col2, col3 = st.columns([2, 3, 1])
            with col1:
                novo_codigo = st.text_input("Código", placeholder="EST-001")
            with col2:
                novo_nome = st.text_input("Nome do Estudo", placeholder="Estudo Fase III")
            with col3:
                st.write("")
                if st.form_submit_button("➕ Criar", type="primary", use_container_width=True):
                    if novo_codigo and novo_nome:
                        if criar_estudo(novo_codigo, novo_nome):
                            st.success(f"Estudo '{novo_codigo}' criado!")
                            load_todos_estudos.clear()
                            st.rerun()
                    else:
                        st.warning("Preencha código e nome.")

        st.markdown("")
        st.markdown("")
        st.markdown("")

        col_title, col_reload = st.columns([6, 1])
        with col_title:
            st.markdown("### Estudos Cadastrados")
        with col_reload:
            if st.button("🔄", key="reload_estudos_adm", help="Atualizar lista"):
                load_todos_estudos.clear()
                st.rerun()

        with st.spinner("Carregando estudos..."):
            df_estudos = load_todos_estudos()

        if not df_estudos.empty:
            col_filtro1, col_filtro2 = st.columns([3, 1])
            with col_filtro1:
                opcoes_estudos = ["Todos"] + [f"{row['codigo']} - {row['nome']}" for _, row in df_estudos.iterrows()]
                filtro_estudo = st.selectbox("Filtrar estudo", opcoes_estudos, key="filtro_estudos", label_visibility="collapsed")
            with col_filtro2:
                filtro_status = st.selectbox("Status", ["Todos", "Ativos", "Inativos"], key="filtro_status_estudos", label_visibility="collapsed")

            df_filtrado = df_estudos.copy()
            if filtro_estudo != "Todos":
                codigo_selecionado = filtro_estudo.split(" - ")[0]
                df_filtrado = df_filtrado[df_filtrado['codigo'] == codigo_selecionado]
            if filtro_status == "Ativos":
                df_filtrado = df_filtrado[df_filtrado['status'] == 'ativo']
            elif filtro_status == "Inativos":
                df_filtrado = df_filtrado[df_filtrado['status'] == 'inativo']

            for _, est in df_filtrado.iterrows():
                with st.expander(f"{'🟢' if est['status'] == 'ativo' else '🔴'} {est['codigo']} - {est['nome']}"):
                    gm_estudo = get_gerente_medico_do_estudo(est['id'])
                    if gm_estudo:
                        st.caption(f"🩺 {gm_estudo['nome']}")
                    else:
                        st.caption("🩺 _Nenhum GM alocado_")

                    with st.form(f"form_edit_estudo_{est['id']}"):
                        col1, col2, col3 = st.columns([2, 3, 2])
                        with col1:
                            edit_codigo = st.text_input("Código", value=est['codigo'], key=f"cod_{est['id']}")
                        with col2:
                            edit_nome = st.text_input("Nome", value=est['nome'], key=f"nome_{est['id']}")
                        with col3:
                            edit_status = st.selectbox("Status", ["ativo", "inativo"], index=0 if est['status'] == 'ativo' else 1, key=f"status_{est['id']}")
                        if st.form_submit_button("Salvar", use_container_width=True):
                            if atualizar_estudo(est['id'], edit_codigo, edit_nome, edit_status):
                                st.success("Atualizado!")
                                load_todos_estudos.clear()
                                st.rerun()

    # ----- TAB: Alocar Monitores -----
    with tab_monitores:
        df_estudos = load_todos_estudos()
        if df_estudos.empty:
            st.info("Nenhum estudo cadastrado.")
        else:
            opcoes_estudos = {f"{row['codigo']} - {row['nome']}": row['id'] for _, row in df_estudos.iterrows()}
            estudo_selecionado = st.selectbox("Estudo", options=list(opcoes_estudos.keys()), key="select_estudo_monitores")
            estudo_id_selecionado = opcoes_estudos[estudo_selecionado]

            col1, col2 = st.columns([3, 1])
            df_usuarios = load_usuarios()
            if not df_usuarios.empty:
                with col1:
                    opcoes_usuarios = {f"{row['nome']} ({row['email']})": row['email'] for _, row in df_usuarios.iterrows()}
                    usuario_selecionado = st.selectbox("Usuário", options=list(opcoes_usuarios.keys()), key="select_usuario_alocar", label_visibility="collapsed")
                    email_selecionado = opcoes_usuarios[usuario_selecionado]
                with col2:
                    if st.button("➕ Alocar", type="primary", use_container_width=True):
                        if alocar_monitor(estudo_id_selecionado, email_selecionado):
                            st.success("Monitor alocado!")
                            load_monitores_do_estudo.clear()
                            st.rerun()
            
            st.markdown("")
            st.markdown("")
            st.markdown("### Monitores Alocados")

            df_monitores = load_monitores_do_estudo(estudo_id_selecionado)
            if df_monitores.empty:
                st.caption("Nenhum monitor alocado neste estudo.")
            else:
                for _, mon in df_monitores.iterrows():
                    col1, col2, col3 = st.columns([4, 5, 1])
                    with col1:
                        st.write(mon['monitor_nome'] or "(sem nome)")
                    with col2:
                        st.caption(mon['monitor_email'])
                    with col3:
                        if st.button("🗑️", key=f"rem_mon_{mon['id']}"):
                            if remover_monitor(mon['id']):
                                load_monitores_do_estudo.clear()
                                st.rerun()

    # ----- TAB: Gerentes Médicos -----
    with tab_gerentes:
        st.caption("Acesso ao sistema externo de avaliação")

        with st.form("form_criar_gerente", clear_on_submit=True):
            col1, col2, col3, col4 = st.columns([3, 3, 3, 1])
            with col1:
                novo_nome_gm = st.text_input("Nome", placeholder="Nome completo")
            with col2:
                novo_email_gm = st.text_input("Email", placeholder="gerente@hospital.com")
            with col3:
                novo_patrocinador_gm = st.text_input("Patrocinador", placeholder="Patrocinador")
            with col4:
                st.write("")
                if st.form_submit_button("➕", type="primary", use_container_width=True):
                    if novo_nome_gm and novo_email_gm and novo_patrocinador_gm:
                        if criar_gerente_medico(novo_nome_gm, novo_email_gm, novo_patrocinador_gm):
                            load_gerentes_medicos.clear()
                            st.rerun()
                    else:
                        st.warning("Preencha todos os campos.")

        st.markdown("")
        st.markdown("")
        st.markdown("### Alocar Gerente Médico a Estudo")

        df_estudos = load_todos_estudos()
        df_gerentes = load_gerentes_medicos()

        if df_estudos.empty:
            st.info("Nenhum estudo cadastrado.")
        elif df_gerentes.empty:
            st.info("Nenhum gerente médico cadastrado. Cadastre um gerente acima para alocar.")
        else:
            col1, col2, col3 = st.columns([3, 3, 2])
            with col1:
                opcoes_estudos_gm = {f"{row['codigo']} - {row['nome']}": row['id'] for _, row in df_estudos.iterrows()}
                estudo_sel_gm = st.selectbox("Estudo", options=list(opcoes_estudos_gm.keys()), key="select_estudo_gerente")
                estudo_id_gm = opcoes_estudos_gm[estudo_sel_gm]
            with col2:
                opcoes_gerentes = {f"{row['nome']}": row['id'] for _, row in df_gerentes.iterrows()}
                gerente_sel = st.selectbox("Gerente Médico", options=list(opcoes_gerentes.keys()), key="select_gerente_alocar")
                gerente_id_sel = opcoes_gerentes[gerente_sel]
            with col3:
                gerente_atual = get_gerente_medico_do_estudo(estudo_id_gm)
                st.write("")
                col_a, col_b = st.columns(2)
                with col_a:
                    if st.button("Alocar", type="primary", use_container_width=True):
                        if alocar_gerente_medico(estudo_id_gm, gerente_id_sel):
                            st.rerun()
                with col_b:
                    if gerente_atual and st.button("Remover", use_container_width=True):
                        if remover_gerente_medico_do_estudo(estudo_id_gm):
                            st.rerun()

            if gerente_atual:
                st.caption(f"Atual: {gerente_atual['nome']}")
            else:
                st.caption("Nenhum GM alocado neste estudo")
        
        st.markdown("")
        st.markdown("")

        with st.expander(f"📋 Gerentes Médicos Cadastrados ({len(df_gerentes)})", expanded=False):
            if not df_gerentes.empty:
                # Selectbox para filtrar por nome
                opcoes_nomes = ["Todos"] + df_gerentes['nome'].tolist()
                filtro_gm = st.selectbox(
                    "Filtrar por nome",
                    opcoes_nomes,
                    key="filtro_gerente_medico",
                    label_visibility="collapsed"
                )

                # Filtra os gerentes
                df_filtrado = df_gerentes.copy()
                if filtro_gm != "Todos":
                    df_filtrado = df_filtrado[df_filtrado['nome'] == filtro_gm]

                st.markdown("")

                if df_filtrado.empty:
                    st.info("Nenhum gerente médico encontrado.")
                else:
                    for _, gm in df_filtrado.iterrows():
                        estudos_gm = get_estudos_do_gerente_medico_por_id(gm['id'])

                        with st.container(border=True):
                            col_info, col_acoes = st.columns([11, 1])

                            with col_info:
                                st.markdown(f"**🩺 {gm['nome']}**")

                                col_det1, col_det2 = st.columns(2)
                                with col_det1:
                                    st.caption(f"📧 {gm['email']}")
                                with col_det2:
                                    patrocinador = gm.get('patrocinador', '') or ''
                                    if patrocinador:
                                        st.caption(f"🏢 {patrocinador}")

                                if estudos_gm:
                                    estudos_badges = " ".join([f"`{e}`" for e in estudos_gm])
                                    st.markdown(f"📋 Estudos: {estudos_badges}")
                                else:
                                    st.caption("📋 _Nenhum estudo alocado_")

                            with col_acoes:
                                st.markdown("")
                                if st.button("🗑️", key=f"rem_gm_{gm['id']}", help="Remover gerente médico"):
                                    if remover_gerente_medico(gm['id']):
                                        load_gerentes_medicos.clear()
                                        st.rerun()
            else:
                st.info("Nenhum gerente médico cadastrado.")

    # ----- TAB: Usuários -----
    with tab_admins:
        with st.form("form_criar_usuario", clear_on_submit=True):
            col1, col2, col3, col4, col5 = st.columns([3, 3, 2, 2, 1])
            with col1:
                novo_nome = st.text_input("Nome", placeholder="Nome completo")
            with col2:
                novo_email = st.text_input("Email", placeholder="usuario@empresa.com")
            with col3:
                novo_cargo = st.text_input("Cargo", placeholder="Cargo")
            with col4:
                novo_perfil = st.selectbox("Perfil", PERFIS_DISPONIVEIS)
            with col5:
                st.write("")
                if st.form_submit_button("➕", type="primary", use_container_width=True):
                    if novo_nome and novo_email:
                        if criar_usuario(novo_nome, novo_email, novo_cargo, novo_perfil):
                            load_usuarios.clear()
                            st.rerun()
                    else:
                        st.warning("Preencha nome e email.")
        
        st.markdown("")
        st.markdown("")
        st.markdown("### Usuários Cadastrados")

        df_usuarios = load_usuarios()
        if not df_usuarios.empty:
            col_filtro1, col_filtro2 = st.columns([3, 1])
            with col_filtro1:
                opcoes_usuarios = ["Todos"] + [f"{row['nome']} ({row['email']})" for _, row in df_usuarios.iterrows()]
                filtro_usuario = st.selectbox("Filtrar", opcoes_usuarios, key="filtro_usuarios", label_visibility="collapsed")
            with col_filtro2:
                filtro_perfil = st.selectbox("Perfil", ["Todos"] + PERFIS_DISPONIVEIS, key="filtro_perfil_usuarios", label_visibility="collapsed")

            df_filtrado = df_usuarios.copy()
            if filtro_usuario != "Todos":
                email_selecionado = filtro_usuario.split("(")[-1].replace(")", "")
                df_filtrado = df_filtrado[df_filtrado['email'] == email_selecionado]
            if filtro_perfil != "Todos":
                df_filtrado = df_filtrado[df_filtrado['perfil'] == filtro_perfil]

            # Badges de perfil
            perfil_badges = {
                "Administrador": "👑",
                "Usuário": "👤"
            }

            for _, usr in df_filtrado.iterrows():
                badge = perfil_badges.get(usr['perfil'], "")
                with st.expander(f"{badge} {usr['nome']} - {usr['perfil']}"):
                    st.caption(f"📧 {usr['email']}")
                    with st.form(f"form_edit_user_{usr['id']}"):
                        col1, col2 = st.columns(2)
                        with col1:
                            edit_nome = st.text_input("Nome", value=usr['nome'], key=f"unome_{usr['id']}")
                            edit_cargo = st.text_input("Cargo", value=usr['cargo'] or "", key=f"ucargo_{usr['id']}")
                        with col2:
                            perfis_lista = PERFIS_DISPONIVEIS
                            perfil_index = perfis_lista.index(usr['perfil']) if usr['perfil'] in perfis_lista else 0
                            edit_perfil = st.selectbox(
                                "Perfil",
                                perfis_lista,
                                index=perfil_index,
                                key=f"uperfil_{usr['id']}"
                            )

                        col_save, col_del = st.columns([3, 1])
                        with col_save:
                            if st.form_submit_button("Salvar", use_container_width=True):
                                if atualizar_usuario(usr['id'], edit_nome, edit_cargo, edit_perfil, user_email):
                                    st.success("Usuário atualizado!")
                                    st.rerun()
                        with col_del:
                            if st.form_submit_button("🗑️ Remover", type="secondary", use_container_width=True):
                                if remover_usuario(usr['id'], user_email):
                                    st.success("Usuário removido!")
                                    st.rerun()


def render_relatorios():
    st.subheader("📑 Relatórios & Auditoria")
    st.write("Central de relatórios e audit log.")
    st.info("Em desenvolvimento...")


# -------------------------------------------------
# Autenticação Microsoft (MANTIDO)
# -------------------------------------------------
auth = MicrosoftAuth()

logged_in = create_login_page(auth)
if not logged_in:
    st.stop()

AuthManager.check_and_refresh_token(auth)

user = AuthManager.get_current_user() or {}
display_name = user.get("displayName", "Usuário")
user_email = (user.get("mail") or user.get("userPrincipalName") or "").lower()

# Obtém perfil do usuário e exibe no header da sidebar
perfil_usuario = get_user_perfil(user_email) or "Não definido"
create_user_header(perfil=perfil_usuario)

st.session_state["display_name"] = display_name
st.session_state["user_email"] = user_email

# Auto-cadastro no primeiro login (apenas @synvia.com)
if user_email and not st.session_state.get("_auto_cadastro_feito"):
    auto_cadastrar_usuario(user_email, display_name)
    st.session_state["_auto_cadastro_feito"] = True

if "pode_acessar_adm" not in st.session_state:
    st.session_state.pode_acessar_adm = pode_acessar_painel_adm(user_email)

pode_acessar_adm = st.session_state.pode_acessar_adm


# -------------------------------------------------
# Navegação Principal
# -------------------------------------------------

# Sidebar fixa com navegação global
with st.sidebar:
    st.markdown("---")

    if st.button("🏠 Meus Estudos", use_container_width=True):
        st.session_state.pop("estudo_ativo_id", None)
        st.session_state.pop("pagina_estudo", None)
        st.session_state["pagina_home"] = "meus_estudos"
        st.rerun()

    if st.button("📑 Relatórios", use_container_width=True):
        if st.session_state.get("estudo_ativo_id"):
            st.session_state["pagina_estudo"] = "relatorios"
        else:
            st.session_state["pagina_home"] = "relatorios"
        st.rerun()

    if pode_acessar_adm:
        if st.button("🛠 Painel Adm", use_container_width=True):
            if st.session_state.get("estudo_ativo_id"):
                st.session_state["pagina_estudo"] = "painel_adm"
            else:
                st.session_state["pagina_home"] = "painel_adm"
            st.rerun()


# Verifica se tem estudo selecionado
estudo_ativo_id = st.session_state.get("estudo_ativo_id")
estudo = None

if estudo_ativo_id:
    estudo = get_estudo_by_id(estudo_ativo_id)
    if not estudo:
        st.error("Estudo não encontrado.")
        st.session_state.pop("estudo_ativo_id", None)
        st.rerun()

# Renderiza conteúdo da página
pagina = st.session_state.get("pagina_estudo") if estudo else st.session_state.get("pagina_home", "meus_estudos")

# Páginas globais (independem do estudo selecionado)
if pagina == "relatorios":
    render_relatorios()
elif pagina == "painel_adm":
    render_painel_adm(pode_acessar_adm, user_email)
# Páginas relacionadas ao estudo
elif estudo:
    # Coleta todas as informações do estudo
    monitores = get_nomes_monitores_do_estudo(estudo['id'])
    gerente = get_gerente_medico_do_estudo(estudo['id'])
    patrocinador = get_patrocinador_do_estudo(estudo['id'])
    qtd_desvios = contar_desvios_do_estudo(estudo['id'])
    data_criacao = estudo.get('criado_em')
    if data_criacao:
        data_criacao_fmt = data_criacao.strftime('%d/%m/%Y') if hasattr(data_criacao, 'strftime') else str(data_criacao)[:10]
    else:
        data_criacao_fmt = '-'

    # Header com título e botão voltar
    col_title, col_back = st.columns([9, 1])
    with col_title:
        st.markdown(f"### 📂 {estudo['codigo']} - {estudo['nome']}")
    with col_back:
        if st.button("↩️ Voltar", use_container_width=True):
            st.session_state.pop("estudo_ativo_id", None)
            st.session_state.pop("pagina_estudo", None)
            st.rerun()

    # Grid 3x3 de informações
    col1, col2, col3 = st.columns(3)

    with col1:
        st.caption("🏢 Patrocinador")
        st.write(patrocinador or "-")
        st.caption("📅 Criado em")
        st.write(data_criacao_fmt)

    with col2:
        st.caption("🩺 Gerente Médico")
        st.write(gerente['nome'] if gerente else "-")
        st.caption("📊 Total de Desvios")
        st.write(str(qtd_desvios))

    with col3:
        st.caption("👥 Monitores")
        if monitores:
            for monitor in monitores:
                st.write(f"• {monitor}")
        else:
            st.write("-")

    # Botões de ação
    st.write("")
    col_btn1, col_btn2, col_spacer = st.columns([1, 1.5, 3.5])
    with col_btn1:
        if st.button("📋 Ver Desvios", use_container_width=True, type="primary" if pagina != "ver_desvios" else "secondary"):
            st.session_state["pagina_estudo"] = "ver_desvios"
            st.rerun()
    with col_btn2:
        if st.button("➕ Cadastrar Desvio", use_container_width=True, type="primary" if pagina != "cadastrar_desvio" else "secondary"):
            st.session_state["pagina_estudo"] = "cadastrar_desvio"
            st.rerun()

    st.divider()

    # Renderiza página específica do estudo
    if pagina == "ver_desvios":
        render_desvios_estudo(estudo, display_name, user_email)
    elif pagina == "cadastrar_desvio":
        render_cadastro_desvio(estudo, user_email, display_name)

# Página inicial (sem estudo selecionado)
else:
    render_meus_estudos(user_email)   