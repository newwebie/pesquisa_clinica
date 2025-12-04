"""
Sistema Syntox Churn
Dashboard moderno e profissional para análise de retenção de laboratórios
"""
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import calendar
import os
import time
import json
import logging
from typing import Optional, List, Dict, Any, Tuple
from io import BytesIO
from dataclasses import dataclass
from urllib.parse import quote_plus
import warnings
import html as html_utils
warnings.filterwarnings('ignore')

# Configurar logger
logger = logging.getLogger(__name__)
# Importar configurações
from config_churn import *
from pandas.tseries.offsets import BDay
# Importar sistema de autenticação Microsoft
from auth_microsoft import MicrosoftAuth, AuthManager, create_login_page, create_user_header

# ============================================
# HELPERS E TOOLTIPS AMIGÁVEIS (SISTEMA V2)
# ============================================

HELPERS_V2 = {
    'perda_risco_alto': """
    **Perda (Risco Alto)** quando, por exemplo:
    - Queda >50% vs baseline mensal, ou
    - Queda WoW >50%, ou
    - **Dias sem coleta no porte**:
      - **Médio**: ≥2 dias úteis e ≤15 corridos
      - **Médio/Grande**: ≥1 dia útil e ≤15 corridos
      - **Grande**: 1 a 5 dias úteis
      - **Pequeno**: não considerado

    Obs.: Perdas **recentes** começam a partir do **teto do risco** do porte, até 180 corridos; **antigas** >180 corridos.
    """,
    
    'baseline_mensal': """
    **Baseline Mensal**: Média dos maiores meses de coletas em 2024 e 2025 (configurável: 3 ou 6 meses). 
    Representa o volume de referência robusto do laboratório, menos suscetível a sazonalidade.
    """,
    
    'wow': """
    **WoW (Week over Week)**: Variação percentual entre a semana ISO atual e a semana anterior.
    Calcula apenas dias úteis de cada semana, excluindo fins de semana e feriados.
    """,
    
    'porte': """
    **Porte do Laboratório** (média mensal 2025):

    - **Pequeno**: até 40
    - **Médio**: 41–80
    - **Médio/Grande**: 81–150
    - **Grande**: >150
    """,
    
    'concorrencia': """
    **Sinal de Concorrência**: Indica que o CNPJ do laboratório apareceu no sistema do concorrente (Gralab).
    
    **Tipos de Concorrência:**
    - **Movimentações Recentes**: Credenciamento ou Descredenciamento nos últimos 14 dias (prioridade máxima)
    - **Apenas Cadastrados**: Laboratório cadastrado no concorrente, mas sem movimentação recente (monitorar)
    
    Dados são atualizados diretamente do arquivo Excel do Gralab, verificando tanto a aba "EntradaSaida" quanto "Dados Completos".
    """,
    
    'cap_alertas': """
    **Cap de Alertas**: Sistema limita a 30-50 alertas por dia, priorizando os mais severos.
    Isso reduz ruído e permite foco nos casos mais críticos.
    """,
    
    'fechamento_semanal': """
    **Fechamento Semanal (WoW - Week over Week)**
    
    Compara o volume de coletas úteis (segunda a sexta, excluindo feriados) 
    da semana ISO atual com a semana ISO anterior.
    
    **Gatilho de alerta:** Queda > 50% WoW
    
    **Régua de dias sem coleta aplicada conforme porte:**
    - Pequeno (≤40/mês): Monitorado apenas por queda de volume (não aciona por dias sem coleta)
    - Médio (41-80): mín 2 dias úteis sem coleta
    - Médio/Grande (81-150): mín 1 dia útil sem coleta
    - Grande (>150): mín 1 dia útil sem coleta
    """,
    
    'fechamento_mensal': """
    **Fechamento Mensal - Baseline Robusta**
    
    Consolida o mês corrente (até dia 30/31) e compara com baseline 
    calculada pela média dos top 3 maiores meses de 2024+2025.
    
    **Baseline adaptativa:** Cada laboratório tem baseline própria 
    baseada em seu histórico de pico.
    
    **Perda recente (até 180 dias):**
    - Pequeno: 30-180 dias corridos
    - Médio: 15-180 dias corridos
    - Médio/Grande: 15-180 dias corridos
    - Grande: 5+ dias úteis, máx 180 dias corridos
    
    **Perda antiga:** >180 dias corridos sem coleta (todos os portes)
    """,
}

MESES_PT_BR = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]


def _build_tooltip_attr(texto: Optional[str]) -> str:
    if not texto:
        return ""
    texto_limpo = " ".join(texto.replace("**", "").replace("  ", " ").split())
    return f'title="{html_utils.escape(texto_limpo)}"'


def render_info_card(icon: str, titulo: str, destaque: str, descricao: str, tooltip: Optional[str] = None, badge: Optional[str] = None):
    tooltip_attr = _build_tooltip_attr(tooltip)
    badge_html = f'<span style="background-color: rgba(0,0,0,0.08); padding: 0.15rem 0.4rem; border-radius: 999px; font-size: 0.75rem; font-weight: 600; margin-left: 0.3rem;">{badge}</span>' if badge else ""
    card_html = f"""
    <div class="info-card" {tooltip_attr}
         style="
            border-radius: 14px;
            padding: 1rem;
            background: var(--background-color-secondary, #f8f9fb);
            border: 1px solid rgba(0,0,0,0.05);
            min-height: 150px;
            display: flex;
            flex-direction: column;
            gap: 0.4rem;
         ">
        <div style="font-size: 1.6rem; line-height: 1;">{icon}</div>
        <div style="font-weight: 600; font-size: 1rem;">
            {titulo}{badge_html}
        </div>
        <div style="font-size: 1.4rem; font-weight: 700; color: var(--primary-color, #2c7be5);">
            {destaque}
        </div>
        <div style="font-size: 0.9rem; color: rgba(0,0,0,0.65); line-height: 1.45;">
            {descricao}
        </div>
    </div>
    """
    st.markdown(card_html, unsafe_allow_html=True)


def get_baseline_candidate_months() -> List[str]:
    mes_atual = datetime.now().month
    meses_2024 = [f"{mes}/24" for mes in MESES_PT_BR]
    meses_2025 = [f"{mes}/25" for mes in MESES_PT_BR[:mes_atual]]
    return meses_2024 + meses_2025


def get_baseline_window_label() -> str:
    mes_atual_nome = MESES_PT_BR[datetime.now().month - 1]
    return f"Jan-Dez/24 + Jan-{mes_atual_nome}/25"


def resumir_meses(meses: List[str], limite: int = 8) -> str:
    if not meses:
        return "Sem histórico disponível"
    if len(meses) <= limite:
        return ", ".join(meses)
    return f"{meses[0]} … {meses[-1]} (+{len(meses) - 2})"


# Dicionário de wordings atualizados
WORDING_V2 = {
    'possível perda': 'Perda',
    'possivel perda': 'Perda',
    'Alto Risco': 'Perda (Risco Alto)',
    'Médio Risco': 'Atenção',
    'Baixo Risco': 'Normal',
}
# ============================================
# FUNÇÕES DE INTEGRAÇÃO SHAREPOINT/ONEDRIVE
# ============================================
def _get_graph_config() -> Optional[Dict[str, Any]]:
    """Extrai configurações do Graph API dos secrets do Streamlit."""
    try:
        graph = st.secrets.get("graph", {})
        files = st.secrets.get("files", {})
        onedrive = st.secrets.get("onedrive", {})
        if not graph:
            return None
        return {
            "tenant_id": graph.get("tenant_id", ""),
            "client_id": graph.get("client_id", ""),
            "client_secret": graph.get("client_secret", ""),
            "hostname": graph.get("hostname", ""),
            "site_path": graph.get("site_path", ""),
            "library_name": graph.get("library_name", "Documents"),
            "user_upn": onedrive.get("user_upn", ""),
            "arquivo": files.get("arquivo", ""),
        }
    except Exception:
        return None
def _is_valid_csv(path: str) -> bool:
    """Verifica se arquivo CSV é válido."""
    try:
        if not os.path.exists(path):
            return False
        df = pd.read_csv(path, nrows=5)
        return len(df.columns) > 0
    except:
        return False
def _is_valid_parquet(path: str) -> bool:
    """Verifica se arquivo Parquet é válido."""
    try:
        if not os.path.exists(path):
            return False
        df = pd.read_parquet(path)
        return len(df.columns) > 0
    except:
        return False
def should_download_sharepoint(arquivo_remoto: str = None, force: bool = False) -> bool:
    """Verifica se deve baixar arquivo do SharePoint."""
    if force:
        return True
    # Determinar qual arquivo verificar (baseado no arquivo remoto solicitado)
    if arquivo_remoto:
        base_name = os.path.basename(arquivo_remoto)
        if base_name:
            arquivo_local = os.path.join(OUTPUT_DIR, base_name)
        else:
            arquivo_local = os.path.join(OUTPUT_DIR, "churn_analysis_latest.csv")
    else:
        arquivo_local = os.path.join(OUTPUT_DIR, "churn_analysis_latest.csv")
    # Verificar se existe arquivo local recente (< 5 minutos)
    if os.path.exists(arquivo_local):
        import time
        idade_arquivo = time.time() - os.path.getmtime(arquivo_local)
        if idade_arquivo < CACHE_TTL: # CACHE_TTL definido em config_churn.py
            return False
    return True
def baixar_sharepoint(arquivo_remoto: str = None, force: bool = False) -> Optional[str]:
    """
    Baixa arquivo do OneDrive/SharePoint via Microsoft Graph.
 
    Args:
        arquivo_remoto: Caminho do arquivo no OneDrive (usa config padrão se None)
        force: Força download mesmo se cache válido
 
    Returns:
        Caminho local do arquivo baixado ou None se falhar
    """
    cfg = _get_graph_config()
 
    # Sem configuração Graph, retornar arquivo local se existir
    if not cfg or not (cfg.get("tenant_id") and cfg.get("client_id") and cfg.get("client_secret")):
        arquivo_local = os.path.join(OUTPUT_DIR, "churn_analysis_latest.csv")
        if os.path.exists(arquivo_local):
            return arquivo_local
        return None
 
    # Verificar se precisa baixar
    if not should_download_sharepoint(arquivo_remoto=arquivo_remoto, force=force):
        # Retornar o arquivo local correspondente ao solicitado
        if arquivo_remoto:
            base_name = os.path.basename(arquivo_remoto)
            if base_name:
                arquivo_local = os.path.join(OUTPUT_DIR, base_name)
            else:
                arquivo_local = os.path.join(OUTPUT_DIR, "churn_analysis_latest.csv")
        else:
            arquivo_local = os.path.join(OUTPUT_DIR, "churn_analysis_latest.csv")
        if os.path.exists(arquivo_local):
            return arquivo_local
 
    try:
        os.makedirs(OUTPUT_DIR, exist_ok=True)
     
        # Usar ChurnSPConnector
        from churn_sp_connector import ChurnSPConnector
     
        connector = ChurnSPConnector(config=st.secrets)
     
        # Determinar arquivo remoto
        if arquivo_remoto is None:
            arquivo_remoto = cfg.get("arquivo", "Data Analysis/Churn PCLs/churn_analysis_latest.csv")
     
        # Baixar arquivo
        content = connector.download(arquivo_remoto)
     
        # Salvar localmente
        base_name = os.path.basename(arquivo_remoto)
        if not base_name:
            base_name = "churn_analysis_latest.csv"
     
        local_path = os.path.join(OUTPUT_DIR, base_name)
     
        with open(local_path, "wb") as f:
            f.write(content)
     
        # Validar arquivo baixado
        if _is_valid_csv(local_path) or _is_valid_parquet(local_path):
            return local_path
     
        return None
     
    except Exception as e:
        st.warning(f"⚠️ Não foi possível baixar do SharePoint: {e}")
        # Tentar usar arquivo local se existir
        arquivo_local = os.path.join(OUTPUT_DIR, "churn_analysis_latest.csv")
        if os.path.exists(arquivo_local):
            return arquivo_local
        return None

def baixar_excel_gralab(force: bool = False) -> Optional[str]:
    """
    Baixa arquivo Excel do Gralab do SharePoint.
    
    Args:
        force: Força download mesmo se cache válido
    
    Returns:
        Caminho local do arquivo baixado ou None se falhar
    """
    arquivo_remoto = "/personal/washington_gouvea_synvia_com_/Documents/Data Analysis/Churn PCLs/Automations/cunha/relatorio_completo_laboratorios_gralab.xlsx"
    base_name = "relatorio_completo_laboratorios_gralab.xlsx"
    arquivo_local = os.path.join(OUTPUT_DIR, base_name)
    
    # Verificar cache (4 horas = 14400 segundos)
    if not force and os.path.exists(arquivo_local):
        import time
        idade_arquivo = time.time() - os.path.getmtime(arquivo_local)
        if idade_arquivo < 14400:  # 4 horas
            return arquivo_local
    
    cfg = _get_graph_config()
    
    # Sem configuração Graph, retornar arquivo local se existir
    if not cfg or not (cfg.get("tenant_id") and cfg.get("client_id") and cfg.get("client_secret")):
        if os.path.exists(arquivo_local):
            return arquivo_local
        return None
    
    try:
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        
        # Usar ChurnSPConnector
        from churn_sp_connector import ChurnSPConnector
        
        connector = ChurnSPConnector(config=st.secrets)
        
        # Baixar arquivo
        content = connector.download(arquivo_remoto)
        
        # Salvar localmente
        with open(arquivo_local, "wb") as f:
            f.write(content)
        
        # Validar se é Excel válido
        try:
            import openpyxl
            openpyxl.load_workbook(arquivo_local, read_only=True)
            return arquivo_local
        except Exception:
            return None
        
    except Exception as e:
        # Tentar usar arquivo local se existir
        if os.path.exists(arquivo_local):
            return arquivo_local
        return None

def show_overlay_loader(title: str = "Carregando...", subtitle: str = "Por favor, aguarde enquanto processamos os dados."):
    """
    Função helper para criar o overlay loader facilmente.
    
    Args:
        title: Título do loader
        subtitle: Subtítulo/descrição do loader
    
    Returns:
        placeholder que deve ser limpo com .empty() quando terminar
    """
    loader_placeholder = st.empty()
    loader_placeholder.markdown(
        f"""
        <div class="overlay-loader">
            <div class="overlay-loader__content">
                <div class="overlay-loader__spinner"></div>
                <div class="overlay-loader__title">{title}</div>
                <div class="overlay-loader__subtitle">{subtitle}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )
    return loader_placeholder

# Configuração da página
st.set_page_config(
    page_title="📊 Syntox Churn",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        'About': "Syntox Churn - Sistema profissional para monitoramento de retenção de PCLs"
    }
)
# CSS moderno e profissional - Atualizado com melhorias de layout
CSS_STYLES = """
<style>
    /* Tema profissional atualizado - Synvia */
    :root {
        --primary-color: #6BBF47;
        --secondary-color: #52B54B;
        --success-color: #6BBF47;
        --warning-color: #F59E0B;
        --danger-color: #DC2626;
        --info-color: #3B82F6;
        --light-bg: #F5F7FA;
        --dark-bg: #262730;
        --border-radius: 12px; /* Aumentado para visual mais moderno */
        --shadow: 0 4px 8px rgba(0,0,0,0.1); /* Sombra mais suave */
        --transition: all 0.3s ease;
    }
    /* Reset e base */
    * { box-sizing: border-box; }
    /* Header profissional */
    .main-header {
        background: linear-gradient(135deg, var(--primary-color), var(--secondary-color));
        color: white;
        padding: 2.5rem 1.5rem; /* Aumentado padding */
        border-radius: var(--border-radius);
        margin-bottom: 2.5rem;
        text-align: center;
        box-shadow: var(--shadow);
    }
    .main-header h1 {
        margin: 0;
        font-size: 2.8rem; /* Aumentado tamanho */
        font-weight: 400;
        text-shadow: 0 2px 4px rgba(0,0,0,0.3);
    }
    .main-header p {
        margin: 0.5rem 0 0 0;
        opacity: 0.9;
        font-size: 1.2rem;
    }
    /* Cards de métricas modernas - Melhorados */
    .metric-card {
        background: white;
        border-radius: var(--border-radius);
        padding: 1.5rem;
        box-shadow: var(--shadow);
        border: 1px solid #e9ecef;
        transition: var(--transition);
        text-align: center;
        margin-bottom: 1.5rem; /* Aumentado espaçamento */
        display: flex;               /* Estabilidade de altura */
        flex-direction: column;      /* Empilha valor, label, delta */
        justify-content: center;     /* Centraliza verticalmente */
        min-height: 140px;           /* Altura mínima consistente */
    }
    .metric-card:hover {
        transform: translateY(-4px); /* Mais elevação */
        box-shadow: 0 6px 12px rgba(0,0,0,0.15);
    }
    .metric-value {
        font-size: 2.2rem; /* Aumentado */
        font-weight: 700;
        margin: 0.5rem 0;
        color: var(--primary-color);
    }
    .metric-label {
        font-size: 1rem; /* Ajustado */
        color: #6c757d;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin: 0;
    }
    .metric-delta {
        font-size: 0.9rem;
        margin-top: 0.5rem;
        min-height: 1rem;            /* Reserva espaço mesmo vazia */
    }
    .metric-delta.positive { color: var(--success-color); }
    .metric-delta.negative { color: var(--danger-color); }
    /* Status badges - Ajustados */
    .status-badge {
        display: inline-block;
        padding: 0.35rem 0.85rem; /* Ajustado espaçamento */
        border-radius: 20px;
        font-size: 0.9rem;
        font-weight: 600;
        text-transform: uppercase;
    }
    .status-alto { background-color: #fee; color: var(--danger-color); border: 1px solid #fcc; }
    .status-medio { background-color: #ffeaa7; color: var(--warning-color); border: 1px solid #ffeaa7; }
    .status-baixo { background-color: #d4edda; color: var(--success-color); border: 1px solid #c3e6cb; }
    .status-inativo { background-color: #f8f9fa; color: #6c757d; border: 1px solid #dee2e6; }
    /* Botões modernos */
    .stButton > button {
        background: linear-gradient(135deg, var(--primary-color), var(--secondary-color));
        color: white;
        border: none;
        border-radius: var(--border-radius);
        padding: 0.85rem 1.75rem; /* Ajustado */
        font-weight: 600;
        transition: var(--transition);
        box-shadow: var(--shadow);
    }
    .stButton > button:hover {
        transform: translateY(-2px); /* Mais elevação */
        box-shadow: 0 6px 12px rgba(0,0,0,0.2);
    }
    /* Sidebar moderna */
    .sidebar-header {
        background: var(--light-bg);
        padding: 1.2rem;
        border-radius: var(--border-radius);
        margin-bottom: 1.2rem;
        border-left: 5px solid var(--primary-color);
    }
    .sidebar-header h3 {
        margin: 0;
        color: var(--primary-color);
        font-size: 1.2rem;
        font-weight: 600;
    }
    /* Tabelas modernas */
    .dataframe-container {
        background: white;
        border-radius: var(--border-radius);
        padding: 1.2rem;
        box-shadow: var(--shadow);
        overflow: hidden;
    }
    /* Expander styling */
    .streamlit-expanderHeader {
        background: var(--light-bg);
        border-radius: var(--border-radius);
        font-weight: 600;
        color: var(--primary-color);
    }
    /* Loading states */
    .loading-container {
        text-align: center;
        padding: 3rem;
        color: #6c757d;
    }
    .loading-spinner {
        border: 4px solid #f3f3f3;
        border-top: 4px solid var(--primary-color);
        border-radius: 50%;
        width: 40px;
        height: 40px;
        animation: spin 1s linear infinite;
        margin: 0 auto 1rem;
    }
    .overlay-loader {
        position: fixed;
        inset: 0;
        background: rgba(15, 23, 42, 0.92);
        backdrop-filter: blur(4px);
        z-index: 10000;
        display: flex;
        align-items: center;
        justify-content: center;
    }
    .overlay-loader__content {
        text-align: center;
        color: #f8fafc;
        max-width: 480px;
        padding: 2.5rem;
        border-radius: 20px;
        background: rgba(17, 24, 39, 0.55);
        box-shadow: 0 20px 45px rgba(15, 23, 42, 0.45);
        border: 1px solid rgba(148, 163, 184, 0.2);
    }
    .overlay-loader__spinner {
        border: 6px solid rgba(148, 163, 184, 0.3);
        border-top: 6px solid var(--secondary-color);
        border-radius: 50%;
        width: 90px;
        height: 90px;
        animation: spin 1s linear infinite;
        margin: 0 auto 1.5rem;
    }
    .overlay-loader__title {
        font-size: 1.6rem;
        font-weight: 600;
        margin-bottom: 0.5rem;
    }
    .overlay-loader__subtitle {
        font-size: 1rem;
        opacity: 0.85;
        line-height: 1.5;
    }
    @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
    }
    /* Responsividade */
    @media (max-width: 768px) {
        .metric-card {
            margin-bottom: 1.5rem;
        }
        .main-header h1 {
            font-size: 2.2rem;
        }
        .metric-value {
            font-size: 1.8rem;
        }
    }
    /* Custom scrollbar */
    ::-webkit-scrollbar {
        width: 8px;
        height: 8px;
    }
    ::-webkit-scrollbar-track {
        background: #f1f1f1;
        border-radius: 4px;
    }
    ::-webkit-scrollbar-thumb {
        background: var(--primary-color);
        border-radius: 4px;
    }
    ::-webkit-scrollbar-thumb:hover {
        background: var(--secondary-color);
    }
    /* Footer */
    .footer {
        text-align: center;
        padding: 2rem 0;
        color: #6c757d;
        border-top: 1px solid #e9ecef;
        margin-top: 3rem;
    }
    /* Dark mode support */
    @media (prefers-color-scheme: dark) {
        :root {
            --light-bg: #2d3748;
            --dark-bg: #1a202c;
        }
        .metric-card {
            background: var(--dark-bg);
            border-color: #4a5568;
            color: white;
        }
        .metric-label {
            color: #a0aec0;
        }
    }
    /* Melhorias de espaçamento e layout */
    section[data-testid="stExpander"] > div {
        margin-bottom: 1rem;
    }
    .stTabs [data-testid="stMarkdownContainer"] {
        font-size: 1.1rem;
        font-weight: 600;
    }
    /* Ajuste para gráficos */
    .plotly-chart {
        margin: 1rem 0;
        border-radius: var(--border-radius);
        box-shadow: var(--shadow);
        padding: 1rem;
        background: white;
    }
</style>
"""
# Injetar CSS
st.markdown(CSS_STYLES, unsafe_allow_html=True)
# ========================================
# CLASSES DO SISTEMA v2.0 - Atualizado com correções de bugs
# ========================================
@dataclass
class KPIMetrics:
    """Classe para armazenar métricas calculadas."""
    total_labs: int = 0
    churn_rate: float = 0.0
    total_coletas: int = 0
    labs_em_risco: int = 0
    ativos_7d: float = 0.0
    ativos_30d: float = 0.0
    labs_alto_risco: int = 0
    labs_medio_risco: int = 0
    labs_baixo_risco: int = 0
    labs_inativos: int = 0
    labs_critico: int = 0
    labs_recuperando: int = 0
    labs_sem_coleta_48h: int = 0
    vol_hoje_total: int = 0
    vol_d1_total: int = 0
    ativos_7d_count: int = 0
    ativos_30d_count: int = 0
    labs_normal_count: int = 0
    labs_atencao_count: int = 0
    labs_moderado_count: int = 0
    labs_alto_count: int = 0
    labs_critico_count: int = 0
    labs_abaixo_mm7_br: int = 0
    labs_abaixo_mm7_br_pct: float = 0.0
    labs_abaixo_mm7_uf: int = 0
    labs_abaixo_mm7_uf_pct: float = 0.0
class DataManager:
    """Gerenciador de dados com cache inteligente."""
    @staticmethod
    def normalizar_cnpj(cnpj: str) -> str:
        """Remove formatação do CNPJ (pontos, traços, barras)"""
        if pd.isna(cnpj) or cnpj == '':
            return ''
        # Converter numéricos para string sem decimais (evita sufixo '.0')
        if isinstance(cnpj, (int, float)):
            try:
                cnpj = str(int(cnpj))
            except Exception:
                cnpj = str(cnpj)
        # Remove tudo exceto dígitos
        cnpj_limpo = ''.join(filter(str.isdigit, str(cnpj)))
        # Garantir 14 dígitos
        if len(cnpj_limpo) < 14:
            cnpj_limpo = cnpj_limpo.zfill(14)
        elif len(cnpj_limpo) > 14:
            cnpj_limpo = cnpj_limpo[-14:]
        return cnpj_limpo
    @staticmethod
    @st.cache_data(ttl=CACHE_TTL)
    def carregar_dados_churn() -> Optional[pd.DataFrame]:
        """Carrega dados de análise de churn com cache inteligente."""
        try:
            # PRIMEIRO: Tentar baixar do SharePoint/OneDrive
            arquivo_sharepoint = baixar_sharepoint()
         
            if arquivo_sharepoint and os.path.exists(arquivo_sharepoint):
                # Tentar ler como CSV primeiro
                try:
                    df = pd.read_csv(arquivo_sharepoint, encoding=ENCODING, low_memory=False)
                    return df
                except Exception:
                    # Tentar como Parquet
                    try:
                        df = pd.read_parquet(arquivo_sharepoint, engine='pyarrow')
                        return df
                    except Exception:
                        pass
         
            # FALLBACK: Tentar arquivos locais
            # Primeiro tenta CSV (mais comum)
            arquivo_csv = os.path.join(OUTPUT_DIR, "churn_analysis_latest.csv")
            if os.path.exists(arquivo_csv):
                df = pd.read_csv(arquivo_csv, encoding=ENCODING, low_memory=False)
                return df
         
            # Fallback para parquet
            arquivo_path = os.path.join(OUTPUT_DIR, CHURN_ANALYSIS_FILE)
            if os.path.exists(arquivo_path):
                df = pd.read_parquet(arquivo_path, engine='pyarrow')
                return df
         
            return None
         
        except Exception as e:
            st.error(f"❌ Erro ao carregar dados: {e}")
            return None
    @staticmethod
    def preparar_dados(df: pd.DataFrame) -> pd.DataFrame:
        """Prepara e limpa os dados carregados - Atualizado para coerência entre telas."""
        if df is None or df.empty:
            return pd.DataFrame()
        
        # VALIDAÇÃO: Remover duplicatas baseadas em CNPJ antes de qualquer processamento
        if 'CNPJ_PCL' in df.columns:
            # Criar CNPJ_Normalizado temporariamente para deduplicação se ainda não existir
            if 'CNPJ_Normalizado' not in df.columns:
                df['CNPJ_Normalizado'] = df['CNPJ_PCL'].apply(DataManager.normalizar_cnpj)
            # Remover duplicatas mantendo o primeiro registro
            df = df.drop_duplicates(subset=['CNPJ_Normalizado'], keep='first')
        elif 'CNPJ_Normalizado' in df.columns:
            # Se já existe CNPJ_Normalizado mas não CNPJ_PCL, usar CNPJ_Normalizado
            df = df.drop_duplicates(subset=['CNPJ_Normalizado'], keep='first')
        
        # Removido bloco de debug da sidebar para manter interface limpa
        # Garantir tipos de dados corretos
        if 'Data_Analise' in df.columns:
            df['Data_Analise'] = pd.to_datetime(df['Data_Analise'], errors='coerce')
        # Calcular volume total se não existir (até o mês atual)
        try:
            # Função inline para evitar dependência circular
            meses_ordem = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
            ano_atual = pd.Timestamp.today().year
            limite_mes = pd.Timestamp.today().month if 2025 == ano_atual else 12
            meses_limite = meses_ordem[:limite_mes]
            sufixo = str(2025)[-2:]
            meses_2025_dyn = [m for m in meses_limite if f'N_Coletas_{m}_{sufixo}' in df.columns]
        except Exception:
            meses_2025_dyn = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out']
        colunas_meses = [f'N_Coletas_{mes}_25' for mes in meses_2025_dyn]
        if 'Volume_Total_2025' not in df.columns:
            df['Volume_Total_2025'] = df[colunas_meses].sum(axis=1, skipna=True) if colunas_meses else 0
        # Adicionar coluna CNPJ normalizado para match com dados VIP
        if 'CNPJ_PCL' in df.columns:
            df['CNPJ_Normalizado'] = df['CNPJ_PCL'].apply(DataManager.normalizar_cnpj)
        # Filtro Active == True para coerência
        if 'Active' in df.columns:
            df = df[df['Active'] == True]

        # 2. INTEGRAÇÃO GLOBAL VIP (Para uso nas abas que mostram tudo)
        try:
            df_vip = DataManager.carregar_dados_vip()
            
            if df_vip is not None and not df_vip.empty:
                if 'CNPJ_Normalizado' not in df_vip.columns:
                     df_vip['CNPJ_Normalizado'] = df_vip['CNPJ'].apply(DataManager.normalizar_cnpj)
                
                # Garantir que não há duplicatas no df_vip antes do merge
                df_vip = df_vip.drop_duplicates(subset=['CNPJ_Normalizado'], keep='first')
                
                # Colunas para trazer
                cols_vip = ['CNPJ_Normalizado', 'Rede', 'Ranking', 'Ranking Rede']
                cols_vip = [c for c in cols_vip if c in df_vip.columns]
                
                # Merge mantendo TODOS os laboratórios (Left Join)
                df = df.merge(df_vip[cols_vip], on='CNPJ_Normalizado', how='left', suffixes=('', '_vip'))
                
                # Remover duplicatas após merge (caso o df original já tivesse duplicatas)
                df = df.drop_duplicates(subset=['CNPJ_Normalizado'], keep='first')
                
                # Consolidar Rede
                if 'Rede' not in df.columns and 'Rede_vip' in df.columns:
                    df['Rede'] = df['Rede_vip']
                elif 'Rede' in df.columns and 'Rede_vip' in df.columns:
                    df['Rede'] = df['Rede'].fillna(df['Rede_vip'])
                
                # Criar flags
                df['VIP'] = np.where(df['Ranking'].notna(), 'Sim', 'Não')
                df['Rede'] = df['Rede'].fillna('-')
            else:
                df['VIP'] = 'Não'
                df['Rede'] = '-'
        except Exception as e:
            # Log erro silencioso para não travar app
            print(f"Erro integração VIP: {e}")
            df['VIP'] = 'Não'
            df['Rede'] = '-'

        # === Nova régua de risco diário ===
        colunas_novas = [
            "Vol_Hoje", "Vol_D1", "MM7", "MM30", "MM90", "DOW_Media",
            "Delta_D1", "Delta_MM7", "Delta_MM30", "Delta_MM90",
            "Risco_Diario", "Recuperacao"
        ]
        try:
            registros = []
            for _, r in df.iterrows():
                res = RiskEngine.classificar(r)
                registros.append(res if res else {c: None for c in colunas_novas})
            df_risk = pd.DataFrame(registros, index=df.index)
            for c in colunas_novas:
                df[c] = df_risk.get(c)
        except Exception:
            for c in colunas_novas:
                if c not in df.columns:
                    df[c] = None
        # Opcional: preservar a coluna antiga para auditoria
        if 'Status_Risco' in df.columns and 'Risco_Diario' in df.columns:
            df.rename(columns={'Status_Risco': 'Status_Risco_Legado'}, inplace=True)

        # Normalizar colunas de recoletas
        recoleta_cols = [c for c in df.columns if c.startswith('Recoletas_')]
        for col in recoleta_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
        for col in ['Total_Recoletas_2024', 'Total_Recoletas_2025']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

        # Garantir colunas de preço mesmo quando não presentes no arquivo (fallback SharePoint)
        price_cols_expected: List[str] = []
        price_prefixes: List[str] = []
        for cfg in PRICE_CATEGORIES.values():
            prefix = cfg['prefix']
            price_prefixes.append(prefix)
            price_cols_expected.extend([
                f'Preco_{prefix}_Total',
                f'Preco_{prefix}_Coleta',
                f'Preco_{prefix}_Exame'
            ])
        extra_price_cols = ['Voucher_Commission', 'Data_Preco_Atualizacao']
        precisa_precos = any(col not in df.columns for col in price_cols_expected) or any(
            col not in df.columns for col in extra_price_cols
        )
        if precisa_precos:
            df_prices_extra = DataManager.carregar_prices()
            if df_prices_extra is not None and not df_prices_extra.empty and '_laboratory' in df_prices_extra.columns:
                df_prices_extra = df_prices_extra.copy()
                df_prices_extra['_laboratory'] = df_prices_extra['_laboratory'].astype(str)
                merge_col = None
                for candidate in ['_id', 'Laboratory_ID', 'LaboratoryId', 'LaboratoryID', 'Lab_ID', 'LabId', 'id']:
                    if candidate in df.columns:
                        merge_col = candidate
                        df[merge_col] = df[merge_col].astype(str)
                        break
                if merge_col:
                    lookup = df_prices_extra.set_index('_laboratory')
                    mapping_cache = {}
                    for col in price_cols_expected + extra_price_cols:
                        if col in lookup.columns:
                            mapping_cache[col] = lookup[col].to_dict()
                    for col, mapping in mapping_cache.items():
                        if col not in df.columns:
                            df[col] = np.nan
                        df[col] = df[col].fillna(df[merge_col].map(mapping))
                else:
                    st.warning("⚠️ Não foi possível vincular preços aos laboratórios (ID ausente).")

        for col in price_cols_expected:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')

        if 'Voucher_Commission' in df.columns:
            df['Voucher_Commission'] = pd.to_numeric(df['Voucher_Commission'], errors='coerce')

        if 'Data_Preco_Atualizacao' in df.columns:
            df['Data_Preco_Atualizacao'] = pd.to_datetime(df['Data_Preco_Atualizacao'], errors='coerce', utc=True)
            try:
                df['Data_Preco_Atualizacao'] = df['Data_Preco_Atualizacao'].dt.tz_convert(TIMEZONE)
            except Exception:
                pass

        return df
    @staticmethod
    @st.cache_data(ttl=CACHE_TTL)
    def carregar_matriz_cs_normalizada() -> Optional[pd.DataFrame]:
        """Carrega dados da matriz CS normalizada com cache inteligente."""
        try:
            # PRIMEIRO: Tentar baixar do SharePoint/OneDrive
            arquivo_vip_remoto = "Data Analysis/Churn PCLs/matriz_cs_normalizada.csv"
            arquivo_sharepoint = baixar_sharepoint(arquivo_remoto=arquivo_vip_remoto)
            if arquivo_sharepoint and os.path.exists(arquivo_sharepoint):
                # Tentar ler como CSV
                try:
                    df = pd.read_csv(arquivo_sharepoint, encoding='utf-8-sig', low_memory=False)
                    # Verificar se tem coluna CNPJ ou CNPJ_PCL
                    if 'CNPJ' in df.columns:
                        coluna_cnpj = 'CNPJ'
                    elif 'CNPJ_PCL' in df.columns:
                        coluna_cnpj = 'CNPJ_PCL'
                        # Renomear para CNPJ para compatibilidade
                        df['CNPJ'] = df['CNPJ_PCL']
                    else:
                    # Warning removido - será tratado onde a função é chamada
                        return None
                    # Ler CNPJ como string para preservar zeros à esquerda
                    df['CNPJ'] = df['CNPJ'].astype(str)
                    df['CNPJ_Normalizado'] = df['CNPJ'].apply(DataManager.normalizar_cnpj)
                    # Toast removido - será exibido onde a função é chamada
                    return df
                except Exception as e:
                    # Warning removido - será tratado onde a função é chamada
                    pass
            # FALLBACK: Tentar arquivos locais
            caminhos_possiveis = [
                VIP_CSV_FILE,
                os.path.join(OUTPUT_DIR, VIP_CSV_FILE),
                os.path.join(os.path.dirname(OUTPUT_DIR), VIP_CSV_FILE),
            ]
            arquivo_csv = None
            for caminho in caminhos_possiveis:
                if os.path.exists(caminho):
                    arquivo_csv = caminho
                    break
            if arquivo_csv:
                # Ler CNPJ como string para preservar zeros à esquerda
                df = pd.read_csv(
                    arquivo_csv,
                    encoding='utf-8-sig',
                    dtype={'CNPJ': 'string'},
                    low_memory=False
                )
                # Garantir que CNPJ seja string e normalizar
                df['CNPJ'] = df['CNPJ'].astype(str)
                df['CNPJ_Normalizado'] = df['CNPJ'].apply(DataManager.normalizar_cnpj)
                # Toast removido - será exibido onde a função é chamada
                return df
            return None
        except Exception as e:
            # Error removido - será tratado onde a função é chamada
            return None
    @staticmethod
    @st.cache_data(ttl=VIP_CACHE_TTL)
    def carregar_dados_vip() -> Optional[pd.DataFrame]:
        """Carrega dados VIP do CSV normalizado com cache."""
        try:
            # Tentar baixar matriz CS do SharePoint
            arquivo_vip_remoto = "Data Analysis/Churn PCLs/matriz_cs_normalizada.csv"
            arquivo_sharepoint = baixar_sharepoint(arquivo_remoto=arquivo_vip_remoto, force=False)
            if arquivo_sharepoint and os.path.exists(arquivo_sharepoint):
                # Ler arquivo VIP
                df_vip = pd.read_csv(
                    arquivo_sharepoint,
                    encoding='utf-8-sig'
                )
                # Verificar se tem coluna CNPJ ou CNPJ_PCL
                if 'CNPJ' in df_vip.columns:
                    coluna_cnpj = 'CNPJ'
                elif 'CNPJ_PCL' in df_vip.columns:
                    coluna_cnpj = 'CNPJ_PCL'
                    # Renomear para CNPJ para compatibilidade
                    df_vip['CNPJ'] = df_vip['CNPJ_PCL']
                else:
                    # Warning removido - será tratado onde a função é chamada
                    return None
                # Ler CNPJ como string para preservar zeros à esquerda
                df_vip['CNPJ'] = df_vip['CNPJ'].astype(str)
                df_vip['CNPJ_Normalizado'] = df_vip['CNPJ'].apply(DataManager.normalizar_cnpj)
                # Remover duplicatas baseadas em CNPJ_Normalizado (manter primeiro registro)
                df_vip = df_vip.drop_duplicates(subset=['CNPJ_Normalizado'], keep='first')
                # Toast removido - será exibido onde a função é chamada
                return df_vip
         
            # FALLBACK: Tentar múltiplos caminhos locais
            caminhos_possiveis = [
                VIP_CSV_FILE,
                os.path.join(OUTPUT_DIR, VIP_CSV_FILE),
                os.path.join(os.path.dirname(OUTPUT_DIR), VIP_CSV_FILE),
            ]
            arquivo_csv = None
            for caminho in caminhos_possiveis:
                if os.path.exists(caminho):
                    arquivo_csv = caminho
                    break
            if arquivo_csv:
                # Ler CNPJ como string para preservar zeros à esquerda
                df_vip = pd.read_csv(
                    arquivo_csv,
                    encoding='utf-8-sig',
                    dtype={'CNPJ': 'string'}
                )
                # Garantir que CNPJ seja string e normalizar
                df_vip['CNPJ'] = df_vip['CNPJ'].astype(str)
                df_vip['CNPJ_Normalizado'] = df_vip['CNPJ'].apply(DataManager.normalizar_cnpj)
                # Remover duplicatas baseadas em CNPJ_Normalizado (manter primeiro registro)
                df_vip = df_vip.drop_duplicates(subset=['CNPJ_Normalizado'], keep='first')
                # Toast removido - será exibido onde a função é chamada
                return df_vip
            else:
                # Warning removido - será tratado onde a função é chamada
                return None
        except Exception as e:
            st.warning(f"Erro ao carregar arquivo VIP: {e}")
            return None

    @staticmethod
    def salvar_dados_vip(df_vip: pd.DataFrame) -> None:
        """Salva dados VIP em CSV local (fallback simples)."""
        if df_vip is None or df_vip.empty:
            return
        try:
            caminho = os.path.join(OUTPUT_DIR, VIP_CSV_FILE) if 'VIP_CSV_FILE' in globals() else os.path.join(OUTPUT_DIR, "vip_data.csv")
            os.makedirs(os.path.dirname(caminho), exist_ok=True)
            df_vip.to_csv(caminho, index=False, encoding='utf-8-sig')
            st.toast("✅ Dados VIP salvos localmente.")
        except Exception as e:
            st.warning(f"Não foi possível salvar dados VIP localmente: {e}")
    @staticmethod
    @st.cache_data(ttl=CACHE_TTL)
    def carregar_laboratories() -> Optional[pd.DataFrame]:
        """Carrega dados de laboratories.csv com cache inteligente."""
        try:
            # PRIMEIRO: Tentar baixar do SharePoint/OneDrive
            arquivo_labs_remoto = "Data Analysis/Churn PCLs/laboratories.csv"
            arquivo_sharepoint = baixar_sharepoint(arquivo_remoto=arquivo_labs_remoto)
            
            if arquivo_sharepoint and os.path.exists(arquivo_sharepoint):
                try:
                    df_labs = pd.read_csv(arquivo_sharepoint, encoding=ENCODING, low_memory=False)
                    
                    # Normalizar CNPJ para permitir matching
                    if 'cnpj' in df_labs.columns:
                        df_labs['cnpj'] = df_labs['cnpj'].astype(str)
                        df_labs['CNPJ_Normalizado'] = df_labs['cnpj'].apply(DataManager.normalizar_cnpj)
                    elif 'CNPJ' in df_labs.columns:
                        df_labs['CNPJ'] = df_labs['CNPJ'].astype(str)
                        df_labs['CNPJ_Normalizado'] = df_labs['CNPJ'].apply(DataManager.normalizar_cnpj)
                    
                    return df_labs
                except Exception as e:
                    # Erro silencioso - será tratado onde a função é chamada
                    pass
            
            # FALLBACK: Tentar arquivo local
            arquivo_labs = os.path.join(OUTPUT_DIR, LABORATORIES_FILE)
            if os.path.exists(arquivo_labs):
                df_labs = pd.read_csv(arquivo_labs, encoding=ENCODING, low_memory=False)
                
                # Normalizar CNPJ para permitir matching
                if 'cnpj' in df_labs.columns:
                    df_labs['cnpj'] = df_labs['cnpj'].astype(str)
                    df_labs['CNPJ_Normalizado'] = df_labs['cnpj'].apply(DataManager.normalizar_cnpj)
                elif 'CNPJ' in df_labs.columns:
                    df_labs['CNPJ'] = df_labs['CNPJ'].astype(str)
                    df_labs['CNPJ_Normalizado'] = df_labs['CNPJ'].apply(DataManager.normalizar_cnpj)
                
                return df_labs
            
            return None
        except Exception as e:
            # Erro silencioso - será tratado onde a função é chamada
            return None
    
    @staticmethod
    @st.cache_data(ttl=CACHE_TTL)
    def carregar_prices(force: bool = False) -> Optional[pd.DataFrame]:
        """Carrega dados de prices.csv com fallback para SharePoint."""
        try:
            arquivo_local = os.path.join(OUTPUT_DIR, PRICES_FILE)
            if os.path.exists(arquivo_local) and not force:
                return pd.read_csv(arquivo_local, encoding=ENCODING, low_memory=False)

            files_cfg = {}
            try:
                files_cfg = st.secrets.get('files', {})
            except Exception:
                files_cfg = {}

            arquivo_remoto = files_cfg.get('prices', PRICES_REMOTE_PATH)
            if arquivo_remoto:
                arquivo_remoto = arquivo_remoto.replace("\\", "/")
            caminho = baixar_sharepoint(arquivo_remoto=arquivo_remoto, force=force)
            if caminho and os.path.exists(caminho):
                return pd.read_csv(caminho, encoding=ENCODING, low_memory=False)
        except Exception as e:
            st.warning(f"⚠️ Falha ao carregar tabela de preços: {e}")
        return None
    
    @staticmethod
    @st.cache_data(ttl=14400)  # Cache de 4 horas
    def carregar_dados_gralab() -> Optional[Dict[str, pd.DataFrame]]:
        """
        Carrega dados do Excel do concorrente Gralab com todas as abas.
        
        Returns:
            Dicionário com DataFrames das abas ou None se falhar
        """
        try:
            # Baixar arquivo Excel do SharePoint
            arquivo_excel = baixar_excel_gralab()
            
            if not arquivo_excel or not os.path.exists(arquivo_excel):
                return None
            
            # Ler todas as abas do Excel
            todas_abas = pd.read_excel(arquivo_excel, sheet_name=None, engine='openpyxl')
            
            # Normalizar CNPJ em todas as abas que tenham coluna CNPJ ou Cnpj
            for nome_aba, df in todas_abas.items():
                # Procurar coluna de CNPJ (case insensitive)
                coluna_cnpj = None
                for col in df.columns:
                    if col.upper() == 'CNPJ':
                        coluna_cnpj = col
                        break
                
                if coluna_cnpj:
                    df[coluna_cnpj] = df[coluna_cnpj].astype(str)
                    df['CNPJ_Normalizado'] = df[coluna_cnpj].apply(DataManager.normalizar_cnpj)
            
            return todas_abas
            
        except Exception as e:
            # Erro silencioso - será tratado onde a função é chamada
            return None


class RiskEngine:
    """Calcula MM7/MM30/MM90, D-1, DOW e classifica o risco diário (nova régua)."""

    @staticmethod
    def _serie_diaria_from_json(json_str: str) -> pd.Series:
        """Converte 'Dados_Diarios_2025' (dict 'YYYY-MM' -> {dia:coletas}) em série diária."""
        if pd.isna(json_str) or str(json_str).strip() in ("", "{}", "null"):
            return pd.Series(dtype="float")
        import json
        try:
            j = json.loads(json_str)
        except Exception:
            return pd.Series(dtype="float")
        rows = []
        for ym, dias in j.items():
            try:
                y, m = ym.split("-")
            except Exception:
                continue
            for d_str, v in dias.items():
                try:
                    d = int(d_str)
                    rows.append((pd.Timestamp(int(y), int(m), d), int(v)))
                except Exception:
                    continue
        if not rows:
            return pd.Series(dtype="float")
        s = pd.Series({d: v for d, v in rows}).sort_index()
        return s

    @staticmethod
    def _last_business_day(reference: Optional[pd.Timestamp] = None) -> pd.Timestamp:
        """Retorna a última data útil (considerando TIMEZONE)."""
        if reference is None:
            reference = pd.Timestamp.now(tz=TIMEZONE)
        else:
            if reference.tzinfo is None:
                reference = reference.tz_localize(TIMEZONE)
            else:
                reference = reference.tz_convert(TIMEZONE)
        reference = reference.normalize()
        reference_naive = reference.tz_localize(None)
        while reference_naive.weekday() >= 5:
            reference_naive = (reference_naive - BDay(1))
        return reference_naive

    @staticmethod
    def _serie_business_day(s: pd.Series, ref_date: pd.Timestamp) -> pd.Series:
        """Reindexa série para frequência de dias úteis até ref_date."""
        if s.empty:
            return s
        s = s.sort_index()
        start = s.index.min()
        if ref_date < start:
            ref_date = start
        idx = pd.bdate_range(start, ref_date)
        if len(idx) == 0:
            idx = pd.DatetimeIndex([ref_date])
        return s.reindex(idx, fill_value=0)

    @staticmethod
    def _rolling_means(s: pd.Series, ref_date: pd.Timestamp) -> dict:
        """MM7/MM30/MM90, D-1, média por DOW e contadores auxiliares."""
        if s.empty:
            return dict(MM7=0, MM30=0, MM90=0, D1=0, DOW=0, HOJE=0, zeros_consec=0, quedas50_consec=0)
        if ref_date not in s.index:
            return dict(MM7=0, MM30=0, MM90=0, D1=0, DOW=0, HOJE=0, zeros_consec=0, quedas50_consec=0)

        hoje = float(s.loc[ref_date])
        serie_ate_ref = s.loc[:ref_date]
        if len(serie_ate_ref) > 1:
            d1 = float(serie_ate_ref.iloc[-2])
        else:
            d1 = 0.0
        mm7 = float(serie_ate_ref.tail(7).mean())
        mm30 = float(serie_ate_ref.tail(30).mean())
        mm90 = float(serie_ate_ref.tail(90).mean())
        dow = int(ref_date.weekday())
        dow_vals = serie_ate_ref[serie_ate_ref.index.weekday == dow]
        dow_mean = float(dow_vals.tail(90).mean()) if len(dow_vals) else 0.0
        zeros_consec = 0
        for valor in serie_ate_ref[::-1]:
            if valor == 0:
                zeros_consec += 1
            else:
                break

        def _is_queda50(idx):
            mm7_local = s.loc[:idx].tail(7).mean()
            return s.loc[idx] < 0.5 * mm7_local if mm7_local > 0 else False

        ultimos = s.loc[:ref_date].tail(3)
        quedas50_consec = sum([_is_queda50(idx) for idx in ultimos.index])
        return dict(MM7=mm7, MM30=mm30, MM90=mm90, D1=d1, DOW=dow_mean, HOJE=hoje,
                    zeros_consec=zeros_consec, quedas50_consec=quedas50_consec)

    @staticmethod
    def _to_float(value: Any) -> Optional[float]:
        """Converte valor para float com tratamento de NaN."""
        try:
            if pd.isna(value):
                return None
        except TypeError:
            pass
        try:
            return float(value)
        except (TypeError, ValueError):
            return None

    @staticmethod
    def classificar(row: pd.Series) -> dict:
        """Aplica as regras do anexo e retorna métricas + 'Risco_Diario' e 'Recuperacao'."""
        s = RiskEngine._serie_diaria_from_json(row.get("Dados_Diarios_2025", "{}"))
        if s.empty:
            return {}
        ref_date = RiskEngine._last_business_day()
        s = RiskEngine._serie_business_day(s, ref_date)
        if s.empty:
            return {}
        m = RiskEngine._rolling_means(s, ref_date)
        hoje, d1 = m["HOJE"], m["D1"]
        mm7, mm30, mm90, dow = m["MM7"], m["MM30"], m["MM90"], m["DOW"]

        def pct(a, b):
            return (a - b) / b * 100 if b and b != 0 else 0.0

        # Lógica híbrida para Delta D-1: usar MM7 como fallback quando D-1 = 0
        if d1 > 0:
            d_vs_d1 = pct(hoje, d1)
        elif d1 == 0 and hoje == 0:
            d_vs_d1 = 0.0
        elif d1 == 0 and hoje > 0:
            # Fallback: usar MM7 como referência quando D-1 está zerado
            d_vs_d1 = pct(hoje, mm7) if mm7 > 0 else 0.0
        else:
            d_vs_d1 = 0.0
        
        d_vs_mm7 = pct(hoje, mm7)
        d_vs_mm30 = pct(hoje, mm30)
        d_vs_mm90 = pct(hoje, mm90)

        mm7_br = RiskEngine._to_float(row.get("MM7_BR"))
        mm7_uf = RiskEngine._to_float(row.get("MM7_UF"))
        mm7_cidade = RiskEngine._to_float(row.get("MM7_CIDADE"))
        contexto_mm = [mm for mm in [mm7_br, mm7_uf, mm7_cidade] if mm is not None and mm > 0]

        reducoes = []
        for mm_ctx in contexto_mm:
            reducao = 1 - (hoje / mm_ctx) if mm_ctx > 0 else 0
            reducoes.append(max(0.0, reducao))
        maior_reducao = max(reducoes) if reducoes else 0.0
        reducao_zero_absoluto = any(mm_ctx > 0 and hoje == 0 for mm_ctx in contexto_mm)

        risco = "🟢 Normal"
        limiar_medio = REDUCAO_MEDIO_RISCO
        limiar_alto = REDUCAO_ALTO_RISCO

        if reducao_zero_absoluto or maior_reducao >= 1.0 or m["zeros_consec"] >= 7 or m["quedas50_consec"] >= 3:
            risco = "⚫ Crítico"
        elif maior_reducao >= limiar_alto:
            risco = "🔴 Alto"
        elif maior_reducao >= limiar_medio:
            risco = "🟠 Moderado"
        elif maior_reducao > 0:
            risco = "🟡 Atenção"
        else:
            risco = "🟢 Normal"

        recuperacao = False
        ultimos_4 = s.loc[:ref_date].tail(4)
        if len(ultimos_4) == 4 and hoje >= mm7 and (ultimos_4.iloc[:3].mean() < 0.9 * mm7):
            recuperacao = True
        return {
            "Vol_Hoje": int(hoje), "Vol_D1": int(d1),
            "MM7": round(mm7, 3), "MM30": round(mm30, 3), "MM90": round(mm90, 3), "DOW_Media": round(dow, 1),
            "Delta_D1": round(d_vs_d1, 1), "Delta_MM7": round(d_vs_mm7, 1),
            "Delta_MM30": round(d_vs_mm30, 1), "Delta_MM90": round(d_vs_mm90, 1),
            "Risco_Diario": risco, "Recuperacao": recuperacao
        }


class VIPManager:
    """Gerenciador de dados VIP."""
    @staticmethod
    def buscar_info_vip(cnpj: str, df_vip: pd.DataFrame) -> Optional[dict]:
        """Busca informações VIP para um CNPJ (apenas ranking e rede, sem contato)."""
        if df_vip is None or df_vip.empty or not cnpj:
            return None
     
        cnpj_normalizado = DataManager.normalizar_cnpj(cnpj)
        if not cnpj_normalizado:
            return None
     
        # Buscar match no DataFrame VIP
        match = df_vip[df_vip['CNPJ_Normalizado'] == cnpj_normalizado]
        if not match.empty:
            row = match.iloc[0]
            return {
                'ranking': row.get('Ranking', ''),
                'ranking_rede': row.get('Ranking Rede', ''),
                'rede': row.get('Rede', '')
                # Campos de contato removidos - agora vêm de laboratories.csv
            }
        return None
    
    @staticmethod
    def buscar_info_laboratory(cnpj: str, df_labs: pd.DataFrame) -> Optional[dict]:
        """Busca informações de contato do laboratories.csv por CNPJ."""
        if df_labs is None or df_labs.empty or not cnpj:
            return None
        
        cnpj_normalizado = DataManager.normalizar_cnpj(cnpj)
        if not cnpj_normalizado:
            return None
        
        # Buscar match no DataFrame de laboratories
        match = df_labs[df_labs['CNPJ_Normalizado'] == cnpj_normalizado]
        if match.empty:
            return None
        
        row = match.iloc[0]
        
        # Função auxiliar para extrair valores de campos aninhados ou flattenados
        def extrair_de_dict(valor_dict, chave, subchave=None):
            """Extrai valor de dict aninhado ou string JSON parseada."""
            if pd.isna(valor_dict) or valor_dict == '':
                return ''
            
            # Se é string, tentar parsear como JSON
            if isinstance(valor_dict, str):
                try:
                    import json
                    import re
                    # Limpar ObjectId e outras strings não-JSON
                    valor_limpo = re.sub(r'ObjectId\([^)]+\)', 'null', valor_dict)
                    valor_limpo = valor_limpo.replace("'", '"')
                    valor_dict = json.loads(valor_limpo)
                except:
                    return ''
            
            # Se é dict, extrair valor
            if isinstance(valor_dict, dict):
                if subchave:
                    # Acessar chave.subchave (ex: contact.telephone)
                    if chave in valor_dict:
                        sub_dict = valor_dict.get(chave, {})
                        if isinstance(sub_dict, dict) and subchave in sub_dict:
                            valor = sub_dict.get(subchave, '')
                            return str(valor).strip() if valor else ''
                else:
                    # Acessar chave diretamente
                    if chave in valor_dict:
                        valor = valor_dict.get(chave, '')
                        return str(valor).strip() if valor else ''
            
            return ''
        
        # Extrair nome do contato (preferência: director.name, fallback: manager.name)
        contato = ''
        
        # Tentar director.name
        if 'director' in df_labs.columns:
            director_data = row.get('director', '')
            contato = extrair_de_dict(director_data, 'name')
        
        # Tentar manager.name como fallback
        if not contato and 'manager' in df_labs.columns:
            manager_data = row.get('manager', '')
            contato = extrair_de_dict(manager_data, 'name')
        
        # Tentar colunas flattenadas (caso o CSV tenha sido flattenado)
        if not contato and 'director.name' in df_labs.columns:
            contato = str(row.get('director.name', '')).strip()
        if not contato and 'manager.name' in df_labs.columns:
            contato = str(row.get('manager.name', '')).strip()
        
        # Extrair telefone de contact.telephone
        telefone = ''
        if 'contact' in df_labs.columns:
            contact_data = row.get('contact', '')
            telefone = extrair_de_dict(contact_data, 'telephone')
        
        # Tentar coluna flattenada
        if not telefone and 'contact.telephone' in df_labs.columns:
            telefone = str(row.get('contact.telephone', '')).strip()
        
        # Extrair email de contact.email
        email = ''
        if 'contact' in df_labs.columns:
            contact_data = row.get('contact', '')
            email = extrair_de_dict(contact_data, 'email')
        
        # Tentar coluna flattenada
        if not email and 'contact.email' in df_labs.columns:
            email = str(row.get('contact.email', '')).strip()
        
        # Função auxiliar para extrair array de strings
        def extrair_array(campo_dict, chave_array):
            """Extrai array de strings de um dict aninhado ou string JSON."""
            if pd.isna(campo_dict) or campo_dict == '':
                return []
            
            # Se é string, tentar parsear como JSON
            if isinstance(campo_dict, str):
                try:
                    import json
                    import re
                    valor_limpo = re.sub(r'ObjectId\([^)]+\)', 'null', campo_dict)
                    valor_limpo = valor_limpo.replace("'", '"')
                    campo_dict = json.loads(valor_limpo)
                except:
                    return []
            
            # Se é dict, extrair array
            if isinstance(campo_dict, dict):
                if chave_array in campo_dict:
                    array_data = campo_dict.get(chave_array, [])
                    if isinstance(array_data, list):
                        return [str(item).strip() for item in array_data if item]
                    elif isinstance(array_data, str):
                        try:
                            import json
                            return json.loads(array_data)
                        except:
                            return [array_data] if array_data else []
            
            return []
        
        # Função auxiliar para extrair campos booleanos True
        def extrair_booleanos(campo_dict, prefixo):
            """Extrai lista de chaves onde valor é True."""
            lista = []
            if pd.isna(campo_dict) or campo_dict == '':
                return lista
            
            # Se é string, tentar parsear como JSON
            if isinstance(campo_dict, str):
                try:
                    import json
                    import re
                    valor_limpo = re.sub(r'ObjectId\([^)]+\)', 'null', campo_dict)
                    valor_limpo = valor_limpo.replace("'", '"')
                    campo_dict = json.loads(valor_limpo)
                except:
                    return lista
            
            # Se é dict, extrair booleanos True
            if isinstance(campo_dict, dict):
                for chave, valor in campo_dict.items():
                    if valor is True or (isinstance(valor, str) and valor.lower() == 'true'):
                        lista.append(chave)
            
            return lista
        
        # Extrair endereço completo (address)
        endereco_completo = {
            'postalCode': '',
            'address': '',
            'addressComplement': '',
            'number': '',
            'neighbourhood': '',
            'city': '',
            'state_code': '',
            'state_name': ''
        }
        
        if 'address' in df_labs.columns:
            address_data = row.get('address', '')
            
            # Extrair campos do endereço usando função auxiliar
            campos_endereco = ['postalCode', 'address', 'addressComplement', 'number', 'neighbourhood', 'city']
            for campo in campos_endereco:
                valor = extrair_de_dict(address_data, campo)
                if valor:
                    endereco_completo[campo] = valor
            
            # Tentar colunas flattenadas
            for campo in campos_endereco:
                coluna_flatten = f'address.{campo}'
                if coluna_flatten in df_labs.columns and not endereco_completo[campo]:
                    valor = str(row.get(coluna_flatten, '')).strip()
                    if valor and valor.lower() != 'nan':
                        endereco_completo[campo] = valor
            
            # Extrair state (objeto aninhado)
            if isinstance(address_data, str):
                try:
                    import json
                    import re
                    valor_limpo = re.sub(r'ObjectId\([^)]+\)', 'null', address_data)
                    valor_limpo = valor_limpo.replace("'", '"')
                    address_dict = json.loads(valor_limpo)
                    if isinstance(address_dict, dict) and 'state' in address_dict:
                        state_data = address_dict.get('state', {})
                        if isinstance(state_data, dict):
                            endereco_completo['state_code'] = str(state_data.get('code', '')).strip()
                            endereco_completo['state_name'] = str(state_data.get('name', '')).strip()
                except:
                    pass
            elif isinstance(address_data, dict) and 'state' in address_data:
                state_data = address_data.get('state', {})
                if isinstance(state_data, dict):
                    endereco_completo['state_code'] = str(state_data.get('code', '')).strip()
                    endereco_completo['state_name'] = str(state_data.get('name', '')).strip()
            
            # Tentar colunas flattenadas para state
            if 'address.state.code' in df_labs.columns and not endereco_completo['state_code']:
                endereco_completo['state_code'] = str(row.get('address.state.code', '')).strip()
            if 'address.state.name' in df_labs.columns and not endereco_completo['state_name']:
                endereco_completo['state_name'] = str(row.get('address.state.name', '')).strip()
        
        # Extrair dados de logistic (days, openingHours, comments)
        logistic_data = {
            'days': [],
            'openingHours': '',
            'comments': ''
        }
        
        # Tentar colunas flattenadas primeiro
        if 'logistic.days' in df_labs.columns:
            days_data = row.get('logistic.days', '')
            if isinstance(days_data, list):
                logistic_data['days'] = [str(item).strip() for item in days_data if item]
            elif isinstance(days_data, str) and days_data:
                try:
                    import json
                    parsed = json.loads(days_data)
                    if isinstance(parsed, list):
                        logistic_data['days'] = [str(item).strip() for item in parsed if item]
                except:
                    logistic_data['days'] = [days_data.strip()] if days_data.strip() else []
        
        if 'logistic.openingHours' in df_labs.columns:
            valor = row.get('logistic.openingHours', '')
            if pd.notna(valor) and str(valor).strip() != '' and str(valor).strip().lower() != 'nan':
                logistic_data['openingHours'] = str(valor).strip()
        
        if 'logistic.comments' in df_labs.columns:
            valor = row.get('logistic.comments', '')
            if pd.notna(valor) and str(valor).strip() != '' and str(valor).strip().lower() != 'nan':
                logistic_data['comments'] = str(valor).strip()
        
        # Fallback: tentar objeto aninhado
        if 'logistic' in df_labs.columns:
            logistic_dict = row.get('logistic', '')
            
            if isinstance(logistic_dict, str) and logistic_dict:
                try:
                    import json
                    import re
                    valor_limpo = re.sub(r'ObjectId\([^)]+\)', 'null', logistic_dict)
                    valor_limpo = valor_limpo.replace("'", '"')
                    logistic_dict = json.loads(valor_limpo)
                except:
                    pass
            
            if isinstance(logistic_dict, dict):
                if not logistic_data['days'] and 'days' in logistic_dict:
                    days_val = logistic_dict.get('days', [])
                    if isinstance(days_val, list):
                        logistic_data['days'] = [str(item).strip() for item in days_val if item]
                
                if not logistic_data['openingHours'] and 'openingHours' in logistic_dict:
                    opening_val = logistic_dict.get('openingHours', '')
                    if opening_val:
                        logistic_data['openingHours'] = str(opening_val).strip()
                
                if not logistic_data['comments'] and 'comments' in logistic_dict:
                    comments_val = logistic_dict.get('comments', '')
                    if comments_val:
                        logistic_data['comments'] = str(comments_val).strip()
        
        # Extrair licensed (booleanos) - tentar todas as possibilidades
        licensed_list = []
        campos_licensed = ['clt', 'cnh', 'cltCnh', 'other', 'online', 'civilService', 'civilServiceAnalysis50', 'otherAnalysis50']
        
        def valor_eh_true(valor):
            """Verifica se um valor deve ser considerado True."""
            if pd.isna(valor) or valor == '':
                return False
            if valor is True:
                return True
            if isinstance(valor, bool):
                return valor
            if isinstance(valor, str):
                return valor.lower() in ['true', '1', 'yes', 't', 'y']
            if isinstance(valor, (int, float)):
                return valor != 0 and valor != 0.0
            return False
        
        # Primeiro: tentar colunas flattenadas com ponto
        for campo in campos_licensed:
            coluna_flatten = f'licensed.{campo}'
            if coluna_flatten in df_labs.columns:
                valor = row.get(coluna_flatten, False)
                if valor_eh_true(valor):
                    licensed_list.append(campo)
        
        # Segundo: tentar sem ponto (caso já esteja flattenado no CSV)
        if not licensed_list:
            for campo in campos_licensed:
                if campo in df_labs.columns:
                    valor = row.get(campo, False)
                    if valor_eh_true(valor):
                        licensed_list.append(campo)
        
        # Terceiro: tentar objeto aninhado (dict ou string JSON)
        # Verificar se coluna licensed existe
        if 'licensed' in df_labs.columns and not licensed_list:
            licensed_dict = row.get('licensed', '')
            
            # Se é string, tentar parsear
            if isinstance(licensed_dict, str) and licensed_dict and licensed_dict.strip():
                try:
                    import json
                    import re
                    import ast
                    # Tentar parsear como JSON
                    valor_limpo = re.sub(r'ObjectId\([^)]+\)', 'null', licensed_dict)
                    valor_limpo = valor_limpo.replace("'", '"')
                    licensed_dict = json.loads(valor_limpo)
                except:
                    try:
                        # Tentar ast.literal_eval (mais seguro que eval)
                        licensed_dict = ast.literal_eval(licensed_dict)
                    except:
                        try:
                            # Última tentativa: eval (menos seguro, mas necessário em alguns casos)
                            licensed_dict = eval(licensed_dict)
                        except:
                            pass
            
            # Se é dict, extrair booleanos True
            if isinstance(licensed_dict, dict):
                for campo in campos_licensed:
                    if campo in licensed_dict:
                        valor = licensed_dict.get(campo, False)
                        if valor_eh_true(valor):
                            if campo not in licensed_list:
                                licensed_list.append(campo)
        
        # Quarto: tentar verificar todas as colunas que contenham "licensed" no nome
        if not licensed_list:
            for col in df_labs.columns:
                if 'licensed' in str(col).lower():
                    # Tentar extrair valor da coluna
                    valor_col = row.get(col, '')
                    # Se a coluna contém um dict/string, tentar parsear
                    if isinstance(valor_col, str) and valor_col and '{' in valor_col:
                        try:
                            import ast
                            valor_col = ast.literal_eval(valor_col)
                        except:
                            pass
                    # Se agora é dict, processar
                    if isinstance(valor_col, dict):
                        for campo in campos_licensed:
                            if campo in valor_col and valor_eh_true(valor_col.get(campo)):
                                if campo not in licensed_list:
                                    licensed_list.append(campo)
        
        # Extrair allowedMethods (booleanos) - tentar todas as possibilidades
        allowed_methods_list = []
        campos_methods = ['cash', 'credit', 'debit', 'billing_laboratory', 'billing_company', 'billing', 'bank_billet', 'eCredit', 'pix']
        
        # Primeiro: tentar colunas flattenadas com ponto
        for campo in campos_methods:
            coluna_flatten = f'allowedMethods.{campo}'
            if coluna_flatten in df_labs.columns:
                valor = row.get(coluna_flatten, False)
                if valor_eh_true(valor):
                    allowed_methods_list.append(campo)
        
        # Segundo: tentar sem ponto (caso já esteja flattenado no CSV)
        if not allowed_methods_list:
            for campo in campos_methods:
                if campo in df_labs.columns:
                    valor = row.get(campo, False)
                    if valor_eh_true(valor):
                        allowed_methods_list.append(campo)
        
        # Terceiro: tentar objeto aninhado (dict ou string JSON)
        # Verificar se coluna allowedMethods existe
        if 'allowedMethods' in df_labs.columns and not allowed_methods_list:
            allowed_methods_dict = row.get('allowedMethods', '')
            
            # Se é string, tentar parsear
            if isinstance(allowed_methods_dict, str) and allowed_methods_dict and allowed_methods_dict.strip():
                try:
                    import json
                    import re
                    import ast
                    # Tentar parsear como JSON
                    valor_limpo = re.sub(r'ObjectId\([^)]+\)', 'null', allowed_methods_dict)
                    valor_limpo = valor_limpo.replace("'", '"')
                    allowed_methods_dict = json.loads(valor_limpo)
                except:
                    try:
                        # Tentar ast.literal_eval (mais seguro que eval)
                        allowed_methods_dict = ast.literal_eval(allowed_methods_dict)
                    except:
                        try:
                            # Última tentativa: eval (menos seguro, mas necessário em alguns casos)
                            allowed_methods_dict = eval(allowed_methods_dict)
                        except:
                            pass
            
            # Se é dict, extrair booleanos True
            if isinstance(allowed_methods_dict, dict):
                for campo in campos_methods:
                    if campo in allowed_methods_dict:
                        valor = allowed_methods_dict.get(campo, False)
                        if valor_eh_true(valor):
                            if campo not in allowed_methods_list:
                                allowed_methods_list.append(campo)
        
        # Quarto: tentar verificar todas as colunas que contenham "allowedmethods" ou "allowed_methods" no nome
        if not allowed_methods_list:
            for col in df_labs.columns:
                col_lower = str(col).lower()
                if 'allowedmethods' in col_lower or 'allowed_methods' in col_lower:
                    # Tentar extrair valor da coluna
                    valor_col = row.get(col, '')
                    # Se a coluna contém um dict/string, tentar parsear
                    if isinstance(valor_col, str) and valor_col and '{' in valor_col:
                        try:
                            import ast
                            valor_col = ast.literal_eval(valor_col)
                        except:
                            pass
                    # Se agora é dict, processar
                    if isinstance(valor_col, dict):
                        for campo in campos_methods:
                            if campo in valor_col and valor_eh_true(valor_col.get(campo)):
                                if campo not in allowed_methods_list:
                                    allowed_methods_list.append(campo)
        
        return {
            'contato': contato if contato else '',
            'telefone': telefone if telefone else '',
            'email': email if email else '',
            'endereco': endereco_completo,
            'logistic': logistic_data,
            'licensed': licensed_list,
            'allowedMethods': allowed_methods_list
        }
def _formatar_df_exibicao(df: pd.DataFrame) -> pd.DataFrame:
    """Padroniza exibição: números sem NaN/None (0), textos como '—'."""
    if df is None or df.empty:
        return df
    df_fmt = df.copy()
    for col in df_fmt.columns:
        if pd.api.types.is_numeric_dtype(df_fmt[col]):
            df_fmt[col] = pd.to_numeric(df_fmt[col], errors='coerce').fillna(0)
        else:
            df_fmt[col] = df_fmt[col].astype(object)
            df_fmt[col] = df_fmt[col].where(df_fmt[col].notna(), '—')
    return df_fmt


DETALHE_QUERY_PARAM = "detalhe_lab"
DETALHES_COLUMN_CONFIG = st.column_config.TextColumn(
    "Selecionar",
    width="small",
    help="Use a caixa de seleção da tabela para abrir a Análise Detalhada do laboratório."
)
VARIACAO_QUEDA_FAIXAS = {
    "Acima de 50%": (50, None),
    "Entre 40% e 50%": (40, 50),
    "Entre 30% e 40%": (30, 40),
    "Entre 20% e 30%": (20, 30),
    "Abaixo de 20%": (None, 20),
}


def calcular_variacao_percentual(valor_atual: Optional[float], valor_anterior: Optional[float]) -> Optional[float]:
    """Retorna variação percentual (atual vs anterior).
    
    Retorna None quando:
    - Valores são inválidos ou não podem ser convertidos para float
    - Valor anterior é NaN, None, zero ou negativo (não há referência válida)
    - Valor atual é NaN ou None
    """
    try:
        atual = float(valor_atual) if valor_atual is not None else None
        anterior = float(valor_anterior) if valor_anterior is not None else None
    except (TypeError, ValueError):
        return None
    
    # Validar valores
    if atual is None or pd.isna(atual):
        return None
    if anterior is None or pd.isna(anterior):
        return None
    
    # Valor anterior deve ser > 0 para cálculo percentual válido
    if anterior <= 0:
        return None
    
    return (atual - anterior) / anterior * 100


def calcular_queda_percentual(valor_atual: Optional[float], valor_anterior: Optional[float]) -> Optional[float]:
    """Retorna magnitude da queda em % (sempre positiva)."""
    variacao = calcular_variacao_percentual(valor_atual, valor_anterior)
    if variacao is None:
        return None
    return max(0.0, -variacao)


def _normalizar_cnpj_str(cnpj: Optional[str]) -> str:
    if cnpj is None:
        return ""
    return ''.join(filter(str.isdigit, str(cnpj)))


def _build_detalhe_url(cnpj: Optional[str]) -> str:
    cnpj_norm = _normalizar_cnpj_str(cnpj)
    if not cnpj_norm:
        return ""
    return f"?{DETALHE_QUERY_PARAM}={quote_plus(cnpj_norm)}"


def adicionar_coluna_detalhes(df: pd.DataFrame, cnpj_col: str = 'CNPJ_Normalizado') -> pd.DataFrame:
    """Mantém DataFrame como está; navegação ocorre via seleção de linha/checkbox."""
    return df


def preparar_dataframe_risco(df: pd.DataFrame) -> pd.DataFrame:
    """Garante colunas derivadas para a lista de risco."""
    if df is None or df.empty:
        return df
    work = df.copy()
    if 'CNPJ_Normalizado' not in work.columns and 'CNPJ_PCL' in work.columns:
        work['CNPJ_Normalizado'] = work['CNPJ_PCL'].apply(DataManager.normalizar_cnpj)

    for col in [
        'WoW_Semana_Atual',
        'WoW_Semana_Anterior',
        'Media_Semanal_2025',
        'Media_Semanal_2024',
        'Controle_Semanal_Estado_Atual',
        'Controle_Semanal_Estado_Anterior'
    ]:
        if col in work.columns:
            work[col] = pd.to_numeric(work[col], errors='coerce')
    
    # Calcular Media_Semanal_2024 se não existir
    if 'Media_Semanal_2024' not in work.columns and 'Total_Coletas_2024' in work.columns:
        work['Media_Semanal_2024'] = pd.to_numeric(work['Total_Coletas_2024'], errors='coerce') / 52
        work['Media_Semanal_2024'] = work['Media_Semanal_2024'].fillna(0.0)

    # Calcular média mensal dos top 3 meses de 2024
    meses_nomes = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
    colunas_2024 = [f'N_Coletas_{m}_24' for m in meses_nomes]
    colunas_2024_existentes = [col for col in colunas_2024 if col in work.columns]
    
    if colunas_2024_existentes:
        def calcular_media_top3(row, colunas):
            """Calcula a média dos top 3 meses com coletas > 0."""
            valores = []
            for col in colunas:
                val = pd.to_numeric(row.get(col, 0), errors='coerce')
                if pd.notna(val) and val > 0:
                    valores.append(val)
            
            if not valores:
                return 0.0
            
            # Ordenar decrescente e pegar top 3 (ou menos se não tiver 3)
            valores_ordenados = sorted(valores, reverse=True)
            top3 = valores_ordenados[:3]
            return sum(top3) / len(top3)
        
        work['Media_Mensal_Top3_2024'] = work.apply(
            lambda row: calcular_media_top3(row, colunas_2024_existentes), 
            axis=1
        )
    else:
        work['Media_Mensal_Top3_2024'] = 0.0
    
    # Calcular média mensal dos top 3 meses de 2025 (apenas até o mês atual)
    from datetime import datetime
    mes_atual = datetime.now().month
    meses_2025_disponiveis = meses_nomes[:mes_atual]
    colunas_2025 = [f'N_Coletas_{m}_25' for m in meses_2025_disponiveis]
    colunas_2025_existentes = [col for col in colunas_2025 if col in work.columns]
    
    if colunas_2025_existentes:
        work['Media_Mensal_Top3_2025'] = work.apply(
            lambda row: calcular_media_top3(row, colunas_2025_existentes), 
            axis=1
        )
    else:
        work['Media_Mensal_Top3_2025'] = 0.0

    vol_ant = work['WoW_Semana_Anterior'] if 'WoW_Semana_Anterior' in work.columns else pd.Series(0, index=work.index, dtype=float)
    vol_atual = work['WoW_Semana_Atual'] if 'WoW_Semana_Atual' in work.columns else pd.Series(0, index=work.index, dtype=float)
    
    # Converter para numérico e tratar NaN - mas não preencher com 0 para manter distinção
    vol_ant = pd.to_numeric(vol_ant, errors='coerce')
    vol_atual = pd.to_numeric(vol_atual, errors='coerce')
    
    # Para queda absoluta, usar 0 quando faltar valor
    vol_ant_fill = vol_ant.fillna(0)
    vol_atual_fill = vol_atual.fillna(0)
    work['Queda_Semanal_Abs'] = vol_ant_fill - vol_atual_fill

    # Calcular porcentagens apenas quando temos valores válidos (anterior > 0)
    # Evita calcular 0% ou 0,2% incorretos quando WoW_Semana_Anterior está zerado
    with np.errstate(divide='ignore', invalid='ignore'):
        # Máscara: anterior deve ser > 0 E ambos devem ser não-nulos
        mask_valido = (vol_ant > 0) & vol_ant.notna() & vol_atual.notna()
        
        work['Variacao_Semanal_Pct'] = np.where(
            mask_valido,
            (vol_atual - vol_ant) / vol_ant * 100,
            np.nan
        )
        work['Queda_Semanal_Pct'] = np.where(
            mask_valido,
            (vol_ant - vol_atual) / vol_ant * 100,
            np.nan
        )

    media_estado_ant = work['Controle_Semanal_Estado_Anterior'] if 'Controle_Semanal_Estado_Anterior' in work.columns else pd.Series(0, index=work.index, dtype=float)
    media_estado_atual = work['Controle_Semanal_Estado_Atual'] if 'Controle_Semanal_Estado_Atual' in work.columns else pd.Series(0, index=work.index, dtype=float)
    
    # Converter para numérico e tratar NaN
    media_estado_ant = pd.to_numeric(media_estado_ant, errors='coerce')
    media_estado_atual = pd.to_numeric(media_estado_atual, errors='coerce')
    
    # Calcular porcentagem apenas quando temos valores válidos (anterior > 0)
    with np.errstate(divide='ignore', invalid='ignore'):
        # Máscara: anterior deve ser > 0 E ambos devem ser não-nulos
        mask_estado_valido = (media_estado_ant > 0) & media_estado_ant.notna() & media_estado_atual.notna()
        
        work['Variacao_Media_Estado_Pct'] = np.where(
            mask_estado_valido,
            (media_estado_atual - media_estado_ant) / media_estado_ant * 100,
            np.nan
        )

    # Calcular variação vs média mensal top 3 de 2024
    media_top3_2024 = pd.to_numeric(work['Media_Mensal_Top3_2024'], errors='coerce')
    with np.errstate(divide='ignore', invalid='ignore'):
        mask_top3_2024_valido = (media_top3_2024 > 0) & media_top3_2024.notna() & vol_atual.notna()
        work['Variacao_vs_Top3_2024_Pct'] = np.where(
            mask_top3_2024_valido,
            (vol_atual - media_top3_2024) / media_top3_2024 * 100,
            np.nan
        )
    
    # Calcular variação vs média mensal top 3 de 2025
    media_top3_2025 = pd.to_numeric(work['Media_Mensal_Top3_2025'], errors='coerce')
    with np.errstate(divide='ignore', invalid='ignore'):
        mask_top3_2025_valido = (media_top3_2025 > 0) & media_top3_2025.notna() & vol_atual.notna()
        work['Variacao_vs_Top3_2025_Pct'] = np.where(
            mask_top3_2025_valido,
            (vol_atual - media_top3_2025) / media_top3_2025 * 100,
            np.nan
        )

    if 'Risco_Por_Dias_Sem_Coleta' in work.columns:
        risco_dias = work['Risco_Por_Dias_Sem_Coleta'].fillna(False)
    else:
        risco_dias = pd.Series(False, index=work.index)

    # Risco por queda: apenas considerar quando Queda_Semanal_Pct é válido (não NaN) e >= 50%
    # Não considerar NaN como 0% (evita falsos positivos quando vol_ant está zerado)
    queda_valida = work['Queda_Semanal_Pct'].notna() & (work['Queda_Semanal_Pct'] >= 50)
    
    # Calcular perdas PRIMEIRO (recentes + antigas) para excluir da classificação de risco
    cnpjs_perdas = set()
    if 'Classificacao_Perda_V2' in work.columns:
        # Identificar CNPJs de perdas recentes e antigas
        perdas_recentes = work[work['Classificacao_Perda_V2'] == 'Perda Recente']
        perdas_antigas = work[work['Classificacao_Perda_V2'].isin(['Perda Antiga', 'Perda Consolidada'])]
        
        if 'CNPJ_Normalizado' in work.columns:
            cnpjs_perdas = set(perdas_recentes['CNPJ_Normalizado'].dropna()) | set(perdas_antigas['CNPJ_Normalizado'].dropna())
    
    # Calcular Em_Risco normalmente
    work['Em_Risco'] = risco_dias | queda_valida
    
    # Ao classificar risco, pule se for perda (forçar Em_Risco = False para perdas)
    if cnpjs_perdas and 'CNPJ_Normalizado' in work.columns:
        mask_perdas = work['CNPJ_Normalizado'].isin(cnpjs_perdas)
        work.loc[mask_perdas, 'Em_Risco'] = False
    
    return work


def aplicar_filtro_variacao(df_src: pd.DataFrame, faixas: List[str]) -> pd.DataFrame:
    """Filtra DataFrame conforme faixas de queda percentual."""
    if df_src is None or df_src.empty or not faixas:
        return df_src
    mask_total = pd.Series(False, index=df_src.index)
    serie_queda = df_src['Queda_Semanal_Pct'].fillna(np.nan)
    for faixa in faixas:
        limites = VARIACAO_QUEDA_FAIXAS.get(faixa)
        if not limites:
            continue
        min_val, max_val = limites
        cond = serie_queda.notna()
        if min_val is not None:
            cond &= serie_queda >= min_val
        if max_val is not None:
            cond &= serie_queda < max_val
        mask_total = mask_total | cond
    if not mask_total.any():
        return df_src.iloc[0:0]
    return df_src[mask_total]


def aplicar_filtro_variacao_generica(
    df_src: pd.DataFrame,
    coluna: str,
    faixas: List[str],
    usar_valor_absoluto: bool = True
) -> pd.DataFrame:
    """Filtra DataFrame por faixas de variação usando a mesma tabela de ranges do WoW."""
    if df_src is None or df_src.empty or not faixas or coluna not in df_src.columns:
        return df_src
    serie = pd.to_numeric(df_src[coluna], errors='coerce')
    mask_total = pd.Series(False, index=df_src.index)
    for faixa in faixas:
        limites = VARIACAO_QUEDA_FAIXAS.get(faixa)
        if not limites:
            continue
        min_val, max_val = limites
        valores = serie.abs() if usar_valor_absoluto else serie
        cond = valores.notna()
        if min_val is not None:
            cond &= valores >= min_val
        if max_val is not None:
            cond &= valores < max_val
        mask_total = mask_total | cond
    if not mask_total.any():
        return df_src.iloc[0:0]
    return df_src[mask_total]


# ============================================
# FUNÇÕES DE FECHAMENTO SEMANAL E MENSAL
# ============================================

@st.cache_data(ttl=300)
def calcular_metricas_fechamento_semanal(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Calcula métricas para fechamento semanal do mês corrente.
    
    Args:
        df: DataFrame com dados de churn
        
    Returns:
        Dicionário com métricas semanais agregadas
    """
    from datetime import datetime
    import calendar
    
    metricas = {
        'total_semanas': 0,
        'semanas_fechadas': 0,
        'volume_atual': 0,
        'volume_anterior': 0,
        'wow_medio': 0.0,
        'volume_semana_atual_total': 0,
        'volume_semana_atual_sem_risco': 0,
        'labs_semana_processados': 0,
        'semanas_detalhes': [],
        'media_semanal_2024': 0.0,
        'media_semanal_2025': 0.0,
        'medias_por_uf': {},
        'labs_com_queda_wow': []
    }

    resumo_cards = {
        'volume_semana_atual': 0,
        'volume_semana_anterior': 0,
        'media_semana_atual': 0.0,
        'media_semana_anterior': 0.0,
        'variacao_media_pct': None,
        'total_labs': 0
    }
    
    if df.empty:
        metricas['resumo_cards'] = resumo_cards
        return metricas

    serie_atual = (
        pd.to_numeric(df['WoW_Semana_Atual'], errors='coerce')
        if 'WoW_Semana_Atual' in df.columns else pd.Series(0, index=df.index, dtype=float)
    ).fillna(0)
    serie_anterior = (
        pd.to_numeric(df['WoW_Semana_Anterior'], errors='coerce')
        if 'WoW_Semana_Anterior' in df.columns else pd.Series(0, index=df.index, dtype=float)
    ).fillna(0)
    resumo_cards.update({
        'volume_semana_atual': float(serie_atual.sum()),
        'volume_semana_anterior': float(serie_anterior.sum()),
        'media_semana_atual': float(serie_atual.mean()) if len(serie_atual) else 0.0,
        'media_semana_anterior': float(serie_anterior.mean()) if len(serie_anterior) else 0.0,
        'total_labs': int(df['CNPJ_Normalizado'].nunique()) if 'CNPJ_Normalizado' in df.columns else len(df)
    })
    resumo_cards['variacao_media_pct'] = calcular_variacao_percentual(
        resumo_cards['media_semana_atual'],
        resumo_cards['media_semana_anterior']
    )
    metricas['resumo_cards'] = resumo_cards
    
    # Calcular métricas básicas diretamente do DataFrame
    hoje = datetime.now()
    mes_atual = hoje.month
    ano_atual = hoje.year
    
    # Calcular total de semanas no mês atual
    primeiro_dia = datetime(ano_atual, mes_atual, 1)
    ultimo_dia = datetime(ano_atual, mes_atual, calendar.monthrange(ano_atual, mes_atual)[1])
    metricas['total_semanas'] = (ultimo_dia.isocalendar()[1] - primeiro_dia.isocalendar()[1]) + 1
    
    # Semanas fechadas = semanas completas até hoje
    semana_hoje = hoje.isocalendar()[1]
    semana_inicio_mes = primeiro_dia.isocalendar()[1]
    metricas['semanas_fechadas'] = max(0, semana_hoje - semana_inicio_mes)
    
    # Calcular médias históricas
    if 'Total_Coletas_2024' in df.columns:
        total_2024 = df['Total_Coletas_2024'].sum()
        metricas['media_semanal_2024'] = float(total_2024 / 52) if total_2024 > 0 else 0.0
    
    if 'Total_Coletas_2025' in df.columns:
        total_2025 = df['Total_Coletas_2025'].sum()
        semanas_decorridas_2025 = semana_hoje
        metricas['media_semanal_2025'] = float(total_2025 / semanas_decorridas_2025) if semanas_decorridas_2025 > 0 else 0.0
    
    # Tentar carregar metadados de fechamento (se existirem)
    try:
        import os
        meta_path = os.path.join(OUTPUT_DIR, "fechamentos_meta.json")
        
        # Tentar baixar do SharePoint se não existir localmente
        if not os.path.exists(meta_path):
            try:
                cfg = _get_graph_config()
                if cfg:
                    arquivo_remoto = cfg.get("arquivo", "Data Analysis/Churn PCLs/churn_analysis_latest.csv")
                    remote_dir = os.path.dirname(arquivo_remoto)
                    remote_meta_path = f"{remote_dir}/fechamentos_meta.json" if remote_dir else "fechamentos_meta.json"
                    remote_meta_path = remote_meta_path.replace("\\", "/")
                    baixar_sharepoint(arquivo_remoto=remote_meta_path)
            except Exception as e_download:
                logger.warning(f"Falha ao tentar baixar metadados do SharePoint: {e_download}")

        if os.path.exists(meta_path):
            with open(meta_path, 'r', encoding='utf-8') as f:
                meta_fechamento = json.load(f)
                
            # Sobrescrever com dados do arquivo se disponíveis
            metricas['total_semanas'] = meta_fechamento.get('total_semanas', metricas['total_semanas'])
            metricas['semanas_fechadas'] = meta_fechamento.get('semanas_fechadas', metricas['semanas_fechadas'])
            metricas['semanas_detalhes'] = meta_fechamento.get('weeks', [])
            metricas['media_semanal_2024'] = meta_fechamento.get('media_semanal_pais_2024', metricas['media_semanal_2024'])
            metricas['media_semanal_2025'] = meta_fechamento.get('media_semanal_pais_2025', metricas['media_semanal_2025'])
            metricas['medias_por_uf'] = meta_fechamento.get('media_semanal_por_uf', {})
    except Exception as e:
        logger.warning(f"Metadados de fechamento não disponíveis, usando cálculos básicos: {e}")
    
    # FALLBACK: Se semanas_detalhes estiver vazio (comum em PROD/SharePoint), reconstruir a partir da coluna Semanas_Mes_Atual
    if not metricas['semanas_detalhes'] and 'Semanas_Mes_Atual' in df.columns:
        try:
            # Dicionário para agregar: (iso_year, iso_week) -> {dados}
            agg_weeks = {}
            
            for val in df['Semanas_Mes_Atual']:
                if pd.isna(val): continue
                try:
                    # Parsear JSON (se for string) ou usar direto se já for lista
                    dados_lab = json.loads(val) if isinstance(val, str) else val
                    if not isinstance(dados_lab, list): continue
                    
                    for semana in dados_lab:
                        iso_year = semana.get('iso_year')
                        iso_week = semana.get('iso_week')
                        
                        if not iso_year or not iso_week:
                            continue
                            
                        key = (iso_year, iso_week)
                        if key not in agg_weeks:
                            agg_weeks[key] = {
                                'semana': semana.get('semana'),
                                'iso_week': iso_week,
                                'iso_year': iso_year,
                                'volume_total': 0,
                                'volume_semana_anterior': 0,
                                'fechada': semana.get('fechada', False)
                            }
                        
                        # Somar volumes
                        vol_util = semana.get('volume_util')
                        if vol_util:
                            agg_weeks[key]['volume_total'] += int(vol_util)
                            
                        vol_ant = semana.get('volume_semana_anterior')
                        if vol_ant:
                            agg_weeks[key]['volume_semana_anterior'] += int(vol_ant)
                        
                except Exception:
                    continue
            
            # Converter para lista ordenada e atualizar métricas
            if agg_weeks:
                weeks_list = sorted(agg_weeks.values(), key=lambda x: (x['iso_year'], x['iso_week']))
                metricas['semanas_detalhes'] = weeks_list
                metricas['total_semanas'] = len(weeks_list)
                metricas['semanas_fechadas'] = sum(1 for w in weeks_list if w['fechada'])
                logger.info(f"Semanas reconstruídas via fallback: {len(weeks_list)} semanas encontradas")
                
        except Exception as e:
            logger.error(f"Erro ao reconstruir semanas_detalhes no fallback: {e}")

    # Calcular volume atual e anterior das semanas
    if metricas['semanas_detalhes']:
        # Última semana (atual ou mais recente)
        ultima_semana = metricas['semanas_detalhes'][-1]
        metricas['volume_atual'] = ultima_semana.get('volume_total', 0)
        metricas['volume_anterior'] = ultima_semana.get('volume_semana_anterior', 0)
        
        # Calcular WoW médio
        if metricas['volume_anterior'] > 0:
            metricas['wow_medio'] = ((metricas['volume_atual'] - metricas['volume_anterior']) / 
                                     metricas['volume_anterior'] * 100)
    
    # Parsear coluna Semanas_Mes_Atual para identificar labs com queda WoW > 50%
    if 'Semanas_Mes_Atual' in df.columns and 'Nome_Fantasia_PCL' in df.columns:
        for _, row in df.iterrows():
            try:
                semanas_json = row.get('Semanas_Mes_Atual', '[]')
                if isinstance(semanas_json, str) and semanas_json and semanas_json != '[]':
                    semanas = json.loads(semanas_json)
                    if semanas:
                        # Verificar última semana do lab
                        ultima_semana_lab = semanas[-1]
                        vol_atual_lab = ultima_semana_lab.get('volume_util', 0)
                        vol_anterior_lab = ultima_semana_lab.get('volume_semana_anterior')
                        
                        metricas['labs_semana_processados'] += 1
                        metricas['volume_semana_atual_total'] += vol_atual_lab or 0
                        if not row.get('Risco_Por_Dias_Sem_Coleta', False):
                            metricas['volume_semana_atual_sem_risco'] += vol_atual_lab or 0
                        
                        if vol_anterior_lab and vol_anterior_lab > 0:
                            wow_pct = ((vol_atual_lab - vol_anterior_lab) / vol_anterior_lab) * 100
                            if wow_pct < -50:  # Queda > 50%
                                # Buscar CNPJ normalizado
                                cnpj_norm = row.get('CNPJ_Normalizado', '')
                                if not cnpj_norm and 'CNPJ_PCL' in row:
                                    cnpj_norm = DataManager.normalizar_cnpj(row.get('CNPJ_PCL', ''))
                                
                                metricas['labs_com_queda_wow'].append({
                                    'nome': row.get('Nome_Fantasia_PCL', 'N/A'),
                                    'cnpj': cnpj_norm,
                                    'vip': row.get('VIP', ''),
                                    'uf': row.get('Estado', 'N/A'),
                                    'representante': row.get('Representante_Nome', 'N/A'),
                                    'wow_pct': wow_pct,
                                    'vol_atual': vol_atual_lab,
                                    'vol_anterior': vol_anterior_lab,
                                    'porte': row.get('Porte', 'N/A'),
                                    'ranking': row.get('Ranking', ''),
                                    'ranking_rede': row.get('Ranking Rede', ''),
                                    'rede': row.get('Rede', ''),
                                    'baseline_mensal': row.get('Baseline_Mensal', 0),
                                    'coletas_mes_atual': row.get('Coletas_Mes_Atual', 0),
                                    'dias_sem_coleta': row.get('Dias_Sem_Coleta_Uteis', row.get('Dias_Sem_Coleta', 0)),
                                    'risco_dias': bool(row.get('Risco_Por_Dias_Sem_Coleta', False)),
                                    'status_risco': row.get('Status_Risco_V2', '—'),
                                    'motivo_risco': row.get('Motivo_Risco_V2', 'Queda WoW > 50%'),
                                    'apareceu_gralab': bool(row.get('Apareceu_Gralab', False)),
                                    'data_ultima_coleta': row.get('Data_Ultima_Coleta', None)
                                })
            except Exception as e:
                continue
    
    # Ordenar labs por WoW (maior queda primeiro)
    metricas['labs_com_queda_wow'].sort(key=lambda x: x.get('wow_pct', 0))
    
    if metricas['labs_semana_processados'] == 0:
        metricas['volume_semana_atual_total'] = metricas['volume_atual']
        metricas['volume_semana_atual_sem_risco'] = metricas['volume_atual']
    
    return metricas


@st.cache_data(ttl=300)
def calcular_metricas_fechamento_mensal(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Calcula métricas para fechamento mensal consolidado.
    
    Args:
        df: DataFrame com dados de churn
        
    Returns:
        Dicionário com métricas mensais agregadas
    """
    metricas = {
        'volume_mes_atual': 0,
        'volume_mes_anterior': 0,
        'baseline_total': 0.0,
        'delta_pct': None,
        'variacao_mes_anterior_pct': None,
        'dia_atual': datetime.now().day,
        'labs_detalhados': [],
        'media_mensal_2024': 0.0,
        'media_mensal_2025': 0.0,
        'medias_por_uf_2024': {},
        'medias_por_uf_2025': {},
        'total_labs': 0
    }
    
    if df.empty:
        return metricas
    
    meses_nomes = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
    hoje = datetime.now()
    mes_atual = hoje.month
    ano_atual = hoje.year
    mes_anterior = 12 if mes_atual == 1 else mes_atual - 1
    ano_mes_anterior = ano_atual - 1 if mes_atual == 1 else ano_atual
    col_mes_atual_hist = f"N_Coletas_{meses_nomes[mes_atual - 1]}_{str(ano_atual)[-2:]}"

    colunas_2024 = [f"N_Coletas_{m}_24" for m in meses_nomes]
    colunas_2025 = [f"N_Coletas_{m}_25" for m in meses_nomes[:mes_atual]]
    colunas_historicas = [c for c in (colunas_2024 + colunas_2025) if c in df.columns]

    totais_historicos: List[float] = []
    for col in colunas_historicas:
        totais_historicos.append(pd.to_numeric(df[col], errors='coerce').sum())

    # Caso o mês corrente ainda não tenha sido fechado nas colunas históricas,
    # usa Coletas_Mes_Atual como fallback para entrar no cálculo da média dos top-N.
    if col_mes_atual_hist not in colunas_historicas and 'Coletas_Mes_Atual' in df.columns:
        totais_historicos.append(pd.to_numeric(df['Coletas_Mes_Atual'], errors='coerce').sum())
    
    # Calcular volume total do mês atual
    if 'Coletas_Mes_Atual' in df.columns:
        metricas['volume_mes_atual'] = int(df['Coletas_Mes_Atual'].sum())
    
    # Calcular volume total do mês anterior (mês fechado)
    sufixo_mes_anterior = str(ano_mes_anterior)[-2:]
    col_mes_anterior = f"N_Coletas_{meses_nomes[mes_anterior - 1]}_{sufixo_mes_anterior}"
    if col_mes_anterior in df.columns:
        metricas['volume_mes_anterior'] = int(pd.to_numeric(df[col_mes_anterior], errors='coerce').sum())
    
    # Calcular baseline total: média dos top-N meses históricos (2024 + 2025)
    if totais_historicos:
        top_n = min(BASELINE_TOP_N, len(totais_historicos))
        metricas['baseline_total'] = float(pd.Series(totais_historicos, dtype=float).nlargest(top_n).mean())
    
    # Delta percentual
    metricas['delta_pct'] = calcular_variacao_percentual(
        metricas['volume_mes_atual'],
        metricas['baseline_total']
    )
    metricas['variacao_mes_anterior_pct'] = calcular_variacao_percentual(
        metricas['volume_mes_atual'],
        metricas['volume_mes_anterior']
    )
    
    # Calcular médias mensais 2024 e 2025
    if 'Total_Coletas_2024' in df.columns:
        metricas['media_mensal_2024'] = float(df['Total_Coletas_2024'].sum() / 12)
    
    if 'Total_Coletas_2025' in df.columns:
        mes_atual = datetime.now().month
        if mes_atual > 0:
            metricas['media_mensal_2025'] = float(df['Total_Coletas_2025'].sum() / mes_atual)
    
    # Médias por UF
    if 'Estado' in df.columns:
        for uf in df['Estado'].unique():
            if pd.notna(uf) and uf != '':
                df_uf = df[df['Estado'] == uf]
                if 'Total_Coletas_2024' in df.columns:
                    metricas['medias_por_uf_2024'][uf] = float(df_uf['Total_Coletas_2024'].sum() / 12)
                if 'Total_Coletas_2025' in df.columns:
                    mes_atual = datetime.now().month
                    if mes_atual > 0:
                        metricas['medias_por_uf_2025'][uf] = float(df_uf['Total_Coletas_2025'].sum() / mes_atual)

    metricas['total_labs'] = int(df['CNPJ_Normalizado'].nunique()) if 'CNPJ_Normalizado' in df.columns else len(df)
    metricas['resumo_mensal'] = {
        'volume_mes_atual': metricas['volume_mes_atual'],
        'volume_mes_anterior': metricas['volume_mes_anterior'],
        'baseline_total': metricas['baseline_total'],
        'variacao_pct': metricas['delta_pct'],
        'variacao_mes_anterior_pct': metricas['variacao_mes_anterior_pct'],
        'total_labs': metricas['total_labs']
    }
    
    # Parsear Baseline_Componentes para detalhamento por laboratório
    colunas_necessarias = ['Nome_Fantasia_PCL', 'Estado', 'Baseline_Componentes', 'Baseline_Mensal',
                          'Coletas_Mes_Atual', 'Queda_Baseline_Pct', 'Representante_Nome', 'Porte']
    
    if all(col in df.columns for col in colunas_necessarias):
        # Filtrar apenas labs com baseline > 0
        df_com_baseline = df[df['Baseline_Mensal'] > 0].copy()
        
        for _, row in df_com_baseline.iterrows():
            try:
                baseline_json = row.get('Baseline_Componentes', '[]')
                baseline_meses = []
                
                if isinstance(baseline_json, str) and baseline_json and baseline_json != '[]':
                    baseline_meses = json.loads(baseline_json)
                
                # Buscar CNPJ normalizado
                cnpj_norm = row.get('CNPJ_Normalizado', '')
                if not cnpj_norm and 'CNPJ_PCL' in row:
                    cnpj_norm = DataManager.normalizar_cnpj(row.get('CNPJ_PCL', ''))
                
                lab_info = {
                    'nome': row.get('Nome_Fantasia_PCL', 'N/A'),
                    'cnpj': cnpj_norm,
                    'vip': row.get('VIP', ''),
                    'uf': row.get('Estado', 'N/A'),
                    'representante': row.get('Representante_Nome', 'N/A'),
                    'porte': row.get('Porte', 'N/A'),
                    'baseline_mensal': float(row.get('Baseline_Mensal', 0)),
                    'coletas_mes_atual': int(row.get('Coletas_Mes_Atual', 0)),
                    'queda_baseline_pct': float(row.get('Queda_Baseline_Pct', 0)),
                    'baseline_meses': baseline_meses,
                    'ranking': row.get('Ranking', ''),
                    'ranking_rede': row.get('Ranking Rede', ''),
                    'rede': row.get('Rede', ''),
                    'apareceu_gralab': bool(row.get('Apareceu_Gralab', False)),
                    'data_ultima_coleta': row.get('Data_Ultima_Coleta', None)
                }
                
                metricas['labs_detalhados'].append(lab_info)
            except Exception as e:
                continue
    
    # Ordenar labs por queda baseline (maior queda primeiro)
    metricas['labs_detalhados'].sort(key=lambda x: x.get('queda_baseline_pct', 0))
    
    return metricas


# ============================================
# NOVAS FUNÇÕES VISUAIS (STORYTELLING)
# ============================================

def highlight_risco_row(row):
    """Aplica cor de fundo sutil apenas para linhas de risco alto."""
    if row.get('Status_Risco_V2') == 'Perda (Risco Alto)':
        return ['background-color: rgba(255, 0, 0, 0.05)'] * len(row)
    return [''] * len(row)

def aplicar_coloracao_variacao_semanal(df: pd.DataFrame) -> pd.DataFrame:
    """
    Aplica coloração condicional visual na coluna Variacao_Semanal_Pct.
    Como o Streamlit não suporta coloração direta em st.dataframe, 
    aplicamos styling via pandas Styler que será exibido separadamente.
    
    Retorna tupla: (DataFrame original, DataFrame estilizado para exibição)
    """
    if 'Variacao_Semanal_Pct' not in df.columns:
        return df
    
    df_copy = df.copy()
    
    # Função para determinar estilo baseado no valor
    def estilo_variacao(valor):
        """Retorna estilo CSS baseado no valor da variação."""
        if pd.isna(valor) or valor is None:
            return ''
        
        try:
            valor_float = float(valor)
            if valor_float >= -20:
                return 'background-color: #d4edda; font-weight: 600;'  # Verde claro
            elif valor_float >= -40:
                return 'background-color: #fff3cd; font-weight: 600;'  # Amarelo
            elif valor_float >= -60:
                return 'background-color: #ffeaa7; font-weight: 600;'  # Laranja
            else:
                return 'background-color: #f8d7da; color: #721c24; font-weight: 700;'  # Vermelho forte
        except (ValueError, TypeError):
            return ''
    
    # Criar coluna auxiliar para styling
    # O styling será aplicado quando exibirmos a tabela
    return df_copy

def criar_styler_com_coloracao(df: pd.DataFrame) -> 'pd.io.formats.style.Styler':
    """
    Cria um pandas Styler com coloração condicional na coluna Variacao_Semanal_Pct.
    """
    if 'Variacao_Semanal_Pct' not in df.columns:
        return df.style
    
    def estilo_variacao(row):
        """Retorna estilo CSS para a célula de Variacao_Semanal_Pct."""
        valor = row.get('Variacao_Semanal_Pct')
        if pd.isna(valor) or valor is None:
            return [''] * len(row)
        
        try:
            valor_float = float(valor)
            estilos = [''] * len(row)
            # Encontrar índice da coluna Variacao_Semanal_Pct
            idx_col = list(row.index).index('Variacao_Semanal_Pct') if 'Variacao_Semanal_Pct' in row.index else -1
            
            if idx_col >= 0:
                if valor_float >= -20:
                    estilos[idx_col] = 'background-color: #d4edda; font-weight: 600;'  # Verde claro
                elif valor_float >= -40:
                    estilos[idx_col] = 'background-color: #fff3cd; font-weight: 600;'  # Amarelo
                elif valor_float >= -60:
                    estilos[idx_col] = 'background-color: #ffeaa7; font-weight: 600;'  # Laranja
                else:
                    estilos[idx_col] = 'background-color: #f8d7da; color: #721c24; font-weight: 700;'  # Vermelho forte
            
            return estilos
        except (ValueError, TypeError):
            return [''] * len(row)
    
    # Aplicar styling
    styled_df = df.style.apply(estilo_variacao, axis=1)
    return styled_df

def formatar_tabela_storytelling(df, config_colunas):
    """Aplica estilo limpo e focado."""
    return st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
        column_config=config_colunas
    )


def _navegar_para_analise_detalhada(cnpj: Optional[str]):
    """Atualiza o estado para abrir a Análise Detalhada do laboratório selecionado."""
    if not cnpj:
        return
    cnpj_str = ''.join(filter(str.isdigit, str(cnpj))) or str(cnpj)
    st.session_state['lab_cnpj_selecionado'] = cnpj_str
    st.session_state['page'] = "📋 Análise Detalhada"
    st.session_state['busca_avancada'] = cnpj_str
    st.rerun()


def _processar_query_param_detalhes():
    """Detecta query param de detalhe e direciona para a análise detalhada."""
    try:
        params = st.experimental_get_query_params()
    except Exception:
        return
    cnpj_param = params.get(DETALHE_QUERY_PARAM)
    if not cnpj_param:
        return
    cnpj_val = _normalizar_cnpj_str(cnpj_param[0])
    st.experimental_set_query_params()
    if cnpj_val:
        _navegar_para_analise_detalhada(cnpj_val)


def _processar_evento_selecao(evento, df_display: pd.DataFrame):
    """Processa seleção de linha nas tabelas interativas - funciona ao clicar na linha ou na lupa."""
    try:
        rows = evento.selection.get("rows", [])  # type: ignore[attr-defined]
    except Exception:
        rows = []
    if not rows:
        return
    idx = rows[0]
    if idx is None:
        return
    if idx >= len(df_display):
        return
    cnpj_val = df_display.iloc[idx].get('CNPJ_Normalizado')
    if cnpj_val:
        _navegar_para_analise_detalhada(cnpj_val)

# ============================================
# ABA 1: FECHAMENTO SEMANAL (TÁTICO)
# ============================================

def renderizar_aba_fechamento_semanal(
    df: pd.DataFrame,
    metrics: KPIMetrics,
    filtros: Dict[str, Any],
    df_total: Optional[pd.DataFrame] = None
):
    st.markdown("## 📅 Fechamento Semanal (Visão Tática)")
    st.caption("Monitoramento de todos os laboratórios. Ordenado por maior queda de volume.")

    # -------------------------------------------------------------
    # Seleção de fechamento semanal (sem “tempo real”)
    # -------------------------------------------------------------
    metricas_sem = calcular_metricas_fechamento_semanal(df)
    met = metricas_sem
    semanas_meta = sorted(
        met.get('semanas_detalhes', []),
        key=lambda x: (x.get('iso_year', 0), x.get('iso_week', 0))
    )
    semanas_map = {
        (sem.get('iso_year'), sem.get('iso_week')): sem
        for sem in semanas_meta
        if sem.get('iso_year') and sem.get('iso_week')
    }

    def _label_semana(sem):
        iso_y = sem.get('iso_year')
        iso_w = sem.get('iso_week')
        try:
            inicio = datetime.fromisocalendar(int(iso_y), int(iso_w), 1)
            fim = inicio + timedelta(days=4)  # corte até sexta
            intervalo = f"{inicio:%d/%m}–{fim:%d/%m}"
        except Exception:
            intervalo = "—"
        num = sem.get('semana', iso_w)
        status = "✅" if sem.get('fechada') else "⏳"
        return f"Semana {num} ({intervalo}) · ISO {iso_w}/{iso_y} {status}"

    semanas_options = [(sem.get('iso_year'), sem.get('iso_week')) for sem in semanas_meta if sem.get('iso_week') and sem.get('iso_year')]
    idx_default = 0
    if semanas_options:
        # Padrão: última semana fechada; se nenhuma fechada, última da lista
        fechadas_idx = [i for i, key in enumerate(semanas_options) if semanas_map.get(key, {}).get('fechada')]
        idx_default = (fechadas_idx[-1] if fechadas_idx else len(semanas_options) - 1)
    semana_state_key = "semana_fechamento_tatica"
    if semana_state_key not in st.session_state and semanas_options:
        st.session_state[semana_state_key] = semanas_options[idx_default]

    if semanas_options:
        sel_default = st.session_state.get(semana_state_key, semanas_options[idx_default])
        try:
            idx_current = semanas_options.index(sel_default)
        except ValueError:
            idx_current = idx_default

        sel_key = st.selectbox(
            "Fechamento semanal (relatório fechado, corte sexta-feira)",
            options=semanas_options,
            format_func=lambda key: _label_semana(semanas_map.get(key, {})),
            index=idx_current,
            key="select_semana_fechamento",
            help="Escolha apenas semanas fechadas; semanas em andamento são ocultadas."
        )
        sel_info = semanas_map.get(sel_key, {})
        if not sel_info.get('fechada'):
            st.warning("Esta semana ainda não está fechada. Aguarde o fechamento (sexta 18h) para o relatório oficial.")
        if st.session_state.get(semana_state_key) != sel_key:
            st.session_state[semana_state_key] = sel_key
            st.rerun()
        semana_escolhida = sel_info if sel_info.get('fechada') else None
        idx_sel = semanas_options.index(sel_key) if sel_key in semanas_options else 0
        semana_anterior_meta = semanas_map.get(semanas_options[idx_sel - 1]) if idx_sel > 0 else None
    else:
        semana_escolhida = None
        semana_anterior_meta = None
    allowed_keys = set()
    if semana_escolhida:
        allowed_keys.add(sel_key)
        if 'idx_sel' in locals() and idx_sel > 0:
            allowed_keys.add(semanas_options[idx_sel - 1])

    def _aplicar_semana(df_src: pd.DataFrame, semana_meta: Optional[dict], semana_prev_meta: Optional[dict]) -> Tuple[pd.DataFrame, float, float]:
        """Cria visão do fechamento da semana selecionada, com comparação vs semana imediatamente anterior."""
        if not semana_meta:
            return df_src.copy(), float(df_src.get('WoW_Semana_Atual', 0).sum()), float(df_src.get('WoW_Semana_Anterior', 0).sum())

        iso_y_sel = int(semana_meta.get('iso_year', 0) or 0)
        iso_w_sel = int(semana_meta.get('iso_week', 0) or 0)
        if semana_prev_meta:
            iso_y_prev = int(semana_prev_meta.get('iso_year', 0) or 0)
            iso_w_prev = int(semana_prev_meta.get('iso_week', 0) or 0)
            prev_key = (iso_y_prev, iso_w_prev)
        else:
            try:
                prev_monday = datetime.fromisocalendar(iso_y_sel, iso_w_sel, 1) - timedelta(days=7)
                prev_iso = prev_monday.isocalendar()
                prev_key = (int(prev_iso.year), int(prev_iso.week))
            except Exception:
                prev_key = None

        def _map_weeks(val):
            try:
                semanas_json = json.loads(val) if isinstance(val, str) else (val if isinstance(val, list) else [])
            except Exception:
                return {}
            mapa = {}
            for s in semanas_json:
                ky = (s.get('iso_year'), s.get('iso_week'))
                if ky[0] and ky[1]:
                    mapa[ky] = s
            return mapa

        df_out = df_src.copy()
        vols_atual = []
        vols_ant = []
        for _, row in df_out.iterrows():
            semanas_map = _map_weeks(row.get('Semanas_Mes_Atual', '[]'))
            atual_info = semanas_map.get((iso_y_sel, iso_w_sel), {})
            vol_atual = atual_info.get('volume_util', 0) or 0
            vol_ant = atual_info.get('volume_semana_anterior')
            if (vol_ant is None or vol_ant == 0) and prev_key and prev_key in semanas_map:
                vol_ant = semanas_map[prev_key].get('volume_util', 0) or vol_ant
            vols_atual.append(float(vol_atual))
            vols_ant.append(float(vol_ant) if vol_ant is not None else 0.0)

        df_out['WoW_Semana_Atual'] = vols_atual
        df_out['WoW_Semana_Anterior'] = vols_ant
        return df_out, sum(vols_atual), sum(vols_ant)

    df_semana_view, total_semana_atual, total_semana_anterior = _aplicar_semana(df, semana_escolhida, semana_anterior_meta)
    df = df_semana_view
    if 'WoW_Semana_Atual' in df.columns and 'WoW_Semana_Anterior' in df.columns:
        df['Variacao_Semanal_Pct'] = np.where(
            df['WoW_Semana_Anterior'] > 0,
            ((df['WoW_Semana_Atual'] - df['WoW_Semana_Anterior']) / df['WoW_Semana_Anterior']) * 100,
            None
        )
    # Atualizar metadados para uso em todas as seções com a semana selecionada
    if semana_escolhida and allowed_keys:
        met['semanas_detalhes'] = [sem for sem in semanas_meta if (sem.get('iso_year'), sem.get('iso_week')) in allowed_keys]
        if not met['semanas_detalhes']:
            met['semanas_detalhes'] = semanas_meta

    # 1. Cálculos Específicos da Aba
    if 'CNPJ_Normalizado' not in df.columns and 'CNPJ_PCL' in df.columns:
        df['CNPJ_Normalizado'] = df['CNPJ_PCL'].apply(DataManager.normalizar_cnpj)

    # Reforçar filtro de porte (garantia mesmo se chamado sem df_view filtrado)
    portes_sel = filtros.get('portes')
    if portes_sel and 'Porte' in df.columns:
        df = df[df['Porte'].isin(portes_sel)].copy()

    cols_num = ['WoW_Semana_Atual', 'WoW_Semana_Anterior', 'Media_Semanal_2025']
    for c in cols_num:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
            
    # Queda Absoluta (Anterior - Atual). Positivo = Perda de volume.
    df['Queda_Semanal_Abs'] = df['WoW_Semana_Anterior'] - df['WoW_Semana_Atual']
    
    # % Diferença da Média Histórica (vs Média 2025)
    df['Pct_Dif_Media_Historica'] = np.where(
        df['Media_Semanal_2025'] > 0,
        ((df['WoW_Semana_Atual'] - df['Media_Semanal_2025']) / df['Media_Semanal_2025']) * 100,
        np.nan
    )
    
    # Calcular Media_Semanal_2024 = Total_Coletas_2024 / 52
    if 'Total_Coletas_2024' in df.columns:
        df['Media_Semanal_2024'] = pd.to_numeric(df['Total_Coletas_2024'], errors='coerce') / 52
    else:
        df['Media_Semanal_2024'] = 0.0
    
    # Calcular Variacao_Media_24_Pct = (WoW_Semana_Atual - Media_Semanal_2024) / Media_Semanal_2024 * 100
    with np.errstate(divide='ignore', invalid='ignore'):
        mask_valido_2024 = (df['Media_Semanal_2024'] > 0) & df['Media_Semanal_2024'].notna() & df['WoW_Semana_Atual'].notna()
        df['Variacao_Media_24_Pct'] = np.where(
            mask_valido_2024,
            ((df['WoW_Semana_Atual'] - df['Media_Semanal_2024']) / df['Media_Semanal_2024']) * 100,
            np.nan
        )

    # Variáveis de Controle (Média do Estado)
    # Semana Atual
    if 'Controle_Semanal_Estado_Atual' not in df.columns:
        media_estado_atual = df.groupby('Estado')['WoW_Semana_Atual'].transform('mean')
        df['Controle_Semanal_Estado_Atual'] = media_estado_atual
    
    # Semana Anterior
    if 'Controle_Semanal_Estado_Anterior' not in df.columns:
        media_estado_anterior = df.groupby('Estado')['WoW_Semana_Anterior'].transform('mean')
        df['Controle_Semanal_Estado_Anterior'] = media_estado_anterior
    
    # Calcular Variacao_vs_Estado_Pct = (WoW_Semana_Atual - Controle_Semanal_Estado_Atual) / Controle_Semanal_Estado_Atual * 100
    with np.errstate(divide='ignore', invalid='ignore'):
        mask_valido_estado = (df['Controle_Semanal_Estado_Atual'] > 0) & df['Controle_Semanal_Estado_Atual'].notna() & df['WoW_Semana_Atual'].notna()
        df['Variacao_vs_Estado_Pct'] = np.where(
            mask_valido_estado,
            ((df['WoW_Semana_Atual'] - df['Controle_Semanal_Estado_Atual']) / df['Controle_Semanal_Estado_Atual']) * 100,
            np.nan
        )

    # Configuração das Colunas
    col_config = {
        "Nome_Fantasia_PCL": st.column_config.TextColumn("Laboratório", width="medium", help="Nome fantasia utilizado pelo time comercial"),
        "CNPJ_Normalizado": st.column_config.TextColumn("CNPJ", width="medium", help="CNPJ normalizado (apenas números)"),
        "VIP": st.column_config.TextColumn("VIP", width="small", help="Indica se o laboratório faz parte da lista VIP"),
        "Rede": st.column_config.TextColumn("Rede", width="medium", help="Rede à qual o laboratório pertence"),
        "Estado": st.column_config.TextColumn("UF", width="small", help="Unidade Federativa do laboratório"),
        "Porte": st.column_config.TextColumn("Porte", width="small", help="Segmentação por volume médio (Grande, Médio/Grande, etc.)"),
        "Media_Semanal_2025": st.column_config.NumberColumn("Média Semanal 25", format="%.0f", help="Média semanal de coletas no ano 2025"),
        "Pct_Dif_Media_Historica": st.column_config.NumberColumn("Var. % vs Média 25", format="%.1f%%", help="Variação percentual do volume atual vs média semanal de 2025. Exibe '—' quando média está zerada ou ausente.", default=None),
        "Media_Semanal_2024": st.column_config.NumberColumn("Média Semanal 24", format="%.0f", help="Média semanal de coletas no ano 2024 (Total_Coletas_2024 / 52)"),
        "Variacao_Media_24_Pct": st.column_config.NumberColumn("Var. % vs Média 24", format="%.1f%%", help="Variação percentual do volume atual vs média semanal de 2024. Exibe '—' quando média está zerada ou ausente.", default=None),
        "Media_Mensal_Top3_2024": st.column_config.NumberColumn("Média Mensal Top 3 (2024)", format="%.0f", help="Média dos 3 melhores meses de coleta em 2024 (usa menos se não tiver 3 meses)"),
        "Variacao_vs_Top3_2024_Pct": st.column_config.NumberColumn("Var. % vs Top 3 Meses 2024", format="%.1f%%", help="Variação percentual do volume da semana atual vs média dos top 3 meses de 2024. Exibe '—' quando média está zerada ou ausente.", default=None),
        "Media_Mensal_Top3_2025": st.column_config.NumberColumn("Média Mensal Top 3 (2025)", format="%.0f", help="Média dos 3 melhores meses de coleta em 2025 (usa menos se não tiver 3 meses)"),
        "Variacao_vs_Top3_2025_Pct": st.column_config.NumberColumn("Var. % vs Top 3 Meses 2025", format="%.1f%%", help="Variação percentual do volume da semana atual vs média dos top 3 meses de 2025. Exibe '—' quando média está zerada ou ausente.", default=None),
        "WoW_Semana_Anterior": st.column_config.NumberColumn("Vol. Ant.", format="%d", help="Volume realizado na semana anterior"),
        "WoW_Semana_Atual": st.column_config.NumberColumn("Vol. Atual", format="%d", help="Volume realizado na semana atual"),
        "Queda_Semanal_Abs": st.column_config.NumberColumn("Queda de volume", format="%d", help="Diferença absoluta de volume entre semana anterior e atual"),
        "Controle_Semanal_Estado_Atual": st.column_config.NumberColumn("Média do Estado", format="%.1f", help="Média de todos os labs do mesmo estado nesta semana"),
        "Variacao_vs_Estado_Pct": st.column_config.NumberColumn("Var. % semana atual vs média UF semana atual", format="%.1f%%", help="Variação percentual do volume da semana atual vs média do estado na semana atual. Exibe '—' quando média está zerada ou ausente.", default=None),
        "Variacao_Semanal_Pct": st.column_config.NumberColumn(
            "Variação WoW (%)", 
            format="%.1f%%", 
            help="🔴 COLUNA MAIS IMPORTANTE - Percentual de variação da semana atual vs anterior (WoW). Cores: ≥-20% verde, -20% a -40% amarelo, -40% a -60% laranja, <-60% vermelho. Exibe '—' quando volume anterior está zerado ou ausente.", 
            default=None
        ),
        "Dias_Sem_Coleta": st.column_config.NumberColumn("Dias Off", format="%d ⚠️", help="Dias úteis consecutivos sem coleta registrados"),
        "Em_Risco": st.column_config.TextColumn("Em Risco?", width="small", help="Aplica regra: queda ≥50% ou dias sem coleta conforme porte"),
        "Controle_Semanal_Estado_Anterior": st.column_config.NumberColumn("Média UF (Ant)", format="%.1f", help="Média de todos os labs do mesmo porte no estado na semana anterior"),
        "Variacao_Media_Estado_Pct": st.column_config.NumberColumn("Var. % média UF semana anterior vs atual", format="%.1f%%", help="Variação percentual da média estadual (semana atual vs anterior). Exibe '—' quando média anterior está zerada ou ausente.", default=None),
        "Data_Ultima_Coleta": st.column_config.DateColumn("Última Coleta", format="DD/MM/YYYY", help="Última coleta registrada (qualquer ano)")
    }
    
    cols_view = ['Nome_Fantasia_PCL', 'CNPJ_Normalizado', 'VIP', 'Rede', 'Estado', 'Porte', 
                 'Media_Semanal_2025', 'WoW_Semana_Anterior', 'WoW_Semana_Atual', 
                 'Queda_Semanal_Abs', 'Dias_Sem_Coleta', 
                 'Controle_Semanal_Estado_Anterior', 'Controle_Semanal_Estado_Atual',
                 'Data_Ultima_Coleta']

    # Preparar dados de risco ANTES de calcular métricas
    df_risco_base = preparar_dataframe_risco(df)
    
    # --- CÁLCULO DE QUEDAS ADICIONAIS (deve vir ANTES dos filtros) ---
    # Calcular magnitude da queda (positivo) para os novos filtros
    # Queda vs Média 25 (Pct_Dif_Media_Historica)
    if 'Pct_Dif_Media_Historica' in df_risco_base.columns:
        serie_queda_25 = pd.to_numeric(df_risco_base['Pct_Dif_Media_Historica'], errors='coerce')
        df_risco_base['Queda_Media_25_Pct'] = np.where(serie_queda_25 < 0, -serie_queda_25, np.nan)
    else:
        df_risco_base['Queda_Media_25_Pct'] = np.nan

    # Queda vs Média 24 (Variacao_Media_24_Pct)
    if 'Variacao_Media_24_Pct' in df_risco_base.columns:
        serie_queda_24 = pd.to_numeric(df_risco_base['Variacao_Media_24_Pct'], errors='coerce')
        df_risco_base['Queda_Media_24_Pct'] = np.where(serie_queda_24 < 0, -serie_queda_24, np.nan)
    else:
        df_risco_base['Queda_Media_24_Pct'] = np.nan

    # Queda vs Estado (Variacao_vs_Estado_Pct)
    if 'Variacao_vs_Estado_Pct' in df_risco_base.columns:
        serie_queda_estado = pd.to_numeric(df_risco_base['Variacao_vs_Estado_Pct'], errors='coerce')
        df_risco_base['Queda_Estado_Pct'] = np.where(serie_queda_estado < 0, -serie_queda_estado, np.nan)
    else:
        df_risco_base['Queda_Estado_Pct'] = np.nan

    # --- LISTA DE RISCO ---
    st.markdown("### 📊 Resumo Semanal")
    st.caption("Todos os laboratórios ordenados pela maior queda semanal. Reﬁne usando os filtros.")
    
    # Obter filtro de variação do estado da sessão
    variacoes_opcoes = list(VARIACAO_QUEDA_FAIXAS.keys())
    # Inicializar com padrão "Acima de 50%" se não houver no session_state
    # Mas se o usuário explicitamente removeu tudo (lista vazia), respeitar isso
    if 'filtro_variacao_risco' not in st.session_state:
        # Primeira vez: usar padrão
        variacoes_sel = ["Acima de 50%"]
    else:
        # Já existe no session_state: usar o valor (pode ser lista vazia se usuário removeu tudo)
        variacoes_sel = st.session_state.get('filtro_variacao_risco', ["Acima de 50%"])
        if not isinstance(variacoes_sel, list):
            variacoes_sel = ["Acima de 50%"]
    
    col_f1, col_f2, col_f3, col_f4 = st.columns(4)
    
    with col_f1:
        variacoes_sel = st.multiselect(
            "Filtro de Queda Semanal (WoW)",
            options=variacoes_opcoes,
            default=variacoes_sel if isinstance(variacoes_sel, list) else ["Acima de 50%"],
            help="Seleciona a faixa de queda percentual (semana atual vs semana anterior). Padrão: 'Acima de 50%'.",
            key="filtro_variacao_risco"
        )
    
    with col_f2:
        queda_media_25_sel = st.multiselect(
            "Filtro de Queda vs Média 25",
            options=variacoes_opcoes,
            default=[],
            help="Seleciona a faixa de queda percentual (semana atual vs média 2025).",
            key="filtro_queda_media_25"
        )

    with col_f3:
        queda_media_24_sel = st.multiselect(
            "Filtro de Queda vs Média 24",
            options=variacoes_opcoes,
            default=[],
            help="Seleciona a faixa de queda percentual (semana atual vs média 2024).",
            key="filtro_queda_media_24"
        )

    with col_f4:
        queda_estado_sel = st.multiselect(
            "Filtro de Queda vs Estado",
            options=variacoes_opcoes,
            default=[],
            help="Seleciona a faixa de queda percentual (semana atual vs média do estado).",
            key="filtro_queda_estado"
        )
    filtros['faixas_variacao_risco'] = variacoes_sel
    
    filtros_ativos = []
    if variacoes_sel: filtros_ativos.append("WoW")
    if queda_media_25_sel: filtros_ativos.append("Média 25")
    if queda_media_24_sel: filtros_ativos.append("Média 24")
    if queda_estado_sel: filtros_ativos.append("Estado")

    if filtros_ativos:
        st.caption(f"Filtros de queda ativos: {', '.join(filtros_ativos)}.")
    else:
        st.caption("Nenhum filtro de variação aplicado. Mostrando todos os laboratórios do porte selecionado, ordenados por dias sem coleta (maior primeiro).")
    
    # Calcular perdas PRIMEIRO (recentes + antigas) para excluir da classificação de risco
    # Isso garante que NENHUM laboratório de perda apareça na Lista de Risco
    cnpjs_perdas = set()
    if 'Classificacao_Perda_V2' in df.columns and 'CNPJ_Normalizado' in df.columns:
        # Identificar CNPJs de perdas recentes e antigas
        perdas_recentes = df[df['Classificacao_Perda_V2'] == 'Perda Recente']
        perdas_antigas = df[df['Classificacao_Perda_V2'].isin(['Perda Antiga', 'Perda Consolidada'])]
        cnpjs_perdas = set(perdas_recentes['CNPJ_Normalizado'].dropna()) | set(perdas_antigas['CNPJ_Normalizado'].dropna())
    
    # Aplicar filtro de variação WoW
    df_risco_filtrado = aplicar_filtro_variacao(df_risco_base, variacoes_sel)
    
    # Aplicar filtros adicionais
    df_risco_filtrado = aplicar_filtro_variacao_generica(df_risco_filtrado, 'Queda_Media_25_Pct', queda_media_25_sel, usar_valor_absoluto=False)
    df_risco_filtrado = aplicar_filtro_variacao_generica(df_risco_filtrado, 'Queda_Media_24_Pct', queda_media_24_sel, usar_valor_absoluto=False)
    df_risco_filtrado = aplicar_filtro_variacao_generica(df_risco_filtrado, 'Queda_Estado_Pct', queda_estado_sel, usar_valor_absoluto=False)
    
    # EXCLUIR perdas da lista de risco - CRÍTICO para evitar overlaps
    if cnpjs_perdas and 'CNPJ_Normalizado' in df_risco_filtrado.columns:
        df_risco_filtrado = df_risco_filtrado[~df_risco_filtrado['CNPJ_Normalizado'].isin(cnpjs_perdas)].copy()
    
    df_risco_ordenado = df_risco_filtrado.sort_values('Dias_Sem_Coleta', ascending=False, na_position='last')
    
    # Calcular totais APÓS filtros aplicados
    total_labs_filtrado = len(df_risco_filtrado) if not df_risco_filtrado.empty else 0
    
    # Recalcular volumes totais baseado nos labs filtrados
    total_semana_atual_filtrado = float(df_risco_filtrado['WoW_Semana_Atual'].sum()) if not df_risco_filtrado.empty else 0.0
    total_semana_anterior_filtrado = float(df_risco_filtrado['WoW_Semana_Anterior'].sum()) if not df_risco_filtrado.empty else 0.0
    
    # ============================================================
    # KPI BOX EXECUTIVO - TOPO DA ABA SEMANAL (calculado APÓS filtros)
    # ============================================================
    cards = {
        'volume_semana_atual': total_semana_atual_filtrado,
        'volume_semana_anterior': total_semana_anterior_filtrado,
        'media_semana_atual': float(df_risco_filtrado['WoW_Semana_Atual'].mean()) if len(df_risco_filtrado) else 0.0,
        'media_semana_anterior': float(df_risco_filtrado['WoW_Semana_Anterior'].mean()) if len(df_risco_filtrado) else 0.0,
        'variacao_media_pct': calcular_variacao_percentual(
            float(df_risco_filtrado['WoW_Semana_Atual'].mean()) if len(df_risco_filtrado) else 0.0,
            float(df_risco_filtrado['WoW_Semana_Anterior'].mean()) if len(df_risco_filtrado) else 0.0
        ),
        'total_labs': total_labs_filtrado
    }

    col1, col2, col3, col4 = st.columns(4)
    col1.metric(
        "Volume semanal (selecionada)",
        f"{cards.get('volume_semana_atual', 0):,.0f}"
    )
    col2.metric(
        "Volume semana anterior",
        f"{cards.get('volume_semana_anterior', 0):,.0f}"
    )
    variacao_media = cards.get('variacao_media_pct')
    if variacao_media is None:
        col3.metric("Variação da média semanal", "—")
    else:
        col3.metric(
            "Variação da média semanal",
            f"{variacao_media:+.1f}%"
        )
    col4.metric(
        "Labs na listagem",
        f"{cards.get('total_labs', 0):,}"
    )

    if semana_escolhida:
        label_atual = _label_semana(semana_escolhida)
        label_prev = _label_semana(semana_anterior_meta) if semana_anterior_meta else "Sem semana anterior disponível"
        st.caption(f"Fechamento selecionado: {label_atual} · Comparação automática: {label_prev}")
    else:
        st.caption("Sem semanas fechadas disponíveis para selecionar.")
    
    st.markdown("---")
    
    # Continuar com exibição da lista de risco
    st.subheader("🚨 Lista de Risco (queda WoW ou dias off)")
    
    df_total_processado = preparar_dataframe_risco(df_total) if df_total is not None else None
    
    # Ordem das colunas conforme especificado
    cols_risco_view = [
        "Nome_Fantasia_PCL",                 # 1. lab
        "Rede",                              # 2. rede
        "Estado",                            # 3. uf
        "Porte",                             # 4. porte
        "VIP",                               # 5. vip
        "Data_Ultima_Coleta",                # 6. última coleta
        "Dias_Sem_Coleta",                   # 7. dias off
        "WoW_Semana_Anterior",               # 8. volume anterior
        "WoW_Semana_Atual",                  # 9. volume atual
        "Variacao_Semanal_Pct",              # 10. variação semana anterior e atual
        "Media_Semanal_2025",                # 11. média semanal 25
        "Pct_Dif_Media_Historica",           # 12. variação média semanal 25 com semana atual
        "Media_Mensal_Top3_2025",            # 13. média mensal top 3 meses 2025
        "Variacao_vs_Top3_2025_Pct",         # 14. variação vs top 3 meses 2025
        "Media_Semanal_2024",                # 15. média semanal 24
        "Variacao_Media_24_Pct",             # 16. variação média semanal 24 com semana atual
        "Media_Mensal_Top3_2024",            # 17. média mensal top 3 meses 2024
        "Variacao_vs_Top3_2024_Pct",         # 18. variação vs top 3 meses 2024
        "Variacao_vs_Estado_Pct",            # 19. variação semana atual vs média do estado semana atual
        "Variacao_Media_Estado_Pct",         # 20. variação média estado
        "Em_Risco",                          # útil quando mostrar todos os labs
        "CNPJ_Normalizado",                  # técnico – pode ficar por último
    ]
    
    if df_risco_ordenado.empty:
        st.info("Nenhum laboratório encontrado para os filtros selecionados.")
    else:
        df_risco_display = df_risco_ordenado.copy()
        df_risco_display['Em_Risco'] = df_risco_display['Em_Risco'].apply(lambda x: "Sim" if bool(x) else "—")
        df_risco_display = adicionar_coluna_detalhes(df_risco_display, 'CNPJ_Normalizado')
        
        # Filtrar apenas as colunas de exibição que existem no DataFrame, mantendo a ordem de cols_risco_view
        cols_display_risco = [c for c in cols_risco_view if c in df_risco_display.columns]
        df_risco_display = df_risco_display[cols_display_risco].reset_index(drop=True)
        
        evento_risco = st.dataframe(
            df_risco_display,
            use_container_width=True,
            hide_index=True,
            column_config=col_config,
            selection_mode="single-row",
            on_select="rerun",
            key="tbl_risco_sem"
        )
        _processar_evento_selecao(evento_risco, df_risco_display)
        st.caption("Selecione a caixa ao lado do laboratório para abrir automaticamente a Análise Detalhada.")
        
    # Botão de exportação Excel (apenas colunas de exibição)
    if not df_risco_ordenado.empty:
        # Usar o mesmo dataframe que é exibido na tela (com as mesmas transformações)
        # Aplicar as mesmas transformações do df_risco_display
        df_export_risco = df_risco_ordenado.copy()
        # Converter Em_Risco para formato de exibição (igual ao que é mostrado na tela)
        df_export_risco['Em_Risco'] = df_export_risco['Em_Risco'].apply(lambda x: "Sim" if bool(x) else "—")
        
        # Filtrar apenas as colunas de exibição que existem no DataFrame
        cols_export_risco = [c for c in cols_risco_view if c in df_export_risco.columns]
        df_export_risco = df_export_risco[cols_export_risco].copy()
        
        excel_buffer = BytesIO()
        df_export_risco.to_excel(excel_buffer, index=False, engine='openpyxl')
        excel_data = excel_buffer.getvalue()
        st.download_button(
            label="📊 Download Excel (Lista de Risco)",
            data=excel_data,
            file_name=f"lista_risco_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Exporta a tabela com filtros aplicados no formato Excel, contendo exatamente as mesmas colunas e dados exibidos na tela."
        )

    st.markdown("---")

    # --- FUNÇÕES AUXILIARES PARA CÁLCULO DE MÉTRICAS DE PERDAS ---
    def calcular_metricas_perda(row):
        """Calcula métricas específicas para cada perda."""
        metricas = {
            'media_mensal_perdas': 0.0,
            'volume_ultimo_mes_coleta': 0,
            'mes_ultimo_coleta': None,
            'volume_ultima_semana_coleta': 0,
            'semana_ultima_coleta': None,
            'variacao_semanal_pct_ultima_semana': -100.0,  # Default: queda completa se não houver dados
            'maxima_coletas': 0,
            'mes_maxima': None,
            'ano_maxima': None
        }
        
        # Calcular média mensal de perdas (média dos meses com coleta)
        meses_nomes = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
        meses_2024 = [f'N_Coletas_{m}_24' for m in meses_nomes]
        meses_2025 = [f'N_Coletas_{m}_25' for m in meses_nomes]
        todas_colunas_meses = meses_2024 + meses_2025
        
        volumes_mensais = []
        for col in todas_colunas_meses:
            if col in row.index:
                valor = pd.to_numeric(row.get(col, 0), errors='coerce')
                if pd.notna(valor) and valor > 0:
                    volumes_mensais.append(valor)
        
        if volumes_mensais:
            metricas['media_mensal_perdas'] = float(np.mean(volumes_mensais))
        
        # Encontrar último mês com coleta (não necessariamente o mês atual)
        ultimo_mes_coleta = None
        ultimo_volume_mes = 0
        for col in reversed(todas_colunas_meses):  # Começar pelos mais recentes
            if col in row.index:
                valor = pd.to_numeric(row.get(col, 0), errors='coerce')
                if pd.notna(valor) and valor > 0:
                    ultimo_mes_coleta = col
                    ultimo_volume_mes = int(valor)
                    break
        
        if ultimo_mes_coleta:
            metricas['volume_ultimo_mes_coleta'] = ultimo_volume_mes
            # Extrair mês e ano do nome da coluna
            partes = ultimo_mes_coleta.split('_')
            if len(partes) >= 3:
                mes_codigo = partes[2]
                ano = partes[3] if len(partes) > 3 else None
                meses_map = {"Jan": "Janeiro", "Fev": "Fevereiro", "Mar": "Março", "Abr": "Abril",
                            "Mai": "Maio", "Jun": "Junho", "Jul": "Julho", "Ago": "Agosto",
                            "Set": "Setembro", "Out": "Outubro", "Nov": "Novembro", "Dez": "Dezembro"}
                metricas['mes_ultimo_coleta'] = f"{meses_map.get(mes_codigo, mes_codigo)}/{ano}" if ano else meses_map.get(mes_codigo, mes_codigo)
        
        # Encontrar última semana com coleta e calcular variação semanal
        if 'Semanas_Mes_Atual' in row.index:
            try:
                semanas_json = row.get('Semanas_Mes_Atual', '[]')
                if isinstance(semanas_json, str) and semanas_json and semanas_json != '[]':
                    semanas = json.loads(semanas_json)
                elif isinstance(semanas_json, list):
                    semanas = semanas_json
                else:
                    semanas = []
                
                # Procurar última semana com volume > 0
                ultima_semana_encontrada = False
                for semana in reversed(semanas):
                    vol_util = semana.get('volume_util', 0)
                    if vol_util and vol_util > 0:
                        ultima_semana_encontrada = True
                        metricas['volume_ultima_semana_coleta'] = int(vol_util)
                        iso_week = semana.get('iso_week')
                        iso_year = semana.get('iso_year')
                        if iso_week and iso_year:
                            metricas['semana_ultima_coleta'] = f"Semana {iso_week}/{iso_year}"
                        
                        # Calcular variação semanal da última semana coletada
                        vol_anterior = semana.get('volume_semana_anterior', None)
                        
                        # Converter volumes para float
                        vol_util_float = float(vol_util) if vol_util else 0.0
                        
                        # Verificar se temos semana anterior válida
                        if vol_anterior is not None:
                            try:
                                vol_anterior = float(vol_anterior) if vol_anterior else 0.0
                                
                                # Calcular variação percentual quando temos semana anterior válida
                                if vol_anterior > 0:
                                    variacao_pct = ((vol_util_float - vol_anterior) / vol_anterior) * 100
                                    metricas['variacao_semanal_pct_ultima_semana'] = round(variacao_pct, 2)
                                else:
                                    # Se não há semana anterior ou volume = 0, queda completa (-100%)
                                    metricas['variacao_semanal_pct_ultima_semana'] = -100.0
                            except (ValueError, TypeError):
                                # Erro ao converter, assumir queda completa
                                metricas['variacao_semanal_pct_ultima_semana'] = -100.0
                        else:
                            # Se não há dados da semana anterior, assumir queda completa (-100%)
                            metricas['variacao_semanal_pct_ultima_semana'] = -100.0
                        break
                
                # Se não encontrou nenhuma semana com volume > 0, assumir queda completa
                if not ultima_semana_encontrada:
                    metricas['variacao_semanal_pct_ultima_semana'] = -100.0
            except Exception as e:
                # Em caso de erro, assumir queda completa
                metricas['variacao_semanal_pct_ultima_semana'] = -100.0
        
        # Encontrar máxima de coletas (valor + mês/ano)
        max_volume = 0
        max_mes_col = None
        for col in todas_colunas_meses:
            if col in row.index:
                valor = pd.to_numeric(row.get(col, 0), errors='coerce')
                if pd.notna(valor) and valor > max_volume:
                    max_volume = int(valor)
                    max_mes_col = col
        
        if max_mes_col and max_volume > 0:
            metricas['maxima_coletas'] = max_volume
            partes = max_mes_col.split('_')
            if len(partes) >= 3:
                mes_codigo = partes[2]
                ano = partes[3] if len(partes) > 3 else None
                meses_map = {"Jan": "Janeiro", "Fev": "Fevereiro", "Mar": "Março", "Abr": "Abril",
                            "Mai": "Maio", "Jun": "Junho", "Jul": "Julho", "Ago": "Agosto",
                            "Set": "Setembro", "Out": "Outubro", "Nov": "Novembro", "Dez": "Dezembro"}
                metricas['mes_maxima'] = meses_map.get(mes_codigo, mes_codigo)
                metricas['ano_maxima'] = ano
        
        return metricas
    
    # --- TABELAS SEPARADAS POR CATEGORIA DE PERDA (LAYOUT EM COLUNA) ---
    
    # Perdas Recentes
    st.subheader("📉 Perdas Recentes (<6 meses)")
    df_perda_recente = df[df['Classificacao_Perda_V2'] == 'Perda Recente'].copy()
    
    # Garantir que colunas de contexto existam (alinhar com Lista de Risco)
    # Como df_perda_recente é um filtro de df, as colunas já devem existir, 
    # mas garantimos que estejam presentes e preenchidas
    if not df_perda_recente.empty:
        # Garantir que colunas de contexto existam com valores padrão se necessário
        colunas_contexto_defaults = {
            'Rede': '-',
            'Estado': '',
            'Porte': 'Pequeno',
            'VIP': 'Não'
        }
        
        for col, default_val in colunas_contexto_defaults.items():
            if col not in df_perda_recente.columns:
                df_perda_recente[col] = default_val
            else:
                # Preencher valores NaN com padrão
                df_perda_recente[col] = df_perda_recente[col].fillna(default_val)
    
    if not df_perda_recente.empty:
        # Calcular métricas para cada perda
        metricas_list = []
        for idx, row in df_perda_recente.iterrows():
            metricas = calcular_metricas_perda(row)
            metricas_list.append(metricas)
        
        # Adicionar colunas de métricas ao DataFrame
        df_perda_recente['Media_Mensal_Perdas'] = [m['media_mensal_perdas'] for m in metricas_list]
        df_perda_recente['Volume_Ultimo_Mes_Coleta'] = [m['volume_ultimo_mes_coleta'] for m in metricas_list]
        df_perda_recente['Mes_Ultimo_Coleta'] = [m['mes_ultimo_coleta'] or 'N/A' for m in metricas_list]
        df_perda_recente['Volume_Ultima_Semana_Coleta'] = [m['volume_ultima_semana_coleta'] for m in metricas_list]
        df_perda_recente['Semana_Ultima_Coleta'] = [m['semana_ultima_coleta'] or 'N/A' for m in metricas_list]
        df_perda_recente['Maxima_Coletas'] = [m['maxima_coletas'] for m in metricas_list]
        df_perda_recente['Mes_Maxima'] = [m['mes_maxima'] or 'N/A' for m in metricas_list]
        df_perda_recente['Ano_Maxima'] = [m['ano_maxima'] or 'N/A' for m in metricas_list]
        
        # Preencher Variacao_Semanal_Pct usando a métrica calculada da última semana coletada
        # Isso calcula a variação da última semana coletada em relação à semana anterior
        # Garantir que None seja convertido para -100.0 (queda completa)
        df_perda_recente['Variacao_Semanal_Pct'] = [
            m.get('variacao_semanal_pct_ultima_semana', -100.0) if m.get('variacao_semanal_pct_ultima_semana') is not None else -100.0
            for m in metricas_list
        ]
        
        # Ordenação: maior para menor (por dias sem coleta ou volume de perda), desempate por data última coleta
        df_perda_recente['_ordem_dias'] = df_perda_recente.get('Dias_Sem_Coleta', 0).fillna(0)
        df_perda_recente['_ordem_queda'] = df_perda_recente.get('Queda_Semanal_Abs', 0).fillna(0)
        df_perda_recente['_ordem_data'] = pd.to_datetime(df_perda_recente.get('Data_Ultima_Coleta', pd.NaT), errors='coerce')
        df_perda_recente['_ordem_data'] = df_perda_recente['_ordem_data'].fillna(pd.Timestamp.min)
        
        # Ordenar: primeiro por dias sem coleta (desc), depois por queda (desc), depois por data última coleta (desc - mais recente primeiro)
        df_perda_recente = df_perda_recente.sort_values(
            ['_ordem_dias', '_ordem_queda', '_ordem_data'],
            ascending=[False, False, False]
        )
        df_perda_recente = df_perda_recente.drop(columns=['_ordem_dias', '_ordem_queda', '_ordem_data'])
        
        # Ordem ideal das colunas para Perdas (2025 - versão definitiva)
        # Alinhada com a Lista de Risco - colunas de contexto essenciais adicionadas
        cols_perda_extended = [
            'Nome_Fantasia_PCL',
            'Rede',                      # ela olha muito
            'Estado',                    # essencial
            'Porte',                     # essencial para contexto
            'VIP',                       # ela filtra por VIP direto
            'Data_Ultima_Coleta',        # a coluna mais importante em perdas – tem que estar logo no começo
            'Dias_Sem_Coleta',           # logo depois da data
            'Volume_Ultimo_Mes_Coleta',  # volume antes de morrer
            'Mes_Ultimo_Coleta',         # mês/ano da última coleta
            'Maxima_Coletas',            # pico histórico
            'Mes_Maxima',                # quando foi o pico
            'Ano_Maxima',
            'Media_Mensal_Perdas',       # média enquanto estava vivo
            'Queda_Semanal_Abs',         # queda bruta
            'CNPJ_Normalizado',          # técnico – final
        ]
        # Filtrar apenas colunas que existem no DataFrame
        cols_perda_extended = [c for c in cols_perda_extended if c in df_perda_recente.columns]
        
        df_perda_recente_display = adicionar_coluna_detalhes(df_perda_recente, 'CNPJ_Normalizado')
        df_perda_recente_display = df_perda_recente_display[cols_perda_extended].reset_index(drop=True)
        
        # Configuração de colunas estendida (inclui colunas específicas de perdas + colunas de contexto)
        col_config_extended = col_config.copy()
        col_config_extended.update({
            # Colunas de contexto (alinhadas com Lista de Risco)
            "Rede": st.column_config.TextColumn("Rede", width="medium", help="Rede à qual o laboratório pertence"),
            "Estado": st.column_config.TextColumn("UF", width="small", help="Unidade Federativa do laboratório"),
            "Porte": st.column_config.TextColumn("Porte", width="small", help="Segmentação por volume médio (Grande, Médio/Grande, etc.)"),
            "VIP": st.column_config.TextColumn("VIP", width="small", help="Indica se o laboratório faz parte da lista VIP"),
            # Colunas específicas de perdas
            "Media_Mensal_Perdas": st.column_config.NumberColumn("Média Mensal", format="%.1f", help="Média mensal de perdas enquanto estava ativo"),
            "Volume_Ultimo_Mes_Coleta": st.column_config.NumberColumn("Vol. Último Mês", format="%d", help="Volume do último mês com coleta"),
            "Mes_Ultimo_Coleta": st.column_config.TextColumn("Mês Última Coleta", help="Mês/ano do último mês com coleta"),
            "Volume_Ultima_Semana_Coleta": st.column_config.NumberColumn("Vol. Última Semana", format="%d", help="Volume da última semana com coleta"),
            "Semana_Ultima_Coleta": st.column_config.TextColumn("Semana Última Coleta", help="Semana ISO da última semana com coleta"),
            "Maxima_Coletas": st.column_config.NumberColumn("Máxima Coletas", format="%d", help="Máxima de coletas do cliente (pico histórico)"),
            "Mes_Maxima": st.column_config.TextColumn("Mês Máxima", help="Mês da máxima de coletas"),
            "Ano_Maxima": st.column_config.TextColumn("Ano Máxima", help="Ano da máxima de coletas"),
        })
        
        evento_recente = st.dataframe(
            df_perda_recente_display,
            use_container_width=True,
            hide_index=True,
            column_config=col_config_extended,
            selection_mode="single-row",
            on_select="rerun",
            key="tbl_perda_recente"
        )
        _processar_evento_selecao(evento_recente, df_perda_recente_display)
        st.caption("Use a caixa de seleção para abrir automaticamente a Análise Detalhada.")
        
        # Botão de exportação Excel (apenas colunas de exibição)
        cols_export_perda_recente = [c for c in cols_perda_extended if c in df_perda_recente.columns]
        df_export_perda_recente = df_perda_recente[cols_export_perda_recente].copy()
        excel_buffer = BytesIO()
        df_export_perda_recente.to_excel(excel_buffer, index=False, engine='openpyxl')
        excel_data = excel_buffer.getvalue()
        st.download_button(
            label="📊 Download Excel (Perdas Recentes)",
            data=excel_data,
            file_name=f"perdas_recentes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Exporta a tabela de perdas recentes no formato Excel, contendo apenas as colunas relevantes."
        )
    else:
        st.info("Nenhuma perda recente.")
    
    st.markdown("---")
    
    # Perdas Antigas (logo abaixo de Perdas Recentes, mesmo layout)
    st.subheader("🗄️ Perdas Antigas (>6 meses)")
    df_antigas = df[df['Classificacao_Perda_V2'].isin(['Perda Antiga', 'Perda Consolidada'])].copy()
    
    # Garantir que colunas de contexto existam (alinhar com Lista de Risco)
    # Como df_antigas é um filtro de df, as colunas já devem existir, 
    # mas garantimos que estejam presentes e preenchidas
    if not df_antigas.empty:
        # Garantir que colunas de contexto existam com valores padrão se necessário
        colunas_contexto_defaults = {
            'Rede': '-',
            'Estado': '',
            'Porte': 'Pequeno',
            'VIP': 'Não'
        }
        
        for col, default_val in colunas_contexto_defaults.items():
            if col not in df_antigas.columns:
                df_antigas[col] = default_val
            else:
                # Preencher valores NaN com padrão
                df_antigas[col] = df_antigas[col].fillna(default_val)
    
    if not df_antigas.empty:
        # Calcular métricas para cada perda (mesma estrutura)
        metricas_list_antigas = []
        for idx, row in df_antigas.iterrows():
            metricas = calcular_metricas_perda(row)
            metricas_list_antigas.append(metricas)
        
        # Adicionar colunas de métricas ao DataFrame
        df_antigas['Media_Mensal_Perdas'] = [m['media_mensal_perdas'] for m in metricas_list_antigas]
        df_antigas['Volume_Ultimo_Mes_Coleta'] = [m['volume_ultimo_mes_coleta'] for m in metricas_list_antigas]
        df_antigas['Mes_Ultimo_Coleta'] = [m['mes_ultimo_coleta'] or 'N/A' for m in metricas_list_antigas]
        df_antigas['Volume_Ultima_Semana_Coleta'] = [m['volume_ultima_semana_coleta'] for m in metricas_list_antigas]
        df_antigas['Semana_Ultima_Coleta'] = [m['semana_ultima_coleta'] or 'N/A' for m in metricas_list_antigas]
        df_antigas['Maxima_Coletas'] = [m['maxima_coletas'] for m in metricas_list_antigas]
        df_antigas['Mes_Maxima'] = [m['mes_maxima'] or 'N/A' for m in metricas_list_antigas]
        df_antigas['Ano_Maxima'] = [m['ano_maxima'] or 'N/A' for m in metricas_list_antigas]
        
        # Preencher Variacao_Semanal_Pct usando a métrica calculada da última semana coletada
        # Isso calcula a variação da última semana coletada em relação à semana anterior
        # Garantir que None seja convertido para -100.0 (queda completa)
        df_antigas['Variacao_Semanal_Pct'] = [
            m.get('variacao_semanal_pct_ultima_semana', -100.0) if m.get('variacao_semanal_pct_ultima_semana') is not None else -100.0
            for m in metricas_list_antigas
        ]
        
        # Ordenação: maior para menor (por dias sem coleta ou volume de perda), desempate por data última coleta
        df_antigas['_ordem_dias'] = df_antigas.get('Dias_Sem_Coleta', 0).fillna(0)
        df_antigas['_ordem_queda'] = df_antigas.get('Queda_Semanal_Abs', 0).fillna(0)
        df_antigas['_ordem_data'] = pd.to_datetime(df_antigas.get('Data_Ultima_Coleta', pd.NaT), errors='coerce')
        df_antigas['_ordem_data'] = df_antigas['_ordem_data'].fillna(pd.Timestamp.min)
        
        # Ordenar: primeiro por dias sem coleta (desc), depois por queda (desc), depois por data última coleta (desc)
        df_antigas = df_antigas.sort_values(
            ['_ordem_dias', '_ordem_queda', '_ordem_data'],
            ascending=[False, False, False]
        )
        df_antigas = df_antigas.drop(columns=['_ordem_dias', '_ordem_queda', '_ordem_data'])
        
        cols_perda_antigas = [
            'Nome_Fantasia_PCL',
            'Rede',
            'Estado',
            'Porte',
            'VIP',
            'Data_Ultima_Coleta',
            'Dias_Sem_Coleta',
            'Volume_Ultimo_Mes_Coleta',
            'Mes_Ultimo_Coleta',
            'Maxima_Coletas',
            'Mes_Maxima',
            'Ano_Maxima',
            'Media_Mensal_Perdas',
            'Queda_Semanal_Abs',
            'CNPJ_Normalizado',
        ]
        # Filtrar apenas colunas que existem no DataFrame
        cols_perda_antigas = [c for c in cols_perda_antigas if c in df_antigas.columns]
        
        df_antigas_display = adicionar_coluna_detalhes(df_antigas, 'CNPJ_Normalizado')
        df_antigas_display = df_antigas_display[cols_perda_antigas].reset_index(drop=True)
        
        evento_antiga = st.dataframe(
            df_antigas_display,
            use_container_width=True,
            hide_index=True,
            column_config=col_config_extended,
            selection_mode="single-row",
            on_select="rerun",
            key="tbl_perda_antiga"
        )
        _processar_evento_selecao(evento_antiga, df_antigas_display)
        st.caption("Use a caixa de seleção para abrir automaticamente a Análise Detalhada.")
        
        # Botão de exportação Excel (apenas colunas de exibição)
        cols_export_perda_antiga = [c for c in cols_perda_antigas if c in df_antigas.columns]
        df_export_perda_antiga = df_antigas[cols_export_perda_antiga].copy()
        excel_buffer = BytesIO()
        df_export_perda_antiga.to_excel(excel_buffer, index=False, engine='openpyxl')
        excel_data = excel_buffer.getvalue()
        st.download_button(
            label="📊 Download Excel (Perdas Antigas)",
            data=excel_data,
            file_name=f"perdas_antigas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Exporta a tabela de perdas antigas no formato Excel, contendo apenas as colunas relevantes."
        )
    else:
        st.info("Nenhuma perda antiga.")

    # ============================================================
    # UNION DEDUP: Unificar todas as listas com classificação unificada
    # ============================================================
    # Filtrar apenas CNPJs válidos (se necessário) e fazer union deduplicada
    df_risco_filtrado_final = df_risco_filtrado.copy() if not df_risco_filtrado.empty else pd.DataFrame()
    df_perda_recente_filtrado = df_perda_recente.copy() if not df_perda_recente.empty else pd.DataFrame()
    df_perda_antiga_filtrado = df_antigas.copy() if not df_antigas.empty else pd.DataFrame()
    
    # Garantir que todas tenham CNPJ_Normalizado
    listas_para_union = []
    if not df_risco_filtrado_final.empty and 'CNPJ_Normalizado' in df_risco_filtrado_final.columns:
        listas_para_union.append(df_risco_filtrado_final)
    if not df_perda_recente_filtrado.empty and 'CNPJ_Normalizado' in df_perda_recente_filtrado.columns:
        listas_para_union.append(df_perda_recente_filtrado)
    if not df_perda_antiga_filtrado.empty and 'CNPJ_Normalizado' in df_perda_antiga_filtrado.columns:
        listas_para_union.append(df_perda_antiga_filtrado)
    
    # Union dedup
    if listas_para_union:
        df_total_dedup = pd.concat(listas_para_union, ignore_index=True).drop_duplicates(subset=['CNPJ_Normalizado'], keep='first')
        total_union = len(df_total_dedup)
    else:
        df_total_dedup = pd.DataFrame()
        total_union = 0
    
    # Validação no App
    if total_union > 5634:
        st.error("❌ > Banco! Cheque vazamentos.")
    else:
        st.success(f"✅ {total_union}/5634 labs monitorados.")

    st.markdown("---")

    # ============================================================
    # 3. DETALHAMENTO POR SEMANA (O "Filme do Mês")
    # ============================================================
    
    # Obter mês de referência dos metadados
    mes_ref_num = met.get('referencia', {}).get('mes', datetime.now().month)
    ano_ref = met.get('referencia', {}).get('ano', datetime.now().year)
    meses_nomes_completos = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
                             "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    mes_nome = meses_nomes_completos[mes_ref_num - 1] if 1 <= mes_ref_num <= 12 else "Mês Atual"
    
    st.subheader(f"📆 Evolução do Mês (Semana a Semana) - {mes_nome}/{ano_ref}")
    
    # met já foi calculado acima nos KPIs executivos

    if met.get('semanas_detalhes'):
        dados_semanas = []

        # Indexar semanas vindas da métrica
        semanas_indexadas = {}
        for s in met['semanas_detalhes']:
            iso_week = s.get('iso_week')
            iso_year = s.get('iso_year')
            if not iso_week or not iso_year:
                continue
            key = (iso_year, iso_week)
            semanas_indexadas[key] = {
                'semana': s.get('semana', iso_week),
                'iso_week': iso_week,
                'iso_year': iso_year,
                'volume_total': s.get('volume_total', 0) or 0,
                'volume_semana_anterior': s.get('volume_semana_anterior', 0) or 0,
                'fechada': s.get('fechada', False)
            }

        # Complementar com o que vier do DF (Semanas_Mes_Atual)
        if 'Semanas_Mes_Atual' in df.columns:
            for val in df['Semanas_Mes_Atual']:
                if pd.isna(val):
                    continue
                try:
                    semanas_json = json.loads(val) if isinstance(val, str) else val
                except Exception:
                    continue
                if not isinstance(semanas_json, list):
                    continue

                for semana in semanas_json:
                    iso_week = semana.get('iso_week')
                    iso_year = semana.get('iso_year')
                    if not iso_week or not iso_year:
                        continue
                    if allowed_keys and (iso_year, iso_week) not in allowed_keys:
                        continue
                    key = (iso_year, iso_week)
                    entrada = semanas_indexadas.get(key, {
                        'semana': semana.get('semana', iso_week),
                        'iso_week': iso_week,
                        'iso_year': iso_year,
                        'volume_total': 0,
                        'volume_semana_anterior': 0,
                        'fechada': False
                    })
                    entrada['volume_total'] += semana.get('volume_util', 0) or 0
                    entrada['volume_semana_anterior'] += semana.get('volume_semana_anterior', 0) or 0
                    entrada['fechada'] = entrada['fechada'] or semana.get('fechada', False)
                    semanas_indexadas[key] = entrada

        # Garantir todas as semanas do mês atual (mesmo que zeradas)
        if not allowed_keys:
            hoje = datetime.now()
            primeiro_dia = datetime(hoje.year, hoje.month, 1)
            ultimo_dia = datetime(hoje.year, hoje.month, calendar.monthrange(hoje.year, hoje.month)[1])
            dia_corrente = primeiro_dia
            expected_keys = set()
            while dia_corrente <= ultimo_dia:
                iso = dia_corrente.isocalendar()
                expected_keys.add((iso.year, iso.week))
                dia_corrente += timedelta(days=1)
        else:
            expected_keys = allowed_keys

        for iso_year, iso_week in expected_keys:
            if (iso_year, iso_week) not in semanas_indexadas:
                fim_semana = datetime.fromisocalendar(iso_year, iso_week, 7)
                semanas_indexadas[(iso_year, iso_week)] = {
                    'semana': iso_week,
                    'iso_week': iso_week,
                    'iso_year': iso_year,
                    'volume_total': 0,
                    'volume_semana_anterior': 0,
                    'fechada': fim_semana.date() < datetime.now().date()
                }

        # Se houver seleção, garantir que só as semanas permitidas sejam exibidas
        if allowed_keys:
            semanas_indexadas = {k: v for k, v in semanas_indexadas.items() if k in allowed_keys}

        # Montar tabela final ordenada
        for iso_year, iso_week in sorted(semanas_indexadas.keys()):
            info = semanas_indexadas[(iso_year, iso_week)]
            vol_atual = info.get('volume_total', 0) or 0
            vol_ant = info.get('volume_semana_anterior', 0) or 0

            try:
                prev_monday = datetime.fromisocalendar(iso_year, iso_week, 1) - timedelta(days=7)
                prev_iso = prev_monday.isocalendar()
                prev_key = (prev_iso.year, prev_iso.week)
                prev_label = f"{prev_iso.week}/{prev_iso.year}"
            except Exception:
                prev_key = None
                prev_label = "-"

            if vol_ant == 0 and prev_key and prev_key in semanas_indexadas:
                vol_ant = semanas_indexadas[prev_key].get('volume_total', 0) or vol_ant

            if vol_ant == 0:
                variacao_pct = 0.0 if vol_atual == 0 else 100.0
            else:
                variacao_pct = ((vol_atual - vol_ant) / vol_ant) * 100

            try:
                semana_inicio = datetime.fromisocalendar(iso_year, iso_week, 1)
                semana_fim = semana_inicio + timedelta(days=6)
                intervalo_str = f"{semana_inicio:%d/%m}–{semana_fim:%d/%m}"
            except Exception:
                intervalo_str = "-"

            dados_semanas.append({
                "Semana": f"Semana {info.get('semana', iso_week)}",
                "ISO": iso_week,
                "Intervalo (seg-dom)": intervalo_str,
                "Semana Anterior (ISO)": prev_label,
                "Volume Útil": round(float(vol_atual), 1),
                "Volume Anterior": round(float(vol_ant), 1),
                "WoW %": round(variacao_pct, 2),
                "Status": "✅ Fechada" if info.get('fechada') else "🔄 Em Andamento",
                "is_active": not info.get('fechada')
            })

        df_semanas = pd.DataFrame(dados_semanas)
        
        def highlight_active_week(row):
            if row['is_active']:
                return ['background-color: #fff3cd; color: #856404'] * len(row)
            return [''] * len(row)
            
        st.dataframe(
            df_semanas.style.apply(highlight_active_week, axis=1),
            use_container_width=True,
            hide_index=True,
            column_config={
                "Semana": st.column_config.TextColumn("Semana", width="small", help="Semana na visão do mês corrente (1, 2, 3...)."),
                "ISO": st.column_config.NumberColumn("ISO Week", format="%d", help="Semana ISO correspondente."),
                "Intervalo (seg-dom)": st.column_config.TextColumn("Dias da Semana", help="Intervalo (segunda a domingo) que compõe a semana ISO."),
                "Semana Anterior (ISO)": st.column_config.TextColumn("Semana ISO anterior", help="ISO week imediatamente anterior (pode ser do mês/ano anterior)."),
                "Volume Útil": st.column_config.NumberColumn("Volume Realizado", format="%d"),
                "Volume Anterior": st.column_config.NumberColumn("Volume Anterior", format="%d"),
                "WoW %": st.column_config.NumberColumn("Variação WoW (%)", format="%.1f%%"),
                "Status": st.column_config.TextColumn("Status"),
                "is_active": None
            }
        )
    else:
        st.info("Detalhes semanais indisponíveis.")

    st.markdown("---")

    # --- GRÁFICOS: EVOLUÇÃO POR PORTE (SEPARADOS) ---
    st.subheader("📊 Evolução de Volume por Porte (Semana Anterior vs Atual)")
    
    # Definir ordem e lista de portes
    portes_config = [
        'Grande', 
        'Médio/Grande', 
        'Médio', 
        'Pequeno'
    ]
    
    # Criar containers verticais para cada gráfico
    for porte in portes_config:
        df_porte = df[df['Porte'] == porte]
        if not df_porte.empty:
            vol_ant = df_porte['WoW_Semana_Anterior'].sum()
            vol_atu = df_porte['WoW_Semana_Atual'].sum()
            
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=['Semana Ant.', 'Semana Atual'],
                y=[vol_ant, vol_atu],
                text=[f"{vol_ant:,.0f}", f"{vol_atu:,.0f}"],
                textposition='auto',
                marker_color=['#d1d5db', '#6BBF47' if vol_atu >= vol_ant else '#ef4444']
            ))
            fig.update_layout(
                title=f"<b>{porte}</b>",
                height=300,
                margin=dict(l=20, r=20, t=40, b=20),
                yaxis_visible=False
            )
            st.plotly_chart(fig, use_container_width=True, key=f"chart_porte_sem_{porte}")
        else:
            st.info(f"Sem dados para porte: {porte}")


# ============================================
# ABA 2: FECHAMENTO MENSAL (ESTRATÉGICO)
# ============================================

def renderizar_aba_fechamento_mensal(df: pd.DataFrame, metrics: KPIMetrics, filtros: Dict[str, Any]):
    st.markdown("## 📊 Fechamento Mensal (Estratégico)")
    st.caption("Comparativo: Realizado Mês Atual vs Baseline Mensal (Média dos Melhores Meses Históricos).")

    if 'CNPJ_Normalizado' not in df.columns and 'CNPJ_PCL' in df.columns:
        df['CNPJ_Normalizado'] = df['CNPJ_PCL'].apply(DataManager.normalizar_cnpj)

    portes_sel = filtros.get('portes')
    if portes_sel and 'Porte' in df.columns:
        df = df[df['Porte'].isin(portes_sel)].copy()

    # Preparar colunas e filtros específicos do mensal
    meses_nomes = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
    hoje = datetime.now()
    mes_atual = hoje.month
    ano_atual = hoje.year
    mes_anterior = 12 if mes_atual == 1 else mes_atual - 1
    ano_mes_anterior = ano_atual - 1 if mes_atual == 1 else ano_atual
    col_mes_anterior = f"N_Coletas_{meses_nomes[mes_anterior - 1]}_{str(ano_mes_anterior)[-2:]}"

    df_mensal = df.copy()
    # Remover laboratórios que não devem aparecer (ex.: Laboratório Cairo)
    if 'Nome_Fantasia_PCL' in df_mensal.columns:
        df_mensal = df_mensal[~df_mensal['Nome_Fantasia_PCL'].str.contains('cairo', case=False, na=False)]

    if col_mes_anterior in df_mensal.columns:
        df_mensal['Coletas_Mes_Anterior'] = pd.to_numeric(df_mensal[col_mes_anterior], errors='coerce').fillna(0).astype(int)
    else:
        df_mensal['Coletas_Mes_Anterior'] = 0

    if 'Baseline_Mensal' not in df_mensal.columns:
        df_mensal['Baseline_Mensal'] = 0
    if 'Coletas_Mes_Atual' not in df_mensal.columns:
        df_mensal['Coletas_Mes_Atual'] = 0

    # Gap no mesmo padrão do semanal: positivo quando o realizado está abaixo da referência
    df_mensal['Gap_Baseline'] = df_mensal['Baseline_Mensal'] - df_mensal['Coletas_Mes_Atual']
    df_mensal['Variacao_Baseline_Pct'] = df_mensal.apply(
        lambda r: calcular_variacao_percentual(r.get('Coletas_Mes_Atual'), r.get('Baseline_Mensal')),
        axis=1
    )
    df_mensal['Variacao_Mes_Anterior_Pct'] = df_mensal.apply(
        lambda r: calcular_variacao_percentual(r.get('Coletas_Mes_Atual'), r.get('Coletas_Mes_Anterior')),
        axis=1
    )

    if all(col in df_mensal.columns for col in ['Baseline_Mensal', 'Estado', 'Porte']):
        df_mensal['Baseline_Estado_Porte'] = df_mensal.groupby(['Estado', 'Porte'])['Baseline_Mensal'].transform('mean')
    else:
        df_mensal['Baseline_Estado_Porte'] = np.nan
    df_mensal['Variacao_Estado_Pct'] = df_mensal.apply(
        lambda r: calcular_variacao_percentual(r.get('Coletas_Mes_Atual'), r.get('Baseline_Estado_Porte')),
        axis=1
    )

    def _status_mensal(row):
        status = row.get('Status_Risco_V2', '')
        if isinstance(status, str) and status.strip():
            if 'Perda' in status:
                return 'Perda'
            return 'Sim'
        if bool(row.get('Risco_Por_Dias_Sem_Coleta', False)) or bool(row.get('Em_Risco', False)):
            return 'Sim'
        return 'Não'

    df_mensal['Status_Exibicao'] = df_mensal.apply(_status_mensal, axis=1)

    variacoes_opcoes = list(VARIACAO_QUEDA_FAIXAS.keys())
    variacao_baseline_sel = st.multiselect(
        'Filtro "Variação % vs Baseline Mensal"',
        options=variacoes_opcoes,
        default=st.session_state.get('filtro_variacao_baseline_mensal', ["Acima de 50%"]),
        help="Faixa de variação do realizado vs baseline.",
        key="filtro_variacao_baseline_mensal"
    )
    variacao_mes_ant_sel = st.multiselect(
        'Filtro "Variação % vs Mês Anterior"',
        options=variacoes_opcoes,
        default=st.session_state.get('filtro_variacao_mes_anterior', ["Acima de 50%"]),
        help="Faixa de variação do realizado vs mês anterior.",
        key="filtro_variacao_mes_anterior"
    )

    df_mensal_filtrado = aplicar_filtro_variacao_generica(
        df_mensal, 'Variacao_Baseline_Pct', variacao_baseline_sel, usar_valor_absoluto=True
    )
    df_mensal_filtrado = aplicar_filtro_variacao_generica(
        df_mensal_filtrado, 'Variacao_Mes_Anterior_Pct', variacao_mes_ant_sel, usar_valor_absoluto=True
    )

    # ============================================================
    # KPI BOX EXECUTIVO - TOPO DA ABA MENSAL (com filtros aplicados)
    # ============================================================
    metricas_mensal = calcular_metricas_fechamento_mensal(df_mensal_filtrado)
    met_mensal = metricas_mensal
    resumo_mensal = met_mensal.get('resumo_mensal', {})

    # Garantir que total_labs reflete o DataFrame filtrado atual
    if 'CNPJ_Normalizado' in df_mensal_filtrado.columns:
        total_labs_atual = int(df_mensal_filtrado['CNPJ_Normalizado'].nunique())
    else:
        total_labs_atual = len(df_mensal_filtrado)
    resumo_mensal['total_labs'] = total_labs_atual

    st.markdown("### 📈 KPIs Executivos")
    col1, col2, col3 = st.columns(3)
    col4, col5, col6 = st.columns(3)
    col1.metric(
        "Volume Mês Atual (realizado)",
        f"{resumo_mensal.get('volume_mes_atual', 0):,.0f}",
        help="Volume total realizado no mês corrente (até hoje)."
    )
    col2.metric(
        "Volume Mês Anterior",
        f"{resumo_mensal.get('volume_mes_anterior', 0):,.0f}",
        help="Volume total do mês anterior completo."
    )
    variacao_mes_anterior = resumo_mensal.get('variacao_mes_anterior_pct')
    col3.metric(
        "Variação % vs Mês Anterior",
        "—" if variacao_mes_anterior is None else f"{variacao_mes_anterior:+.1f}%",
        help="% de diferença do mês atual vs mês anterior."
    )
    col4.metric(
        "Baseline Mensal (top 3 meses)",
        f"{resumo_mensal.get('baseline_total', 0):,.0f}",
        help="Média real dos 3 maiores meses históricos de coleta (2024 + 2025)."
    )
    variacao_baseline = resumo_mensal.get('variacao_pct')
    col5.metric(
        "Variação % vs Baseline Mensal",
        "—" if variacao_baseline is None else f"{variacao_baseline:+.1f}%",
        help="% de diferença do Volume mês atual vs Baseline Mensal."
    )
    col6.metric(
        "Labs na Listagem (após filtro)",
        f"{resumo_mensal.get('total_labs', 0):,}",
        help="Quantos laboratórios aparecem na tabela abaixo após o filtro aplicado."
    )

    st.markdown("---")
    
    # ============================================================
    # GRÁFICO: VOLUME MENSAL EVOLUTIVO 2024-2025
    # ============================================================
    st.subheader("📊 Volume Mensal Evolutivo 2024-2025")
    
    if df_mensal_filtrado.empty:
        st.info("Nenhum laboratório encontrado com os filtros selecionados.")
        return
    
    meses = []
    volumes = []
    cols_meses_hist = [c for c in df_mensal_filtrado.columns if c.startswith('N_Coletas_') and ('_24' in c or '_25' in c)]
    
    for c in cols_meses_hist:
        vol_total = df_mensal_filtrado[c].sum()
        if vol_total > 0:
            partes = c.split('_')
            if len(partes) >= 4:
                mes = partes[2]
                ano = partes[3]
                label = f"{mes}/{ano[-2:]}"
                meses.append(label)
                volumes.append(vol_total)
    
    if meses and volumes:
        df_hist = pd.DataFrame({"Mês": meses, "Volume": volumes})
        
        def criar_chave_ordenacao(mes_str):
            meses_map = {
                'Jan': 1, 'Fev': 2, 'Mar': 3, 'Abr': 4, 'Mai': 5, 'Jun': 6,
                'Jul': 7, 'Ago': 8, 'Set': 9, 'Out': 10, 'Nov': 11, 'Dez': 12
            }
            try:
                mes, ano = mes_str.split('/')
                return (int(ano), meses_map.get(mes, 0))
            except:
                return (0, 0)
        
        df_hist['_ordem'] = df_hist['Mês'].apply(criar_chave_ordenacao)
        df_hist = df_hist.sort_values('_ordem').drop('_ordem', axis=1)
        
        fig = px.bar(
            df_hist, 
            x="Mês", 
            y="Volume", 
            text="Volume",
            title="<b>Volume Mensal 2024-2025</b>",
            labels={"Volume": "Volume de Coletas", "Mês": "Mês/Ano"}
        )
        fig.update_traces(
            texttemplate='%{text:,.0f}', 
            textposition='outside',
            marker_color='#6BBF47'
        )
        
        fig.update_layout(
            height=400,
            xaxis_tickangle=-45,
            showlegend=True
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Dados históricos mensais não disponíveis para visualização.")
    
    # Variável Controle Mensal (Média Estado)
    if 'Controle_Mensal_Estado' not in df_mensal_filtrado.columns:
        media_estado = df_mensal_filtrado.groupby('Estado')['Coletas_Mes_Atual'].transform('mean')
        df_mensal_filtrado['Controle_Mensal_Estado'] = media_estado

    st.subheader("📋 Listagem Geral (Fechamento Mensal)")
    df_sorted = df_mensal_filtrado.sort_values('Gap_Baseline', ascending=True)

    def _formatar_pico(valor, mes):
        if pd.isna(valor):
            return "—"
        try:
            val_fmt = f"{float(valor):,.0f}"
        except Exception:
            val_fmt = str(valor)
        if mes:
            return f"{val_fmt} - {mes}"
        return val_fmt

    df_sorted['Pico_2025'] = df_sorted.apply(
        lambda r: _formatar_pico(r.get('Maior_N_Coletas_Mes_2025'), r.get('Mes_Maior_Coleta_2025')),
        axis=1
    )
    df_sorted['Pico_2024'] = df_sorted.apply(
        lambda r: _formatar_pico(r.get('Maior_N_Coletas_Mes_2024'), r.get('Mes_Historico')),
        axis=1
    )

    cols_mensal = [
        'Nome_Fantasia_PCL', 'CNPJ_Normalizado', 'VIP', 'Rede', 'Estado', 'Porte',
        'Baseline_Mensal', 'Pico_2025', 'Pico_2024',
        'Coletas_Mes_Anterior', 'Coletas_Mes_Atual',
        'Gap_Baseline', 'Variacao_Baseline_Pct', 'Variacao_Mes_Anterior_Pct',
        'Baseline_Estado_Porte', 'Variacao_Estado_Pct', 'Status_Exibicao'
    ]
    
    cols_final = [c for c in cols_mensal if c in df_sorted.columns]

    df_mensal_display = adicionar_coluna_detalhes(df_sorted[cols_final], 'CNPJ_Normalizado')
    df_mensal_display = df_mensal_display[cols_final].reset_index(drop=True)
    evento_mensal = st.dataframe(
        df_mensal_display,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Nome_Fantasia_PCL": st.column_config.TextColumn("Laboratório", width="medium", help="Nome fantasia cadastrado no CRM"),
            "CNPJ_Normalizado": st.column_config.TextColumn("CNPJ", width="medium", help="Identificador único (apenas números)"),
            "VIP": st.column_config.TextColumn("VIP", width="small", help="Sinaliza se o laboratório está na carteira VIP"),
            "Rede": st.column_config.TextColumn("Rede", width="medium", help="Rede ou grupo ao qual o laboratório pertence"),
            "Estado": st.column_config.TextColumn("UF", width="small", help="Estado do laboratório"),
            "Porte": st.column_config.TextColumn("Porte", width="small", help="Segmentação de porte"),
            "Baseline_Mensal": st.column_config.NumberColumn("Baseline Mensal", format="%d", help="Média dos 3 maiores meses (2024+2025)"),
            "Pico_2025": st.column_config.TextColumn("Pico 2025", width="medium", help="Maior mês de 2025 (volume + mês)"),
            "Pico_2024": st.column_config.TextColumn("Pico 2024", width="medium", help="Maior mês de 2024 (volume + mês)"),
            "Coletas_Mes_Anterior": st.column_config.NumberColumn("Volume Mês Anterior", format="%d", help="Volume completo do mês anterior"),
            "Coletas_Mes_Atual": st.column_config.NumberColumn("Volume Mês Atual", format="%d", help="Volume realizado até hoje no mês corrente"),
            "Gap_Baseline": st.column_config.NumberColumn("Gap Volume vs Baseline", format="%d", help="Diferença absoluta (realizado - baseline)"),
            "Variacao_Baseline_Pct": st.column_config.NumberColumn("Variação % vs Baseline", format="%.1f%%", help="% do realizado vs baseline"),
            "Variacao_Mes_Anterior_Pct": st.column_config.NumberColumn("Variação % vs Mês Anterior", format="%.1f%%", help="% do mês atual vs mês anterior"),
            "Baseline_Estado_Porte": st.column_config.NumberColumn("Média Mensal Estado (baseline)", format="%d", help="Média dos 3 maiores meses dos labs do mesmo porte no mesmo estado"),
            "Variacao_Estado_Pct": st.column_config.NumberColumn("Variação % vs Média Estado", format="%.1f%%", help="Comparação do realizado com a média do estado"),
            "Status_Exibicao": st.column_config.TextColumn("Em_Risco / Status", width="small", help="Sim/Não ou Perda, quando aplicável")
        },
        selection_mode="single-row",
        on_select="rerun",
        key="tbl_mensal_master"
    )
    _processar_evento_selecao(evento_mensal, df_mensal_display)
    st.caption("Use a caixa de seleção para abrir automaticamente a Análise Detalhada.")

    # Botão Exportação Excel Mensal (apenas colunas de exibição)
    # cols_final já foi filtrado para incluir apenas colunas que existem em df_sorted
    df_export_mensal = df_sorted[cols_final].copy()
    excel_buffer = BytesIO()
    df_export_mensal.to_excel(excel_buffer, index=False, engine='openpyxl')
    excel_data = excel_buffer.getvalue()
    st.download_button(
        label="📊 Download Excel (Fechamento Mensal)",
        data=excel_data,
        file_name=f"relatorio_fechamento_mensal_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Exporta a listagem geral do fechamento mensal no formato Excel, contendo apenas as colunas relevantes."
    )

    st.markdown("---")


# ============================================
# FUNÇÕES AUXILIARES DO SISTEMA V2
# ============================================

def exibir_bloco_concorrencia(row: pd.Series):
    """
    Exibe bloco de alerta de concorrência se laboratório apareceu no Gralab.
    
    Args:
        row: Série com dados do laboratório
    """
    if row.get('Apareceu_Gralab', False):
        data_gralab = row.get('Gralab_Data')
        tipo_gralab = row.get('Gralab_Tipo', 'Não especificado')
        
        if pd.notna(data_gralab):
            data_formatada = pd.to_datetime(data_gralab).strftime('%d/%m/%Y')
        else:
            data_formatada = 'Data não disponível'
        
        st.warning(f"""
        ⚠️ **ALERTA DE CONCORRÊNCIA**
        - **Data**: {data_formatada}
        - **Tipo**: {tipo_gralab}
        - CNPJ apareceu no sistema Gralab nos últimos 14 dias
        """)
        
        with st.expander("ℹ️ O que isso significa?"):
            st.info(HELPERS_V2['concorrencia'])


def exibir_metricas_v2(row: pd.Series):
    """
    Exibe métricas do sistema v2 (Baseline, WoW, Porte, etc).
    
    Args:
        row: Série com dados do laboratório
    """
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        baseline = row.get('Baseline_Mensal', 0)
        st.metric(
            "Baseline Mensal",
            f"{baseline:.0f}",
            help=HELPERS_V2['baseline_mensal']
        )
    
    with col2:
        wow_pct = row.get('WoW_Percentual', 0)
        cor_wow = "normal" if wow_pct >= 0 else "inverse"
        st.metric(
            "WoW",
            f"{wow_pct:+.1f}%",
            delta=f"{wow_pct:.1f}%",
            delta_color=cor_wow,
            help=HELPERS_V2['wow']
        )
    
    with col3:
        porte = row.get('Porte', 'Desconhecido')
        st.metric(
            "Porte",
            porte,
            help=HELPERS_V2['porte']
        )


def exibir_helper_icone(chave_helper: str, label: str = "ℹ️"):
    """
    Exibe ícone de ajuda com tooltip.
    
    Args:
        chave_helper: Chave do dicionário HELPERS_V2
        label: Rótulo do ícone
    """
    if chave_helper in HELPERS_V2:
        st.markdown(f"**{label}**")
        with st.expander("Saiba mais"):
            st.info(HELPERS_V2[chave_helper])


class FilterManager:
    """Gerenciador de filtros da interface."""
    def __init__(self):
        self.filtros = {}
    def renderizar_sidebar_filtros(self, df: pd.DataFrame) -> Dict[str, Any]:
        """Renderiza filtros otimizados na sidebar."""
        st.sidebar.markdown('<div class="sidebar-header" style="font-size: 1rem; font-weight: 600; color: var(--primary-color);">🔧 Filtros</div>', unsafe_allow_html=True)
        filtros = {}
        
        # Filtro por Porte
        if 'Porte' in df.columns:
            portes_lista = (
                df['Porte']
                .astype(str)
                .str.strip()
                .replace({'nan': '', 'None': ''})
            )
            # Ordem desejada
            ordem_porte = {'Grande': 0, 'Médio/Grande': 1, 'Médio': 2, 'Pequeno': 3}
            portes_opcoes = sorted({p for p in portes_lista if p}, key=lambda x: ordem_porte.get(x, 99))
            # Default: Grande e Médio/Grande
            default_portes = [p for p in ['Grande', 'Médio/Grande'] if p in portes_opcoes]
            if not default_portes:  # Fallback se nenhum dos dois existir
                default_portes = portes_opcoes
        else:
            portes_opcoes = []
            default_portes = []
        
        filtros['portes'] = st.sidebar.multiselect(
            "🏗️ Porte do Laboratório",
            options=portes_opcoes,
            default=default_portes,
            help="Selecione um ou mais portes para filtrar os laboratórios exibidos."
        )
        
        # ===== FILTRO UF (SISTEMA V2 - PRIORITÁRIO) =====
        if 'Estado' in df.columns:
            # Filtrar valores válidos de UF (não vazios, não NaN, não None)
            ufs_disponiveis = sorted([
                str(uf).strip() 
                for uf in df['Estado'].unique() 
                if pd.notna(uf) and str(uf).strip() and str(uf).strip().upper() != 'NAN'
            ])
            if ufs_disponiveis:
                filtros['uf_selecionada'] = st.sidebar.selectbox(
                    "🗺️ Filtrar por UF",
                    options=['Todas'] + ufs_disponiveis,
                    index=0,
                    help="Visualizar alertas e dados segmentados por estado"
                )
            else:
                filtros['uf_selecionada'] = 'Todas'
        else:
            filtros['uf_selecionada'] = 'Todas'
        
        # Filtro VIP com opção de alternar (desativado por padrão)
        filtros['apenas_vip'] = st.sidebar.toggle(
            "🌟 Apenas Clientes VIP",
            value=False,
            help="Ative para mostrar apenas clientes VIP, desative para mostrar todos"
        )
     
        # Separador visual
        st.sidebar.markdown("---")
        
        # Filtro por representante
        if 'Representante_Nome' in df.columns:
            representantes_lista = (
                df['Representante_Nome']
                .astype(str)
                .str.strip()
                .replace({'nan': '', 'None': ''})
            )
            representantes_opcoes = sorted({r for r in representantes_lista if r})
        else:
            representantes_opcoes = []
        filtros['representantes'] = st.sidebar.multiselect(
            "👤 Representantes",
            options=representantes_opcoes,
            help="Selecione um ou mais representantes para filtrar os laboratórios exibidos."
        )

        st.sidebar.markdown("---")
        

        st.sidebar.markdown("---")
        # Filtro por período - Anos e Meses (dados mensais)
        st.sidebar.markdown("**📅 Período de Análise (Mensal)**")
        # Verificar anos disponíveis nos dados
        anos_disponiveis = []
        if 'N_Coletas_Jan_24' in df.columns:
            anos_disponiveis.append(2024)
        if 'N_Coletas_Jan_25' in df.columns:
            anos_disponiveis.append(2025)
        if not anos_disponiveis:
            st.sidebar.warning("⚠️ Nenhum dado mensal encontrado")
            anos_disponiveis = [2024, 2025] # fallback
        # Seleção de ano
        ano_selecionado = st.sidebar.selectbox(
            "📊 Ano de Análise:",
            options=anos_disponiveis,
            index=len(anos_disponiveis)-1, # Padrão: último ano disponível
            help="Selecione o ano para análise mensal"
        )
        # Mapeamento de meses
        meses_map = {
            'Jan': 'Janeiro', 'Fev': 'Fevereiro', 'Mar': 'Março', 'Abr': 'Abril',
            'Mai': 'Maio', 'Jun': 'Junho', 'Jul': 'Julho', 'Ago': 'Agosto',
            'Set': 'Setembro', 'Out': 'Outubro', 'Nov': 'Novembro', 'Dez': 'Dezembro'
        }
        # Meses disponíveis para o ano selecionado
        sufixo_ano = str(ano_selecionado)[-2:] # '24' ou '25'
        meses_disponiveis = []
        for mes_codigo in ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun',
                          'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']:
            coluna_mes = f'N_Coletas_{mes_codigo}_{sufixo_ano}'
            if coluna_mes in df.columns:
                meses_disponiveis.append(mes_codigo)
        if not meses_disponiveis:
            st.sidebar.warning(f"⚠️ Nenhum mês encontrado para {ano_selecionado}")
            meses_disponiveis = ['Jan', 'Fev', 'Mar'] # fallback
        # Seleção de meses
        meses_opcoes = [f"{mes} - {meses_map.get(mes, mes)}" for mes in meses_disponiveis]
        meses_selecionados_opcoes = st.sidebar.multiselect(
            f"📅 Meses de {ano_selecionado}:",
            options=meses_opcoes,
            default=meses_opcoes, # Todos selecionados por padrão
            help=f"Selecione os meses de {ano_selecionado} para análise. Deixe todos selecionados para visão completa do ano.",
            key=f"meses_{ano_selecionado}"
        )
        # Converter para códigos de mês
        meses_selecionados = []
        for opcao in meses_selecionados_opcoes:
            mes_codigo = opcao.split(' - ')[0]
            if mes_codigo in meses_disponiveis:
                meses_selecionados.append(mes_codigo)
        # Armazenar filtros para uso posterior
        filtros['ano_selecionado'] = ano_selecionado
        filtros['meses_selecionados'] = meses_selecionados
        filtros['sufixo_ano'] = sufixo_ano
        # Mostrar período selecionado (texto discreto)
        meses_nomes = [meses_map.get(mes, mes) for mes in meses_selecionados]
        periodo_texto = f"{ano_selecionado}: {', '.join(meses_nomes[:3])}" # Max 3 meses no texto
        if len(meses_selecionados) > 3:
            periodo_texto += f" +{len(meses_selecionados)-3}..."
        st.sidebar.markdown(f"<small>📊 {periodo_texto}</small>", unsafe_allow_html=True)
        self.filtros = filtros
        return filtros
    def aplicar_filtros(self, df: pd.DataFrame, filtros: Dict[str, Any]) -> pd.DataFrame:
        """Aplica filtros otimizados ao DataFrame - Atualizado para coerência."""
        if df.empty:
            return df
        df_filtrado = df.copy()
        
        # ===== FILTRO UF (SISTEMA V2 - PRIORITÁRIO) =====
        if filtros.get('uf_selecionada') and filtros['uf_selecionada'] != 'Todas':
            if 'Estado' in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado['Estado'] == filtros['uf_selecionada']]
        
        # Filtro VIP (sempre ativo)
        if filtros.get('apenas_vip', False):
            try:
                # Carregar dados VIP
                df_vip = DataManager.carregar_dados_vip()
                if df_vip is not None and not df_vip.empty:
                    # Normalizar CNPJs para match com tratamento de erro
                    df_filtrado['CNPJ_Normalizado'] = df_filtrado['CNPJ_PCL'].apply(
                        lambda x: ''.join(filter(str.isdigit, str(x))) if pd.notna(x) and str(x).strip() != '' else ''
                    )
                    df_vip['CNPJ_Normalizado'] = df_vip['CNPJ'].apply(
                        lambda x: ''.join(filter(str.isdigit, str(x))) if pd.notna(x) and str(x).strip() != '' else ''
                    )
                 
                    # Filtrar apenas registros que estão na lista VIP (com validação)
                    if 'CNPJ_Normalizado' in df_filtrado.columns and 'CNPJ_Normalizado' in df_vip.columns:
                        # Remover CNPJs vazios antes do match
                        df_filtrado = df_filtrado[df_filtrado['CNPJ_Normalizado'] != '']
                        df_vip_clean = df_vip[df_vip['CNPJ_Normalizado'] != '']
                     
                        if not df_vip_clean.empty:
                            df_filtrado = df_filtrado[df_filtrado['CNPJ_Normalizado'].isin(df_vip_clean['CNPJ_Normalizado'])]
                        else:
                            # Se não há CNPJs válidos na lista VIP, retornar DataFrame vazio
                            return pd.DataFrame()
                    else:
                        # Se as colunas não existem, retornar DataFrame vazio
                        return pd.DataFrame()
                else:
                    # Se não há dados VIP, retornar DataFrame vazio
                    return pd.DataFrame()
            except Exception as e:
                # Em caso de erro, retornar DataFrame vazio e log do erro
                st.error(f"Erro ao aplicar filtro VIP: {str(e)}")
                return pd.DataFrame()
        # Filtro por período (compatibilidade com filtros antigos)
        if 'Data_Analise' in df_filtrado.columns and filtros.get('data_inicio') and filtros.get('data_fim'):
            try:
                # Garantir que as datas sejam do tipo date
                data_inicio = filtros['data_inicio']
                data_fim = filtros['data_fim']
                # Se for datetime, converter para date
                if hasattr(data_inicio, 'date'):
                    data_inicio = data_inicio.date()
                if hasattr(data_fim, 'date'):
                    data_fim = data_fim.date()
                # Verificar se a coluna Data_Analise é do tipo datetime
                if df_filtrado['Data_Analise'].dtype == 'object':
                    # Tentar converter para datetime
                    df_filtrado['Data_Analise'] = pd.to_datetime(df_filtrado['Data_Analise'], errors='coerce')
                # Aplicar filtro apenas se a conversão foi bem-sucedida
                if df_filtrado['Data_Analise'].dtype.name.startswith('datetime'):
                    df_filtrado = df_filtrado[
                        (df_filtrado['Data_Analise'].dt.date >= data_inicio) &
                        (df_filtrado['Data_Analise'].dt.date <= data_fim)
                    ]
            except Exception as e:
                # Em caso de erro no filtro de data, continuar sem filtrar
                st.warning(f"Aviso: Erro ao aplicar filtro de período: {str(e)}")
                pass
        # Filtro por representante
        representantes_sel = filtros.get('representantes', [])
        if representantes_sel and 'Representante_Nome' in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado['Representante_Nome'].isin(representantes_sel)]
        
        # Filtro por porte
        portes_sel = filtros.get('portes', [])
        if portes_sel and 'Porte' in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado['Porte'].isin(portes_sel)]
        
        # Para dados mensais, o filtro principal será usado nos cálculos dos gráficos
        # Os filtros 'ano_selecionado', 'meses_selecionados' e 'sufixo_ano' são usados
        # diretamente nas funções de cálculo dos gráficos
        return df_filtrado
class KPIManager:
    """Gerenciador de cálculos de KPIs - Atualizado para coerência entre telas."""
    @staticmethod
    def calcular_kpis(df: pd.DataFrame) -> KPIMetrics:
        if df.empty:
            return KPIMetrics()
        metrics = KPIMetrics()
        df_recent = df[df['Dias_Sem_Coleta'] <= 90].copy() if 'Dias_Sem_Coleta' in df.columns else df.copy()
        metrics.total_labs = len(df_recent)
        # Distribuição por Risco_Diario
        labs_normal = labs_atencao = labs_moderado = labs_alto = labs_critico = 0
        if 'Risco_Diario' in df_recent.columns:
            c = df_recent['Risco_Diario'].value_counts()
            labs_normal = c.get('🟢 Normal', 0)
            labs_atencao = c.get('🟡 Atenção', 0)
            labs_moderado = c.get('🟠 Moderado', 0)
            labs_alto = c.get('🔴 Alto', 0)
            labs_critico = c.get('⚫ Crítico', 0)
        metrics.labs_baixo_risco = labs_normal + labs_atencao
        metrics.labs_medio_risco = labs_moderado
        metrics.labs_alto_risco = labs_alto + labs_critico
        metrics.labs_em_risco = labs_moderado + labs_alto + labs_critico
        metrics.labs_critico = labs_critico
        metrics.labs_normal_count = labs_normal
        metrics.labs_atencao_count = labs_atencao
        metrics.labs_moderado_count = labs_moderado
        metrics.labs_alto_count = labs_alto
        metrics.labs_critico_count = labs_critico
        metrics.churn_rate = (metrics.labs_em_risco / metrics.total_labs * 100) if metrics.total_labs else 0
        # Labs abaixo de MM7 por contexto
        def _count_below(column_name: str) -> int:
            if column_name not in df_recent.columns or 'Vol_Hoje' not in df_recent.columns:
                return 0
            serie_ref = pd.to_numeric(df_recent[column_name], errors='coerce')
            vol_hoje = pd.to_numeric(df_recent['Vol_Hoje'], errors='coerce').fillna(0)
            mask_valid = serie_ref.notna() & (serie_ref > 0)
            return int(((vol_hoje < serie_ref) & mask_valid).sum())

        metrics.labs_abaixo_mm7_br = _count_below('MM7_BR')
        metrics.labs_abaixo_mm7_uf = _count_below('MM7_UF')
        denominator = metrics.total_labs or 0
        metrics.labs_abaixo_mm7_br_pct = (
            metrics.labs_abaixo_mm7_br / denominator * 100 if denominator else 0.0
        )
        metrics.labs_abaixo_mm7_uf_pct = (
            metrics.labs_abaixo_mm7_uf / denominator * 100 if denominator else 0.0
        )
        # Total coletas 2025
        meses_2025 = ChartManager._meses_ate_hoje(df_recent, 2025)
        cols = [f'N_Coletas_{m}_25' for m in meses_2025 if f'N_Coletas_{m}_25' in df_recent.columns]
        metrics.total_coletas = int(df_recent[cols].sum().sum()) if cols else 0
        # Volumes diários
        metrics.vol_hoje_total = int(df_recent['Vol_Hoje'].fillna(0).sum()) if 'Vol_Hoje' in df_recent.columns else 0
        metrics.vol_d1_total = int(df_recent['Vol_D1'].fillna(0).sum()) if 'Vol_D1' in df_recent.columns else 0
        # Recuperação e zeros consecutivos
        if 'Recuperacao' in df_recent.columns:
            metrics.labs_recuperando = int(df_recent['Recuperacao'].fillna(False).sum())
        if {'Vol_Hoje', 'Vol_D1'}.issubset(df_recent.columns):
            zeros_48h = df_recent[
                df_recent['Vol_Hoje'].fillna(0).eq(0) &
                df_recent['Vol_D1'].fillna(0).eq(0)
            ]
            metrics.labs_sem_coleta_48h = len(zeros_48h)
        # Ativos recentes
        if 'Dias_Sem_Coleta' in df_recent.columns and metrics.total_labs > 0:
            ativos_7d_df = df_recent[df_recent['Dias_Sem_Coleta'] <= 7]
            ativos_30d_df = df_recent[df_recent['Dias_Sem_Coleta'] <= 30]
            metrics.ativos_7d_count = len(ativos_7d_df)
            metrics.ativos_30d_count = len(ativos_30d_df)
            metrics.ativos_7d = metrics.ativos_7d_count / metrics.total_labs * 100
            metrics.ativos_30d = metrics.ativos_30d_count / metrics.total_labs * 100
        return metrics
class ChartManager:
    """Gerenciador de criação de gráficos - Atualizado com correções de bugs e layouts."""
    @staticmethod
    def _meses_ate_hoje(df: pd.DataFrame, ano: int) -> list:
        """Retorna lista de códigos de meses disponíveis até o mês corrente para o ano informado.
        - Garante ordem cronológica correta
        - Considera apenas colunas que existem no DataFrame
        - Para anos anteriores ao corrente, considera até Dezembro
        """
        meses_ordem = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
        ano_atual = pd.Timestamp.today().year
        limite_mes = pd.Timestamp.today().month if ano == ano_atual else 12
        meses_limite = meses_ordem[:limite_mes]
        sufixo = str(ano)[-2:]
        return [m for m in meses_limite if f'N_Coletas_{m}_{sufixo}' in df.columns]
    
    @staticmethod
    def _obter_ultimo_mes_fechado(df: pd.DataFrame, dia_corte: int = 5) -> dict:
        """
        Retorna informações sobre o último mês fechado disponível.
        
        Se estamos antes do dia 'dia_corte' do mês atual, considera o mês anterior.
        Pode retornar dados de 2024 se não houver dados de 2025.
        
        Args:
            df: DataFrame com os dados de coletas
            dia_corte: Dia do mês a partir do qual considera o mês atual válido (default: 5)
        
        Returns:
            dict com:
                - 'mes': código do mês (ex: 'Out')
                - 'ano': ano (ex: 2025)
                - 'sufixo': sufixo de 2 dígitos do ano (ex: '25')
                - 'coluna': nome da coluna no DataFrame (ex: 'N_Coletas_Out_25')
                - 'display': string para exibição (ex: 'Out/2025')
                - 'ano_comparacao': ano a ser usado para comparações
        """
        meses_ordem = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
        hoje = pd.Timestamp.today()
        dia_atual = hoje.day
        mes_atual = hoje.month
        ano_atual = hoje.year
        
        # Determinar qual mês considerar baseado no dia de corte
        if dia_atual < dia_corte:
            # Se estamos antes do dia de corte, usar o mês anterior
            if mes_atual == 1:
                # Se estamos em janeiro, voltar para dezembro do ano anterior
                mes_referencia = 12
                ano_referencia = ano_atual - 1
            else:
                mes_referencia = mes_atual - 1
                ano_referencia = ano_atual
        else:
            # Usar o mês atual
            mes_referencia = mes_atual
            ano_referencia = ano_atual
        
        # Tentar encontrar coluna começando pelo ano de referência
        mes_codigo = meses_ordem[mes_referencia - 1]
        sufixo_ref = str(ano_referencia)[-2:]
        coluna_ref = f'N_Coletas_{mes_codigo}_{sufixo_ref}'
        
        # Verificar se a coluna existe no DataFrame
        if coluna_ref in df.columns:
            return {
                'mes': mes_codigo,
                'ano': ano_referencia,
                'sufixo': sufixo_ref,
                'coluna': coluna_ref,
                'display': f'{mes_codigo}/{ano_referencia}',
                'ano_comparacao': ano_referencia
            }
        
        # Se não encontrou, tentar buscar o último mês disponível retroativamente
        # Tentar meses anteriores no mesmo ano
        for mes_idx in range(mes_referencia - 1, 0, -1):
            mes_codigo = meses_ordem[mes_idx - 1]
            coluna_teste = f'N_Coletas_{mes_codigo}_{sufixo_ref}'
            if coluna_teste in df.columns:
                return {
                    'mes': mes_codigo,
                    'ano': ano_referencia,
                    'sufixo': sufixo_ref,
                    'coluna': coluna_teste,
                    'display': f'{mes_codigo}/{ano_referencia}',
                    'ano_comparacao': ano_referencia
                }
        
        # Se não encontrou no ano atual/referência, tentar ano anterior
        ano_anterior = ano_referencia - 1
        sufixo_anterior = str(ano_anterior)[-2:]
        for mes_idx in range(12, 0, -1):
            mes_codigo = meses_ordem[mes_idx - 1]
            coluna_teste = f'N_Coletas_{mes_codigo}_{sufixo_anterior}'
            if coluna_teste in df.columns:
                return {
                    'mes': mes_codigo,
                    'ano': ano_anterior,
                    'sufixo': sufixo_anterior,
                    'coluna': coluna_teste,
                    'display': f'{mes_codigo}/{ano_anterior}',
                    'ano_comparacao': ano_anterior
                }
        
        # Se não encontrou nenhuma coluna, retornar vazio
        return {
            'mes': None,
            'ano': None,
            'sufixo': None,
            'coluna': None,
            'display': 'N/A',
            'ano_comparacao': ano_atual
        }
    
    @staticmethod
    def criar_grafico_distribuicao_risco(df: pd.DataFrame):
        if df.empty:
            st.info("📊 Nenhum dado disponível para o gráfico")
            return
        if 'Risco_Diario' not in df.columns:
            st.warning("⚠️ Coluna 'Risco_Diario' não encontrada nos dados.")
            return
        status_counts = df['Risco_Diario'].value_counts()
        cores_map = {
            '🟢 Normal': '#16A34A',
            '🟡 Atenção': '#F59E0B',
            '🟠 Moderado': '#FB923C',
            '🔴 Alto': '#DC2626',
            '⚫ Crítico': '#111827'
        }
        fig = px.pie(
            values=status_counts.values,
            names=status_counts.index,
            title="📊 Distribuição de Risco Diário<br><sup>Baseado em dias úteis e reduções vs. MM7_BR/MM7_UF/MM7_CIDADE</sup>",
            color=status_counts.index,
            color_discrete_map=cores_map
        )
        fig.update_traces(
            textposition='inside',
            textinfo='percent+label+value',
            texttemplate='%{label}<br>%{value} labs<br>(%{percent})',
            hovertemplate='<b>%{label}</b><br>%{value} laboratórios<br>%{percent}<extra></extra>'
        )
        fig.update_layout(
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5),
            height=500,
            margin=dict(l=40, r=40, t=40, b=40)
        )
        st.plotly_chart(fig, width='stretch')
    @staticmethod
    def criar_grafico_top_labs(df: pd.DataFrame, top_n: int = 10):
        if df.empty:
            st.info("📊 Nenhum dado disponível para o gráfico")
            return
        if 'Risco_Diario' not in df.columns:
            st.warning("⚠️ Coluna 'Risco_Diario' não encontrada nos dados.")
            return
        labs_risco = df[df['Risco_Diario'].isin(['🟠 Moderado', '🔴 Alto', '⚫ Crítico'])].copy()
        if labs_risco.empty:
            st.info("✅ Nenhum laboratório em risco encontrado!")
            return
        # Ordenar por maior queda vs MM7 e menor volume do dia
        if 'Delta_MM7' in labs_risco.columns:
            labs_risco = labs_risco.sort_values(['Delta_MM7', 'Vol_Hoje'], ascending=[True, True])
        else:
            labs_risco = labs_risco.sort_values('Vol_Hoje', ascending=True)
        cores_map = {'🟠 Moderado': '#FB923C', '🔴 Alto': '#DC2626', '⚫ Crítico': '#111827'}
        fig = px.bar(
            labs_risco.head(top_n),
            x='Vol_Hoje',
            y='Nome_Fantasia_PCL',
            orientation='h',
            title=f"🚨 Top {top_n} Laboratórios em Risco (Diário)<br><sup>Classificação baseada em dias úteis</sup>",
            color='Risco_Diario',
            color_discrete_map=cores_map,
            text='Delta_MM7'
        )
        fig.update_traces(texttemplate='%{text:.1f}% vs MM7', textposition='outside')
        fig.update_layout(
            yaxis={'categoryorder': 'total ascending'},
            xaxis_title="Coletas (Último Dia Útil)",
            yaxis_title="Laboratório",
            showlegend=True,
            height=500,
            margin=dict(l=40, r=40, t=40, b=100)
        )
        st.plotly_chart(fig, width='stretch')
    @staticmethod
    def criar_grafico_media_diaria(
        df: pd.DataFrame,
        lab_cnpj: Optional[str] = None,
        lab_nome: Optional[str] = None
    ):
        """Cria gráfico de média diária por mês usando dados reais de 2025."""
        if df.empty:
            st.info("📊 Nenhum dado disponível para o gráfico")
            return
        if not lab_cnpj and not lab_nome:
            st.info("📊 Selecione um laboratório para visualizar a média diária")
            return

        df_ref = df
        if lab_cnpj and 'CNPJ_Normalizado' not in df_ref.columns and 'CNPJ_PCL' in df_ref.columns:
            df_ref = df_ref.copy()
            df_ref['CNPJ_Normalizado'] = df_ref['CNPJ_PCL'].apply(DataManager.normalizar_cnpj)

        if lab_cnpj and 'CNPJ_Normalizado' in df_ref.columns:
            lab_data = df_ref[df_ref['CNPJ_Normalizado'] == lab_cnpj]
        else:
            lab_data = df_ref[df_ref['Nome_Fantasia_PCL'] == lab_nome]

        if lab_data.empty:
            st.info("📊 Laboratório não encontrado")
            return

        lab = lab_data.iloc[0]
        nome_exibicao = lab_nome or lab.get('Nome_Fantasia_PCL') or lab_cnpj
        
        # Verificar se temos dados diários reais de 2025
        if 'Dados_Diarios_2025' not in lab or pd.isna(lab['Dados_Diarios_2025']) or lab['Dados_Diarios_2025'] == '{}':
            st.info("📊 Nenhum dado diário disponível para 2025 para este laboratório.")
            return
        
        import json
        try:
            # Carregar dados diários reais
            dados_diarios = json.loads(lab['Dados_Diarios_2025'])
        except (json.JSONDecodeError, TypeError):
            st.info("📊 Erro ao processar dados diários para este laboratório.")
            return
        
        if not dados_diarios:
            st.info("📊 Nenhum dado diário disponível para 2025.")
            return
        
        # Calcular média diária real baseada em dias com coleta
        meses_ordem = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
        medias_diarias = []
        meses_com_dados = []
        
        for mes_key, dias_mes in dados_diarios.items():
            # Extrair mês do formato "2025-10"
            try:
                ano, mes_num = mes_key.split('-')
                mes_num = int(mes_num)
                if mes_num >= 1 and mes_num <= 12:
                    mes_nome = meses_ordem[mes_num - 1]
                    
                    # Calcular total de coletas e dias com coleta para este mês
                    total_coletas = sum(int(coletas) for coletas in dias_mes.values())
                    dias_com_coleta = len(dias_mes)
                    
                    # Média diária = total de coletas / dias com coleta (não dias do mês)
                    if dias_com_coleta > 0:
                        media_diaria = total_coletas / dias_com_coleta
                        medias_diarias.append(media_diaria)
                        meses_com_dados.append(mes_nome)
            except (ValueError, IndexError):
                continue
        
        if not medias_diarias:
            st.info("📊 Nenhuma coleta encontrada nos dados diários de 2025.")
            return
        
        # Criar gráfico
        fig = px.bar(
            x=meses_com_dados,
            y=medias_diarias,
            title=f"📊 Média Diária Real por Mês - {nome_exibicao}<br><sup>Baseado em dias com coleta real</sup>",
            color=medias_diarias,
            color_continuous_scale='Greens',
            text=[f"{val:.1f}" for val in medias_diarias]
        )
     
        fig.update_traces(
            texttemplate='%{text} coletas',
            textposition='outside',
            hovertemplate='<b>Mês:</b> %{x}<br><b>Média Diária:</b> %{y:.1f} coletas<br><sup>Baseado em dias com coleta real</sup><extra></extra>'
        )
     
        fig.update_layout(
            xaxis_title="Mês",
            yaxis_title="Média Diária (Coletas)",
            showlegend=False,
            height=600,
            margin=dict(l=60, r=60, t=80, b=80),
            autosize=True,
            font=dict(size=14)
        )
     
        st.plotly_chart(fig, width='stretch')
        
        # Explicação metodológica
        with st.expander("ℹ️ Sobre Esta Análise", expanded=False):
            st.markdown(f"""
            **Como é calculada a média diária real:**
            1. **Base de dados**: Dados reais de coletas de 2025 por dia
            2. **Cálculo**: Total de coletas do mês ÷ dias com coleta (não dias do mês)
            3. **Vantagem**: Mostra a produtividade real nos dias de trabalho
            4. **Exemplo**: Se em Outubro houve 8 coletas em 4 dias diferentes, a média é 2.0 coletas/dia
            
            **💡 Insight**: Esta análise mostra:
            - Produtividade real nos dias de coleta
            - Padrões de intensidade de trabalho
            - Comparação mais precisa entre meses
            """)
    @staticmethod
    def criar_grafico_coletas_por_dia(
        df: pd.DataFrame,
        lab_cnpj: Optional[str] = None,
        lab_nome: Optional[str] = None
    ):
        """Cria gráfico de coletas por dia do mês usando dados reais de 2025."""
        if df.empty:
            st.info("📊 Nenhum dado disponível para o gráfico")
            return
        if not lab_cnpj and not lab_nome:
            st.info("📊 Selecione um laboratório para visualizar as coletas por dia")
            return

        df_ref = df
        if lab_cnpj and 'CNPJ_Normalizado' not in df_ref.columns and 'CNPJ_PCL' in df_ref.columns:
            df_ref = df_ref.copy()
            df_ref['CNPJ_Normalizado'] = df_ref['CNPJ_PCL'].apply(DataManager.normalizar_cnpj)

        if lab_cnpj and 'CNPJ_Normalizado' in df_ref.columns:
            lab_data = df_ref[df_ref['CNPJ_Normalizado'] == lab_cnpj]
        else:
            lab_data = df_ref[df_ref['Nome_Fantasia_PCL'] == lab_nome]

        if lab_data.empty:
            st.info("📊 Laboratório não encontrado")
            return

        lab = lab_data.iloc[0]
        nome_exibicao = lab_nome or lab.get('Nome_Fantasia_PCL') or lab_cnpj

        # Verificar se temos dados diários reais de 2025
        if 'Dados_Diarios_2025' not in lab or pd.isna(lab['Dados_Diarios_2025']) or lab['Dados_Diarios_2025'] == '{}':
            st.info("📊 Nenhum dado diário disponível para 2025 para este laboratório.")
            return

        import json
        try:
            # Carregar dados diários reais
            dados_diarios = json.loads(lab['Dados_Diarios_2025'])
        except (json.JSONDecodeError, TypeError):
            st.info("📊 Erro ao processar dados diários para este laboratório.")
            return

        if not dados_diarios:
            st.info("📊 Nenhum dado diário disponível para 2025.")
            return

        # Converter dados para DataFrame
        dados_grafico = []
        meses_ordem = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']

        for mes_key, dias_mes in dados_diarios.items():
            # Extrair mês do formato "2025-10"
            try:
                ano, mes_num = mes_key.split('-')
                mes_num = int(mes_num)
                if mes_num >= 1 and mes_num <= 12:
                    mes_nome = meses_ordem[mes_num - 1]

                    # Adicionar apenas dias com coletas reais
                    for dia_str, coletas in dias_mes.items():
                        dia = int(dia_str)
                        if coletas > 0:  # Só mostrar dias com coletas
                            dados_grafico.append({
                                'Dia': dia,
                                'Mês': mes_nome,
                                'Coletas': int(coletas)
                            })
            except (ValueError, IndexError):
                continue

        if not dados_grafico:
            st.info("📊 Nenhuma coleta encontrada nos dados diários de 2025.")
            return

        df_grafico = pd.DataFrame(dados_grafico)

        # Criar gráfico de linha interativo
        fig = px.line(
            df_grafico,
            x='Dia',
            y='Coletas',
            color='Mês',
            title=f"📅 Coletas por Dia Útil do Mês - {nome_exibicao}",
            markers=True,
            line_shape='linear'
        )

        # Configurar tooltip personalizado com nome correto do mês
        fig.update_traces(
            hovertemplate='<b>Dia:</b> %{x}<br><b>Mês:</b> %{fullData.name}<br><b>Coletas:</b> %{y:.0f}<extra></extra>'
        )

        fig.update_layout(
            xaxis_title="Dia do Mês (dias úteis disponíveis)",
            yaxis_title="Número de Coletas (dias úteis)",
            xaxis=dict(tickmode='linear', tick0=1, dtick=5),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=-0.15,
                xanchor="center",
                x=0.5,
                bgcolor="rgba(255,255,255,0.8)",
                bordercolor="rgba(0,0,0,0.2)",
                borderwidth=1
            ),
            height=600,
            margin=dict(l=60, r=60, t=80, b=120),  # Margem inferior maior para legenda
            autosize=True,
            font=dict(size=14),
            # Tornar o gráfico mais interativo
            hovermode='x unified',
            # Melhorar a aparência das linhas
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)'
        )

        # Adicionar anotação explicativa (dica persistente)
        fig.add_annotation(
            text="💡 Dica: dê duplo clique no mês na legenda para focar apenas aquela série. Clique simples mostra/oculta linhas.",
            xref="paper", yref="paper",
            x=0.5, y=-0.25,
            showarrow=False,
            font=dict(size=12, color="gray"),
            xanchor="center"
        )

        st.plotly_chart(fig, width='stretch')
    @staticmethod
    def criar_grafico_media_dia_semana_novo(
        df: pd.DataFrame,
        lab_cnpj: Optional[str] = None,
        lab_nome: Optional[str] = None,
        filtros: dict = None
    ):
        """NOVA VERSÃO - Cria gráfico de distribuição de coletas por dia da semana usando dados reais de 2025."""
        if df.empty:
            st.info("📊 Nenhum dado disponível para o gráfico")
            return
        if not lab_cnpj and not lab_nome:
            st.info("📊 Selecione um laboratório para visualizar a distribuição semanal")
            return

        df_ref = df
        if lab_cnpj and 'CNPJ_Normalizado' not in df_ref.columns and 'CNPJ_PCL' in df_ref.columns:
            df_ref = df_ref.copy()
            df_ref['CNPJ_Normalizado'] = df_ref['CNPJ_PCL'].apply(DataManager.normalizar_cnpj)

        if lab_cnpj and 'CNPJ_Normalizado' in df_ref.columns:
            lab_data = df_ref[df_ref['CNPJ_Normalizado'] == lab_cnpj]
        else:
            lab_data = df_ref[df_ref['Nome_Fantasia_PCL'] == lab_nome]

        if lab_data.empty:
            st.info("📊 Laboratório não encontrado")
            return

        lab = lab_data.iloc[0]
        nome_exibicao = lab_nome or lab.get('Nome_Fantasia_PCL') or lab_cnpj
        
        # Verificar se temos dados semanais reais de 2025
        if 'Dados_Semanais_2025' not in lab or pd.isna(lab['Dados_Semanais_2025']) or lab['Dados_Semanais_2025'] == '{}':
            st.info("📊 Nenhum dado semanal disponível para 2025 para este laboratório.")
            return
        
        import json
        try:
            dados_semanais = json.loads(lab['Dados_Semanais_2025'])
        except (json.JSONDecodeError, TypeError):
            st.info("📊 Erro ao processar dados semanais para este laboratório.")
            return
        
        if not dados_semanais:
            st.info("📊 Nenhum dado semanal disponível para 2025.")
            return
        
        # NOVA IMPLEMENTAÇÃO - Criar dados de forma mais simples e direta
        dias_uteis = ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta']
        cores_dias = {
            'Segunda': '#6BBF47', 'Terça': '#ff7f0e', 'Quarta': '#52B54B', 'Quinta': '#d62728',
            'Sexta': '#9467bd'
        }
        
        # Criar lista de dados de forma mais direta
        dados_grafico = []
        total_coletas = 0
        
        for dia in dias_uteis:
            coletas = dados_semanais.get(dia, 0)
            total_coletas += coletas
            dados_grafico.append({
                'dia': dia,
                'coletas': coletas,
                'cor': cores_dias[dia]
            })
        
        max_coletas = max((item['coletas'] for item in dados_grafico), default=0)
        y_axis_max = max_coletas * 1.2 if max_coletas > 0 else 10

        # Calcular percentuais
        for item in dados_grafico:
            if total_coletas > 0:
                item['percentual'] = round((item['coletas'] / total_coletas) * 100, 1)
            else:
                item['percentual'] = 0.0
        
        # CRIAR GRÁFICO NOVO DO ZERO
        import plotly.graph_objects as go
        
        fig = go.Figure()
        
        # Adicionar barras uma por uma para ter controle total
        for i, row in enumerate(dados_grafico):
            fig.add_trace(go.Bar(
                x=[row['dia']],
                y=[row['coletas']],
                name=row['dia'],
                marker_color=row['cor'],
                text=[f"{row['coletas']} coletas<br>({row['percentual']:.1f}%)"],
                textposition='outside',
                hovertemplate=f"<b>{row['dia']}</b><br>" +
                             f"Coletas: {row['coletas']}<br>" +
                             f"Percentual: {row['percentual']:.1f}% da semana<extra></extra>",
                showlegend=False
            ))
        
        # Configurar layout
        fig.update_layout(
            title=f"📅 Distribuição Real de Coletas por Dia Útil da Semana<br><sup>{nome_exibicao} | Total semanal: {total_coletas} coletas úteis</sup>",
            xaxis_title="Dia da Semana (dias úteis)",
            yaxis_title="Coletas por Dia Útil",
            height=600,
            margin=dict(l=60, r=60, t=100, b=80),
            font=dict(size=14),
            title_font_size=18,
            yaxis=dict(range=[0, y_axis_max])
        )
        
        # Adicionar linha de média diária
        if total_coletas > 0:
            media_diaria = total_coletas / len(dias_uteis)
            fig.add_hline(
                y=media_diaria,
                line_dash="dash",
                line_color="red",
                annotation_text=f"Média por dia útil: {media_diaria:.1f} coletas",
                annotation_position="top right"
            )
        
        st.plotly_chart(fig, width='stretch')
        
        # Métricas
        col1, col2, col3 = st.columns(3)
        with col1:
            dia_max = max(dados_grafico, key=lambda x: x['coletas'])
            st.metric("📈 Dia Mais Forte", dia_max['dia'], f"{dia_max['coletas']:.0f} coletas")
        with col2:
            dia_min = min(dados_grafico, key=lambda x: x['coletas'])
            st.metric("📉 Dia Mais Fraco", dia_min['dia'], f"{dia_min['coletas']:.0f} coletas")
        with col3:
            max_coletas = max(item['coletas'] for item in dados_grafico)
            min_coletas = min(item['coletas'] for item in dados_grafico)
            variacao = ((max_coletas - min_coletas) / max_coletas * 100) if max_coletas > 0 else 0
            st.metric("📊 Variação (forte vs fraco)", f"{variacao:.1f}%", "forte vs fraco")
        
        # Debug removido após validação dos percentuais

    @staticmethod
    def criar_grafico_media_dia_semana(df: pd.DataFrame, lab_selecionado: str = None, filtros: dict = None):
        """Cria gráfico de distribuição de coletas por dia da semana usando dados reais de 2025."""
        if df.empty:
            st.info("📊 Nenhum dado disponível para o gráfico")
            return
        if not lab_selecionado:
            st.info("📊 Selecione um laboratório para visualizar a distribuição semanal")
            return
        lab_data = df[df['Nome_Fantasia_PCL'] == lab_selecionado]
        if not lab_data.empty:
            lab = lab_data.iloc[0]
            
            # Verificar se temos dados semanais reais de 2025
            if 'Dados_Semanais_2025' not in lab or pd.isna(lab['Dados_Semanais_2025']) or lab['Dados_Semanais_2025'] == '{}':
                st.info("📊 Nenhum dado semanal disponível para 2025 para este laboratório.")
                return
            
            import json
            try:
                # Carregar dados semanais reais
                dados_semanais = json.loads(lab['Dados_Semanais_2025'])
            except (json.JSONDecodeError, TypeError):
                st.info("📊 Erro ao processar dados semanais para este laboratório.")
                return
            
            if not dados_semanais:
                st.info("📊 Nenhum dado semanal disponível para 2025.")
                return
            
            # Converter dados para DataFrame
            dias_uteis = ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta']
            cores_dias = {
                'Segunda': '#6BBF47', # Verde Synvia
                'Terça': '#ff7f0e', # Laranja
                'Quarta': '#52B54B', # Verde Synvia Escuro
                'Quinta': '#d62728', # Vermelho
                'Sexta': '#9467bd' # Roxo
            }
            
            dados_semana = []
            total_coletas_semana = 0
            
            for dia in dias_uteis:
                coletas_dia = dados_semanais.get(dia, 0)
                total_coletas_semana += coletas_dia
                dados_semana.append({
                    'Dia_Semana': dia,
                    'Coletas_Reais': coletas_dia,
                    'Cor': cores_dias[dia]
                })
            
            df_semana = pd.DataFrame(dados_semana)
            
            # Calcular percentuais corretos baseados nos dados reais
            if total_coletas_semana > 0:
                df_semana['Percentual'] = (df_semana['Coletas_Reais'] / total_coletas_semana * 100).round(1)
            else:
                df_semana['Percentual'] = 0.0
            # Criar título informativo
            periodo_texto = "dados reais de 2025"
            
            # Calcular média diária correta (soma das coletas semanais / 5 dias úteis)
            media_diaria = total_coletas_semana / len(dias_uteis) if total_coletas_semana > 0 else 0
            
            # Gráfico de barras
            max_coletas_semana = df_semana['Coletas_Reais'].max() if not df_semana.empty else 0
            y_axis_max = max_coletas_semana * 1.2 if max_coletas_semana > 0 else 10
            fig = px.bar(
                df_semana,
                x='Dia_Semana',
                y='Coletas_Reais',
                title=f"📅 Distribuição Real de Coletas por Dia Útil da Semana<br><sup>{lab_selecionado} | Baseado em: {periodo_texto} | Total semanal: {total_coletas_semana:.0f} coletas úteis</sup>",
                color='Dia_Semana',
                color_discrete_map=cores_dias,
                text='Coletas_Reais'
            )
            # Usar hovertemplate com cálculo direto do percentual
            fig.update_traces(
                texttemplate='%{text:.0f} coletas<br>(%{customdata:.1f}%)',
                textposition='outside',
                customdata=df_semana['Percentual'],
                hovertemplate='<b>%{x}</b><br>Coletas: %{y:.0f}<br>Percentual: %{customdata:.1f}% da semana<extra></extra>'
            )
            fig.update_layout(
                xaxis_title="Dia da Semana (dias úteis)",
                yaxis_title="Coletas por Dia Útil",
                showlegend=False,
                coloraxis_showscale=False,
                height=700,  # Aumentado significativamente para destaque
                margin=dict(l=60, r=60, t=100, b=80),  # Margens aumentadas
                autosize=True,  # Responsivo
                font=dict(size=14),  # Fonte maior para melhor legibilidade
                title_font_size=18,  # Título maior
                yaxis=dict(range=[0, y_axis_max])
            )
            # Adicionar linha de referência da média diária
            if media_diaria > 0:
                fig.add_hline(
                    y=media_diaria,
                    line_dash="dash",
                    line_color="red",
                    annotation_text=f"Média por dia útil: {media_diaria:.1f} coletas",
                    annotation_position="top right"
                )
            st.plotly_chart(fig, width='stretch')
            # Métricas adicionais
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric(
                    "📈 Dia Mais Forte",
                    df_semana.loc[df_semana['Coletas_Reais'].idxmax(), 'Dia_Semana'],
                    f"{df_semana['Coletas_Reais'].max():.0f} coletas"
                )
            with col2:
                st.metric(
                    "📉 Dia Mais Fraco",
                    df_semana.loc[df_semana['Coletas_Reais'].idxmin(), 'Dia_Semana'],
                    f"{df_semana['Coletas_Reais'].min():.0f} coletas"
                )
            with col3:
                variacao_semanal = (df_semana['Coletas_Reais'].max() - df_semana['Coletas_Reais'].min()) / df_semana['Coletas_Reais'].max() * 100 if df_semana['Coletas_Reais'].max() > 0 else 0
                st.metric(
                    "📊 Variação (forte vs fraco)",
                    f"{variacao_semanal:.1f}%",
                    "forte vs fraco"
                )
            # Explicação metodológica
            with st.expander("ℹ️ Sobre Esta Análise", expanded=False):
                st.markdown(f"""
                **Como é calculada a distribuição semanal:**
                1. **Base de dados**: Dados reais de coletas de 2025 ({periodo_texto})
                2. **Distribuição real**: Baseada nas datas exatas das coletas (createdAt)
                   - **Total semanal**: {total_coletas_semana:.0f} coletas úteis
                   - **Percentuais**: Calculados considerando apenas dias úteis
                3. **Média diária útil**: {media_diaria:.1f} coletas (total semanal ÷ {len(dias_uteis)})
                **💡 Insight**: Esta análise mostra:
                - Padrões reais de coleta do laboratório
                - Dias com maior/menor movimento baseado em dados históricos
                - Oportunidades de otimização de recursos
                **⚠️ Importante**: Estes são valores estimados baseados em padrões históricos.
                Dados diários reais forneceriam análise mais precisa.
                """)
    @staticmethod
    def criar_grafico_controle_br_uf_cidade(
        df: pd.DataFrame,
        df_filtrado: Optional[pd.DataFrame] = None,
        lab_cnpj: Optional[str] = None,
        lab_nome: Optional[str] = None,
        usar_mm30: bool = False
    ):
        """
        Cria gráfico comparativo de controle BR × UF × Cidade × Lab atual.
        Mostra séries temporais de MM7 ou MM30 (dias úteis) para cada contexto.
        
        Args:
            df: DataFrame completo (para calcular contextos BR/UF/Cidade)
            df_filtrado: DataFrame filtrado (para série atual quando não há lab específico)
            lab_cnpj: CNPJ do laboratório específico (opcional)
            lab_nome: Nome do laboratório específico (opcional)
            usar_mm30: Se True, usa MM30; se False, usa MM7
        """
        if df.empty:
            st.info("📊 Nenhum dado disponível para o gráfico de controle")
            return
        
        import json
        from pandas.tseries.offsets import BDay
        
        # Determinar contexto atual (lab específico ou conjunto filtrado)
        serie_atual = pd.Series(dtype="float")
        nome_serie_atual = "Conjunto Filtrado"
        lab_data = pd.DataFrame()
        
        if lab_cnpj or lab_nome:
            # Buscar lab específico
            df_ref = df.copy()
            if lab_cnpj and 'CNPJ_Normalizado' not in df_ref.columns and 'CNPJ_PCL' in df_ref.columns:
                df_ref['CNPJ_Normalizado'] = df_ref['CNPJ_PCL'].apply(DataManager.normalizar_cnpj)
            
            if lab_cnpj and 'CNPJ_Normalizado' in df_ref.columns:
                lab_data = df_ref[df_ref['CNPJ_Normalizado'] == lab_cnpj]
            else:
                lab_data = df_ref[df_ref['Nome_Fantasia_PCL'] == lab_nome]
            
            if not lab_data.empty:
                lab = lab_data.iloc[0]
                nome_serie_atual = lab.get('Nome_Fantasia_PCL', lab_cnpj or lab_nome)
                if 'Dados_Diarios_2025' in lab and pd.notna(lab['Dados_Diarios_2025']):
                    serie_atual = RiskEngine._serie_diaria_from_json(lab['Dados_Diarios_2025'])
        else:
            # Agregar série do conjunto filtrado (usar df_filtrado se disponível, senão df)
            df_para_serie = df_filtrado if df_filtrado is not None and not df_filtrado.empty else df
            todas_series = []
            for _, row in df_para_serie.iterrows():
                if 'Dados_Diarios_2025' in row and pd.notna(row['Dados_Diarios_2025']):
                    s = RiskEngine._serie_diaria_from_json(row['Dados_Diarios_2025'])
                    if not s.empty:
                        todas_series.append(s)
            
            if todas_series:
                # Agregar todas as séries por data
                todas_datas = set()
                for s in todas_series:
                    todas_datas.update(s.index)
                
                serie_agregada = pd.Series(index=sorted(todas_datas), dtype="float")
                for data in serie_agregada.index:
                    total = sum(s.get(data, 0) for s in todas_series)
                    serie_agregada[data] = total
                
                serie_atual = serie_agregada
        
        # Agregar séries por contexto (BR, UF, Cidade)
        def agregar_por_contexto(df_contexto: pd.DataFrame) -> pd.Series:
            """Agrega coletas diárias por contexto."""
            todas_series = []
            for _, row in df_contexto.iterrows():
                if 'Dados_Diarios_2025' in row and pd.notna(row['Dados_Diarios_2025']):
                    s = RiskEngine._serie_diaria_from_json(row['Dados_Diarios_2025'])
                    if not s.empty:
                        todas_series.append(s)
            
            if not todas_series:
                return pd.Series(dtype="float")
            
            # Agregar todas as séries por data
            todas_datas = set()
            for s in todas_series:
                todas_datas.update(s.index)
            
            serie_agregada = pd.Series(index=sorted(todas_datas), dtype="float")
            for data in serie_agregada.index:
                total = sum(s.get(data, 0) for s in todas_series)
                serie_agregada[data] = total
            
            return serie_agregada
        
        # Agregar por BR (todos os labs do DataFrame completo)
        serie_br = agregar_por_contexto(df)
        
        # Agregar por UF (se temos info de UF)
        serie_uf = pd.Series(dtype="float")
        uf_nome = ""
        if lab_cnpj or lab_nome:
            if not lab_data.empty and 'Estado' in lab_data.columns:
                uf_nome = lab_data.iloc[0]['Estado']
                if pd.notna(uf_nome) and uf_nome:
                    df_uf = df[df['Estado'] == uf_nome]
                    serie_uf = agregar_por_contexto(df_uf)
        else:
            # Para conjunto filtrado, usar UF do primeiro lab (se disponível)
            if not df.empty and 'Estado' in df.columns:
                uf_nome = df.iloc[0]['Estado']
                if pd.notna(uf_nome) and uf_nome:
                    df_uf = df[df['Estado'] == uf_nome]
                    serie_uf = agregar_por_contexto(df_uf)
        
        # Agregar por Cidade (se temos info de Cidade)
        serie_cidade = pd.Series(dtype="float")
        cidade_nome = ""
        if lab_cnpj or lab_nome:
            if not lab_data.empty and 'Cidade' in lab_data.columns:
                cidade_nome = lab_data.iloc[0]['Cidade']
                if pd.notna(cidade_nome) and cidade_nome:
                    df_cidade = df[df['Cidade'] == cidade_nome]
                    serie_cidade = agregar_por_contexto(df_cidade)
        else:
            # Para conjunto filtrado, usar Cidade do primeiro lab (se disponível)
            if not df.empty and 'Cidade' in df.columns:
                cidade_nome = df.iloc[0]['Cidade']
                if pd.notna(cidade_nome) and cidade_nome:
                    df_cidade = df[df['Cidade'] == cidade_nome]
                    serie_cidade = agregar_por_contexto(df_cidade)
        
        # Calcular médias móveis ao longo do tempo (apenas dias úteis)
        def calcular_mm_serie(serie: pd.Series, janela: int) -> pd.Series:
            """Calcula média móvel de janela dias úteis ao longo da série."""
            if serie.empty:
                return pd.Series(dtype="float")
            
            # Garantir que temos apenas dias úteis
            serie = serie.sort_index()
            serie_uteis = serie[serie.index.weekday < 5]  # Segunda=0 a Sexta=4
            
            if serie_uteis.empty:
                return pd.Series(dtype="float")
            
            # Calcular MM ao longo do tempo
            mm_serie = serie_uteis.rolling(window=janela, min_periods=1).mean()
            return mm_serie
        
        janela = 30 if usar_mm30 else 7
        mm_label = "MM30" if usar_mm30 else "MM7"
        
        # Calcular MMs para cada contexto
        mm_br = calcular_mm_serie(serie_br, janela)
        mm_uf = calcular_mm_serie(serie_uf, janela)
        mm_cidade = calcular_mm_serie(serie_cidade, janela)
        mm_atual = calcular_mm_serie(serie_atual, janela)
        
        # Preparar dados para o gráfico
        dados_grafico = []
        
        # Adicionar série BR
        for data, valor in mm_br.items():
            dados_grafico.append({
                'Data': data,
                'Valor': valor,
                'Serie': '🇧🇷 MM7_BR' if not usar_mm30 else '🇧🇷 MM30_BR'
            })
        
        # Adicionar série UF
        if not mm_uf.empty and uf_nome:
            uf_label = f"📍 MM7_UF ({uf_nome})" if not usar_mm30 else f"📍 MM30_UF ({uf_nome})"
            for data, valor in mm_uf.items():
                dados_grafico.append({
                    'Data': data,
                    'Valor': valor,
                    'Serie': uf_label
                })
        
        # Adicionar série Cidade
        if not mm_cidade.empty and cidade_nome:
            cidade_label = f"🏙️ MM7_CIDADE ({cidade_nome})" if not usar_mm30 else f"🏙️ MM30_CIDADE ({cidade_nome})"
            for data, valor in mm_cidade.items():
                dados_grafico.append({
                    'Data': data,
                    'Valor': valor,
                    'Serie': cidade_label
                })
        
        # Adicionar série atual
        if not mm_atual.empty:
            atual_label = f"📊 {nome_serie_atual} ({mm_label})"
            for data, valor in mm_atual.items():
                dados_grafico.append({
                    'Data': data,
                    'Valor': valor,
                    'Serie': atual_label
                })
        
        if not dados_grafico:
            st.info("📊 Nenhum dado disponível para gerar o gráfico de controle")
            return
        
        df_grafico = pd.DataFrame(dados_grafico)
        
        # Criar gráfico de linha
        cores_map = {
            '🇧🇷 MM7_BR': '#DC2626',
            '🇧🇷 MM30_BR': '#DC2626',
            '📍 MM7_UF': '#3B82F6',
            '📍 MM30_UF': '#3B82F6',
            '🏙️ MM7_CIDADE': '#10B981',
            '🏙️ MM30_CIDADE': '#10B981'
        }
        
        # Adicionar cores dinâmicas para série atual
        for serie_nome in df_grafico['Serie'].unique():
            if serie_nome not in cores_map:
                cores_map[serie_nome] = '#6BBF47'  # Cor padrão verde
        
        fig = px.line(
            df_grafico,
            x='Data',
            y='Valor',
            color='Serie',
            title=f"📊 Controle BR × UF × Cidade × {nome_serie_atual}<br><sup>{mm_label} - Apenas dias úteis</sup>",
            markers=True,
            line_shape='linear',
            color_discrete_map=cores_map
        )
        
        fig.update_traces(
            hovertemplate='<b>%{fullData.name}</b><br>Data: %{x|%d/%m/%Y}<br>Valor: %{y:.2f}<extra></extra>',
            line=dict(width=2.5)
        )
        
        fig.update_layout(
            xaxis_title="Data (dias úteis)",
            yaxis_title=f"Média Móvel ({mm_label})",
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=-0.2,
                xanchor="center",
                x=0.5,
                bgcolor="rgba(255,255,255,0.9)",
                bordercolor="rgba(0,0,0,0.2)",
                borderwidth=1
            ),
            height=600,
            margin=dict(l=60, r=60, t=100, b=120),
            hovermode='x unified',
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(size=12)
        )
        
        st.plotly_chart(fig, width='stretch')
    
    @staticmethod
    def criar_grafico_evolucao_mensal(
        df: pd.DataFrame,
        lab_cnpj: Optional[str] = None,
        lab_nome: Optional[str] = None,
        chart_key: str = "default"
    ):
        """Cria gráfico de evolução mensal - Atualizado com correções de diferença 2024/2025."""
        if df.empty:
            st.info("📊 Nenhum dado disponível para o gráfico")
            return
        meses = ChartManager._meses_ate_hoje(df, 2025)
        if not meses:
            st.info("📊 Nenhum mês disponível até a data atual")
            return
        colunas_meses = [f'N_Coletas_{mes}_25' for mes in meses]
        if lab_cnpj or lab_nome:
            # Gráfico para laboratório específico
            df_ref = df
            if lab_cnpj and 'CNPJ_Normalizado' not in df_ref.columns and 'CNPJ_PCL' in df_ref.columns:
                df_ref = df_ref.copy()
                df_ref['CNPJ_Normalizado'] = df_ref['CNPJ_PCL'].apply(DataManager.normalizar_cnpj)

            if lab_cnpj and 'CNPJ_Normalizado' in df_ref.columns:
                lab_data = df_ref[df_ref['CNPJ_Normalizado'] == lab_cnpj]
            else:
                lab_data = df_ref[df_ref['Nome_Fantasia_PCL'] == lab_nome]
            if not lab_data.empty:
                lab = lab_data.iloc[0]
                nome_exibicao = lab_nome or lab.get('Nome_Fantasia_PCL') or lab_cnpj
                valores_2025 = [lab.get(col, 0) for col in colunas_meses]
             
                # Dados 2024 (mesmos meses para comparação direta)
                colunas_2024 = [f'N_Coletas_{mes}_24' for mes in meses]
                valores_2024 = [lab.get(col, 0) for col in colunas_2024]
             
                # Calcular médias - Corrigido agrupamento temporal
                media_2025 = sum(valores_2025) / len(valores_2025) if valores_2025 else 0
                media_2024 = sum(valores_2024) / len(valores_2024) if valores_2024 else 0
             
                # Criar DataFrame para o gráfico
                df_grafico = pd.DataFrame({
                    'Mês': meses,
                    '2025': valores_2025,
                    '2024': valores_2024,
                    'Média 2025': [media_2025] * len(meses),
                    'Média 2024': [media_2024] * len(meses)
                })
             
                # Criar gráfico com múltiplas linhas
                fig = px.line(
                    df_grafico,
                    x='Mês',
                    y=['2025', '2024', 'Média 2025', 'Média 2024'],
                    title=f"📈 Evolução Mensal - {nome_exibicao}",
                    markers=True,
                    line_shape='spline'
                )
                
                # Personalizar cores e estilos
                fig.update_traces(
                    mode='lines+markers',
                    hovertemplate='<b>Mês:</b> %{x}<br><b>Coletas:</b> %{y}<extra></extra>'
                )
                
                # Garantir que não há valores negativos exibidos (ajustar eixo Y para começar em 0 ou acima)
                fig.update_layout(
                    yaxis=dict(
                        rangemode='tozero',  # Garante que o eixo Y começa em 0 ou acima
                        showgrid=True,
                        gridcolor='rgba(128, 128, 128, 0.2)'
                    )
                )
             
                # Cores personalizadas
                fig.data[0].line.color = '#6BBF47' # Verde Synvia para 2025
                fig.data[1].line.color = '#ff7f0e' # Laranja para 2024
                fig.data[2].line.color = '#6BBF47' # Verde Synvia para média 2025
                fig.data[2].line.dash = 'dash'
                fig.data[3].line.color = '#ff7f0e' # Laranja para média 2024
                fig.data[3].line.dash = 'dash'
                # Ajustar textos de hover para diferenciar coletas x médias
                fig.data[0].hovertemplate = '<b>Mês:</b> %{x}<br><b>Coletas 2025:</b> %{y:.0f}<extra></extra>'
                fig.data[1].hovertemplate = '<b>Mês:</b> %{x}<br><b>Coletas 2024:</b> %{y:.0f}<extra></extra>'
                fig.data[2].hovertemplate = '<b>Mês:</b> %{x}<br><b>Média 2025:</b> %{y:.1f}<extra></extra>'
                fig.data[3].hovertemplate = '<b>Mês:</b> %{x}<br><b>Média 2024:</b> %{y:.1f}<extra></extra>'
                fig.update_layout(
                    xaxis_title="Mês",
                    yaxis_title="Número de Coletas",
                    hovermode='x unified',
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",
                        y=-0.15,
                        xanchor="center",
                        x=0.5
                    ),
                    height=600,  # Aumentado conforme solicitado
                    margin=dict(l=60, r=60, t=60, b=80),  # Margens aumentadas para evitar cortes
                    autosize=True,  # Responsivo
                    showlegend=True
                )
                st.plotly_chart(fig, width='stretch', key=f"evolucao_mensal_lab_{chart_key}")
        else:
            # Gráfico agregado
            valores_agregados = [df[col].sum() for col in colunas_meses]
            fig = px.line(
                x=meses,
                y=valores_agregados,
                title="📈 Evolução Mensal Agregada (2025)",
                markers=True,
                line_shape='spline'
            )
            fig.update_traces(
                mode='lines+markers+text',
                text=valores_agregados,
                textposition="top center",
                hovertemplate='<b>Mês:</b> %{x}<br><b>Total Coletas:</b> %{y}<extra></extra>'
            )
            fig.update_layout(
                xaxis_title="Mês",
                yaxis_title="Total de Coletas",
                hovermode='x unified',
                height=600,  # Aumentado conforme solicitado
                margin=dict(l=60, r=60, t=60, b=80),  # Margens aumentadas
                autosize=True  # Responsivo
            )
            st.plotly_chart(fig, width='stretch', key=f"evolucao_mensal_agregado_{chart_key}")
class UIManager:
    """Gerenciador da interface do usuário - Atualizado com tabs."""
    @staticmethod
    def renderizar_header():
        """Renderiza o cabeçalho principal."""
        st.markdown("""
        <div class="main-header">
            <h1>📊 Syntox Churn</h1>
            <p>Dashboard profissional para análise de retenção de laboratórios</p>
        </div>
        """, unsafe_allow_html=True)
    @staticmethod
    def renderizar_kpi_cards(metrics: KPIMetrics):
        """Renderiza cards de KPIs modernos V2 - Otimizado para eliminar redundâncias."""
        # Primeira linha: Métricas principais
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            # Card 1: Labs monitorados com breakdown de risco
            risco_breakdown = []
            if metrics.labs_moderado_count > 0:
                risco_breakdown.append(f"🟠 {metrics.labs_moderado_count:,}")
            if metrics.labs_alto_count > 0:
                risco_breakdown.append(f"🔴 {metrics.labs_alto_count:,}")
            if metrics.labs_critico_count > 0:
                risco_breakdown.append(f"⚫ {metrics.labs_critico_count:,}")
            
            risco_text = " | ".join(risco_breakdown) if risco_breakdown else "Nenhum em risco"
            recuperacao_text = f"🔄 Recuperação: {metrics.labs_recuperando:,}" if metrics.labs_recuperando > 0 else ""
            delta_text = f"{risco_text}" + (f" | {recuperacao_text}" if recuperacao_text else "")
            
            st.markdown(f"""
            <div class="metric-card" title="Total de laboratórios ativos nos últimos 90 dias. Breakdown por nível de risco (🟠 Moderado, 🔴 Alto, ⚫ Crítico). Recuperação: laboratórios que voltaram a operar acima da MM7 após período de queda.">
                <div class="metric-value">{metrics.total_labs:,}</div>
                <div class="metric-label">Labs Monitorados (≤90 dias)</div>
                <div class="metric-delta" style="font-size:0.85rem;">{delta_text}</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            # Card 2: Coletas do dia
            delta_text = f"D-1: {metrics.vol_d1_total:,} | YTD: {metrics.total_coletas:,}"
            st.markdown(f"""
            <div class="metric-card" title="Total de coletas registradas no último dia útil. D-1: volume de coletas do dia útil anterior. YTD (Year To Date): soma total de coletas em 2025 até o momento.">
                <div class="metric-value">{metrics.vol_hoje_total:,}</div>
                <div class="metric-label">Coletas Hoje</div>
                <div class="metric-delta">{delta_text}</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            # Card 3: Risco crítico consolidado (sem redundância)
            risco_alto_critico = metrics.labs_alto_count + metrics.labs_critico_count
            if risco_alto_critico > 0:
                if metrics.labs_alto_count > 0 and metrics.labs_critico_count > 0:
                    delta_text = f"🔴 {metrics.labs_alto_count:,} | ⚫ {metrics.labs_critico_count:,}"
                elif metrics.labs_critico_count > 0:
                    delta_text = f"⚫ {metrics.labs_critico_count:,} críticos"
                else:
                    delta_text = f"🔴 {metrics.labs_alto_count:,} alto"
            else:
                delta_text = "✅ Nenhum"
            
            st.markdown(f"""
            <div class="metric-card" title="Laboratórios em risco alto (🔴) ou crítico (⚫) pela régua baseada em dias úteis e reduções vs. MM7_BR/MM7_UF/MM7_CIDADE.">
                <div class="metric-value">{risco_alto_critico:,}</div>
                <div class="metric-label">Risco Alto + Crítico</div>
                <div class="metric-delta">{delta_text}</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col4:
            # Card 4: Sem coleta 48h com ativos 7D
            delta_class = "positive" if metrics.ativos_7d >= 80 else "negative"
            ativos_label = f"Ativos 7D: {metrics.ativos_7d:.1f}% ({metrics.ativos_7d_count}/{metrics.total_labs})" if metrics.total_labs else "Ativos 7D: --"
            st.markdown(f"""
            <div class="metric-card" title="Laboratórios com dois dias úteis consecutivos sem registrar coletas (Vol_Hoje = 0 e Vol_D1 = 0). Ativos 7D: percentual de laboratórios com pelo menos uma coleta nos últimos 7 dias úteis.">
                <div class="metric-value">{metrics.labs_sem_coleta_48h:,}</div>
                <div class="metric-label">Sem Coleta (48h)</div>
                <div class="metric-delta {delta_class}">{ativos_label}</div>
            </div>
            """, unsafe_allow_html=True)
        
        # Segunda linha: Distribuição de risco e comparação MM7
        col5, col6 = st.columns([1.5, 1])
        
        with col5:
            # Card 5: Distribuição completa de risco (consolidado)
            total_risco = metrics.labs_moderado_count + metrics.labs_alto_count + metrics.labs_critico_count
            risco_dist_text = f"🟢 {metrics.labs_normal_count:,} | 🟡 {metrics.labs_atencao_count:,} | 🟠 {metrics.labs_moderado_count:,} | 🔴 {metrics.labs_alto_count:,} | ⚫ {metrics.labs_critico_count:,}"
            st.markdown(f"""
            <div class="metric-card" title="Distribuição completa de risco diário calculada sobre dias úteis e reduções máximas vs. MM7 dos contextos BR/UF/Cidade.">
                <div class="metric-value" style="font-size:1.3rem; margin-bottom:0.3rem;">Distribuição de Risco</div>
                <div class="metric-delta" style="display:flex; flex-wrap:wrap; gap:0.4rem; font-weight:600; font-size:0.9rem; justify-content:center;">
                    <span>🟢 {metrics.labs_normal_count:,}</span>
                    <span>🟡 {metrics.labs_atencao_count:,}</span>
                    <span>🟠 {metrics.labs_moderado_count:,}</span>
                    <span>🔴 {metrics.labs_alto_count:,}</span>
                    <span>⚫ {metrics.labs_critico_count:,}</span>
                </div>
                <div class="metric-label" style="margin-top:0.5rem;">Régua de risco (dias úteis)</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col6:
            # Card 6: Comparação MM7_BR vs MM7_UF (consolidado)
            mm7_br_pct = f"{metrics.labs_abaixo_mm7_br_pct:.1f}%" if metrics.total_labs else "--"
            mm7_uf_pct = f"{metrics.labs_abaixo_mm7_uf_pct:.1f}%" if metrics.total_labs else "--"
            delta_text = f"BR: {mm7_br_pct} | UF: {mm7_uf_pct}"
            
            st.markdown(f"""
            <div class="metric-card" title="Laboratórios abaixo da média móvel: MM7_BR (nacional) e MM7_UF (estadual), ambas construídas apenas com dias úteis.">
                <div class="metric-value" style="font-size:1.3rem; margin-bottom:0.3rem;">Abaixo da MM7</div>
                <div class="metric-delta" style="font-size:0.95rem; font-weight:600; margin:0.5rem 0;">
                    <div>🇧🇷 BR: {metrics.labs_abaixo_mm7_br:,} ({mm7_br_pct})</div>
                    <div>📍 UF: {metrics.labs_abaixo_mm7_uf:,} ({mm7_uf_pct})</div>
                </div>
                <div class="metric-label">Comparação Nacional vs Estadual</div>
            </div>
            """, unsafe_allow_html=True)
class MetricasAvancadas:
    """Classe para métricas avançadas de laboratórios - Atualizado organização e comparativos."""
 
    @staticmethod
    def calcular_metricas_lab(
        df: pd.DataFrame,
        lab_cnpj: Optional[str] = None,
        lab_nome: Optional[str] = None
    ) -> dict:
        """Calcula métricas avançadas para um laboratório específico - Atualizado score."""

        df_ref = df
        if lab_cnpj and 'CNPJ_Normalizado' not in df_ref.columns and 'CNPJ_PCL' in df_ref.columns:
            df_ref = df_ref.copy()
            df_ref['CNPJ_Normalizado'] = df_ref['CNPJ_PCL'].apply(DataManager.normalizar_cnpj)

        lab_data = pd.DataFrame()
        if lab_cnpj and 'CNPJ_Normalizado' in df_ref.columns:
            lab_data = df_ref[df_ref['CNPJ_Normalizado'] == lab_cnpj]
        if lab_data.empty and lab_nome:
            lab_data = df_ref[df_ref['Nome_Fantasia_PCL'] == lab_nome]

        if lab_data.empty:
            return {}

        lab = lab_data.iloc[0]
     
        # Total de coletas 2025 (até o mês atual)
        meses_2025 = ChartManager._meses_ate_hoje(df, 2025)
        colunas_2025 = [f'N_Coletas_{mes}_25' for mes in meses_2025]
        total_coletas_2025 = sum(lab.get(col, 0) for col in colunas_2025)
     
        # Média dos últimos 3 meses (dinâmico)
        if len(meses_2025) >= 3:
            ultimos_3_meses = meses_2025[-3:]
        else:
            ultimos_3_meses = meses_2025
        colunas_3_meses = [f'N_Coletas_{mes}_25' for mes in ultimos_3_meses]
        media_3_meses = sum(lab.get(col, 0) for col in colunas_3_meses) / len(colunas_3_meses) if colunas_3_meses else 0
     
        # Média diária (últimos 3 meses)
        dias_3_meses = 90 # Aproximadamente 3 meses
        media_diaria = media_3_meses / 30 if media_3_meses > 0 else 0
     
        # Agudo (7 dias) - coletas nos últimos 7 dias
        dias_sem_coleta = lab.get('Dias_Sem_Coleta', 0)
        agudo = "Ativo" if dias_sem_coleta <= 7 else "Inativo"
     
        # Crônico (fechamentos mensais) - baseado na variação
        variacao = lab.get('Variacao_Percentual', 0)
        if variacao > 20:
            cronico = "Crescimento"
        elif variacao < -20:
            cronico = "Declínio"
        else:
            cronico = "Estável"

        vol_hoje = lab.get('Vol_Hoje', 0)
        vol_hoje = int(vol_hoje) if pd.notna(vol_hoje) else 0
        vol_d1 = lab.get('Vol_D1', 0)
        vol_d1 = int(vol_d1) if pd.notna(vol_d1) else 0
        delta_mm7_val = lab.get('Delta_MM7', None)
        delta_mm7 = round(float(delta_mm7_val), 1) if pd.notna(delta_mm7_val) else None
        delta_d1_val = lab.get('Delta_D1', None)
        delta_d1 = round(float(delta_d1_val), 1) if pd.notna(delta_d1_val) else None
        mm7_val = lab.get('MM7', None)
        mm30_val = lab.get('MM30', None)
        delta_mm30_val = lab.get('Delta_MM30', None)
        mm7 = round(float(mm7_val), 1) if pd.notna(mm7_val) else None
        mm30 = round(float(mm30_val), 1) if pd.notna(mm30_val) else None
        delta_mm30 = round(float(delta_mm30_val), 1) if pd.notna(delta_mm30_val) else None
        risco_diario = lab.get('Risco_Diario', 'N/A')
        if pd.isna(risco_diario):
            risco_diario = 'N/A'
     
        return {
            'total_coletas': int(total_coletas_2025),
            'media_3_meses': round(media_3_meses, 1),
            'media_diaria': round(media_diaria, 1),
            'vol_hoje': vol_hoje,
            'vol_d1': vol_d1,
            'delta_mm7': delta_mm7,
            'delta_d1': delta_d1,
            'mm7': mm7,
            'mm30': mm30,
            'delta_mm30': delta_mm30,
            'agudo': agudo,
            'cronico': cronico,
            'dias_sem_coleta': int(dias_sem_coleta),
            'variacao_percentual': round(variacao, 1),
            'risco_diario': risco_diario
        }
    @staticmethod
    def calcular_metricas_evolucao(
        df: pd.DataFrame,
        lab_cnpj: Optional[str] = None,
        lab_nome: Optional[str] = None
    ) -> dict:
        """Calcula métricas de evolução e comparativos para um laboratório específico - Atualizado organização e comparativo."""

        df_ref = df
        if lab_cnpj and 'CNPJ_Normalizado' not in df_ref.columns and 'CNPJ_PCL' in df_ref.columns:
            df_ref = df_ref.copy()
            df_ref['CNPJ_Normalizado'] = df_ref['CNPJ_PCL'].apply(DataManager.normalizar_cnpj)

        lab_data = pd.DataFrame()
        if lab_cnpj and 'CNPJ_Normalizado' in df_ref.columns:
            lab_data = df_ref[df_ref['CNPJ_Normalizado'] == lab_cnpj]
        if lab_data.empty and lab_nome:
            lab_data = df_ref[df_ref['Nome_Fantasia_PCL'] == lab_nome]

        if lab_data.empty:
            return {}

        lab = lab_data.iloc[0]
        # Total de coletas 2024 (todos os meses disponíveis)
        meses_2024 = ChartManager._meses_ate_hoje(df, 2024)
        colunas_2024 = [f'N_Coletas_{mes}_24' for mes in meses_2024]
        total_coletas_2024 = sum(lab.get(col, 0) for col in colunas_2024)
        # Total de coletas 2025 (até o mês atual)
        meses_2025 = ChartManager._meses_ate_hoje(df, 2025)
        colunas_2025 = [f'N_Coletas_{mes}_25' for mes in meses_2025]
        total_coletas_2025 = sum(lab.get(col, 0) for col in colunas_2025)
        # Média de 2024
        media_2024 = total_coletas_2024 / len(colunas_2024) if colunas_2024 else 0
        # Média de 2025
        media_2025 = total_coletas_2025 / len(colunas_2025) if colunas_2025 else 0
        # Último mês fechado (usando lógica de dia de corte)
        info_ultimo_mes = ChartManager._obter_ultimo_mes_fechado(df)
        media_ultimo_mes = lab.get(info_ultimo_mes['coluna'], 0) if info_ultimo_mes['coluna'] else 0
        
        # Máxima histórica 2024
        max_2024 = max(lab.get(col, 0) for col in colunas_2024) if colunas_2024 else 0
        max_2024_col = max(colunas_2024, key=lambda c: lab.get(c, 0), default=None)
        # Máxima histórica 2025
        max_2025 = max(lab.get(col, 0) for col in colunas_2025) if colunas_2025 else 0
        max_2025_col = max(colunas_2025, key=lambda c: lab.get(c, 0), default=None)

        meses_nomes_completos = {
            "Jan": "Jan", "Fev": "Fev", "Mar": "Mar", "Abr": "Abr",
            "Mai": "Mai", "Jun": "Jun", "Jul": "Jul", "Ago": "Ago",
            "Set": "Set", "Out": "Out", "Nov": "Nov", "Dez": "Dez"
        }

        def _formatar_mes_col(col_name: Optional[str]) -> str:
            if not col_name or '_' not in col_name:
                return "N/A"
            partes = col_name.split('_')
            if len(partes) < 4:
                return "N/A"
            mes = meses_nomes_completos.get(partes[2], partes[2])
            ano = "20" + partes[3][-2:] if len(partes[3]) <= 2 else partes[3]
            return f"{mes}/{ano}"

        max_2024_mes = _formatar_mes_col(max_2024_col)
        max_2025_mes = _formatar_mes_col(max_2025_col)
        
        # Determinar métricas de comparação baseado no ano do último mês
        ano_comparacao = info_ultimo_mes['ano_comparacao']
        if ano_comparacao == 2024:
            # Se o último mês fechado é de 2024, comparar com métricas de 2024
            media_comparacao = media_2024
            max_comparacao = max_2024
        else:
            # Se é 2025, usar métricas de 2025
            media_comparacao = media_2025
            max_comparacao = max_2025
        
        return {
            'total_coletas_2024': int(total_coletas_2024),
            'total_coletas_2025': int(total_coletas_2025),
            'media_2024': round(media_2024, 1),
            'media_2025': round(media_2025, 1),
            'media_ultimo_mes': int(media_ultimo_mes),
            'max_2024': int(max_2024),
            'max_2025': int(max_2025),
            'max_2024_mes': max_2024_mes,
            'max_2025_mes': max_2025_mes,
            'ultimo_mes_display': info_ultimo_mes['display'],
            'ano_comparacao': ano_comparacao,
            'media_comparacao': round(media_comparacao, 1),
            'max_comparacao': int(max_comparacao)
        }
class AnaliseInteligente:
    """Classe para análises inteligentes e insights automáticos - Atualizado score."""
 
    @staticmethod
    def calcular_insights_automaticos(df: pd.DataFrame) -> pd.DataFrame:
        """Calcula insights automáticos para cada laboratório."""
        df_insights = df.copy()
     
        # Volume atual (último mês fechado com lógica de dia de corte)
        info_ultimo_mes = ChartManager._obter_ultimo_mes_fechado(df_insights)
        if info_ultimo_mes['coluna'] and info_ultimo_mes['coluna'] in df_insights.columns:
            df_insights['Volume_Atual_2025'] = df_insights[info_ultimo_mes['coluna']].fillna(0)
        else:
            df_insights['Volume_Atual_2025'] = 0
     
        # Volume máximo do ano passado
        colunas_2024 = [col for col in df_insights.columns if 'N_Coletas_' in col and '24' in col]
        if colunas_2024:
            df_insights['Volume_Maximo_2024'] = df_insights[colunas_2024].max(axis=1).fillna(0)
        else:
            df_insights['Volume_Maximo_2024'] = 0
     
        # Tendência de volume (comparação atual vs máximo histórico)
        df_insights['Tendencia_Volume'] = df_insights.apply(
            lambda row: 'Crescimento' if row['Volume_Atual_2025'] > row['Volume_Maximo_2024']
            else 'Declínio' if row['Volume_Atual_2025'] < row['Volume_Maximo_2024'] * 0.5
            else 'Estável', axis=1
        )
     
        # Insights automáticos
        df_insights['Insights_Automaticos'] = df_insights.apply(
            lambda row: AnaliseInteligente._gerar_insights(row), axis=1
        )
     
        return df_insights
 
    @staticmethod
    def _gerar_insights(row) -> str:
        """Gera insights automáticos baseados nos dados."""
        insights = []
     
        # Análise de dias sem coleta
        dias_sem = row.get('Dias_Sem_Coleta', 0)
        if dias_sem > 90:
            insights.append("🚨 CRÍTICO: Sem coletas há mais de 3 meses")
        elif dias_sem > 60:
            insights.append("⚠️ ALERTA: Sem coletas há mais de 2 meses")
        elif dias_sem > 30:
            insights.append("📉 ATENÇÃO: Sem coletas há mais de 1 mês")
     
        # Análise de volume
        volume_atual = row.get('Volume_Atual_2025', 0)
        volume_max = row.get('Volume_Maximo_2024', 0)
        if volume_max > 0:
            ratio = volume_atual / volume_max
            if ratio > 1.5:
                insights.append("📈 EXCELENTE: Volume 50% acima do histórico")
            elif ratio > 1.2:
                insights.append("📊 POSITIVO: Volume 20% acima do histórico")
            elif ratio < 0.3:
                insights.append("📉 CRÍTICO: Volume 70% abaixo do histórico")
            elif ratio < 0.6:
                insights.append("⚠️ ALERTA: Volume 40% abaixo do histórico")
     
        # Análise de tendência
        variacao = row.get('Variacao_Percentual', 0)
        if variacao > 100:
            insights.append("🚀 CRESCIMENTO: Variação superior a 100%")
        elif variacao > 50:
            insights.append("📈 POSITIVO: Variação superior a 50%")
        elif variacao < -80:
            insights.append("📉 CRÍTICO: Queda superior a 80%")
        elif variacao < -50:
            insights.append("⚠️ ALERTA: Queda superior a 50%")
     
        return " | ".join(insights) if insights else "✅ Estável"
class ReportManager:
    """Gerenciador de geração de relatórios."""
    @staticmethod
    def gerar_relatorio_automatico(df: pd.DataFrame, metrics: KPIMetrics, tipo: str):
        """Gera relatório automático baseado no tipo."""
        if tipo == "semanal":
            ReportManager._gerar_relatorio_semanal(df, metrics)
        elif tipo == "mensal":
            ReportManager._gerar_relatorio_mensal(df, metrics)
    @staticmethod
    def _gerar_relatorio_semanal(df: pd.DataFrame, metrics: KPIMetrics):
        """Gera relatório semanal."""
        sumario = f"""
        📊 **Relatório Semanal de Churn - {datetime.now().strftime('%d/%m/%Y')}**
        **KPIs Principais:**
        • Total de Coletas: {metrics.total_coletas:,}
        • Labs em Risco: {metrics.labs_em_risco:,}
        • Ativos (7d): {metrics.ativos_7d:.1f}%
        **Alertas:**
        • {metrics.labs_alto_risco:,} laboratórios em alto risco
        • {metrics.labs_medio_risco:,} laboratórios em médio risco
        **Recomendações:**
        • Focar nos {metrics.labs_alto_risco} labs de alto risco
        • Monitorar closely os {metrics.labs_medio_risco} labs de médio risco
        """
        st.success("✅ Relatório Semanal Gerado!")
        st.code(sumario, language="markdown")
        # Download do relatório
        st.download_button(
            "📥 Download Relatório Semanal",
            sumario,
            file_name=f"relatorio_semanal_{datetime.now().strftime('%Y%m%d')}.md",
            mime="text/markdown",
            key="download_relatorio_semanal"
        )
    @staticmethod
    def _gerar_relatorio_mensal(df: pd.DataFrame, metrics: KPIMetrics):
        """Gera relatório mensal detalhado."""
        # Calcular top variações
        if 'Variacao_Percentual' in df.columns:
            top_quedas = df.nsmallest(10, 'Variacao_Percentual')[['Nome_Fantasia_PCL', 'Variacao_Percentual', 'Estado']].copy()
            top_quedas['Ranking'] = range(1, len(top_quedas) + 1)
            top_quedas = top_quedas[['Ranking', 'Nome_Fantasia_PCL', 'Variacao_Percentual', 'Estado']]
            
            top_recuperacoes = df.nlargest(10, 'Variacao_Percentual')[['Nome_Fantasia_PCL', 'Variacao_Percentual', 'Estado']].copy()
            top_recuperacoes['Ranking'] = range(1, len(top_recuperacoes) + 1)
            top_recuperacoes = top_recuperacoes[['Ranking', 'Nome_Fantasia_PCL', 'Variacao_Percentual', 'Estado']]
        sumario = f"""
        📊 **Relatório Mensal de Churn - {datetime.now().strftime('%B/%Y').title()}**
        **KPIs Executivos:**
        • Total de Laboratórios: {metrics.total_labs:,}
        • Taxa de Churn: {metrics.churn_rate:.1f}%
        • Net Revenue Retention: {metrics.nrr:.1f}%
        • Laboratórios em Risco: {metrics.labs_em_risco:,}
        • Ativos (7 dias): {metrics.ativos_7d:.1f}%
        • Ativos (30 dias): {metrics.ativos_30d:.1f}%
        **Distribuição por Risco:**
        • Alto Risco: {metrics.labs_alto_risco:,} ({metrics.labs_alto_risco/metrics.total_labs*100:.1f}%)
        • Médio Risco: {metrics.labs_medio_risco:,} ({metrics.labs_medio_risco/metrics.total_labs*100:.1f}%)
        • Baixo Risco: {metrics.labs_baixo_risco:,} ({metrics.labs_baixo_risco/metrics.total_labs*100:.1f}%)
        • Inativos: {metrics.labs_inativos:,} ({metrics.labs_inativos/metrics.total_labs*100:.1f}%)
        **Análise de Tendências:**
        """
        if 'Variacao_Percentual' in df.columns:
            media_variacao = df['Variacao_Percentual'].mean()
            sumario += f"""
        • Variação Média: {media_variacao:.1f}%
        • Top Recuperações: {len(top_recuperacoes)} laboratórios
        • Top Quedas: {len(top_quedas)} laboratórios
            """
        st.success("✅ Relatório Mensal Gerado!")
        with st.expander("📋 Ver Relatório Completo", expanded=True):
            st.code(sumario, language="markdown")
        # Tabelas detalhadas
        if 'Variacao_Percentual' in df.columns:
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("📉 Top 10 Quedas")
                st.dataframe(
                    top_quedas,
                    width='stretch',
                    column_config={
                        "Ranking": st.column_config.NumberColumn("🏆", width="small", help="Posição no ranking"),
                        "Nome_Fantasia_PCL": st.column_config.TextColumn("Laboratório", help="Nome do laboratório"),
                        "Variacao_Percentual": st.column_config.NumberColumn("Variação %", format="%.2f%%", help="Variação percentual"),
                        "Estado": st.column_config.TextColumn("Estado", help="Estado do laboratório")
                    },
                    hide_index=True
                )
            with col2:
                st.subheader("📈 Top 10 Recuperações")
                st.dataframe(
                    top_recuperacoes,
                    width='stretch',
                    column_config={
                        "Ranking": st.column_config.NumberColumn("🏆", width="small", help="Posição no ranking"),
                        "Nome_Fantasia_PCL": st.column_config.TextColumn("Laboratório", help="Nome do laboratório"),
                        "Variacao_Percentual": st.column_config.NumberColumn("Variação %", format="%.2f%%", help="Variação percentual"),
                        "Estado": st.column_config.TextColumn("Estado", help="Estado do laboratório")
                    },
                    hide_index=True
                )
        # Download do relatório
        st.download_button(
            "📥 Download Relatório Mensal",
            sumario,
            file_name=f"relatorio_mensal_{datetime.now().strftime('%Y%m%d')}.md",
            mime="text/markdown",
            key="download_relatorio_mensal"
        )
def show_toast_once(message: str, key: str):
    """Mostra um toast apenas uma vez por sessão."""
    if key not in st.session_state:
        st.toast(message)
        st.session_state[key] = True

def main():
    """Função principal do dashboard v2.0 - Atualizado com tabs e navegação."""
    # ============================================
    # AUTENTICAÇÃO MICROSOFT
    # ============================================
    try:
        # Inicializar autenticador Microsoft
        auth = MicrosoftAuth()
        # Verificar autenticação
        if not create_login_page(auth):
            # Se não conseguiu fazer login, parar execução
            return
        
        # Verificar e renovar token automaticamente se necessário
        if not AuthManager.check_and_refresh_token(auth):
            # Se falhou ao renovar token, forçar novo login
            st.warning("⚠️ Sua sessão expirou. Por favor, faça login novamente.")
            st.rerun()
            return
        
        # Criar cabeçalho com informações do usuário
        create_user_header()
    except Exception as e:
        st.error(f"❌ Erro no sistema de autenticação: {str(e)}")
        st.warning("Verifique as configurações de autenticação no arquivo secrets.toml")
        return
    # ============================================
    # DASHBOARD PRINCIPAL (APENAS PARA USUÁRIOS AUTENTICADOS)
    # ============================================
    # Removido cabeçalho principal para layout mais discreto
    # Carregar e preparar dados
    loader_placeholder = st.empty()
    loader_placeholder.markdown(
        """
        <div class="overlay-loader">
            <div class="overlay-loader__content">
                <div class="overlay-loader__spinner"></div>
                <div class="overlay-loader__title">Carregando dados atualizados...</div>
                <div class="overlay-loader__subtitle">Estamos sincronizando as coletas mais recentes. Isso pode levar alguns segundos.</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )
    try:
        df_raw = DataManager.carregar_dados_churn()
        if df_raw is None:
            st.error("❌ Não foi possível carregar os dados. Por favor, tente novamente mais tarde.")
            return
        df = DataManager.preparar_dados(df_raw)
        show_toast_once(f"✅ Dados carregados: {len(df):,} laboratórios", "dados_carregados")
    finally:
        loader_placeholder.empty()
    # Indicador de última atualização
    if not df.empty and 'Data_Analise' in df.columns:
        ultima_atualizacao = df['Data_Analise'].max()
        st.markdown(f"**Última Atualização:** {ultima_atualizacao.strftime('%d/%m/%Y %H:%M:%S')}")
    
   
        st.header("🔧 Manutenção de Dados VIP")
        
        # Mensagem de manutenção
        st.error("🚧 **TELA EM MANUTENÇÃO** 🚧\n\nEsta funcionalidade está temporariamente indisponível. Por favor, utilize outras telas do sistema.")
        st.stop()
        
        st.markdown("""
        Gerencie a lista de laboratórios VIPs. Você pode adicionar novos laboratórios,
        editar informações existentes ou remover laboratórios da lista.
        """)
        
        df_vips = DataManager.carregar_dados_vip()
        
        if df_vips.empty:
            st.info("Nenhum laboratório VIP cadastrado. Comece adicionando um novo!")
            df_vips = pd.DataFrame(columns=['CNPJ_PCL', 'Nome_Fantasia_PCL', 'Rede', 'Observacoes'])
        
        st.markdown("---")
        st.subheader("Lista de Laboratórios VIPs")
        
        # Exibir e editar a lista de VIPs
        edited_df = st.data_editor(
            df_vips,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "CNPJ_PCL": st.column_config.TextColumn("CNPJ", help="CNPJ do laboratório (apenas números)", required=True),
                "Nome_Fantasia_PCL": st.column_config.TextColumn("Nome Fantasia", help="Nome fantasia do laboratório", required=True),
                "Rede": st.column_config.TextColumn("Rede", help="Nome da rede a qual o laboratório pertence", required=True),
                "Observacoes": st.column_config.TextColumn("Observações", help="Quaisquer observações relevantes")
            }
        )
        
        if st.button("Salvar Alterações VIPs"):
            # Validar CNPJs (apenas números)
            edited_df['CNPJ_PCL'] = edited_df['CNPJ_PCL'].astype(str).str.replace(r'\D', '', regex=True)
            
            # Remover linhas vazias ou com CNPJ duplicado/vazio
            edited_df = edited_df.dropna(subset=['CNPJ_PCL'])
            edited_df = edited_df[edited_df['CNPJ_PCL'] != '']
            edited_df = edited_df.drop_duplicates(subset=['CNPJ_PCL'])
            
            DataManager.salvar_dados_vip(edited_df)
            st.success("✅ Dados VIPs salvos com sucesso!")
            st.rerun()
    
    elif st.session_state.page == "📋 Análise Detalhada":
        st.header("📋 Análise Detalhada")
        # REGRA: Na análise detalhada, se filtro VIP estiver desmarcado, usar base completa
        # Se marcado, mostrar apenas VIPs; se desmarcado, mostrar tudo (incluindo VIPs)
        if filtros.get('apenas_vip', False):
            # Filtro VIP marcado: usar df_filtrado (já filtrado por VIP)
            df_analise_detalhada = df_filtrado.copy()
        else:
            # Filtro VIP desmarcado: usar base completa (df), mas aplicar outros filtros se houver
            df_analise_detalhada = df.copy()
            # Aplicar outros filtros (UF, porte, representante) mas não VIP
            if filtros.get('uf_selecionada') and filtros['uf_selecionada'] != 'Todas':
                if 'Estado' in df_analise_detalhada.columns:
                    df_analise_detalhada = df_analise_detalhada[df_analise_detalhada['Estado'] == filtros['uf_selecionada']]
            if filtros.get('portes') and 'Porte' in df_analise_detalhada.columns:
                df_analise_detalhada = df_analise_detalhada[df_analise_detalhada['Porte'].isin(filtros['portes'])]
            if filtros.get('representantes') and 'Representante_Nome' in df_analise_detalhada.columns:
                df_analise_detalhada = df_analise_detalhada[df_analise_detalhada['Representante_Nome'].isin(filtros['representantes'])]
        
        # Filtros avançados com design moderno
        st.markdown("""
        <div style="background: linear-gradient(135deg, #6BBF47 0%, #52B54B 100%);
                    color: white; padding: 1.5rem; border-radius: 10px;
                    margin-bottom: 1rem; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
            <h3 style="margin: 0; font-size: 1.3rem;">🔍 Busca Inteligente de Laboratórios</h3>
            <p style="margin: 0.5rem 0 0 0; opacity: 0.9;">
                Busque por CNPJ (com ou sem formatação) ou nome do laboratório. Funciona para qualquer laboratório da base.
            </p>
        </div>
        """, unsafe_allow_html=True)
        # Seleção de laboratório específico
        if not df_analise_detalhada.empty:
            if 'CNPJ_Normalizado' not in df_analise_detalhada.columns:
                df_analise_detalhada['CNPJ_Normalizado'] = df_analise_detalhada['CNPJ_PCL'].apply(DataManager.normalizar_cnpj)
            df_analise_detalhada['CNPJ_Normalizado'] = df_analise_detalhada['CNPJ_Normalizado'].fillna('')

            labs_catalogo = df_analise_detalhada[
                ['CNPJ_PCL', 'CNPJ_Normalizado', 'Nome_Fantasia_PCL', 'Razao_Social_PCL', 'Cidade', 'Estado']
            ].copy()
            labs_catalogo = labs_catalogo[labs_catalogo['CNPJ_Normalizado'] != ""]
            labs_catalogo['CNPJ_Normalizado'] = labs_catalogo['CNPJ_Normalizado'].astype(str)
            labs_catalogo = labs_catalogo.drop_duplicates('CNPJ_Normalizado')

            def formatar_cnpj_display(cnpj_val):
                digitos = ''.join(filter(str.isdigit, str(cnpj_val))) if pd.notna(cnpj_val) else ''
                if len(digitos) == 14:
                    return f"{digitos[:2]}.{digitos[2:5]}.{digitos[5:8]}/{digitos[8:12]}-{digitos[12:]}"
                return digitos or "N/A"

            def montar_rotulo(row):
                nome = row.get('Nome_Fantasia_PCL') or row.get('Razao_Social_PCL') or "Laboratório sem nome"
                cidade = row.get('Cidade') or ''
                estado = row.get('Estado') or ''
                if cidade and estado:
                    local = f"{cidade}/{estado}"
                elif cidade:
                    local = cidade
                elif estado:
                    local = estado
                else:
                    local = "Localidade não informada"
                cnpj_fmt = formatar_cnpj_display(row.get('CNPJ_PCL') or row.get('CNPJ_Normalizado'))
                return f"{nome} - {local} (CNPJ: {cnpj_fmt})"

            lab_display_map = {str(row['CNPJ_Normalizado']): montar_rotulo(row) for _, row in labs_catalogo.iterrows()}
            lab_nome_map = {
                str(row['CNPJ_Normalizado']): row.get('Nome_Fantasia_PCL') or row.get('Razao_Social_PCL') or str(row['CNPJ_Normalizado'])
                for _, row in labs_catalogo.iterrows()
            }
            lista_cnpjs_ordenada = sorted(lab_display_map.keys(), key=lambda cnpj: lab_display_map[cnpj].lower())
            lista_cnpjs_validos = set(lista_cnpjs_ordenada)

            LAB_STATE_KEY = 'lab_cnpj_selecionado'
            lab_cnpj_estado = st.session_state.get(LAB_STATE_KEY, "") or ""
            if lab_cnpj_estado and lab_cnpj_estado not in lista_cnpjs_validos:
                lab_cnpj_estado = ""
                st.session_state[LAB_STATE_KEY] = ""

            opcoes_select = [""] + lista_cnpjs_ordenada
            index_padrao = opcoes_select.index(lab_cnpj_estado) if lab_cnpj_estado in lista_cnpjs_validos else 0

            # Layout melhorado com 3 colunas - ajustado para melhor alinhamento
            col1, col2, col3 = st.columns([4, 1.5, 2.5])
            with col1:
                # Campo de busca aprimorado
                busca_lab = st.text_input(
                    "🔎 Buscar PCL",
                    placeholder="CNPJ (com/sem formatação) ou Nome do laboratório",
                    help="Digite CNPJ (com ou sem pontos/traços) ou nome do laboratório/razão social",
                    key="busca_avancada"
                )
            with col2:
                # Botão de busca funcional (lupa como submit)
                buscar_btn = st.button("🔍", type="primary", help="Clique para buscar o PCL no Gralab/backend", use_container_width=True)
            with col3:
                # Seleção por dropdown como alternativa
                lab_selecionado = st.selectbox(
                    "📋 Lista Rápida:",
                    options=opcoes_select,
                    index=index_padrao,
                    format_func=lambda cnpj: "Selecione um laboratório" if cnpj == "" else lab_display_map.get(cnpj, cnpj),
                    help="Ou selecione um laboratório da lista completa"
                )
                lab_selecionado = lab_selecionado or ""
                if lab_selecionado != st.session_state.get(LAB_STATE_KEY, ""):
                    st.session_state[LAB_STATE_KEY] = lab_selecionado
            lab_cnpj_estado = st.session_state.get(LAB_STATE_KEY, "") or ""
            # Informações de ajuda - Atualizado espaçamento dica busca
            with st.expander("💡 Dicas de Busca", expanded=False):
                st.markdown("""
                **🔢 Para CNPJ:**
                - Apenas números: `51865434001248`
                - Com formatação: `51.865.434/0012-48`
                **🏥 Para Nome:**
                - Nome fantasia ou razão social
                - Busca parcial e sem distinção de maiúsculas/minúsculas
                **📊 Resultados:**
                - 1 resultado: Selecionado automaticamente
                - Múltiplos: Lista para escolher o correto
                """)
            # Estado da busca
            lab_final = None
            lab_final_cnpj = lab_cnpj_estado if lab_cnpj_estado in lista_cnpjs_validos else ""
            # Verificar se há busca ativa ou laboratório selecionado
            busca_ativa = buscar_btn or (busca_lab and len(busca_lab.strip()) > 2)
            tem_selecao = bool(lab_cnpj_estado)
            if busca_ativa or tem_selecao:
                # Lógica de busca aprimorada
                if busca_ativa and busca_lab:
                    busca_normalizada = busca_lab.strip()
                    # Verificar se é CNPJ (com ou sem formatação)
                    cnpj_limpo = ''.join(filter(str.isdigit, busca_normalizada))
                    if len(cnpj_limpo) >= 1:
                        if len(cnpj_limpo) >= 14:
                            lab_encontrado = df_analise_detalhada[df_analise_detalhada['CNPJ_Normalizado'] == cnpj_limpo]
                        else:
                            lab_encontrado = df_analise_detalhada[df_analise_detalhada['CNPJ_Normalizado'].str.startswith(cnpj_limpo)]
                    else:
                        # Buscar por nome (case insensitive e parcial) - apenas nome fantasia e razão social
                        lab_encontrado = df_analise_detalhada[
                            df_analise_detalhada['Nome_Fantasia_PCL'].str.contains(busca_normalizada, case=False, na=False) |
                            df_analise_detalhada['Razao_Social_PCL'].str.contains(busca_normalizada, case=False, na=False)
                        ]
                    lab_encontrado = lab_encontrado[lab_encontrado['CNPJ_Normalizado'] != ""].drop_duplicates('CNPJ_Normalizado')
                    if not lab_encontrado.empty:
                        if len(lab_encontrado) == 1:
                            lab_info_unico = lab_encontrado.iloc[0]
                            lab_final = lab_info_unico.get('Nome_Fantasia_PCL') or lab_info_unico.get('Razao_Social_PCL')
                            lab_final_cnpj = str(lab_info_unico.get('CNPJ_Normalizado', ''))
                            st.toast(
                                f"✅ Laboratório encontrado: {lab_final} (CNPJ: {formatar_cnpj_display(lab_final_cnpj)})"
                            )
                            st.session_state[LAB_STATE_KEY] = lab_final_cnpj
                        else:
                            # Múltiplos resultados - mostrar opções
                            st.info(f"🔍 Encontrados {len(lab_encontrado)} laboratórios. Selecione um:")
                            opcoes_df = lab_encontrado.head(10)
                            opcoes_cnpjs = [""] + opcoes_df['CNPJ_Normalizado'].astype(str).tolist()
                            if 'multiplo_resultados' in st.session_state:
                                valor_multi = st.session_state['multiplo_resultados']
                                if valor_multi not in opcoes_cnpjs:
                                    st.session_state['multiplo_resultados'] = ""
                            lab_selecionado_multiplo = st.selectbox(
                                "Selecione o laboratório correto:",
                                options=opcoes_cnpjs,
                                format_func=lambda cnpj: "Selecione" if cnpj == "" else lab_display_map.get(cnpj, cnpj),
                                key="multiplo_resultados"
                            )
                            if lab_selecionado_multiplo:
                                lab_final_cnpj = str(lab_selecionado_multiplo)
                                lab_final = lab_nome_map.get(lab_final_cnpj, lab_final_cnpj)
                                st.session_state[LAB_STATE_KEY] = lab_final_cnpj
                    else:
                        # Limpar flag de laboratório fora dos filtros ao iniciar nova busca
                        if 'lab_fora_filtros' in st.session_state:
                            del st.session_state['lab_fora_filtros']
                        # Não encontrou - verificar se existe na base completa e qual filtro está impedindo
                        cnpj_limpo = ''.join(filter(str.isdigit, busca_normalizada))
                        lab_na_base_completa = None
                        
                        if len(cnpj_limpo) >= 1:
                            if len(cnpj_limpo) >= 14:
                                lab_na_base_completa = df[df['CNPJ_Normalizado'] == cnpj_limpo]
                            else:
                                lab_na_base_completa = df[df['CNPJ_Normalizado'].str.startswith(cnpj_limpo)]
                        else:
                            # Buscar por nome na base completa
                            lab_na_base_completa = df[
                                df['Nome_Fantasia_PCL'].str.contains(busca_normalizada, case=False, na=False) |
                                df['Razao_Social_PCL'].str.contains(busca_normalizada, case=False, na=False)
                            ]
                        
                        lab_na_base_completa = lab_na_base_completa[lab_na_base_completa['CNPJ_Normalizado'] != ""].drop_duplicates('CNPJ_Normalizado')
                        
                        if not lab_na_base_completa.empty:
                            # Encontrou na base completa mas não nos filtros atuais
                            # Se encontrou exatamente 1, vamos mostrar mesmo assim (usando base completa)
                            if len(lab_na_base_completa) == 1:
                                lab_info_encontrado = lab_na_base_completa.iloc[0]
                                lab_final_cnpj = str(lab_info_encontrado.get('CNPJ_Normalizado', ''))
                                lab_final = lab_info_encontrado.get('Nome_Fantasia_PCL') or lab_info_encontrado.get('Razao_Social_PCL') or lab_final_cnpj
                                st.session_state[LAB_STATE_KEY] = lab_final_cnpj
                                # Marcar que há um laboratório pesquisado fora dos filtros
                                st.session_state['lab_fora_filtros'] = True
                                st.warning("⚠️ Este laboratório está fora dos filtros atuais, mas será exibido mesmo assim.")
                                
                                # Identificar quais filtros estão ativos
                                filtros_ativos = []
                                if filtros.get('apenas_vip', False):
                                    filtros_ativos.append("**Apenas VIPs**")
                                if filtros.get('representantes'):
                                    filtros_ativos.append(f"**Representante(s)**: {', '.join(filtros['representantes'])}")
                                if filtros.get('ufs'):
                                    filtros_ativos.append(f"**UF(s)**: {', '.join(filtros['ufs'])}")
                                
                                if filtros_ativos:
                                    st.info(f"💡 Este laboratório está sendo filtrado por:\n\n" + 
                                           "\n".join([f"- {f}" for f in filtros_ativos]) +
                                           "\n\n💡 **Dica**: Desative os filtros na barra lateral para ver apenas laboratórios que atendem aos critérios.")
                            else:
                                # Múltiplos resultados na base completa
                                st.warning("⚠️ Nenhum laboratório encontrado com os filtros atuais")
                                
                                # Identificar quais filtros estão ativos
                                filtros_ativos = []
                                if filtros.get('apenas_vip', False):
                                    filtros_ativos.append("**Apenas VIPs**")
                                if filtros.get('representantes'):
                                    filtros_ativos.append(f"**Representante(s)**: {', '.join(filtros['representantes'])}")
                                if filtros.get('ufs'):
                                    filtros_ativos.append(f"**UF(s)**: {', '.join(filtros['ufs'])}")
                                
                                if filtros_ativos:
                                    st.info(f"💡 Encontramos **{len(lab_na_base_completa)} laboratório(s)** na base completa, mas está(ão) sendo filtrado(s) por:\n\n" + 
                                           "\n".join([f"- {f}" for f in filtros_ativos]))
                                    
                                    st.caption("💡 **Dica**: Desative os filtros na barra lateral para visualizar estes laboratórios.")
                                else:
                                    st.info(f"💡 Encontramos **{len(lab_na_base_completa)} laboratório(s)** na base completa, mas está(ão) fora do período selecionado ou de outros filtros aplicados.")
                        else:
                            # Não encontrou nem na base completa
                            st.warning("⚠️ Nenhum laboratório encontrado com os critérios informados")
                            st.caption("Este laboratório não está na nossa base de dados.")
                elif tem_selecao:
                    # Laboratório selecionado diretamente da lista
                    lab_final_cnpj = st.session_state.get(LAB_STATE_KEY, "")
                    lab_final = lab_nome_map.get(lab_final_cnpj, lab_final_cnpj)
                    # Limpar flag de laboratório fora dos filtros quando selecionado da lista
                    if 'lab_fora_filtros' in st.session_state:
                        del st.session_state['lab_fora_filtros']
                # Renderizar dados do laboratório encontrado/selecionado
                if lab_final_cnpj:
                    st.markdown("---") # Separador antes dos dados
                    # Verificar se é VIP
                    df_vip = DataManager.carregar_dados_vip()
                    # Tentar buscar primeiro em df_analise_detalhada, se não encontrar, buscar na base completa
                    if lab_final_cnpj:
                        lab_data = df_analise_detalhada[df_analise_detalhada['CNPJ_Normalizado'] == lab_final_cnpj]
                        # Se não encontrou em df_analise_detalhada, buscar na base completa
                        if lab_data.empty:
                            lab_data = df[df['CNPJ_Normalizado'] == lab_final_cnpj]
                    else:
                        lab_data = df_analise_detalhada[df_analise_detalhada['Nome_Fantasia_PCL'] == lab_final]
                        # Se não encontrou em df_analise_detalhada, buscar na base completa
                        if lab_data.empty:
                            lab_data = df[df['Nome_Fantasia_PCL'] == lab_final]
                    info_vip = None
                    if not lab_data.empty and df_vip is not None:
                        cnpj_lab = lab_data.iloc[0].get('CNPJ_PCL', '')
                        info_vip = VIPManager.buscar_info_vip(cnpj_lab, df_vip)
                    # Container principal para informações do laboratório
                    st.markdown(f"""
                        <div style="background: linear-gradient(135deg, #6BBF47 0%, #52B54B 100%);
                                    color: white; padding: 2rem; border-radius: 15px;
                                    margin-bottom: 2rem; box-shadow: 0 8px 25px rgba(0,0,0,0.15);">
                            <div style="display: flex; align-items: center;">
                                <div style="font-size: 2rem; margin-right: 1rem;">🏥</div>
                                <div>
                                    <h2 style="margin: 0; font-size: 1.8rem; font-weight: 600;">{lab_final}</h2>
                                </div>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                    # Armazenar informações da rede para filtro automático na tabela
                    if info_vip and 'rede' in info_vip:
                        st.session_state['rede_lab_pesquisado'] = info_vip['rede']
                    else:
                        st.session_state['rede_lab_pesquisado'] = None
                    # Ficha Técnica Comercial
                    st.markdown("""
                        <div style="background: white; border-radius: 8px; padding: 1.5rem; margin-bottom: 2rem;
                                    border: 1px solid #e9ecef; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                            <h3 style="margin: 0 0 1rem 0; color: #2c3e50; font-weight: 600; border-bottom: 2px solid #6BBF47; padding-bottom: 0.5rem;">
                                📋 Ficha Técnica Comercial
                            </h3>
                        """, unsafe_allow_html=True)
                    # Informações de contato e localização
                    # Tentar buscar primeiro em df_analise_detalhada, se não encontrar, buscar na base completa
                    if lab_final_cnpj:
                        lab_data = df_analise_detalhada[df_analise_detalhada['CNPJ_Normalizado'] == lab_final_cnpj]
                        # Se não encontrou em df_analise_detalhada, buscar na base completa
                        if lab_data.empty:
                            lab_data = df[df['CNPJ_Normalizado'] == lab_final_cnpj]
                    else:
                        lab_data = df_analise_detalhada[df_analise_detalhada['Nome_Fantasia_PCL'] == lab_final]
                        # Se não encontrou em df_analise_detalhada, buscar na base completa
                        if lab_data.empty:
                            lab_data = df[df['Nome_Fantasia_PCL'] == lab_final]
                    if not lab_data.empty:
                            lab_info = lab_data.iloc[0]
                         
                            # CNPJ formatado
                            cnpj_raw = str(lab_info.get('CNPJ_PCL', ''))
                            cnpj_formatado = f"{cnpj_raw[:2]}.{cnpj_raw[2:5]}.{cnpj_raw[5:8]}/{cnpj_raw[8:12]}-{cnpj_raw[12:14]}" if len(cnpj_raw) == 14 else cnpj_raw
                         
                            # Carregar dados de laboratories.csv
                            df_labs = DataManager.carregar_laboratories()
                            info_lab = None
                            if df_labs is not None and not df_labs.empty:
                                info_lab = VIPManager.buscar_info_laboratory(cnpj_raw, df_labs)
                            
                            # Prioridade: 1) laboratories.csv, 2) matriz VIP (legado), 3) dados do laboratório
                            contato = ''
                            telefone = ''
                            email = ''
                            
                            if info_lab:
                                contato = info_lab.get('contato', '')
                                telefone = info_lab.get('telefone', '')
                                email = info_lab.get('email', '')
                            
                            # Fallback para dados VIP (legado) se não encontrado no laboratories.csv
                            if not contato and info_vip:
                                contato = info_vip.get('contato', '')
                            if not telefone and info_vip:
                                telefone = info_vip.get('telefone', '')
                            if not email and info_vip:
                                email = info_vip.get('email', '')
                            
                            # Último fallback: dados do lab_info
                            # Nota: lab_info pode não ter campo 'Contato' separado, então apenas telefone/email
                            if not telefone:
                                telefone = lab_info.get('Telefone', 'N/A')
                            if not email:
                                email = lab_info.get('Email', 'N/A')
                            # Se contato ainda estiver vazio, deixar como 'N/A' (não há fallback no lab_info)
                            
                            # Aceita voucher (onlineVoucher)
                            aceita_voucher_flag = None
                            if info_lab:
                                aceita_voucher_flag = info_lab.get('onlineVoucher')
                            if aceita_voucher_flag is None:
                                aceita_voucher_flag = lab_info.get('onlineVoucher')
                            aceita_voucher_txt = (
                                'Sim' if aceita_voucher_flag is True else
                                'Não' if aceita_voucher_flag is False else
                                'Não informado'
                            )
                            
                            representante = lab_info.get('Representante_Nome', 'N/A')
                            
                            # Limpar dados vazios
                            telefone = telefone if telefone and telefone != 'N/A' and telefone != '' else 'N/A'
                            email = email if email and email != 'N/A' and email != '' else 'N/A'
                            contato = contato if contato and contato != '' else 'N/A'
                            representante = representante if representante and representante != 'N/A' else 'N/A'
                            
                            # Extrair novos dados do info_lab
                            endereco_completo = info_lab.get('endereco', {}) if info_lab else {}
                            logistic_data = info_lab.get('logistic', {}) if info_lab else {}
                            licensed_list = info_lab.get('licensed', []) if info_lab else []
                            allowed_methods_list = info_lab.get('allowedMethods', []) if info_lab else []
                            
                            # Mapear dias da semana para português
                            dias_semana_map = {
                                'mon': 'Segunda', 'tue': 'Terça', 'wed': 'Quarta', 
                                'thu': 'Quinta', 'fri': 'Sexta', 'sat': 'Sábado', 'sun': 'Domingo'
                            }
                            
                            # Formatar dias de funcionamento
                            dias_funcionamento = []
                            if logistic_data.get('days'):
                                for dia in logistic_data.get('days', []):
                                    dias_funcionamento.append(dias_semana_map.get(dia.lower(), dia.capitalize()))
                            dias_funcionamento_str = ', '.join(dias_funcionamento) if dias_funcionamento else 'N/A'
                            horario_funcionamento = logistic_data.get('openingHours', '') if logistic_data.get('openingHours') else 'N/A'
                            
                            # Formatar endereço completo
                            endereco_linha1 = ''
                            endereco_linha2 = ''
                            cep_formatado = 'N/A'
                            if endereco_completo:
                                endereco_parts = []
                                if endereco_completo.get('address'):
                                    endereco_parts.append(endereco_completo.get('address', ''))
                                if endereco_completo.get('number'):
                                    endereco_parts.append(f"nº {endereco_completo.get('number', '')}")
                                if endereco_completo.get('addressComplement'):
                                    endereco_parts.append(endereco_completo.get('addressComplement', ''))
                                endereco_linha1 = ', '.join(endereco_parts) if endereco_parts else 'N/A'
                                
                                endereco_parts2 = []
                                if endereco_completo.get('neighbourhood'):
                                    endereco_parts2.append(endereco_completo.get('neighbourhood', ''))
                                if endereco_completo.get('city'):
                                    endereco_parts2.append(endereco_completo.get('city', ''))
                                if endereco_completo.get('state_code'):
                                    endereco_parts2.append(endereco_completo.get('state_code', ''))
                                endereco_linha2 = ' - '.join(endereco_parts2) if endereco_parts2 else 'N/A'
                                
                                # Formatar CEP
                                if endereco_completo.get('postalCode'):
                                    cep_raw = str(endereco_completo.get('postalCode', '')).strip()
                                    if len(cep_raw) == 8:
                                        cep_formatado = f"{cep_raw[:5]}-{cep_raw[5:]}"
                                    else:
                                        cep_formatado = cep_raw
                            
                            # Fallback para dados do lab_info se endereço completo não disponível
                            if not endereco_linha1 or endereco_linha1 == 'N/A':
                                endereco_linha1 = 'N/A'
                            if not endereco_linha2 or endereco_linha2 == 'N/A':
                                endereco_linha2 = f"{lab_info.get('Cidade', 'N/A')} - {lab_info.get('Estado', 'N/A')}"
                            
                            # Formatar licenças
                            licencas_map = {
                                'clt': 'CLT', 'cnh': 'CNH', 'cltCnh': 'CLT/CNH',
                                'other': 'Outros', 'online': 'Online',
                                'civilService': 'Concurso Público', 
                                'civilServiceAnalysis50': 'Concurso Público (50)',
                                'otherAnalysis50': 'Outros (50)'
                            }
                            licencas_formatadas = [licencas_map.get(l, l) for l in licensed_list] if licensed_list else []
                            licencas_str = ', '.join(licencas_formatadas) if licencas_formatadas else 'N/A'
                            
                            # Formatar métodos de pagamento
                            metodos_map = {
                                'cash': 'Dinheiro', 'credit': 'Crédito', 'debit': 'Débito',
                                'billing_laboratory': 'Faturamento Lab', 'billing_company': 'Faturamento Empresa',
                                'billing': 'Faturamento', 'bank_billet': 'Boleto',
                                'eCredit': 'e-Crédito', 'pix': 'PIX'
                            }
                            metodos_formatados = [metodos_map.get(m, m) for m in allowed_methods_list] if allowed_methods_list else []
                            metodos_str = ', '.join(metodos_formatados) if metodos_formatados else 'N/A'
                         
                            st.markdown(f"""
                            <div style="background: #f8f9fa; border-radius: 6px; padding: 1rem; margin-bottom: 1rem; border-left: 4px solid #6c757d;">
                                <div style="font-size: 0.9rem; color: #666; margin-bottom: 0.5rem; font-weight: 600;">INFORMAÇÕES DE CONTATO</div>
                                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 1rem;">
                                    <div>
                                        <div style="font-size: 0.8rem; color: #666; margin-bottom: 0.3rem;">CNPJ</div>
                                        <div style="font-size: 1rem; font-weight: bold; color: #495057;">{cnpj_formatado}</div>
                                    </div>
                                    <div>
                                        <div style="font-size: 0.8rem; color: #666; margin-bottom: 0.3rem;">CEP</div>
                                        <div style="font-size: 1rem; font-weight: bold; color: #495057;">{cep_formatado}</div>
                                    </div>
                                    <div style="grid-column: 1 / -1;">
                                        <div style="font-size: 0.8rem; color: #666; margin-bottom: 0.3rem;">Endereço</div>
                                        <div style="font-size: 1rem; font-weight: bold; color: #495057;">{endereco_linha1}</div>
                                        <div style="font-size: 0.9rem; color: #6c757d; margin-top: 0.2rem;">{endereco_linha2}</div>
                                    </div>
                                    <div>
                                        <div style="font-size: 0.8rem; color: #666; margin-bottom: 0.3rem;">Localização</div>
                                        <div style="font-size: 1rem; font-weight: bold; color: #495057;">{lab_info.get('Cidade', 'N/A')} - {lab_info.get('Estado', 'N/A')}</div>
                                    </div>
                                    <div>
                                        <div style="font-size: 0.8rem; color: #666; margin-bottom: 0.3rem;">Contato</div>
                                        <div style="font-size: 1rem; font-weight: bold; color: #495057;">{contato}</div>
                                    </div>
                                    <div>
                                        <div style="font-size: 0.8rem; color: #666; margin-bottom: 0.3rem;">Telefone</div>
                                        <div style="font-size: 1rem; font-weight: bold; color: #495057;">{telefone}</div>
                                    </div>
                                    <div>
                                        <div style="font-size: 0.8rem; color: #666; margin-bottom: 0.3rem;">Email</div>
                                        <div style="font-size: 1rem; font-weight: bold; color: #495057;">{email}</div>
                                    </div>
                                    <div>
                                        <div style="font-size: 0.8rem; color: #666; margin-bottom: 0.3rem;">Representante</div>
                                        <div style="font-size: 1rem; font-weight: bold; color: #495057;">{representante}</div>
                                    </div>
                                    <div>
                                        <div style="font-size: 0.8rem; color: #666; margin-bottom: 0.3rem;">Dias de Funcionamento</div>
                                        <div style="font-size: 1rem; font-weight: bold; color: #495057;">{dias_funcionamento_str}</div>
                                    </div>
                                    <div>
                                        <div style="font-size: 0.8rem; color: #666; margin-bottom: 0.3rem;">Horário</div>
                                        <div style="font-size: 1rem; font-weight: bold; color: #495057;">{horario_funcionamento}</div>
                                    </div>
                                    <div>
                                        <div style="font-size: 0.8rem; color: #666; margin-bottom: 0.3rem;">Licenças</div>
                                        <div style="font-size: 1rem; font-weight: bold; color: #495057;">{licencas_str}</div>
                                    </div>
                                    <div>
                                        <div style="font-size: 0.8rem; color: #666; margin-bottom: 0.3rem;">Métodos de Pagamento</div>
                                        <div style="font-size: 1rem; font-weight: bold; color: #495057;">{metodos_str}</div>
                                    </div>
                                </div>
                            </div>
                            """, unsafe_allow_html=True)

                            # Bloco de preços praticados
                            def _formatar_preco_valor(valor):
                                try:
                                    if pd.notna(valor):
                                        return f"R$ {float(valor):.2f}".replace('.', ',')
                                except Exception:
                                    pass
                                return "N/A"

                            price_labels = {
                                'CLT': 'CLT',
                                'CNH': 'CNH',
                                'Civil_Service': 'Concurso Público',
                                'Civil_Service50': 'Concurso Público (50)',
                                'CLT_CNH': 'CLT / CNH',
                                'Outros': 'Outros',
                                'Outros50': 'Outros (50)'
                            }

                            price_cards = []
                            possui_preco = False
                            for key, cfg in PRICE_CATEGORIES.items():
                                prefix = cfg['prefix']
                                label = price_labels.get(prefix, prefix.replace('_', ' '))
                                total = lab_info.get(f'Preco_{prefix}_Total', np.nan)
                                coleta = lab_info.get(f'Preco_{prefix}_Coleta', np.nan)
                                exame = lab_info.get(f'Preco_{prefix}_Exame', np.nan)

                                possui_valores = any(pd.notna(v) for v in [total, coleta, exame])
                                if not possui_valores:
                                    continue

                                possui_preco = True

                                price_cards.append(
                                    f"<div style=\"background: white; border-radius: 8px; padding: 1rem; "
                                    f"box-shadow: 0 2px 6px rgba(0,0,0,0.08);\">"
                                    f"<div style=\"font-size: 0.85rem; color: #6c757d; text-transform: uppercase; "
                                    f"letter-spacing: 0.5px; margin-bottom: 0.6rem; font-weight: 700;\">"
                                    f"{label}"
                                    f"</div>"
                                    f"<div style=\"display: flex; justify-content: space-between; font-size: 0.85rem; "
                                    f"color: #6c757d; margin-bottom: 0.4rem;\">"
                                    f"<span>Coleta</span>"
                                    f"<strong style=\"color: #495057;\">{_formatar_preco_valor(coleta)}</strong>"
                                    f"</div>"
                                    f"<div style=\"display: flex; justify-content: space-between; font-size: 0.85rem; "
                                    f"color: #6c757d; margin-bottom: 0.4rem;\">"
                                    f"<span>Exame</span>"
                                    f"<strong style=\"color: #495057;\">{_formatar_preco_valor(exame)}</strong>"
                                    f"</div>"
                                    f"<div style=\"display: flex; justify-content: space-between; font-size: 0.85rem; "
                                    f"color: #6c757d;\">"
                                    f"<span>Total</span>"
                                    f"<strong style=\"color: #495057;\">{_formatar_preco_valor(total)}</strong>"
                                    f"</div>"
                                    f"</div>"
                                )

                            if possui_preco or pd.notna(lab_info.get('Voucher_Commission', np.nan)) or aceita_voucher_txt != 'Não informado':
                                voucher_valor = lab_info.get('Voucher_Commission', np.nan)
                                voucher_fmt = f"{float(voucher_valor):.0f}%" if pd.notna(voucher_valor) else None
                                data_preco = lab_info.get('Data_Preco_Atualizacao')
                                if isinstance(data_preco, pd.Timestamp):
                                    data_preco_fmt = data_preco.tz_localize(None).strftime("%d/%m/%Y %H:%M")
                                else:
                                    data_preco_fmt = "N/A"

                                if price_cards:
                                    cards_html = "".join(price_cards)
                                else:
                                    cards_html = (
                                        "<div style=\"background: white; border-radius: 8px; padding: 1rem; "
                                        "color: #6c757d; text-align: center;\">Nenhum preço cadastrado.</div>"
                                    )

                                st.markdown(f"""
                                    <div style="background: #f8f9fa; border-radius: 6px; padding: 1rem; margin-bottom: 1rem; border-left: 4px solid #0d6efd;">
                                        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 1rem;">
                                            <div style="font-size: 0.9rem; color: #0d6efd; font-weight: 700; text-transform: uppercase;">Tabela de Preços</div>
                                            <div style="font-size: 0.8rem; color: #6c757d;">
                                                Atualizado em <strong>{data_preco_fmt}</strong> • Aceita voucher: <strong>{aceita_voucher_txt}</strong>{f" • Comissão: <strong>{voucher_fmt}</strong>" if voucher_fmt else ""}
                                            </div>
                                        </div>
                                        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 1rem;">
                                            {cards_html}
                                        </div>
                                    </div>
                                """, unsafe_allow_html=True)

                    # Informações VIP se disponível
                    if info_vip:
                        st.markdown(f"""
                            <div style="background: #f8f9fa; border-radius: 6px; padding: 1rem; margin-bottom: 1rem; border-left: 4px solid #6BBF47;">
                                <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 1rem; text-align: center;">
                                    <div>
                                        <div style="font-size: 0.8rem; color: #666; margin-bottom: 0.3rem;">RANKING GERAL</div>
                                        <div style="font-size: 1.2rem; font-weight: bold; color: #FFD700;">{info_vip.get('ranking', 'N/A')}</div>
                                    </div>
                                    <div>
                                        <div style="font-size: 0.8rem; color: #666; margin-bottom: 0.3rem;">RANKING REDE</div>
                                        <div style="font-size: 1.2rem; font-weight: bold; color: #FFA500;">{info_vip.get('ranking_rede', 'N/A')}</div>
                                    </div>
                                    <div>
                                        <div style="font-size: 0.8rem; color: #666; margin-bottom: 0.3rem;">REDE</div>
                                        <div style="font-size: 1.1rem; font-weight: bold; color: #6BBF47;">{info_vip.get('rede', 'N/A')}</div>
                                    </div>
                                </div>
                            </div>
                            """, unsafe_allow_html=True)
                    # Métricas comerciais essenciais
                    # Usar base completa se não encontrou em df_analise_detalhada
                    df_para_metricas = df_analise_detalhada.copy()
                    if lab_final_cnpj:
                        lab_existe = not df_analise_detalhada[df_analise_detalhada['CNPJ_Normalizado'] == lab_final_cnpj].empty
                        if not lab_existe:
                            df_para_metricas = df.copy()
                    
                    metricas = MetricasAvancadas.calcular_metricas_lab(
                        df_para_metricas,
                        lab_cnpj=lab_final_cnpj,
                        lab_nome=lab_final
                    )
                    if metricas:
                        # Dados de Performance
                        st.markdown(f"""
                            <div style="background: #f8f9fa; border-radius: 6px; padding: 1rem; margin-bottom: 1rem; border-left: 4px solid #28a745;">
                                <div style="font-size: 0.9rem; color: #666; margin-bottom: 0.5rem; font-weight: 600;">PERFORMANCE 2025</div>
                                <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 1rem; text-align: center;">
                                    <div>
                                        <div style="font-size: 0.8rem; color: #666;">Total Coletas</div>
                                        <div style="font-size: 1.3rem; font-weight: bold; color: #28a745;">{metricas['total_coletas']:,}</div>
                                    </div>
                                    <div>
                                        <div style="font-size: 0.8rem; color: #666;">Média 3 Meses</div>
                                        <div style="font-size: 1.3rem; font-weight: bold; color: #28a745;">{metricas['media_3_meses']:.1f}</div>
                                    </div>
                                    <div>
                                        <div style="font-size: 0.8rem; color: #666;">Média Diária</div>
                                        <div style="font-size: 1.3rem; font-weight: bold; color: #28a745;">{metricas['media_diaria']:.1f}</div>
                                    </div>
                                    <div>
                                        <div style="font-size: 0.8rem; color: #666;">Coletas (Hoje)</div>
                                        <div style="font-size: 1.3rem; font-weight: bold; color: #28a745;">{metricas['vol_hoje']:,}</div>
                                    </div>
                                </div>
                            </div>
                            """, unsafe_allow_html=True)
                        # Status (sem classificação de risco alto/médio/baixo, sem médias móveis)
                        status_color = "#28a745" if metricas['agudo'] == "Ativo" else "#dc3545"
                        dias_sem_coleta = metricas.get('dias_sem_coleta', 0)
                        # Remover indicador "0 dias sem coleta" - mostrar apenas o valor
                        dias_color = "#28a745" if dias_sem_coleta == 0 else "#ffc107" if dias_sem_coleta <= 7 else "#dc3545"
                     
                        st.markdown(f"""
                            <div style="background: #f8f9fa; border-radius: 6px; padding: 1rem; margin-bottom: 1rem; border-left: 4px solid {dias_color};">
                                <div style="font-size: 0.9rem; color: #666; margin-bottom: 0.5rem; font-weight: 600;">STATUS</div>
                                <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 1rem; text-align: center;">
                                    <div>
                                        <div style="font-size: 0.8rem; color: #666;">Status Atual</div>
                                        <div style="font-size: 1.3rem; font-weight: bold; color: {status_color};">{metricas['agudo']}</div>
                                    </div>
                                    <div>
                                        <div style="font-size: 0.8rem; color: #666;">Dias sem Coleta</div>
                                        <div style="font-size: 1.3rem; font-weight: bold; color: {dias_color};">{dias_sem_coleta}</div>
                                    </div>
                                </div>
                            </div>
                            """, unsafe_allow_html=True)
                        # Histórico de Performance - Reorganizado conforme solicitação
                        # Calcular máxima de coletas histórica (respeitando meses disponíveis)
                        metricas_evolucao = MetricasAvancadas.calcular_metricas_evolucao(
                            df_para_metricas,
                            lab_cnpj=lab_final_cnpj,
                            lab_nome=lab_final
                        )
                        st.markdown(f"""
                            <div style="background: #f8f9fa; border-radius: 6px; padding: 1rem; margin-bottom: 1rem; border-left: 4px solid #17a2b8;">
                                <div style="font-size: 0.9rem; color: #666; margin-bottom: 0.5rem; font-weight: 600;">HISTÓRICO DE PERFORMANCE</div>
                                <div style="display: grid; grid-template-columns: repeat(4, minmax(0,1fr)); gap: 1rem; text-align: center;">
                                    <div>
                                        <div style="font-size: 0.8rem; color: #666;">Média 2024</div>
                                        <div style="font-size: 1.3rem; font-weight: bold; color: #17a2b8;">{metricas_evolucao['media_2024']:.1f}</div>
                                    </div>
                                    <div>
                                        <div style="font-size: 0.8rem; color: #666;">Média 2025</div>
                                        <div style="font-size: 1.3rem; font-weight: bold; color: #17a2b8;">{metricas_evolucao['media_2025']:.1f}</div>
                                    </div>
                                    <div>
                                        <div style="font-size: 0.8rem; color: #666;">Máxima 2024</div>
                                        <div style="font-size: 1.3rem; font-weight: bold; color: #17a2b8;">{metricas_evolucao['max_2024']:,}</div>
                                        <div style="font-size: 0.75rem; color: #6c757d;">{metricas_evolucao.get('max_2024_mes', 'N/A')}</div>
                                    </div>
                                    <div>
                                        <div style="font-size: 0.8rem; color: #666;">Máxima 2025</div>
                                        <div style="font-size: 1.3rem; font-weight: bold; color: #17a2b8;">{metricas_evolucao['max_2025']:,}</div>
                                        <div style="font-size: 0.75rem; color: #6c757d;">{metricas_evolucao.get('max_2025_mes', 'N/A')}</div>
                                    </div>
                                </div>
                            </div>
                            """, unsafe_allow_html=True)
                        # Adiciona também os cards de Totais e Comparativos ao bloco de Histórico
                        st.markdown(f"""
                        <div style="background: #f8f9fa; border-radius: 6px; padding: 1rem; margin-bottom: 1rem; border-left: 4px solid #28a745;">
                            <div style="font-size: 0.9rem; color: #666; margin-bottom: 0.5rem; font-weight: 600;">TOTAIS DE COLETAS</div>
                        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 1rem; text-align: center;">
                            <div>
                                <div style="font-size: 0.8rem; color: #666;">Total 2024</div>
                                <div style="font-size: 1.4rem; font-weight: bold; color: #28a745;">{metricas_evolucao['total_coletas_2024']:,}</div>
                            </div>
                            <div>
                                <div style="font-size: 0.8rem; color: #666;">Total 2025</div>
                                <div style="font-size: 1.4rem; font-weight: bold; color: #6BBF47;">{metricas_evolucao['total_coletas_2025']:,}</div>
                            </div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

                        st.markdown(f"""
                        <div style="background: #f8f9fa; border-radius: 6px; padding: 1rem; margin-bottom: 1rem; border-left: 4px solid #0d6efd;">
                            <div style="font-size: 0.9rem; color: #666; margin-bottom: 0.5rem; font-weight: 600;">MÊS ATUAL E PICO HISTÓRICO</div>
                            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 1rem; text-align: center;">
                                <div>
                                    <div style="font-size: 0.8rem; color: #666;">Volume mês atual ({metricas_evolucao['ultimo_mes_display']})</div>
                                    <div style="font-size: 1.3rem; font-weight: bold; color: #0d6efd;">{metricas_evolucao['media_ultimo_mes']:,}</div>
                                </div>
                                <div>
                                    <div style="font-size: 0.8rem; color: #666;">Maior mês 2024</div>
                                    <div style="font-size: 1.1rem; font-weight: bold; color: #6c757d;">{metricas_evolucao.get('max_2024_mes', 'N/A')}</div>
                                    <div style="font-size: 0.9rem; color: #6BBF47; font-weight: 700;">{metricas_evolucao['max_2024']:,}</div>
                                </div>
                                <div>
                                    <div style="font-size: 0.8rem; color: #666;">Maior mês 2025</div>
                                    <div style="font-size: 1.1rem; font-weight: bold; color: #6c757d;">{metricas_evolucao.get('max_2025_mes', 'N/A')}</div>
                                    <div style="font-size: 0.9rem; color: #6BBF47; font-weight: 700;">{metricas_evolucao['max_2025']:,}</div>
                                </div>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)

                        variacao_ultimo_vs_media = ((metricas_evolucao['media_ultimo_mes'] - metricas_evolucao['media_comparacao']) / metricas_evolucao['media_comparacao'] * 100) if metricas_evolucao['media_comparacao'] > 0 else 0
                        variacao_ultimo_vs_maxima = ((metricas_evolucao['media_ultimo_mes'] - metricas_evolucao['max_comparacao']) / metricas_evolucao['max_comparacao'] * 100) if metricas_evolucao['max_comparacao'] > 0 else 0
                        cor_variacao = "#28a745" if variacao_ultimo_vs_media >= 0 else "#dc3545"
                        cor_maxima = "#28a745" if variacao_ultimo_vs_maxima >= 0 else "#dc3545"
                        ano_comp = metricas_evolucao['ano_comparacao']
                        st.markdown(f"""
                        <div style="background: #f8f9fa; border-radius: 6px; padding: 1rem; margin-bottom: 1rem; border-left: 4px solid #6f42c1;">
                            <div style="font-size: 0.9rem; color: #666; margin-bottom: 0.5rem; font-weight: 600;">COMPARATIVOS</div>
                            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 1rem; text-align: center;">
                                <div>
                                    <div style="font-size: 0.8rem; color: #666;">Último Mês ({metricas_evolucao['ultimo_mes_display']}) vs Média {ano_comp}</div>
                                    <div style="font-size: 1.2rem; font-weight: bold; color: {cor_variacao};">
                                        {'+' if variacao_ultimo_vs_media >= 0 else ''}{variacao_ultimo_vs_media:.1f}%
                                    </div>
                                    <div style="font-size: 0.7rem; color: #666;">{metricas_evolucao['media_ultimo_mes']:,} vs {metricas_evolucao['media_comparacao']:.1f}</div>
                                </div>
                                <div>
                                    <div style="font-size: 0.8rem; color: #666;">Último Mês ({metricas_evolucao['ultimo_mes_display']}) vs Máxima {ano_comp}</div>
                                    <div style="font-size: 1.2rem; font-weight: bold; color: {cor_maxima};">
                                        {'+' if variacao_ultimo_vs_maxima >= 0 else ''}{variacao_ultimo_vs_maxima:.1f}%
                                    </div>
                                    <div style="font-size: 0.7rem; color: #666;">{metricas_evolucao['media_ultimo_mes']:,} vs {metricas_evolucao['max_comparacao']:,}</div>
                                </div>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                        # Gráfico de coletas diárias (mês atual) + comparação
                        st.markdown("---")
                        # Adicionar helper no título
                        col_title, col_helper = st.columns([4, 1])
                        with col_title:
                            st.subheader("📊 Coletas por Dia (mês atual)")
                        with col_helper:
                            st.markdown(
                                "💡 *Visualize coletas dia a dia e compare com outros meses*",
                                help="Esta seção mostra o volume de coletas realizadas em cada dia do mês. Você pode selecionar outro mês abaixo para comparação lado a lado."
                            )
                        
                        # Buscar dados diários da base completa
                        dados_encontrados = False
                        if lab_final_cnpj and 'Dados_Diarios_2025' in df.columns:
                            lab_dados = df[df['CNPJ_Normalizado'] == lab_final_cnpj]
                            if not lab_dados.empty:
                                dados_diarios_raw = lab_dados.iloc[0].get('Dados_Diarios_2025', '{}')
                                try:
                                    import json  # Import explícito para garantir disponibilidade
                                    # Parse dos dados JSON
                                    dados_diarios = json.loads(dados_diarios_raw) if isinstance(dados_diarios_raw, str) else dados_diarios_raw
                                    
                                    # Obter mês atual
                                    hoje = datetime.now()
                                    mes_atual_num = hoje.month
                                    ano_atual = hoje.year
                                    mes_atual_key = f"{ano_atual}-{mes_atual_num:02d}"
                                    
                                    meses_nomes = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
                                                  "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
                                    mes_atual_nome = meses_nomes[mes_atual_num - 1]
                                    
                                    # Criar lista de meses disponíveis para comparação
                                    meses_disponiveis = []
                                    meses_opcoes_display = ["Nenhum (apenas mês atual)"]
                                    meses_opcoes_keys = [None]
                                    
                                    # Adicionar meses de 2025 (até o mês atual)
                                    for mes_num in range(1, mes_atual_num + 1):
                                        mes_key = f"2025-{mes_num:02d}"
                                        if mes_key in dados_diarios and dados_diarios[mes_key] and mes_key != mes_atual_key:
                                            meses_opcoes_display.append(f"{meses_nomes[mes_num - 1]}/2025")
                                            meses_opcoes_keys.append(mes_key)
                                    
                                    # Adicionar meses de 2024
                                    for mes_num in range(1, 13):
                                        mes_key = f"2024-{mes_num:02d}"
                                        if mes_key in dados_diarios and dados_diarios[mes_key]:
                                            meses_opcoes_display.append(f"{meses_nomes[mes_num - 1]}/2024")
                                            meses_opcoes_keys.append(mes_key)
                                    
                                    # Seletor de mês para comparação
                                    col_select, col_info = st.columns([3, 1])
                                    with col_select:
                                        if len(meses_opcoes_display) > 1:
                                            mes_comparacao_idx = st.selectbox(
                                                "🔄 Comparar com outro mês:",
                                                range(len(meses_opcoes_display)),
                                                format_func=lambda i: meses_opcoes_display[i],
                                                key="select_mes_comparacao",
                                                help="Selecione um mês para comparar lado a lado com o mês atual. As métricas mostrarão a diferença (delta) entre os períodos."
                                            )
                                            mes_comparacao_key = meses_opcoes_keys[mes_comparacao_idx]
                                        else:
                                            mes_comparacao_key = None
                                            st.info("💡 Nenhum outro mês com dados disponível para comparação")
                                    
                                    # Verificar se há dados para o mês atual
                                    if mes_atual_key in dados_diarios and dados_diarios[mes_atual_key]:
                                        dias_mes_atual = dados_diarios[mes_atual_key]
                                        
                                        # Criar DataFrame do mês atual
                                        df_dias_atual = pd.DataFrame([
                                            {"Dia": int(dia), "Coletas": int(volume)}
                                            for dia, volume in sorted(dias_mes_atual.items(), key=lambda x: int(x[0]))
                                        ])
                                        
                                        if not df_dias_atual.empty:
                                            dados_encontrados = True
                                            
                                            # Se houver mês de comparação selecionado
                                            if mes_comparacao_key:
                                                # Obter dados do mês de comparação
                                                dias_mes_comp = dados_diarios.get(mes_comparacao_key, {})
                                                df_dias_comp = pd.DataFrame([
                                                    {"Dia": int(dia), "Coletas": int(volume)}
                                                    for dia, volume in sorted(dias_mes_comp.items(), key=lambda x: int(x[0]))
                                                ]) if dias_mes_comp else pd.DataFrame()
                                                
                                                # Extrair ano e mês da comparação
                                                ano_comp, mes_comp = mes_comparacao_key.split('-')
                                                mes_comp_nome = meses_nomes[int(mes_comp) - 1]
                                                
                                                # Exibir dois gráficos lado a lado
                                                col1, col2 = st.columns(2)
                                                
                                                with col1:
                                                    # Gráfico do mês atual
                                                    fig_atual = px.bar(
                                                        df_dias_atual,
                                                        x="Dia",
                                                        y="Coletas",
                                                        text="Coletas",
                                                        title=f"{mes_atual_nome}/{ano_atual}",
                                                        color="Coletas",
                                                        color_continuous_scale="Blues"
                                                    )
                                                    fig_atual.update_traces(textposition='outside', textfont_size=10)
                                                    fig_atual.update_layout(
                                                        height=360,
                                                        margin=dict(l=10, r=10, t=40, b=10),
                                                        xaxis_title="Dia",
                                                        yaxis_title="Coletas",
                                                        showlegend=False
                                                    )
                                                    fig_atual.update_xaxes(dtick=1)
                                                    st.plotly_chart(fig_atual, use_container_width=True, key="graf_atual")
                                                
                                                with col2:
                                                    if not df_dias_comp.empty:
                                                        # Gráfico do mês de comparação
                                                        fig_comp = px.bar(
                                                            df_dias_comp,
                                                            x="Dia",
                                                            y="Coletas",
                                                            text="Coletas",
                                                            title=f"{mes_comp_nome}/{ano_comp}",
                                                            color="Coletas",
                                                            color_continuous_scale="Greens"
                                                        )
                                                        fig_comp.update_traces(textposition='outside', textfont_size=10)
                                                        fig_comp.update_layout(
                                                            height=360,
                                                            margin=dict(l=10, r=10, t=40, b=10),
                                                            xaxis_title="Dia",
                                                            yaxis_title="Coletas",
                                                            showlegend=False
                                                        )
                                                        fig_comp.update_xaxes(dtick=1)
                                                        st.plotly_chart(fig_comp, use_container_width=True, key="graf_comparacao")
                                                    else:
                                                        st.warning("Sem dados para o mês selecionado")
                                                
                                                # Métricas comparativas
                                                st.markdown("### 📊 Comparação de Métricas")
                                                st.caption(f"Comparando {mes_atual_nome}/{ano_atual} (azul) vs {mes_comp_nome}/{ano_comp} (verde) • Valores em vermelho ↓ indicam redução | Valores em verde ↑ indicam aumento")
                                                col1, col2, col3, col4 = st.columns(4)
                                                
                                                total_atual = int(df_dias_atual['Coletas'].sum())
                                                media_atual = float(df_dias_atual['Coletas'].mean())
                                                max_atual = int(df_dias_atual['Coletas'].max())
                                                
                                                if not df_dias_comp.empty:
                                                    total_comp = int(df_dias_comp['Coletas'].sum())
                                                    media_comp = float(df_dias_comp['Coletas'].mean())
                                                    max_comp = int(df_dias_comp['Coletas'].max())
                                                    
                                                    diff_total = total_atual - total_comp
                                                    diff_total_pct = (diff_total / total_comp * 100) if total_comp > 0 else 0
                                                    diff_media = media_atual - media_comp
                                                    diff_media_pct = (diff_media / media_comp * 100) if media_comp > 0 else 0
                                                    
                                                    with col1:
                                                        st.metric(
                                                            "Total no Mês",
                                                            f"{total_atual:,}",
                                                            delta=f"{diff_total:+,} ({diff_total_pct:+.1f}%)",
                                                            delta_color="normal",
                                                            help=f"Soma de todas as coletas do mês. Delta mostra diferença vs {mes_comp_nome}/{ano_comp}: {total_atual:,} - {total_comp:,} = {diff_total:+,}"
                                                        )
                                                    with col2:
                                                        st.metric(
                                                            "Média Diária",
                                                            f"{media_atual:.1f}",
                                                            delta=f"{diff_media:+.1f} ({diff_media_pct:+.1f}%)",
                                                            delta_color="normal",
                                                            help=f"Média de coletas por dia no mês. Delta vs {mes_comp_nome}/{ano_comp}: {media_atual:.1f} - {media_comp:.1f} = {diff_media:+.1f}"
                                                        )
                                                    with col3:
                                                        st.metric(
                                                            f"Máximo {mes_atual_nome}",
                                                            f"{max_atual:,}",
                                                            help=f"Dia {df_dias_atual.loc[df_dias_atual['Coletas'].idxmax(), 'Dia']}"
                                                        )
                                                    with col4:
                                                        st.metric(
                                                            f"Máximo {mes_comp_nome}",
                                                            f"{max_comp:,}",
                                                            help=f"Dia {df_dias_comp.loc[df_dias_comp['Coletas'].idxmax(), 'Dia']}"
                                                        )
                                                else:
                                                    with col1:
                                                        st.metric("Total Mês Atual", f"{total_atual:,}")
                                                    with col2:
                                                        st.metric("Média Diária", f"{media_atual:.1f}")
                                                    with col3:
                                                        st.metric("Dia com Mais Coletas", f"Dia {df_dias_atual.loc[df_dias_atual['Coletas'].idxmax(), 'Dia']}")
                                                    with col4:
                                                        st.metric("Máximo em um Dia", f"{max_atual:,}")
                                            
                                            else:
                                                # Apenas mês atual (sem comparação)
                                                fig_dias = px.bar(
                                                    df_dias_atual,
                                                    x="Dia",
                                                    y="Coletas",
                                                    text="Coletas",
                                                    title=f"Coletas por dia - {mes_atual_nome}/{ano_atual}",
                                                    color="Coletas",
                                                    color_continuous_scale="Blues"
                                                )
                                                fig_dias.update_traces(textposition='outside', textfont_size=10)
                                                fig_dias.update_layout(
                                                    height=360,
                                                    margin=dict(l=10, r=10, t=60, b=10),
                                                    xaxis_title="Dia do Mês",
                                                    yaxis_title="Número de Coletas",
                                                    showlegend=False
                                                )
                                                fig_dias.update_xaxes(dtick=1)
                                                st.plotly_chart(fig_dias, use_container_width=True, key="graf_dias_mes_detalhe")
                                                
                                                # Métricas (sem comparação)
                                                col1, col2, col3, col4 = st.columns(4)
                                                with col1:
                                                    st.metric("Total no Mês", f"{df_dias_atual['Coletas'].sum():,}")
                                                with col2:
                                                    st.metric("Média Diária", f"{df_dias_atual['Coletas'].mean():.1f}")
                                                with col3:
                                                    st.metric("Dia com Mais Coletas", f"Dia {df_dias_atual.loc[df_dias_atual['Coletas'].idxmax(), 'Dia']}")
                                                with col4:
                                                    st.metric("Máximo em um Dia", f"{df_dias_atual['Coletas'].max():,}")
                                
                                except Exception as e:
                                    import traceback
                                    st.warning(f"⚠️ Erro ao processar dados diários: {e}")
                                    with st.expander("🔍 Detalhes do erro (debug)"):
                                        st.code(f"Tipo de dados: {type(dados_diarios_raw)}\nConteúdo (primeiros 200 chars): {str(dados_diarios_raw)[:200]}\nErro completo: {traceback.format_exc()}")
                        
                        # Mensagem caso não encontre dados
                        if not dados_encontrados:
                            st.info("📊 Sem dados diários para o mês atual deste laboratório.")

                        # Evolução semanal do mês (visão do laboratório)
                        hoje_ref = datetime.now()
                        try:
                            mes_ref_num = met.get('referencia', {}).get('mes', hoje_ref.month)  # usa mesmo mês da visão geral
                            ano_ref = met.get('referencia', {}).get('ano', hoje_ref.year)
                        except Exception:
                            mes_ref_num = hoje_ref.month
                            ano_ref = hoje_ref.year
                        meses_nomes_completos = [
                            "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
                            "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
                        ]
                        mes_ref_nome = meses_nomes_completos[mes_ref_num - 1]
                        st.subheader(f"📆 Evolução do Mês (Semana a Semana) - {mes_ref_nome}/{ano_ref}")

                        semanas_raw = []
                        df_semana_ref = df_para_metricas.copy()
                        if 'CNPJ_Normalizado' not in df_semana_ref.columns and 'CNPJ_PCL' in df_semana_ref.columns:
                            df_semana_ref['CNPJ_Normalizado'] = df_semana_ref['CNPJ_PCL'].apply(
                                lambda x: ''.join(filter(str.isdigit, str(x))) if pd.notna(x) else ''
                            )
                        if lab_final_cnpj:
                            lab_semana = df_semana_ref[df_semana_ref['CNPJ_Normalizado'] == lab_final_cnpj]
                        else:
                            lab_semana = df_semana_ref[df_semana_ref['Nome_Fantasia_PCL'] == lab_final]

                        fallback_vol_semana_anterior = 0
                        if not lab_semana.empty:
                            # Usa a mesma base do fechamento semanal para preencher a primeira semana quando faltarem dados anteriores
                            fallback_vol_semana_anterior = pd.to_numeric(
                                lab_semana.iloc[0].get('WoW_Semana_Anterior', 0), errors='coerce'
                            )
                            if pd.isna(fallback_vol_semana_anterior):
                                fallback_vol_semana_anterior = 0

                        if not lab_semana.empty and 'Semanas_Mes_Atual' in lab_semana.columns:
                            semanas_val = lab_semana.iloc[0].get('Semanas_Mes_Atual', [])
                            if isinstance(semanas_val, str) and semanas_val and semanas_val != '[]':
                                try:
                                    import json
                                    semanas_raw = json.loads(semanas_val)
                                except Exception:
                                    semanas_raw = []
                            elif isinstance(semanas_val, list):
                                semanas_raw = semanas_val

                        if semanas_raw:
                            semanas_indexadas = {}
                            for semana in semanas_raw:
                                iso_week = semana.get('iso_week')
                                iso_year = semana.get('iso_year')
                                if not iso_week or not iso_year:
                                    continue
                                key = (iso_year, iso_week)
                                entrada = semanas_indexadas.get(key, {
                                    'semana': semana.get('semana', iso_week),
                                    'iso_week': iso_week,
                                    'iso_year': iso_year,
                                    'volume_total': 0,
                                    'volume_semana_anterior': 0,
                                    'fechada': False
                                })
                                entrada['volume_total'] += semana.get('volume_util', 0) or 0
                                entrada['volume_semana_anterior'] += semana.get('volume_semana_anterior', 0) or 0
                                entrada['fechada'] = entrada['fechada'] or semana.get('fechada', False)
                                semanas_indexadas[key] = entrada

                            # Garantir semanas do mês corrente mesmo quando zeradas
                            primeiro_dia = datetime(ano_ref, mes_ref_num, 1)
                            ultimo_dia = datetime(ano_ref, mes_ref_num, calendar.monthrange(ano_ref, mes_ref_num)[1])
                            dia_corrente = primeiro_dia
                            expected_keys = set()
                            while dia_corrente <= ultimo_dia:
                                iso_info = dia_corrente.isocalendar()
                                expected_keys.add((iso_info.year, iso_info.week))
                                dia_corrente += timedelta(days=1)

                            for iso_year, iso_week in expected_keys:
                                if (iso_year, iso_week) not in semanas_indexadas:
                                    fim_semana = datetime.fromisocalendar(iso_year, iso_week, 7)
                                    semanas_indexadas[(iso_year, iso_week)] = {
                                        'semana': iso_week,
                                        'iso_week': iso_week,
                                        'iso_year': iso_year,
                                        'volume_total': 0,
                                        'volume_semana_anterior': 0,
                                        'fechada': fim_semana.date() < hoje_ref.date()
                                    }

                            dados_semanas = []
                            prev_vol = fallback_vol_semana_anterior or 0
                            for iso_year, iso_week in sorted(semanas_indexadas.keys()):
                                info = semanas_indexadas[(iso_year, iso_week)]
                                vol_atual = info.get('volume_total', 0) or 0
                                vol_ant_data = info.get('volume_semana_anterior', 0) or 0

                                prev_label = "-"
                                try:
                                    prev_monday = datetime.fromisocalendar(iso_year, iso_week, 1) - timedelta(days=7)
                                    prev_iso = prev_monday.isocalendar()
                                    prev_key = (prev_iso.year, prev_iso.week)
                                    prev_label = f"{prev_iso.week}/{prev_iso.year}"
                                except Exception:
                                    prev_key = None

                                # Se não temos volume anterior na base, usar volume da semana anterior calculado
                                vol_ant_calc = prev_vol if prev_vol else 0
                                vol_ant = vol_ant_data or vol_ant_calc

                                if vol_ant == 0 and prev_key and prev_key in semanas_indexadas:
                                    vol_ant = semanas_indexadas[prev_key].get('volume_total', 0) or vol_ant

                                if vol_ant == 0:
                                    variacao_pct = 100.0 if vol_atual > 0 else 0.0
                                else:
                                    variacao_pct = ((vol_atual - vol_ant) / vol_ant) * 100

                                try:
                                    semana_inicio = datetime.fromisocalendar(iso_year, iso_week, 1)
                                    semana_fim = semana_inicio + timedelta(days=6)
                                    intervalo_str = f"{semana_inicio:%d/%m}-{semana_fim:%d/%m}"
                                except Exception:
                                    intervalo_str = "-"

                                dados_semanas.append({
                                    "Semana": f"Semana {info.get('semana', iso_week)}",
                                    "ISO": iso_week,
                                    "Intervalo (seg-dom)": intervalo_str,
                                    "Semana Anterior (ISO)": prev_label,
                                    "Volume Útil": round(float(vol_atual), 1),
                                    "Volume Anterior": round(float(vol_ant), 1),
                                    "WoW %": round(variacao_pct, 2),
                                    "Status": "✅ Fechada" if info.get('fechada') else "⏳ Em Andamento",
                                    "is_active": not info.get('fechada')
                                })

                                prev_vol = vol_atual

                            df_semanas_lab = pd.DataFrame(dados_semanas)

                            def _highlight_active_week(row):
                                if row['is_active']:
                                    return ['background-color: #fff3cd; color: #856404'] * len(row)
                                return [''] * len(row)

                            st.dataframe(
                                df_semanas_lab.style.apply(_highlight_active_week, axis=1),
                                use_container_width=True,
                                hide_index=True,
                                column_config={
                                    "Semana": st.column_config.TextColumn("Semana", width="small", help="Semana na visão do mês corrente (1, 2, 3...)."),
                                    "ISO": st.column_config.NumberColumn("ISO Week", format="%d", help="Semana ISO correspondente."),
                                    "Intervalo (seg-dom)": st.column_config.TextColumn("Dias da Semana", help="Intervalo (segunda a domingo) que compõe a semana ISO."),
                                    "Semana Anterior (ISO)": st.column_config.TextColumn("Semana ISO anterior", help="ISO week imediatamente anterior (pode ser do mês/ano anterior)."),
                                    "Volume Útil": st.column_config.NumberColumn("Volume Realizado", format="%d"),
                                    "Volume Anterior": st.column_config.NumberColumn("Volume Anterior", format="%d"),
                                    "WoW %": st.column_config.NumberColumn("Variação WoW (%)", format="%.1f%%"),
                                    "Status": st.column_config.TextColumn("Status"),
                                    "is_active": None
                                }
                            )
                        else:
                            st.info("📆 Sem dados semanais para o mês atual deste laboratório.")

                        st.subheader("📈 Evolução Mensal")
                        ChartManager.criar_grafico_evolucao_mensal(
                            df_para_metricas,
                            lab_cnpj=lab_final_cnpj,
                            lab_nome=lab_final,
                            chart_key="historico"
                        )
                    st.markdown("</div>", unsafe_allow_html=True)
                    
                    # Seção de Gráficos com Abas - Refatorado conforme solicitação
                    st.markdown("""
                        <div style="background: white; border-radius: 8px; padding: 1.5rem; margin-bottom: 2rem;
                                    border: 1px solid #e9ecef; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                            <h3 style="margin: 0 0 1rem 0; color: #2c3e50; font-weight: 600; border-bottom: 2px solid #6BBF47; padding-bottom: 0.5rem;">
                                📊 Análise Visual Detalhada
                            </h3>
                        """, unsafe_allow_html=True)
                    
                    # Criar abas para organizar os gráficos (Resumo Executivo removido)
                    tab_distribuicao, tab_media_diaria, tab_coletas_dia = st.tabs([
                        "📊 Distribuição por Dia Útil", "📅 Média Diária", "📈 Coletas por Dia Útil"
                    ])
                    
                    with tab_distribuicao:
                        st.subheader("📊 Distribuição de Coletas por Dia Útil da Semana")
                        # Gráfico com destaque maior conforme solicitado
                        ChartManager.criar_grafico_media_dia_semana_novo(
                            df_para_metricas,
                            lab_cnpj=lab_final_cnpj,
                            lab_nome=lab_final,
                            filtros=filtros
                        )
                    
                    with tab_media_diaria:
                        st.subheader("📊 Média Diária por Mês")
                        ChartManager.criar_grafico_media_diaria(
                            df_para_metricas,
                            lab_cnpj=lab_final_cnpj,
                            lab_nome=lab_final
                        )

                    with tab_coletas_dia:
                        st.subheader("📈 Coletas por Dia Útil do Mês")
                        ChartManager.criar_grafico_coletas_por_dia(
                            df_para_metricas,
                            lab_cnpj=lab_final_cnpj,
                            lab_nome=lab_final
                        )

        # Conteúdo único da análise detalhada
        # Carregar dados VIP para análise de rede
        df_vip_tabela = DataManager.carregar_dados_vip()
        # Adicionar informações de rede se disponível
        df_tabela = df_analise_detalhada.copy()
        mostrar_rede = False
        # Garantir que CNPJ_Normalizado existe em df_tabela
        if 'CNPJ_Normalizado' not in df_tabela.columns and 'CNPJ_PCL' in df_tabela.columns:
            df_tabela['CNPJ_Normalizado'] = df_tabela['CNPJ_PCL'].apply(
                lambda x: ''.join(filter(str.isdigit, str(x))) if pd.notna(x) else ''
            )
        # Se há um laboratório pesquisado que está fora dos filtros, adicioná-lo à tabela
        lab_fora_filtros = st.session_state.get('lab_fora_filtros', False)
        if lab_fora_filtros and lab_final_cnpj and 'CNPJ_Normalizado' in df_tabela.columns:
            # Verificar se o laboratório não está na tabela
            if lab_final_cnpj not in df_tabela['CNPJ_Normalizado'].values:
                # Buscar o laboratório na base completa
                lab_na_base = df[df['CNPJ_Normalizado'] == lab_final_cnpj]
                if not lab_na_base.empty:
                    # Adicionar o laboratório à tabela
                    df_tabela = pd.concat([df_tabela, lab_na_base], ignore_index=True)
        if df_vip_tabela is not None and not df_vip_tabela.empty:
            df_vip_tabela = df_vip_tabela.copy()
            df_vip_tabela['CNPJ_Normalizado'] = df_vip_tabela['CNPJ'].apply(
                lambda x: ''.join(filter(str.isdigit, str(x))) if pd.notna(x) else ''
            )
            colunas_vip_disponiveis = ['CNPJ_Normalizado']
            for col in ['Rede', 'Ranking', 'Ranking Rede']:
                if col in df_vip_tabela.columns:
                    colunas_vip_disponiveis.append(col)
            if len(colunas_vip_disponiveis) > 1:
                df_tabela = df_tabela.merge(
                    df_vip_tabela[colunas_vip_disponiveis],
                    on='CNPJ_Normalizado',
                    how='left'
                )
                mostrar_rede = 'Rede' in df_tabela.columns

        # Determinar rede do lab selecionado (quando existir)
        rede_padrao = st.session_state.get('rede_lab_pesquisado')
        if not rede_padrao and lab_final_cnpj and 'Rede' in df_tabela.columns:
            rede_lab_row = df_tabela[df_tabela['CNPJ_Normalizado'] == lab_final_cnpj]
            if not rede_lab_row.empty:
                rede_padrao = rede_lab_row.iloc[0].get('Rede')
                st.session_state['rede_lab_pesquisado'] = rede_padrao

        if mostrar_rede and 'Rede' in df_tabela.columns:
            redes_disponiveis = ["Todas"] + sorted(df_tabela['Rede'].dropna().unique().tolist())
            if rede_padrao and rede_padrao in redes_disponiveis:
                rede_filtro = rede_padrao
                st.markdown(f"""
                <div style="background: linear-gradient(135deg, #e8f5e9, #f1f8f1); border-radius: 6px; padding: 0.8rem; margin-bottom: 1rem;">
                    <span style="color: #52B54B; font-size: 0.9rem;">🎯 <strong>Filtro automático ativo:</strong> mostrando apenas laboratórios da rede <strong>"{rede_filtro}"</strong></span>
                </div>
                """, unsafe_allow_html=True)
            else:
                rede_filtro = st.selectbox(
                    "🏢 Filtrar por Rede:",
                    options=redes_disponiveis,
                    index=0,
                    help="Selecione uma rede para filtrar",
                    key="filtro_rede_tabela"
                )
        else:
            rede_filtro = "Todas"

        df_tabela_filtrada = df_tabela.copy()
        # Filtro por rede ou apenas o lab atual
        if rede_filtro != "Todas" and mostrar_rede:
            df_tabela_filtrada = df_tabela_filtrada[df_tabela_filtrada['Rede'] == rede_filtro]
        elif not mostrar_rede and lab_final_cnpj:
            df_tabela_filtrada = df_tabela_filtrada[df_tabela_filtrada['CNPJ_Normalizado'] == lab_final_cnpj]
        # Mostrar informações da rede se filtrada
        if rede_filtro != "Todas" and mostrar_rede and not df_tabela_filtrada.empty:
            # Estatísticas da rede
            stats_rede = {
                'total_labs': len(df_tabela_filtrada),
                'volume_total': df_tabela_filtrada['Volume_Total_2025'].sum() if 'Volume_Total_2025' in df_tabela_filtrada.columns else 0,
                'media_volume': df_tabela_filtrada['Volume_Total_2025'].mean() if 'Volume_Total_2025' in df_tabela_filtrada.columns else 0,
                'labs_risco_alto': (
                    df_tabela_filtrada['Risco_Diario'].isin(['🔴 Alto', '⚫ Crítico']).sum()
                    if 'Risco_Diario' in df_tabela_filtrada.columns else 0
                ),
                'labs_ativos': len(df_tabela_filtrada[df_tabela_filtrada['Dias_Sem_Coleta'] <= 30]) if 'Dias_Sem_Coleta' in df_tabela_filtrada.columns else 0
            }
            st.markdown(f"""
            <div style="background: linear-gradient(135deg, #e8f5e9, #f1f8f1); border-radius: 8px; padding: 1rem; margin-bottom: 1rem;">
                <h4 style="margin: 0 0 0.5rem 0; color: #52B54B;">📊 Estatísticas da Rede: {rede_filtro}</h4>
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 1rem;">
                    <div style="text-align: center;">
                        <div style="font-size: 1.5rem; font-weight: bold; color: #6BBF47;">{stats_rede['total_labs']}</div>
                        <div style="font-size: 0.8rem; color: #666;">Laboratórios</div>
                    </div>
                    <div style="text-align: center;">
                        <div style="font-size: 1.5rem; font-weight: bold; color: #6BBF47;">{stats_rede['volume_total']:,.0f}</div>
                        <div style="font-size: 0.8rem; color: #666;">Volume Total</div>
                    </div>
                    <div style="text-align: center;">
                        <div style="font-size: 1.5rem; font-weight: bold; color: #6BBF47;">{stats_rede['media_volume']:.0f}</div>
                        <div style="font-size: 0.8rem; color: #666;">Média por Lab</div>
                    </div>
                    <div style="text-align: center;">
                        <div style="font-size: 1.5rem; font-weight: bold; color: #f44336;">{stats_rede['labs_risco_alto']}</div>
                        <div style="font-size: 0.8rem; color: #666;">Alto Risco</div>
                    </div>
                    <div style="text-align: center;">
                        <div style="font-size: 1.5rem; font-weight: bold; color: #4caf50;">{stats_rede['labs_ativos']}</div>
                        <div style="font-size: 0.8rem; color: #666;">Ativos (30d)</div>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)

        # Configurar colunas da tabela
        # Se faz parte de rede: mostrar visão mensal por coluna (janeiro, fevereiro...)
        # Sempre mostrar planilhão de labs (rede ou não)
        colunas_principais = [
            'CNPJ_PCL', 'Nome_Fantasia_PCL', 'Estado', 'Cidade', 'Representante_Nome',
            'Dias_Sem_Coleta'
        ]

        # Adicionar colunas de coletas mensais (2024 e 2025)
        meses_nomes = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
        # Mapeamento dos códigos dos meses para nomes completos em português
        meses_nomes_completos = {
            "Jan": "Janeiro", "Fev": "Fevereiro", "Mar": "Março", "Abr": "Abril",
            "Mai": "Maio", "Jun": "Junho", "Jul": "Julho", "Ago": "Agosto",
            "Set": "Setembro", "Out": "Outubro", "Nov": "Novembro", "Dez": "Dezembro"
        }
        mes_limite_2025 = min(datetime.now().month, 12)
        
        # Colunas de 2024 (todos os meses)
        cols_2024 = [f'N_Coletas_{m}_24' for m in meses_nomes]
        # Colunas de 2025 (até o mês atual)
        cols_2025 = [f'N_Coletas_{m}_25' for m in meses_nomes[:mes_limite_2025]]

        # Calcular métricas de mês atual e picos para todos
        df_tabela_filtrada = df_tabela_filtrada.copy()
        mes_atual_codigo = meses_nomes[datetime.now().month - 1]
        col_mes_atual = f'N_Coletas_{mes_atual_codigo}_25'
        if col_mes_atual in df_tabela_filtrada.columns:
            df_tabela_filtrada['Volume_Mes_Vigente'] = pd.to_numeric(df_tabela_filtrada[col_mes_atual], errors='coerce').fillna(0)
        else:
            df_tabela_filtrada['Volume_Mes_Vigente'] = 0

        if cols_2024:
            df_tabela_filtrada['Maior_Mes_2024_Valor'] = df_tabela_filtrada[cols_2024].max(axis=1, skipna=True).fillna(0)
            df_tabela_filtrada['Maior_Mes_2024_Nome'] = df_tabela_filtrada[cols_2024].idxmax(axis=1).apply(
                lambda x: meses_nomes_completos.get(x.split('_')[2], x.split('_')[2]) if pd.notna(x) and '_' in str(x) else 'N/A'
            )
        else:
            df_tabela_filtrada['Maior_Mes_2024_Valor'] = 0
            df_tabela_filtrada['Maior_Mes_2024_Nome'] = 'N/A'

        if cols_2025:
            df_tabela_filtrada['Maior_Mes_2025_Valor'] = df_tabela_filtrada[cols_2025].max(axis=1, skipna=True).fillna(0)
            df_tabela_filtrada['Maior_Mes_2025_Nome'] = df_tabela_filtrada[cols_2025].idxmax(axis=1).apply(
                lambda x: meses_nomes_completos.get(x.split('_')[2], x.split('_')[2]) if pd.notna(x) and '_' in str(x) else 'N/A'
            )
        else:
            df_tabela_filtrada['Maior_Mes_2025_Valor'] = 0
            df_tabela_filtrada['Maior_Mes_2025_Nome'] = 'N/A'

        total_mes_vigente = int(df_tabela_filtrada['Volume_Mes_Vigente'].sum())

        colunas_principais.extend(['Volume_Mes_Vigente', 'Maior_Mes_2024_Valor', 'Maior_Mes_2024_Nome',
                                  'Maior_Mes_2025_Valor', 'Maior_Mes_2025_Nome'])
        colunas_principais.extend(cols_2024 + cols_2025)
        if mostrar_rede:
            colunas_principais.extend(['Rede', 'Ranking', 'Ranking Rede'])

        # Se não houver rede, mostrar apenas o laboratório selecionado
        if not mostrar_rede and lab_final_cnpj:
            df_tabela_filtrada = df_tabela_filtrada[df_tabela_filtrada['CNPJ_Normalizado'] == lab_final_cnpj]

        if mostrar_rede and rede_filtro != "Todas":
            total_labs_rede = len(df_tabela_filtrada)
            volume_total_rede = df_tabela_filtrada[cols_2025 + cols_2024].sum().sum() if not df_tabela_filtrada.empty else 0
            st.markdown(f"""
            <div style="background: #e8f5e9; border-radius: 8px; padding: 1rem; margin-bottom: 1rem;">
                <h4 style="margin: 0 0 0.5rem 0; color: #52B54B;">📊 Métricas da Rede: {rede_filtro}</h4>
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 1rem;">
                    <div style="text-align: center;">
                        <div style="font-size: 1.2rem; font-weight: bold; color: #6BBF47;">{total_labs_rede}</div>
                        <div style="font-size: 0.8rem; color: #666;">Laboratórios da rede</div>
                    </div>
                    <div style="text-align: center;">
                        <div style="font-size: 1.2rem; font-weight: bold; color: #6BBF47;">{total_mes_vigente:,.0f}</div>
                        <div style="font-size: 0.8rem; color: #666;">Total coletas mês vigente</div>
                    </div>
                    <div style="text-align: center;">
                        <div style="font-size: 1.2rem; font-weight: bold; color: #6BBF47;">{volume_total_rede:,.0f}</div>
                        <div style="font-size: 0.8rem; color: #666;">Total 2024-2025 (rede)</div>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)
        
        colunas_existentes = [col for col in colunas_principais if col in df_tabela_filtrada.columns]
        if not df_tabela_filtrada.empty and colunas_existentes:
            df_exibicao = df_tabela_filtrada[colunas_existentes].copy()
            # Formatação de colunas
            # Criar configuração de colunas de forma mais explícita (sem risco diário, sem médias móveis)
            column_config = {
                "CNPJ_PCL": st.column_config.TextColumn(
                    "📄 CNPJ",
                    help="CNPJ (Cadastro Nacional de Pessoa Jurídica) do laboratório. Identificador único para busca e identificação"
                ),
                "Nome_Fantasia_PCL": st.column_config.TextColumn(
                    "🏥 Nome Fantasia",
                    help="Nome comercial/fantasia do laboratório. Use para busca rápida por nome"
                ),
                "Estado": st.column_config.TextColumn(
                    "🗺️ Estado",
                    help="Estado (UF) onde o laboratório está localizado. Permite filtrar e agrupar por região geográfica"
                ),
                "Cidade": st.column_config.TextColumn(
                    "🏙️ Cidade",
                    help="Cidade onde o laboratório está localizado. Permite análise mais granular por localização"
                ),
                "Representante_Nome": st.column_config.TextColumn(
                    "👤 Representante",
                    help="Nome do representante comercial responsável pelo laboratório. Útil para contato direto e gestão de relacionamento"
                ),
                "Dias_Sem_Coleta": st.column_config.NumberColumn(
                    "Dias Sem Coleta",
                    help="Número consecutivo de dias sem registrar coletas. Valores altos indicam possível inatividade do laboratório"
                ),
                "Volume_Mes_Vigente": st.column_config.NumberColumn(
                    "Volume Mês Vigente",
                    format="%d",
                    help="Volume de coletas do mês atual (vigente)"
                ),
                "Maior_Mes_2024_Valor": st.column_config.NumberColumn(
                    "Maior Mês 2024 (Valor)",
                    format="%d",
                    help="Maior volume mensal de coletas em 2024"
                ),
                "Maior_Mes_2024_Nome": st.column_config.TextColumn(
                    "Maior Mês 2024 (Mês)",
                    help="Mês em que ocorreu o maior volume de coletas em 2024"
                ),
                "Maior_Mes_2025_Valor": st.column_config.NumberColumn(
                    "Maior Mês 2025 (Valor)",
                    format="%d",
                    help="Maior volume mensal de coletas em 2025"
                ),
                "Maior_Mes_2025_Nome": st.column_config.TextColumn(
                    "Maior Mês 2025 (Mês)",
                    help="Mês em que ocorreu o maior volume de coletas em 2025"
                )
            }
            
            # Adicionar configurações para colunas mensais de 2024
            for col in cols_2024:
                if col in df_exibicao.columns:
                    mes_codigo = col.split('_')[2]  # Corrigido: pegar o terceiro elemento (índice 2)
                    mes_nome = meses_nomes_completos.get(mes_codigo, mes_codigo)
                    # Usar configuração mais simples
                    column_config[col] = st.column_config.NumberColumn(
                        f"{mes_nome}/24",
                        help=f"Total de coletas realizadas em {mes_nome} de 2024. Permite análise de sazonalidade e comparação ano a ano"
                    )
            
            # Adicionar configurações para colunas mensais de 2025
            for col in cols_2025:
                if col in df_exibicao.columns:
                    mes_codigo = col.split('_')[2]  # Corrigido: pegar o terceiro elemento (índice 2)
                    mes_nome = meses_nomes_completos.get(mes_codigo, mes_codigo)
                    # Usar configuração mais simples
                    column_config[col] = st.column_config.NumberColumn(
                        f"{mes_nome}/25",
                        help=f"Total de coletas realizadas em {mes_nome} de 2025. Permite acompanhamento do desempenho mensal atual"
                    )
            
            # Adicionar colunas de rede se disponível
            if 'Rede' in df_exibicao.columns:
                column_config["Rede"] = st.column_config.TextColumn(
                    "🏢 Rede",
                    help="Nome da rede à qual o laboratório pertence. Permite agrupar e comparar laboratórios da mesma rede"
                )
            if 'Ranking' in df_exibicao.columns:
                column_config["Ranking"] = st.column_config.TextColumn(
                    "🏆 Ranking",
                    help="Posição do laboratório no ranking geral por volume de coletas. Ranking 1 = maior volume"
                )
            if 'Ranking_Rede' in df_exibicao.columns:
                column_config["Ranking_Rede"] = st.column_config.TextColumn(
                    "🏅 Ranking Rede",
                    help="Posição do laboratório no ranking dentro de sua própria rede. Permite identificar líderes regionais por rede"
                )
            
            # Renomear as colunas diretamente no dataframe para exibir nomes completos dos meses
            df_exibicao_renamed = _formatar_df_exibicao(df_exibicao)
            rename_dict = {}
            
            # Renomear colunas principais para nomes mais legíveis
            rename_dict.update({
                "CNPJ_PCL": "CNPJ",
                "Nome_Fantasia_PCL": "Nome Fantasia",
                "Representante_Nome": "Representante",
                "Dias_Sem_Coleta": "Dias Sem Coleta",
                "Volume_Mes_Vigente": "Volume Mês Vigente",
                "Maior_Mes_2024_Valor": "Maior Mês 2024 (Valor)",
                "Maior_Mes_2024_Nome": "Maior Mês 2024 (Mês)",
                "Maior_Mes_2025_Valor": "Maior Mês 2025 (Valor)",
                "Maior_Mes_2025_Nome": "Maior Mês 2025 (Mês)",
                "Ranking_Rede": "Ranking Rede"
            })
            
            # Renomear colunas de 2024
            for col in cols_2024:
                if col in df_exibicao_renamed.columns:
                    mes_codigo = col.split('_')[2]  # Corrigido: pegar o terceiro elemento (índice 2)
                    mes_nome = meses_nomes_completos.get(mes_codigo, mes_codigo)
                    rename_dict[col] = f"{mes_nome}/24"
            
            # Renomear colunas de 2025
            for col in cols_2025:
                if col in df_exibicao_renamed.columns:
                    mes_codigo = col.split('_')[2]  # Corrigido: pegar o terceiro elemento (índice 2)
                    mes_nome = meses_nomes_completos.get(mes_codigo, mes_codigo)
                    rename_dict[col] = f"{mes_nome}/25"
            
            df_exibicao_renamed = df_exibicao_renamed.rename(columns=rename_dict)
            
            # Atualizar column_config com os nomes renomeados
            column_config_renamed = {}
            for col_original, config in column_config.items():
                if col_original in rename_dict:
                    col_nomeada = rename_dict[col_original]
                    column_config_renamed[col_nomeada] = config
                elif col_original in df_exibicao_renamed.columns:
                    column_config_renamed[col_original] = config
            
            # Mostrar tabela com contador
            st.markdown(f"**Mostrando {len(df_exibicao_renamed)} laboratórios**")
            st.dataframe(
                df_exibicao_renamed,
                width='stretch',
                height=500,
                hide_index=True,
                column_config=column_config_renamed
            )
            
            # Botões de download
            col_download1, col_download2 = st.columns(2)
            with col_download1:
                csv_data = df_exibicao.to_csv(index=False, encoding='utf-8')
                st.download_button(
                    label="📥 Download CSV",
                    data=csv_data,
                    file_name=f"dados_laboratorios_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    key="download_csv_tabela"
                )
            with col_download2:
                excel_buffer = BytesIO()
                df_exibicao.to_excel(excel_buffer, index=False, engine='openpyxl')
                excel_data = excel_buffer.getvalue()
                st.download_button(
                    label="📥 Download Excel",
                    data=excel_data,
                    file_name=f"dados_laboratorios_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_excel_tabela"
                )
        else:
            # Não mostrar mensagem se há um laboratório pesquisado que está fora dos filtros
            if not (st.session_state.get('lab_fora_filtros', False) and lab_final_cnpj):
                st.info("📋 Nenhum laboratório encontrado com os filtros aplicados.")
        
        # ========================================
        # DADOS DO CONCORRENTE GRALAB (FINAL DA PÁGINA)
        # ========================================
        if lab_final_cnpj:  # Só mostrar se houver laboratório selecionado
            st.markdown("---")
            st.markdown("""
            <div style="background: linear-gradient(135deg, #ffd700 0%, #ffed4e 100%);
                        color: #333; padding: 2rem; border-radius: 12px;
                        margin: 2rem 0; box-shadow: 0 6px 12px rgba(0,0,0,0.15);">
                <h2 style="margin: 0; font-size: 1.8rem; color: #b8860b; font-weight: 700;">
                    🏆 Dados no Concorrente Gralab (CunhaLab)
                </h2>
                <p style="margin: 0.5rem 0 0 0; font-size: 1.1rem; color: #666;">
                    Compare os dados deste laboratório com a base do concorrente
                </p>
            </div>
            """, unsafe_allow_html=True)
            
            try:
                loader = show_overlay_loader(
                    "Carregando dados do Gralab...",
                    "Buscando informações do concorrente. Isso pode levar alguns segundos."
                )
                try:
                    dados_gralab = DataManager.carregar_dados_gralab()
                finally:
                    loader.empty()
                
                if dados_gralab and 'Dados Completos' in dados_gralab:
                    df_gralab = dados_gralab['Dados Completos']
                    
                    # Buscar laboratório pelo CNPJ normalizado
                    lab_gralab = df_gralab[df_gralab['CNPJ_Normalizado'] == lab_final_cnpj]
                    
                    if not lab_gralab.empty:
                        lab_g = lab_gralab.iloc[0]
                        
                        # Verificar se está na aba EntradaSaida
                        status_movimentacao = ""
                        if 'EntradaSaida' in dados_gralab:
                            df_entrada_saida = dados_gralab['EntradaSaida']
                            lab_entrada = df_entrada_saida[df_entrada_saida['CNPJ_Normalizado'] == lab_final_cnpj]
                            if not lab_entrada.empty:
                                tipo_mov = lab_entrada.iloc[0].get('Tipo Movimentação', '')
                                status_lab = lab_entrada.iloc[0].get('Status', '')
                                if tipo_mov or status_lab:
                                    status_movimentacao = f"<div style='margin-top: 1rem; padding: 1rem; background: #fff3cd; border-radius: 8px; border-left: 4px solid #ffc107;'>"
                                    status_movimentacao += f"<strong style='font-size: 1.1rem;'>Movimentação:</strong> {tipo_mov} | <strong>Status:</strong> {status_lab}</div>"
                        
                        # Extrair preços
                        preco_cnh = lab_g.get('Preço CNH', 'N/A')
                        preco_concurso = lab_g.get('Preço Concurso', 'N/A')
                        preco_clt = lab_g.get('Preço CLT', 'N/A')
                        
                        # Formatar preços
                        def formatar_preco(preco):
                            try:
                                if pd.notna(preco) and preco != '' and preco != 'N/A':
                                    return f"R$ {float(preco):.2f}"
                                return "N/A"
                            except:
                                return "N/A"
                        
                        preco_cnh_fmt = formatar_preco(preco_cnh)
                        preco_concurso_fmt = formatar_preco(preco_concurso)
                        preco_clt_fmt = formatar_preco(preco_clt)
                        
                        st.markdown(f"""
                        <div style="background: linear-gradient(135deg, #fff8dc 0%, #fffacd 100%); 
                                    border-radius: 12px; padding: 2rem; margin: 1rem 0 2rem 0;
                                    border: 3px solid #ffd700; box-shadow: 0 4px 12px rgba(255,215,0,0.4);">
                            <h3 style="margin: 0 0 1.5rem 0; color: #b8860b; font-weight: 700; font-size: 1.5rem;">
                                ✅ Laboratório Encontrado na Base do Gralab (CunhaLab)
                            </h3>
                            <div style="background: white; border-radius: 10px; padding: 1.5rem; margin-bottom: 1rem; box-shadow: 0 2px 6px rgba(0,0,0,0.1);">
                                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 1.5rem;">
                                    <div>
                                        <div style="font-size: 0.9rem; color: #666; margin-bottom: 0.5rem; font-weight: 600;">NOME</div>
                                        <div style="font-size: 1.2rem; font-weight: 700; color: #2c3e50;">{lab_g.get('Nome', 'N/A')}</div>
                                    </div>
                                    <div>
                                        <div style="font-size: 0.9rem; color: #666; margin-bottom: 0.5rem; font-weight: 600;">CIDADE / UF</div>
                                        <div style="font-size: 1.2rem; font-weight: 700; color: #2c3e50;">{lab_g.get('Cidade', 'N/A')} / {lab_g.get('UF', 'N/A')}</div>
                                    </div>
                                    <div>
                                        <div style="font-size: 0.9rem; color: #666; margin-bottom: 0.5rem; font-weight: 600;">TELEFONE</div>
                                        <div style="font-size: 1.1rem; font-weight: 600; color: #2c3e50;">{lab_g.get('Telefone', 'N/A')}</div>
                                    </div>
                                    <div>
                                        <div style="font-size: 0.9rem; color: #666; margin-bottom: 0.5rem; font-weight: 600;">ENDEREÇO</div>
                                        <div style="font-size: 1rem; color: #2c3e50;">{lab_g.get('Endereco', 'N/A')[:60]}...</div>
                                    </div>
                                </div>
                            </div>
                            <div style="background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%); 
                                        border-radius: 10px; padding: 1.5rem; margin-top: 1rem; 
                                        border-left: 5px solid #2196f3; box-shadow: 0 2px 6px rgba(33,150,243,0.3);">
                                <div style="font-size: 1.1rem; color: #1565c0; margin-bottom: 1rem; font-weight: 700;">💰 PREÇOS PRATICADOS PELO CONCORRENTE</div>
                                <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 1.5rem; text-align: center;">
                                    <div style="background: white; border-radius: 8px; padding: 1rem; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                                        <div style="font-size: 0.9rem; color: #666; margin-bottom: 0.5rem; font-weight: 600;">🎫 CNH</div>
                                        <div style="font-size: 1.5rem; font-weight: bold; color: #2196f3;">{preco_cnh_fmt}</div>
                                    </div>
                                    <div style="background: white; border-radius: 8px; padding: 1rem; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                                        <div style="font-size: 0.9rem; color: #666; margin-bottom: 0.5rem; font-weight: 600;">📝 Concurso</div>
                                        <div style="font-size: 1.5rem; font-weight: bold; color: #2196f3;">{preco_concurso_fmt}</div>
                                    </div>
                                    <div style="background: white; border-radius: 8px; padding: 1rem; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                                        <div style="font-size: 0.9rem; color: #666; margin-bottom: 0.5rem; font-weight: 600;">👔 CLT</div>
                                        <div style="font-size: 1.5rem; font-weight: bold; color: #2196f3;">{preco_clt_fmt}</div>
                                    </div>
                                </div>
                            </div>
                            {status_movimentacao}
                        </div>
                        """, unsafe_allow_html=True)

                        # Comparativo com nossos preços
                        def _to_float_preco(valor):
                            if isinstance(valor, str):
                                valor = valor.replace('R$', '').replace(' ', '').replace('.', '').replace(',', '.')
                            return pd.to_numeric(valor, errors='coerce')

                        def _to_float_local(valor):
                            return pd.to_numeric(valor, errors='coerce')

                        comparacoes_precos = []
                        nosso_preco_clt = _to_float_local(lab_info.get('Preco_CLT_Total'))
                        nosso_preco_cnh = _to_float_local(lab_info.get('Preco_CNH_Total'))
                        nosso_preco_civil = _to_float_local(lab_info.get('Preco_Civil_Service_Total'))
                        if pd.isna(nosso_preco_civil):
                            nosso_preco_civil = _to_float_local(lab_info.get('Preco_Civil_Service50_Total'))

                        comparacoes_dados = [
                            ("CLT", nosso_preco_clt, _to_float_preco(preco_clt)),
                            ("CNH", nosso_preco_cnh, _to_float_preco(preco_cnh)),
                            ("Concurso Público", nosso_preco_civil, _to_float_preco(preco_concurso))
                        ]

                        def _formatar_delta(valor):
                            if pd.isna(valor):
                                return "—"
                            sinal = "+" if valor > 0 else ""
                            return f"{sinal}R$ {abs(valor):.2f}".replace('.', ',')

                        for label, nosso_val, conc_val in comparacoes_dados:
                            if pd.isna(nosso_val) and pd.isna(conc_val):
                                continue
                            delta_val = nosso_val - conc_val if pd.notna(nosso_val) and pd.notna(conc_val) else np.nan
                            cor_delta = '#6c757d'
                            if pd.notna(delta_val):
                                if delta_val > 0:
                                    cor_delta = '#dc3545'
                                elif delta_val < 0:
                                    cor_delta = '#198754'
                            comparacoes_precos.append(
                                f"<div style=\"background: white; border-radius: 10px; padding: 1rem; box-shadow: 0 2px 6px rgba(0,0,0,0.08);\">"
                                f"<div style=\"font-size: 0.9rem; color: #6c757d; font-weight: 700; text-transform: uppercase; margin-bottom: 0.6rem;\">{label}</div>"
                                f"<div style=\"display: flex; justify-content: space-between; font-size: 0.85rem; color: #6c757d; margin-bottom: 0.3rem;\">"
                                f"<span>Nosso preço</span>"
                                f"<strong style=\"color: #2c3e50;\">{_formatar_preco_valor(nosso_val)}</strong>"
                                f"</div>"
                                f"<div style=\"display: flex; justify-content: space-between; font-size: 0.85rem; color: #6c757d; margin-bottom: 0.3rem;\">"
                                f"<span>Concorrente</span>"
                                f"<strong style=\"color: #2c3e50;\">{_formatar_preco_valor(conc_val)}</strong>"
                                f"</div>"
                                f"<div style=\"display: flex; justify-content: space-between; font-size: 0.85rem; color: #6c757d;\">"
                                f"<span>Diferença</span>"
                                f"<strong style=\"color: {cor_delta};\">{_formatar_delta(delta_val)}</strong>"
                                f"</div>"
                                f"</div>"
                            )

                        if comparacoes_precos:
                            st.markdown(f"""
                                <div style="background: #eef2ff; border-radius: 10px; padding: 1rem 1.5rem; margin-top: 1rem; border-left: 5px solid #6366f1;">
                                    <div style="font-size: 0.95rem; color: #4338ca; font-weight: 700; margin-bottom: 0.8rem; text-transform: uppercase;">
                                        Comparativo de Preços
                                    </div>
                                    <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 1rem;">
                                        {''.join(comparacoes_precos)}
                                    </div>
                                </div>
                            """, unsafe_allow_html=True)

                    else:
                        st.info("ℹ️ Este laboratório não está cadastrado na base do Gralab (CunhaLab)")
                else:
                    st.warning("⚠️ Não foi possível carregar dados do Gralab (CunhaLab)")
            except Exception as e:
                st.error(f"❌ Erro ao carregar dados do Gralab (CunhaLab): {e}")

        # Fechar container principal
        st.markdown("</div>", unsafe_allow_html=True)
    elif st.session_state.page == "🏢 Ranking Rede":
        st.header("🏢 Ranking por Rede")
        # Carregar dados VIP para análise de rede
        df_vip = DataManager.carregar_dados_vip()
        if df_vip is not None and not df_vip.empty:
            # Merge dos dados principais com dados VIP
            df_com_rede = df_filtrado.copy()
            # Adicionar coluna CNPJ normalizado para match
            df_com_rede['CNPJ_Normalizado'] = df_com_rede['CNPJ_PCL'].apply(
                lambda x: ''.join(filter(str.isdigit, str(x))) if pd.notna(x) else ''
            )
            df_vip['CNPJ_Normalizado'] = df_vip['CNPJ'].apply(
                lambda x: ''.join(filter(str.isdigit, str(x))) if pd.notna(x) else ''
            )
            # Merge dos dados
            df_com_rede = df_com_rede.merge(
                df_vip[['CNPJ_Normalizado', 'Rede', 'Ranking', 'Ranking Rede']],
                on='CNPJ_Normalizado',
                how='left'
            )
            # Filtros específicos para ranking de rede
            st.markdown("""
            <div style="background: linear-gradient(135deg, #6BBF47 0%, #52B54B 100%);
                        color: white; padding: 1rem; border-radius: 8px;
                        margin: 1rem 0; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                <h4 style="margin: 0;">🔍 Filtros Gerais de Redes</h4>
            </div>
            """, unsafe_allow_html=True)
            col1, col2, col3 = st.columns(3)
            with col1:
                redes_disponiveis = sorted(df_com_rede['Rede'].dropna().unique())
                rede_selecionada = st.multiselect(
                    "🏢 Redes:",
                    options=redes_disponiveis,
                    default=redes_disponiveis if len(redes_disponiveis) <= 5 else [],
                    help="Selecione as redes para análise"
                )
            with col2:
                rankings_rede = sorted(df_com_rede['Ranking Rede'].dropna().unique())
                ranking_rede_selecionado = st.multiselect(
                    "🏅 Ranking Rede:",
                    options=rankings_rede,
                    default=rankings_rede if len(rankings_rede) <= 5 else [],
                    help="Selecione os rankings de rede"
                )
            with col3:
                # Categorias de redes (ouro, prata, bronze, diamante)
                categorias_rede = []
                if 'Ranking Rede' in df_com_rede.columns:
                    df_cats = df_com_rede.copy()
                    df_cats['Categoria_Rede'] = df_cats['Ranking Rede'].apply(
                        lambda x: 'Diamante' if str(x).upper() in ['DIAMANTE', 'DIAMOND'] else
                                 'Ouro' if str(x).upper() in ['OURO', 'GOLD', 'ORO'] else
                                 'Prata' if str(x).upper() in ['PRATA', 'SILVER', 'PLATA'] else
                                 'Bronze' if str(x).upper() in ['BRONZE', 'BRONCE'] else
                                 'Outros'
                    )
                    categorias_rede = sorted(df_cats['Categoria_Rede'].unique())
                categoria_selecionada = st.multiselect(
                    "🏆 Categoria Rede:",
                    options=categorias_rede,
                    default=categorias_rede if len(categorias_rede) <= 4 else [],
                    help="Filtrar por categoria da rede (Diamante, Ouro, Prata, Bronze)"
                )
            # Quarta coluna para tipo de análise
            col4 = st.columns(1)[0]
            with col4:
                tipo_analise = st.selectbox(
                    "📊 Tipo de Análise:",
                    options=["Visão Geral", "Por Volume", "Por Performance", "Por Risco", "🔄 Comparação de Redes"],
                    help="Escolha o tipo de análise a ser realizada"
                )
            # Aplicar filtros
            df_rede_filtrado = df_com_rede.copy()
            # Nota explicativa sobre filtros
            st.info("💡 **Dica:** Use os filtros acima para análise geral. Para exploração detalhada de uma rede específica, role para baixo até a seção 'Explorador Detalhado por Rede'.")
            if rede_selecionada:
                df_rede_filtrado = df_rede_filtrado[df_rede_filtrado['Rede'].isin(rede_selecionada)]
            if ranking_rede_selecionado:
                df_rede_filtrado = df_rede_filtrado[df_rede_filtrado['Ranking Rede'].isin(ranking_rede_selecionado)]
            # Aplicar filtro de categoria de rede
            if categoria_selecionada:
                df_cats_filtro = df_rede_filtrado.copy()
                df_cats_filtro['Categoria_Rede'] = df_cats_filtro['Ranking Rede'].apply(
                    lambda x: 'Diamante' if str(x).upper() in ['DIAMANTE', 'DIAMOND'] else
                             'Ouro' if str(x).upper() in ['OURO', 'GOLD', 'ORO'] else
                             'Prata' if str(x).upper() in ['PRATA', 'SILVER', 'PLATA'] else
                             'Bronze' if str(x).upper() in ['BRONZE', 'BRONCE'] else
                             'Outros'
                )
                df_rede_filtrado = df_cats_filtro[df_cats_filtro['Categoria_Rede'].isin(categoria_selecionada)]
            # ========================================
            # CÁLCULO GLOBAL DE ESTATÍSTICAS DE REDES
            # ========================================
            # Calcular rede_stats para uso em todas as análises
            rede_stats = pd.DataFrame() # Inicializar vazio por segurança
            if not df_rede_filtrado.empty and 'Rede' in df_rede_filtrado.columns:
                # Remover duplicatas baseado no CNPJ antes da contagem
                df_sem_duplicatas_rede = df_rede_filtrado.drop_duplicates(subset=['CNPJ_PCL'], keep='first')
                # Estatísticas expandidas por rede
                rede_stats = df_sem_duplicatas_rede.groupby('Rede').agg(
                    Qtd_Labs=('Nome_Fantasia_PCL', 'count'),
                    Volume_Total=('Volume_Total_2025', 'sum'),
                    Volume_Medio=('Volume_Total_2025', 'mean'),
                    Volume_Std=('Volume_Total_2025', 'std'),
                    Estado_Principal=('Estado', lambda x: x.mode().iloc[0] if not x.mode().empty else 'N/A'),
                    Cidades_Unicas=('Cidade', 'nunique'),
                    Labs_Churn=('Risco_Diario', lambda x: x.isin(['🟠 Moderado', '🔴 Alto', '⚫ Crítico']).sum())
                ).reset_index()
                # Adicionar mais métricas calculadas
                rede_stats['Taxa_Churn'] = (rede_stats['Labs_Churn'] / rede_stats['Qtd_Labs'] * 100).round(1)
                rede_stats['Volume_por_Lab'] = (rede_stats['Volume_Total'] / rede_stats['Qtd_Labs']).round(0)
                # Adicionar categoria da rede se disponível
                if 'Ranking Rede' in df_sem_duplicatas_rede.columns:
                    rede_ranking = df_sem_duplicatas_rede.groupby('Rede')['Ranking Rede'].first().reset_index()
                    rede_stats = rede_stats.merge(rede_ranking, on='Rede', how='left')
                    # Adicionar categoria
                    rede_stats['Categoria_Rede'] = rede_stats['Ranking Rede'].apply(
                        lambda x: 'Diamante' if str(x).upper() in ['DIAMANTE', 'DIAMOND'] else
                                 'Ouro' if str(x).upper() in ['OURO', 'GOLD', 'ORO'] else
                                 'Prata' if str(x).upper() in ['PRATA', 'SILVER', 'PLATA'] else
                                 'Bronze' if str(x).upper() in ['BRONZE', 'BRONCE'] else
                                 'Outros'
                    )
                else:
                    rede_stats['Ranking Rede'] = 'N/A'
                    rede_stats['Categoria_Rede'] = 'N/A'
            if not df_rede_filtrado.empty:
                # Análise baseada no tipo selecionado
                if tipo_analise == "Visão Geral":
                    # Cards de métricas gerais
                    col1, col2, col3, col4 = st.columns(4)
                    total_redes = len(rede_stats) if not rede_stats.empty else 0
                    total_labs_rede = rede_stats['Qtd_Labs'].sum() if not rede_stats.empty else 0
                    volume_total_rede = rede_stats['Volume_Total'].sum() if not rede_stats.empty else 0
                    with col1:
                        st.metric("🏢 Total de Redes", total_redes)
                    with col2:
                        st.metric("🏥 Labs nas Redes", f"{total_labs_rede:,}")
                    with col3:
                        st.metric("📦 Volume Total", f"{volume_total_rede:,}")
                    with col4:
                        media_por_rede = volume_total_rede / total_redes if total_redes > 0 else 0
                        st.metric("📊 Média por Rede", f"{media_por_rede:,.0f}")
                    # ========================================
                    # CARDS DE LOCALIDADE E VOLUMES
                    # ========================================
                    st.markdown("""
                    <div style="background: linear-gradient(135deg, #6BBF47 0%, #8FD968 100%);
                                color: white; padding: 1rem; border-radius: 8px;
                                margin: 1rem 0; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                        <h4 style="margin: 0;">📍 Distribuição por Localidade</h4>
                    </div>
                    """, unsafe_allow_html=True)
                    # Cards de localidade
                    col1, col2, col3, col4, col5, col6 = st.columns(6)
                    # Calcular métricas por estado
                    df_sem_duplicatas_local = df_rede_filtrado.drop_duplicates(subset=['CNPJ_PCL'], keep='first')
                    # Número total de laboratórios
                    total_labs = len(df_sem_duplicatas_local)
                    # Por estado
                    estados_stats = df_sem_duplicatas_local.groupby('Estado').agg({
                        'Nome_Fantasia_PCL': 'count',
                        'Volume_Total_2025': ['sum', 'mean']
                    }).round(2)
                    # Achatar colunas multi-índice
                    estados_stats.columns = ['Qtd_Labs', 'Volume_Total', 'Volume_Medio']
                    estados_stats = estados_stats.reset_index()
                    # Top 5 estados por quantidade
                    top_estados = estados_stats.nlargest(5, 'Qtd_Labs')
                    # Por cidade
                    cidades_stats = df_sem_duplicatas_local.groupby('Cidade').agg({
                        'Nome_Fantasia_PCL': 'count',
                        'Volume_Total_2025': ['sum', 'mean']
                    }).round(2)
                    cidades_stats.columns = ['Qtd_Labs', 'Volume_Total', 'Volume_Medio']
                    cidades_stats = cidades_stats.reset_index()
                    # Top 5 cidades por quantidade
                    top_cidades = cidades_stats.nlargest(5, 'Qtd_Labs')
                    with col1:
                        st.metric("🏥 Total Labs", f"{total_labs:,}")
                    with col2:
                        total_estados = df_sem_duplicatas_local['Estado'].nunique()
                        st.metric("🗺️ Estados", f"{total_estados}")
                    with col3:
                        total_cidades = df_sem_duplicatas_local['Cidade'].nunique()
                        st.metric("🏙️ Cidades", f"{total_cidades}")
                    with col4:
                        volume_total_3m = df_sem_duplicatas_local['Volume_Total_2025'].sum()
                        st.metric("📦 Vol. Total 2025", f"{volume_total_3m:,.0f}")
                    with col5:
                        volume_medio_3m = df_sem_duplicatas_local['Volume_Total_2025'].mean()
                        st.metric("📊 Vol. Médio 2025", f"{volume_medio_3m:,.0f}")
                    with col6:
                        volume_medio_por_lab = volume_total_3m / total_labs if total_labs > 0 else 0
                        st.metric("📈 Vol/Lab", f"{volume_medio_por_lab:,.0f}")
                    # Tabelas detalhadas por localidade
                    col1, col2 = st.columns(2)
                    with col1:
                        st.subheader("📍 Top Estados")
                        # Adicionar ranking para top_estados
                        top_estados_display = top_estados.copy()
                        top_estados_display['Ranking'] = range(1, len(top_estados_display) + 1)
                        top_estados_display = top_estados_display[['Ranking', 'Estado', 'Qtd_Labs', 'Volume_Total', 'Volume_Medio']]
                        st.dataframe(
                            top_estados_display,
                            width='stretch',
                            column_config={
                                "Ranking": st.column_config.NumberColumn("🏆", width="small", help="Posição no ranking"),
                                "Estado": st.column_config.TextColumn("🏛️ Estado"),
                                "Qtd_Labs": st.column_config.NumberColumn("🏥 Labs"),
                                "Volume_Total": st.column_config.NumberColumn("📦 Vol. Total", format="%.0f"),
                                "Volume_Medio": st.column_config.NumberColumn("📊 Vol. Médio", format="%.0f")
                            },
                            hide_index=True
                        )
                    with col2:
                        st.subheader("🏙️ Top Cidades")
                        # Adicionar ranking para top_cidades
                        top_cidades_display = top_cidades.copy()
                        top_cidades_display['Ranking'] = range(1, len(top_cidades_display) + 1)
                        top_cidades_display = top_cidades_display[['Ranking', 'Cidade', 'Qtd_Labs', 'Volume_Total', 'Volume_Medio']]
                        st.dataframe(
                            top_cidades_display,
                            width='stretch',
                            column_config={
                                "Ranking": st.column_config.NumberColumn("🏆", width="small", help="Posição no ranking"),
                                "Cidade": st.column_config.TextColumn("🏙️ Cidade"),
                                "Qtd_Labs": st.column_config.NumberColumn("🏥 Labs"),
                                "Volume_Total": st.column_config.NumberColumn("📦 Vol. Total", format="%.0f"),
                                "Volume_Medio": st.column_config.NumberColumn("📊 Vol. Médio", format="%.0f")
                            },
                            hide_index=True
                        )
                elif tipo_analise == "Por Volume":
                    st.subheader("📦 Análise por Volume de Coletas")
                    # Ranking de redes por volume - remover duplicatas antes da contagem
                    df_sem_duplicatas_volume = df_rede_filtrado.drop_duplicates(subset=['CNPJ_PCL'], keep='first')
                    volume_por_rede = df_sem_duplicatas_volume.groupby('Rede')['Volume_Total_2025'].agg(['sum', 'mean', 'count']).reset_index()
                    volume_por_rede.columns = ['Rede', 'Volume_Total', 'Volume_Medio', 'Qtd_Labs']
                    volume_por_rede = volume_por_rede.sort_values('Volume_Total', ascending=False)
                    # Gráfico de ranking
                    fig_ranking = px.bar(
                        volume_por_rede.head(10),
                        x='Rede',
                        y='Volume_Total',
                        title="🏆 Top 10 Redes por Volume Total",
                        color='Volume_Medio',
                        color_continuous_scale='Viridis',
                        text='Volume_Total'
                    )
                    fig_ranking.update_traces(texttemplate='%{text:.0f}', textposition='outside')
                    max_volume = volume_por_rede.head(10)['Volume_Total'].max()
                    y_axis_max = max_volume * 1.2 if max_volume > 0 else 10
                    fig_ranking.update_layout(
                        xaxis_tickangle=-45,
                        height=500,
                        margin=dict(l=60, r=60, t=80, b=80),
                        yaxis=dict(range=[0, y_axis_max])
                    )
                    st.plotly_chart(fig_ranking, width='stretch')
                    # Tabela detalhada
                    # Adicionar ranking para volume_por_rede
                    volume_por_rede_display = volume_por_rede.round(2).copy()
                    volume_por_rede_display['Ranking'] = range(1, len(volume_por_rede_display) + 1)
                    volume_por_rede_display = volume_por_rede_display[['Ranking', 'Rede', 'Volume_Total', 'Volume_Medio', 'Qtd_Labs']]
                    st.dataframe(
                        volume_por_rede_display,
                        width='stretch',
                        column_config={
                            "Ranking": st.column_config.NumberColumn("🏆", width="small", help="Posição no ranking"),
                            "Rede": st.column_config.TextColumn("🏢 Rede"),
                            "Volume_Total": st.column_config.NumberColumn("📦 Volume Total", format="%.0f"),
                            "Volume_Medio": st.column_config.NumberColumn("📊 Volume Médio", format="%.1f"),
                            "Qtd_Labs": st.column_config.NumberColumn("🏥 Qtd Labs")
                        },
                        hide_index=True
                    )
                elif tipo_analise == "Por Performance":
                    st.subheader("📈 Análise de Performance por Rede")
                    # Performance por rede (baseado em crescimento/variacao) - remover duplicatas
                    if 'Variacao_Percentual' in df_rede_filtrado.columns:
                        df_sem_duplicatas_perf = df_rede_filtrado.drop_duplicates(subset=['CNPJ_PCL'], keep='first')
                        perf_rede = df_sem_duplicatas_perf.groupby('Rede').agg({
                            'Variacao_Percentual': ['mean', 'count'],
                            'Volume_Total_2025': 'sum'
                        }).reset_index()
                        perf_rede.columns = ['Rede', 'Variacao_Media', 'Qtd_Labs', 'Volume_Total']
                        perf_rede = perf_rede.sort_values('Variacao_Media', ascending=False)
                        col1, col2 = st.columns(2)
                        with col1:
                            # Performance por variação
                            fig_perf = px.bar(
                                perf_rede.head(10),
                                x='Rede',
                                y='Variacao_Media',
                                title="📈 Top 10 Redes por Performance (Variação %)",
                                color='Variacao_Media',
                                color_continuous_scale='RdYlGn',
                                text='Variacao_Media'
                            )
                            fig_perf.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
                            max_var = perf_rede.head(10)['Variacao_Media'].max()
                            y_axis_max = max_var * 1.15 if max_var > 0 else 1
                            fig_perf.update_layout(
                                xaxis_tickangle=-45,
                                height=500,
                                margin=dict(l=60, r=60, t=80, b=80),
                                yaxis=dict(range=[0, y_axis_max])
                            )
                            st.plotly_chart(fig_perf, width='stretch')
                        with col2:
                            # Scatter plot: Volume vs Performance
                            fig_scatter = px.scatter(
                                perf_rede,
                                x='Volume_Total',
                                y='Variacao_Media',
                                size='Qtd_Labs',
                                color='Rede',
                                title="📊 Volume vs Performance por Rede",
                                labels={'Volume_Total': 'Volume Total', 'Variacao_Media': 'Variação Média %'}
                            )
                            fig_scatter.update_layout(height=500, margin=dict(l=40, r=40, t=40, b=40))
                            st.plotly_chart(fig_scatter, width='stretch')
                        # Tabela de performance
                        st.dataframe(
                            perf_rede.round(2),
                            width='stretch',
                            column_config={
                                "Rede": st.column_config.TextColumn("🏢 Rede"),
                                "Variacao_Media": st.column_config.NumberColumn("📈 Variação Média %", format="%.2f%%"),
                                "Qtd_Labs": st.column_config.NumberColumn("🏥 Qtd Labs"),
                                "Volume_Total": st.column_config.NumberColumn("📦 Volume Total", format="%.0f")
                            },
                            hide_index=True
                        )
                elif tipo_analise == "Por Risco":
                    st.subheader("⚠️ Análise de Risco por Rede")
                    if 'Risco_Diario' not in df_rede_filtrado.columns:
                        st.warning("⚠️ Coluna 'Risco_Diario' não encontrada nos dados.")
                    else:
                        df_risco = df_rede_filtrado.drop_duplicates(subset=['CNPJ_PCL'], keep='first')
                        labs_risco = df_risco[df_risco['Risco_Diario'].isin(['🟠 Moderado', '🔴 Alto', '⚫ Crítico'])]
                        cores_map = {
                            '🟢 Normal': '#16A34A',
                            '🟡 Atenção': '#F59E0B',
                            '🟠 Moderado': '#FB923C',
                            '🔴 Alto': '#DC2626',
                            '⚫ Crítico': '#111827'
                        }
                        if labs_risco.empty:
                            st.success("✅ Nenhuma rede com laboratórios em risco elevado.")
                        else:
                            resumo_rede = labs_risco.groupby('Rede').agg(
                                Labs_Risco=('CNPJ_PCL', 'count'),
                                Vol_Hoje_Medio=('Vol_Hoje', 'mean'),
                                Delta_MM7_Medio=('Delta_MM7', 'mean'),
                            Recuperando=('Recuperacao', lambda x: x.fillna(False).astype(bool).sum())
                            ).reset_index()
                            resumo_rede['Delta_MM7_Medio'] = resumo_rede['Delta_MM7_Medio'].round(1)
                            resumo_rede['Vol_Hoje_Medio'] = resumo_rede['Vol_Hoje_Medio'].round(1)
                            resumo_rede = resumo_rede.sort_values(['Labs_Risco', 'Delta_MM7_Medio'], ascending=[False, True])
                            col1, col2 = st.columns(2)
                            with col1:
                                fig_top = px.bar(
                                    resumo_rede.head(10),
                                    x='Labs_Risco',
                                    y='Rede',
                                    orientation='h',
                                    title="🚨 Redes com Mais Labs em Risco",
                                    color='Delta_MM7_Medio',
                                    color_continuous_scale='Reds',
                                    text='Labs_Risco'
                                )
                                fig_top.update_traces(texttemplate='%{text}', textposition='outside')
                                fig_top.update_layout(xaxis_title="Laboratórios em risco", yaxis_title="Rede",
                                                      height=500, margin=dict(l=40, r=40, t=40, b=40))
                                st.plotly_chart(fig_top, width='stretch')
                            with col2:
                                resumo_rede_delta = resumo_rede.sort_values('Delta_MM7_Medio')
                                fig_delta = px.bar(
                                    resumo_rede_delta.head(10),
                                    x='Delta_MM7_Medio',
                                    y='Rede',
                                    orientation='h',
                                    title="📉 Redes com Maior Queda vs MM7",
                                    color='Labs_Risco',
                                    color_continuous_scale='Reds',
                                    text='Delta_MM7_Medio'
                                )
                                fig_delta.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
                                fig_delta.update_layout(xaxis_title="Δ vs MM7 (%)", yaxis_title="Rede",
                                                        height=500, margin=dict(l=40, r=40, t=40, b=40))
                                st.plotly_chart(fig_delta, width='stretch')
                        st.dataframe(
                            resumo_rede,
                            width='stretch',
                            column_config={
                                "Rede": st.column_config.TextColumn("🏢 Rede"),
                                "Labs_Risco": st.column_config.NumberColumn("🚨 Labs em Risco"),
                                "Vol_Hoje_Medio": st.column_config.NumberColumn("📦 Vol. Médio (Hoje)", format="%.1f"),
                                "Delta_MM7_Medio": st.column_config.NumberColumn("Δ Médio vs MM7", format="%.1f%%"),
                                "Recuperando": st.column_config.NumberColumn("🔁 Em Recuperação")
                            },
                            hide_index=True
                        )
                        risco_status = df_risco.groupby(['Rede', 'Risco_Diario']).size().reset_index(name='Qtd')
                        fig_status = px.bar(
                            risco_status,
                            x='Rede',
                            y='Qtd',
                            color='Risco_Diario',
                            title="📊 Distribuição de Risco Diário por Rede",
                            color_discrete_map=cores_map,
                            barmode='stack'
                        )
                        fig_status.update_layout(xaxis_tickangle=-45, height=500, margin=dict(l=40, r=40, t=40, b=40))
                        st.plotly_chart(fig_status, width='stretch')
                        # Destaques de risco crítico
                        redes_criticas = labs_risco[labs_risco['Risco_Diario'] == '⚫ Crítico']['Rede'].value_counts()
                        if not redes_criticas.empty:
                            st.error("🚨 Redes com laboratórios em risco crítico detectadas!")
                            for rede, qtd in redes_criticas.items():
                                st.write(f"• **{rede}**: {qtd} laboratório(s) crítico(s)")
                elif tipo_analise == "🔄 Comparação de Redes":
                    st.subheader("🔄 Comparação Direta de Redes")
                    # Seletor de redes para comparação (máximo 5 para legibilidade)
                    redes_para_comparar = st.multiselect(
                        "🏢 Selecione até 5 redes para comparar:",
                        options=sorted(rede_stats['Rede'].unique()),
                        default=sorted(rede_stats['Rede'].unique())[:3] if len(rede_stats) >= 3 else sorted(rede_stats['Rede'].unique()),
                        max_selections=5,
                        help="Escolha as redes que deseja comparar diretamente"
                    )
                    if redes_para_comparar:
                        # Filtrar dados apenas das redes selecionadas
                        redes_comparacao = rede_stats[rede_stats['Rede'].isin(redes_para_comparar)].copy()
                        if not redes_comparacao.empty:
                            # ========================================
                            # DASHBOARD DE COMPARAÇÃO
                            # ========================================
                            # Cards de comparação rápida
                            st.markdown("### 📊 Comparação Rápida")
                            col1, col2, col3, col4 = st.columns(4)
                            with col1:
                                maior_qtd = redes_comparacao.loc[redes_comparacao['Qtd_Labs'].idxmax()]
                                st.metric(
                                    "🏥 Maior Qtd Labs",
                                    f"{int(maior_qtd['Qtd_Labs'])}",
                                    f"{maior_qtd['Rede'][:15]}..."
                                )
                            with col2:
                                maior_volume = redes_comparacao.loc[redes_comparacao['Volume_Total'].idxmax()]
                                st.metric(
                                    "📦 Maior Volume",
                                    f"{maior_volume['Volume_Total']:,.0f}",
                                    f"{maior_volume['Rede'][:15]}..."
                                )
                            with col3:
                                menor_churn = redes_comparacao.loc[redes_comparacao['Taxa_Churn'].idxmin()]
                                st.metric(
                                    "✅ Menor Churn",
                                    f"{menor_churn['Taxa_Churn']:.1f}%",
                                    f"{menor_churn['Rede'][:15]}..."
                                )
                            with col4:
                                maior_risco = redes_comparacao.loc[redes_comparacao['Labs_Churn'].idxmax()]
                                st.metric(
                                    "⚠️ Mais Labs em Risco",
                                    f"{int(maior_risco['Labs_Churn'])}",
                                    f"{maior_risco['Rede'][:15]}..."
                                )
                            # ========================================
                            # GRÁFICOS COMPARATIVOS
                            # ========================================
                            st.markdown("### 📈 Comparações Visuais")
                            # Gráfico de barras comparativo - múltiplas métricas
                            col1, col2 = st.columns(2)
                            with col1:
                                # Comparação por quantidade de laboratórios e volume
                                fig_comp1 = go.Figure()
                                for _, rede in redes_comparacao.iterrows():
                                    fig_comp1.add_trace(go.Bar(
                                        name=f"{rede['Rede'][:12]}...",
                                        x=['Labs', 'Volume (k)'],
                                        y=[rede['Qtd_Labs'], rede['Volume_Total']/1000],
                                        text=[f"{int(rede['Qtd_Labs'])}", f"{rede['Volume_Total']/1000:.0f}k"],
                                        textposition='auto',
                                    ))
                                fig_comp1.update_layout(
                                    title="🏥 Labs vs 📦 Volume por Rede",
                                    barmode='group',
                                    height=400
                                )
                                st.plotly_chart(fig_comp1, width='stretch')
                            with col2:
                                # Comparação de performance (volume médio e taxa churn)
                                fig_comp2 = go.Figure()
                                for _, rede in redes_comparacao.iterrows():
                                    fig_comp2.add_trace(go.Scatter(
                                        name=f"{rede['Rede'][:12]}...",
                                        x=[rede['Volume_Medio']],
                                        y=[rede['Taxa_Churn']],
                                        mode='markers+text',
                                        text=f"{rede['Rede'][:8]}...",
                                        textposition="top center",
                                        marker=dict(size=15)
                                    ))
                                fig_comp2.update_layout(
                                    title="💰 Volume Médio vs 📉 Taxa Churn",
                                    xaxis_title="Volume Médio por Lab",
                                    yaxis_title="Taxa Churn (%)",
                                    height=400
                                )
                                st.plotly_chart(fig_comp2, width='stretch')
                            # ========================================
                            # TABELA COMPARATIVA DETALHADA
                            # ========================================
                            st.markdown("### 📋 Comparação Detalhada")
                            # Reordenar colunas para melhor visualização
                            cols_comparacao = [
                                'Rede', 'Categoria_Rede', 'Qtd_Labs', 'Labs_Churn', 'Taxa_Churn',
                                'Volume_Total', 'Volume_Medio', 'Volume_por_Lab'
                            ]
                            # Adicionar indicadores visuais de risco
                            redes_comparacao_display = redes_comparacao[cols_comparacao].copy()
                            # Função para adicionar indicadores de risco
                            def adicionar_indicador_risco(row):
                                indicadores = []
                                # Indicador de alto churn
                                if row['Taxa_Churn'] > 30:
                                    indicadores.append("🔴")
                                elif row['Taxa_Churn'] > 15:
                                    indicadores.append("🟠")
                                else:
                                    indicadores.append("🟢")
                                # Indicador de concentração de labs em risco
                                proporcao_risco = (row['Labs_Churn'] / row['Qtd_Labs']) if row['Qtd_Labs'] else 0
                                if proporcao_risco >= 0.5:
                                    indicadores.append("⚠️")
                                elif proporcao_risco >= 0.3:
                                    indicadores.append("⚡")
                                # Indicador de baixa eficiência (volume por lab)
                                media_geral = redes_comparacao['Volume_por_Lab'].mean()
                                if row['Volume_por_Lab'] < media_geral * 0.7:
                                    indicadores.append("📉")
                                return ' '.join(indicadores) if indicadores else "✅"
                            redes_comparacao_display['🚨 Indicadores'] = redes_comparacao_display.apply(adicionar_indicador_risco, axis=1)
                            # Reordenar para colocar indicadores primeiro
                            cols_final = ['🚨 Indicadores'] + cols_comparacao
                            redes_comparacao_display = redes_comparacao_display[cols_final]
                            st.dataframe(
                                redes_comparacao_display.round(2),
                                width='stretch',
                                column_config={
                                    "🚨 Indicadores": st.column_config.TextColumn("🚨 Alertas", width="small"),
                                    "Rede": st.column_config.TextColumn("🏢 Rede", width="medium"),
                                    "Categoria_Rede": st.column_config.TextColumn("🏆 Categoria", width="small"),
                                    "Qtd_Labs": st.column_config.NumberColumn("🏥 Labs", format="%d"),
                                    "Labs_Churn": st.column_config.NumberColumn("❌ Churn", format="%d"),
                                    "Taxa_Churn": st.column_config.NumberColumn("📉 % Churn", format="%.1f%%"),
                                    "Volume_Total": st.column_config.NumberColumn("📦 Vol. Total", format="%.0f"),
                                    "Volume_Medio": st.column_config.NumberColumn("📊 Vol. Médio", format="%.0f"),
                                    "Volume_por_Lab": st.column_config.NumberColumn("💰 Vol/Lab", format="%.0f")
                                },
                                hide_index=True
                            )
                            # ========================================
                            # RANKING COMPARATIVO
                            # ========================================
                            st.markdown("### 🏆 Rankings Comparativos")
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.subheader("🥇 Por Volume Total")
                                ranking_volume = redes_comparacao.sort_values('Volume_Total', ascending=False)[['Rede', 'Volume_Total']]
                                for idx, row in ranking_volume.iterrows():
                                    medal = "🥇" if idx == 0 else "🥈" if idx == 1 else "🥉" if idx == 2 else "📊"
                                    st.write(f"{medal} {row['Rede'][:20]}...: {row['Volume_Total']:,.0f}")
                            with col2:
                                st.subheader("🥇 Por Eficiência")
                                ranking_eficiencia = redes_comparacao.sort_values('Volume_por_Lab', ascending=False)[['Rede', 'Volume_por_Lab']]
                                for idx, row in ranking_eficiencia.iterrows():
                                    medal = "🥇" if idx == 0 else "🥈" if idx == 1 else "🥉" if idx == 2 else "📊"
                                    st.write(f"{medal} {row['Rede'][:20]}...: {row['Volume_por_Lab']:,.0f}")
                            with col3:
                                st.subheader("🥇 Por Menor Risco")
                                ranking_risco = redes_comparacao.sort_values('Taxa_Churn', ascending=True)[['Rede', 'Taxa_Churn']]
                                for idx, row in ranking_risco.iterrows():
                                    medal = "🥇" if idx == 0 else "🥈" if idx == 1 else "🥉" if idx == 2 else "📊"
                                    st.write(f"{medal} {row['Rede'][:20]}...: {row['Taxa_Churn']:.1f}%")
                        else:
                            st.warning("⚠️ Nenhuma rede encontrada com os critérios selecionados.")
                    else:
                        st.info("ℹ️ Selecione pelo menos uma rede para iniciar a comparação.")
            else:
                st.warning("⚠️ Nenhum dado encontrado com os filtros aplicados.")
        else:
            st.warning("⚠️ Dados VIP não disponíveis. Verifique se o arquivo Excel foi carregado corretamente.")
    elif st.session_state.page == "🔧 Manutenção VIPs":
        st.header("🔧 Manutenção de Dados VIP")
        
        # Mensagem de manutenção
        st.error("🚧 **TELA EM MANUTENÇÃO** 🚧\n\nEsta funcionalidade está temporariamente indisponível. Por favor, utilize outras telas do sistema.")
        st.stop()
        
        st.markdown("""
        <div style="background: linear-gradient(135deg, #6BBF47 0%, #52B54B 100%);
                    color: white; padding: 1rem; border-radius: 8px; margin-bottom: 2rem;">
            <h3 style="margin: 0; color: white;">Gerenciamento de Laboratórios VIP</h3>
            <p style="margin: 0.5rem 0 0 0; opacity: 0.9;">Adicione, edite e gerencie laboratórios VIP com validação completa e histórico de alterações.</p>
        </div>
        """, unsafe_allow_html=True)
     
        # Importar módulos necessários
        try:
            from vip_history_manager import VIPHistoryManager
            from vip_integration import VIPIntegration
            import json
            import shutil
        except ImportError as e:
            st.error(f"Erro ao importar módulos VIP: {e}")
            st.stop()
     
        # Inicializar gerenciadores
        history_manager = VIPHistoryManager(OUTPUT_DIR)
        vip_integration = VIPIntegration(OUTPUT_DIR)
     
        # Sub-abas para diferentes funcionalidades
        sub_tab1, sub_tab2, sub_tab3, sub_tab4 = st.tabs([
            "📋 Visualizar VIPs",
            "➕ Adicionar VIP",
            "✏️ Editar VIP",
            "📊 Histórico"
        ])
     
        with sub_tab1:
            st.subheader("📋 Lista de Laboratórios VIP")
         
            # Carregar dados VIP
            loader = show_overlay_loader(
                "Carregando dados VIP...",
                "Buscando lista de laboratórios VIP. Aguarde um momento."
            )
            try:
                df_vip = DataManager.carregar_dados_vip()
            finally:
                loader.empty()
         
            if df_vip is not None and not df_vip.empty:
                # Filtros
                col1, col2, col3 = st.columns(3)
             
                with col1:
                    ranking_filtro = st.selectbox(
                        "🏆 Ranking:",
                        options=["Todos"] + sorted(df_vip['Ranking'].dropna().unique().tolist()),
                        help="Filtrar por ranking individual"
                    )
             
                with col2:
                    ranking_rede_filtro = st.selectbox(
                        "🏅 Ranking Rede:",
                        options=["Todos"] + sorted(df_vip['Ranking Rede'].dropna().unique().tolist()),
                        help="Filtrar por ranking de rede"
                    )
             
                with col3:
                    rede_filtro = st.selectbox(
                        "🏢 Rede:",
                        options=["Todas"] + sorted(df_vip['Rede'].dropna().unique().tolist()),
                        help="Filtrar por rede"
                    )
             
                # Aplicar filtros
                df_filtrado = df_vip.copy()
             
                if ranking_filtro != "Todos":
                    df_filtrado = df_filtrado[df_filtrado['Ranking'] == ranking_filtro]
             
                if ranking_rede_filtro != "Todos":
                    df_filtrado = df_filtrado[df_filtrado['Ranking Rede'] == ranking_rede_filtro]
             
                if rede_filtro != "Todas":
                    df_filtrado = df_filtrado[df_filtrado['Rede'] == rede_filtro]
             
                # Estatísticas
                col1, col2, col3, col4 = st.columns(4)
             
                with col1:
                    st.metric("📊 Total VIPs", len(df_filtrado))
             
                with col2:
                    st.metric("🏆 Rankings", len(df_filtrado['Ranking'].unique()))
             
                with col3:
                    st.metric("🏢 Redes", len(df_filtrado['Rede'].unique()))
             
                with col4:
                    st.metric("🏅 Rankings Rede", len(df_filtrado['Ranking Rede'].unique()))
             
                # Tabela de dados
                st.subheader("📋 Dados VIP Filtrados")
             
                # Configurar colunas para exibição
                colunas_exibir = ['CNPJ', 'RAZÃO SOCIAL', 'NOME FANTASIA', 'Cidade ', 'UF',
                                'Ranking', 'Ranking Rede', 'Rede', 'STATUS']
             
                colunas_existentes = [col for col in colunas_exibir if col in df_filtrado.columns]
             
                if colunas_existentes:
                    st.dataframe(
                        df_filtrado[colunas_existentes],
                        width='stretch',
                        height=400,
                        column_config={
                            "CNPJ": st.column_config.TextColumn("📄 CNPJ", help="CNPJ do laboratório"),
                            "RAZÃO SOCIAL": st.column_config.TextColumn("🏢 Razão Social"),
                            "NOME FANTASIA": st.column_config.TextColumn("🏥 Nome Fantasia"),
                            "Cidade ": st.column_config.TextColumn("🏙️ Cidade"),
                            "UF": st.column_config.TextColumn("🗺️ Estado"),
                            "Ranking": st.column_config.TextColumn("🏆 Ranking"),
                            "Ranking Rede": st.column_config.TextColumn("🏅 Ranking Rede"),
                            "Rede": st.column_config.TextColumn("🏢 Rede"),
                            "STATUS": st.column_config.TextColumn("📊 Status")
                        },
                        hide_index=True
                    )
                else:
                    st.warning("Nenhuma coluna válida encontrada para exibição")
            else:
                st.warning("⚠️ Nenhum dado VIP encontrado. Execute primeiro o script de normalização.")
     
        with sub_tab2:
            st.subheader("➕ Adicionar Novo Laboratório VIP")
         
            # Formulário para adicionar adicionar VIP
            with st.form("form_adicionar_vip"):
                col1, col2 = st.columns(2)
             
                with col1:
                    cnpj_novo = st.text_input(
                        "📄 CNPJ:",
                        placeholder="00.000.000/0000-00",
                        help="CNPJ do laboratório (será validado automaticamente)"
                    )
                 
                    razao_social = st.text_input(
                        "🏢 Razão Social:",
                        placeholder="Nome da empresa"
                    )
                 
                    nome_fantasia = st.text_input(
                        "🏥 Nome Fantasia:",
                        placeholder="Nome comercial"
                    )
                 
                    cidade = st.text_input(
                        "🏙️ Cidade:",
                        placeholder="Nome da cidade"
                    )
             
                with col2:
                    uf = st.selectbox(
                        "🗺️ Estado:",
                        options=[""] + ESTADOS_BRASIL,
                        help="Selecione o estado"
                    )
             
                    ranking = st.selectbox(
                        "🏆 Ranking:",
                        options=list(CATEGORIAS_RANKING.keys()),
                        help="Ranking individual do laboratório"
                    )
                 
                    ranking_rede = st.selectbox(
                        "🏅 Ranking Rede:",
                        options=list(CATEGORIAS_RANKING_REDE.keys()),
                        help="Ranking da rede"
                    )
                 
                    rede = st.text_input(
                        "🏢 Rede:",
                        placeholder="Nome da rede"
                    )
             
                contato = st.text_input(
                    "👤 Contato:",
                    placeholder="Nome do contato"
                )
             
                telefone = st.text_input(
                    "📞 Telefone/WhatsApp:",
                    placeholder="(00) 00000-0000"
                )
             
                observacoes = st.text_area(
                    "📝 Observações:",
                    placeholder="Observações adicionais (opcional)"
                )
             
                submitted = st.form_submit_button("➕ Adicionar VIP", type="primary")
             
                if submitted:
                    # Validações
                    erros = []
                 
                    # Validar CNPJ
                    if not cnpj_novo:
                        erros.append("CNPJ é obrigatório")
                    else:
                        valido, mensagem = vip_integration.validar_cnpj(cnpj_novo)
                        if not valido:
                            erros.append(f"CNPJ inválido: {mensagem}")
                        elif vip_integration.verificar_cnpj_vip_existe(cnpj_novo):
                            erros.append("CNPJ já existe na lista VIP")
                 
                    # Validar campos obrigatórios
                    if not razao_social:
                        erros.append("Razão Social é obrigatório")
                 
                    if not nome_fantasia:
                        erros.append("Nome Fantasia é obrigatório")
                 
                    if not uf:
                        erros.append("Estado é obrigatório")
                 
                    if not rede:
                        erros.append("Rede é obrigatória")
                 
                    if erros:
                        for erro in erros:
                            st.error(f"❌ {erro}")
                    else:
                        # Auto-completar dados se CNPJ existe nos laboratórios
                        dados_lab = vip_integration.buscar_laboratorio_por_cnpj(cnpj_novo)
                        if dados_lab:
                            if not razao_social:
                                razao_social = dados_lab.get('razao_social', '')
                            if not nome_fantasia:
                                nome_fantasia = dados_lab.get('nome_fantasia', '')
                            if not cidade:
                                cidade = dados_lab.get('cidade', '')
                            if not uf:
                                uf = dados_lab.get('estado', '')
                     
                        # Criar backup antes de adicionar
                        if VIP_AUTO_BACKUP:
                            try:
                                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                backup_path = os.path.join(VIP_BACKUP_DIR, f"vip_backup_{timestamp}.csv")
                                os.makedirs(VIP_BACKUP_DIR, exist_ok=True)
                             
                                if os.path.exists(os.path.join(OUTPUT_DIR, VIP_CSV_FILE)):
                                    shutil.copy2(os.path.join(OUTPUT_DIR, VIP_CSV_FILE), backup_path)
                                    st.toast(f"✅ Backup criado: {backup_path}")
                            except Exception as e:
                                st.warning(f"⚠️ Erro ao criar backup: {e}")
                     
                        # Adicionar novo VIP
                        try:
                            # Carregar dados existentes
                            df_vip_atual = DataManager.carregar_dados_vip()
                            if df_vip_atual is None:
                                df_vip_atual = pd.DataFrame()
                         
                            # Criar novo registro
                            novo_registro = {
                                'CNPJ': cnpj_novo,
                                'RAZÃO SOCIAL': razao_social,
                                'NOME FANTASIA': nome_fantasia,
                                'Cidade ': cidade,
                                'UF': uf,
                                'Contato PCL': contato,
                                'Whatsapp/telefone': telefone,
                                'REP': '', # Será preenchido automaticamente se CNPJ existir
                                'CS': '', # Será preenchido automaticamente se CNPJ existir
                                'STATUS': 'ATIVO',
                                'Ranking': ranking,
                                'Ranking Rede': ranking_rede,
                                'Rede': rede
                            }
                         
                            # Adicionar ao DataFrame
                            df_novo = pd.DataFrame([novo_registro])
                            df_vip_atualizado = pd.concat([df_vip_atual, df_novo], ignore_index=True)
                         
                            # Salvar CSV atualizado
                            caminho_csv = os.path.join(OUTPUT_DIR, VIP_CSV_FILE)
                            df_vip_atualizado.to_csv(caminho_csv, index=False, encoding='utf-8-sig')
                         
                            # Registrar no histórico
                            history_manager.registrar_insercao(
                                cnpj=cnpj_novo,
                                dados_novos=novo_registro,
                                usuario="streamlit_user",
                                observacoes=observacoes
                            )
                         
                            # Limpar cache
                            DataManager.carregar_dados_vip.clear()
                         
                            st.toast(f"✅ Laboratório VIP adicionado com sucesso!")
                            st.success(f"📄 CNPJ: {cnpj_novo}")
                            st.success(f"🏥 Nome: {nome_fantasia}")
                         
                            # Mostrar sugestões de laboratórios similares
                            sugestoes = vip_integration.obter_sugestoes_laboratorios(limite=5)
                            if sugestoes:
                                st.info("💡 Outros laboratórios que ainda não são VIP:")
                                for sug in sugestoes[:3]:
                                    st.write(f"• {sug['nome_fantasia']} ({sug['cnpj']}) - {sug['estado']}")
                         
                        except Exception as e:
                            st.error(f"❌ Erro ao adicionar VIP: {e}")
     
        with sub_tab3:
            st.subheader("✏️ Editar Laboratório VIP")
         
            # Carregar dados VIP
            loader = show_overlay_loader(
                "Carregando dados VIP...",
                "Buscando lista de laboratórios VIP para edição."
            )
            try:
                df_vip = DataManager.carregar_dados_vip()
            finally:
                loader.empty()
         
            if df_vip is not None and not df_vip.empty:
                # Selecionar VIP para editar
                col1, col2 = st.columns([2, 1])
             
                with col1:
                    # Busca por CNPJ ou nome
                    busca = st.text_input(
                        "🔍 Buscar VIP:",
                        placeholder="Digite CNPJ ou nome do laboratório"
                    )
             
                with col2:
                    if busca:
                        # Filtrar resultados
                        mask = (
                            df_vip['CNPJ'].str.contains(busca, case=False, na=False) |
                            df_vip['NOME FANTASIA'].str.contains(busca, case=False, na=False) |
                            df_vip['RAZÃO SOCIAL'].str.contains(busca, case=False, na=False)
                        )
                        df_filtrado = df_vip[mask]
                    else:
                        df_filtrado = df_vip
             
                if not df_filtrado.empty:
                    # Selecionar VIP
                    vip_selecionado = st.selectbox(
                        "📋 Selecionar VIP para editar:",
                        options=df_filtrado.index,
                        format_func=lambda x: f"{df_filtrado.loc[x, 'NOME FANTASIA']} - {df_filtrado.loc[x, 'CNPJ']}",
                        help="Selecione o laboratório VIP para editar"
                    )
                 
                    if vip_selecionado is not None:
                        vip_data = df_filtrado.loc[vip_selecionado]
                     
                        st.markdown("---")
                        st.subheader(f"✏️ Editando: {vip_data['NOME FANTASIA']}")
                     
                        # Formulário de edição
                        with st.form("form_editar_vip"):
                            col1, col2 = st.columns(2)
                         
                            with col1:
                                cnpj_edit = st.text_input(
                                    "📄 CNPJ:",
                                    value=vip_data['CNPJ'],
                                    disabled=True, # CNPJ não pode ser alterado
                                    help="CNPJ não pode ser alterado"
                                )
                             
                                razao_social_edit = st.text_input(
                                    "🏢 Razão Social:",
                                    value=vip_data.get('RAZÃO SOCIAL', '')
                                )
                             
                                nome_fantasia_edit = st.text_input(
                                    "🏥 Nome Fantasia:",
                                    value=vip_data.get('NOME FANTASIA', '')
                                )
                             
                                cidade_edit = st.text_input(
                                    "🏙️ Cidade:",
                                    value=vip_data.get('Cidade ', '')
                                )
                         
                            with col2:
                                uf_edit = st.selectbox(
                                    "🗺️ Estado:",
                                    options=ESTADOS_BRASIL,
                                    index=ESTADOS_BRASIL.index(vip_data.get('UF', '')) if vip_data.get('UF', '') in ESTADOS_BRASIL else 0
                                )
                             
                                ranking_edit = st.selectbox(
                                    "🏆 Ranking:",
                                    options=list(CATEGORIAS_RANKING.keys()),
                                    index=list(CATEGORIAS_RANKING.keys()).index(vip_data.get('Ranking', 'BRONZE')) if vip_data.get('Ranking', '') in CATEGORIAS_RANKING else 0
                                )
                             
                                ranking_rede_edit = st.selectbox(
                                    "🏅 Ranking Rede:",
                                    options=list(CATEGORIAS_RANKING_REDE.keys()),
                                    index=list(CATEGORIAS_RANKING_REDE.keys()).index(vip_data.get('Ranking Rede', 'BRONZE')) if vip_data.get('Ranking Rede', '') in CATEGORIAS_RANKING_REDE else 0
                                )
                             
                                rede_edit = st.text_input(
                                    "🏢 Rede:",
                                    value=vip_data.get('Rede', '')
                                )
                         
                            contato_edit = st.text_input(
                                "👤 Contato:",
                                value=vip_data.get('Contato PCL', '')
                            )
                         
                            telefone_edit = st.text_input(
                                "📞 Telefone/WhatsApp:",
                                value=vip_data.get('Whatsapp/telefone', '')
                            )
                         
                            status_edit = st.selectbox(
                                "📊 Status:",
                                options=['ATIVO', 'INATIVO', 'DELETADO'],
                                index=['ATIVO', 'INATIVO', 'DELETADO'].index(vip_data.get('STATUS', 'ATIVO'))
                            )
                         
                            observacoes_edit = st.text_area(
                                "📝 Observações da Edição:",
                                placeholder="Descreva as alterações realizadas"
                            )
                         
                            submitted_edit = st.form_submit_button("💾 Salvar Alterações", type="primary")
                         
                            if submitted_edit:
                                # Verificar se houve alterações
                                alteracoes = []
                             
                                if razao_social_edit != vip_data.get('RAZÃO SOCIAL', ''):
                                    alteracoes.append(('RAZÃO SOCIAL', vip_data.get('RAZÃO SOCIAL', ''), razao_social_edit))
                             
                                if nome_fantasia_edit != vip_data.get('NOME FANTASIA', ''):
                                    alteracoes.append(('NOME FANTASIA', vip_data.get('NOME FANTASIA', ''), nome_fantasia_edit))
                             
                                if ranking_edit != vip_data.get('Ranking', ''):
                                    alteracoes.append(('Ranking', vip_data.get('Ranking', ''), ranking_edit))
                             
                                if ranking_rede_edit != vip_data.get('Ranking Rede', ''):
                                    alteracoes.append(('Ranking Rede', vip_data.get('Ranking Rede', ''), ranking_rede_edit))
                             
                                if rede_edit != vip_data.get('Rede', ''):
                                    alteracoes.append(('Rede', vip_data.get('Rede', ''), rede_edit))
                             
                                if status_edit != vip_data.get('STATUS', ''):
                                    alteracoes.append(('STATUS', vip_data.get('STATUS', ''), status_edit))
                             
                                if alteracoes:
                                    # Criar backup antes de editar
                                    if VIP_AUTO_BACKUP:
                                        try:
                                            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                            backup_path = os.path.join(VIP_BACKUP_DIR, f"vip_backup_{timestamp}.csv")
                                            os.makedirs(VIP_BACKUP_DIR, exist_ok=True)
                                         
                                            if os.path.exists(os.path.join(OUTPUT_DIR, VIP_CSV_FILE)):
                                                shutil.copy2(os.path.join(OUTPUT_DIR, VIP_CSV_FILE), backup_path)
                                                st.toast(f"✅ Backup criado: {backup_path}")
                                        except Exception as e:
                                            st.warning(f"⚠️ Erro ao criar backup: {e}")
                                 
                                    # Atualizar dados
                                    try:
                                        # Atualizar DataFrame
                                        df_vip_atualizado = df_vip.copy()
                                        df_vip_atualizado.loc[vip_selecionado, 'RAZÃO SOCIAL'] = razao_social_edit
                                        df_vip_atualizado.loc[vip_selecionado, 'NOME FANTASIA'] = nome_fantasia_edit
                                        df_vip_atualizado.loc[vip_selecionado, 'Cidade '] = cidade_edit
                                        df_vip_atualizado.loc[vip_selecionado, 'UF'] = uf_edit
                                        df_vip_atualizado.loc[vip_selecionado, 'Ranking'] = ranking_edit
                                        df_vip_atualizado.loc[vip_selecionado, 'Ranking Rede'] = ranking_rede_edit
                                        df_vip_atualizado.loc[vip_selecionado, 'Rede'] = rede_edit
                                        df_vip_atualizado.loc[vip_selecionado, 'Contato PCL'] = contato_edit
                                        df_vip_atualizado.loc[vip_selecionado, 'Whatsapp/telefone'] = telefone_edit
                                        df_vip_atualizado.loc[vip_selecionado, 'STATUS'] = status_edit
                                     
                                        # Salvar CSV atualizado
                                        caminho_csv = os.path.join(OUTPUT_DIR, VIP_CSV_FILE)
                                        df_vip_atualizado.to_csv(caminho_csv, index=False, encoding='utf-8-sig')
                                     
                                        # Registrar alterações no histórico
                                        for campo, valor_anterior, valor_novo in alteracoes:
                                            history_manager.registrar_edicao(
                                                cnpj=vip_data['CNPJ'],
                                                campo_alterado=campo,
                                                valor_anterior=valor_anterior,
                                                valor_novo=valor_novo,
                                                dados_antes=vip_data.to_dict(),
                                                dados_depois=df_vip_atualizado.loc[vip_selecionado].to_dict(),
                                                usuario="streamlit_user",
                                                observacoes=observacoes_edit
                                            )
                                     
                                        # Limpar cache
                                        DataManager.carregar_dados_vip.clear()
                                     
                                        st.toast(f"✅ Laboratório VIP atualizado com sucesso!")
                                        st.success(f"📝 {len(alteracoes)} campo(s) alterado(s)")
                                     
                                        # Mostrar resumo das alterações
                                        for campo, valor_anterior, valor_novo in alteracoes:
                                            st.info(f"🔄 {campo}: '{valor_anterior}' → '{valor_novo}'")
                                     
                                    except Exception as e:
                                        st.error(f"❌ Erro ao atualizar VIP: {e}")
                                else:
                                    st.info("ℹ️ Nenhuma alteração detectada")
            else:
                st.warning("⚠️ Nenhum dado VIP encontrado. Execute primeiro o script de normalização.")
     
        with sub_tab4:
            st.subheader("📊 Histórico de Alterações")
         
            # Estatísticas do histórico
            stats = history_manager.obter_estatisticas()
         
            if stats.get('total_alteracoes', 0) > 0:
                col1, col2, col3, col4 = st.columns(4)
             
                with col1:
                    st.metric("📊 Total Alterações", stats['total_alteracoes'])
             
                with col2:
                    st.metric("➕ Inserções", stats['por_tipo'].get('insercao', 0))
             
                with col3:
                    st.metric("✏️ Edições", stats['por_tipo'].get('edicao', 0))
             
                with col4:
                    st.metric("🗑️ Exclusões", stats['por_tipo'].get('exclusao', 0))
             
                # Filtros para histórico
                col1, col2, col3 = st.columns(3)
             
                with col1:
                    tipo_filtro = st.selectbox(
                        "🔍 Tipo de Alteração:",
                        options=["Todos"] + list(stats['por_tipo'].keys()),
                        help="Filtrar por tipo de alteração"
                    )
             
                with col2:
                    cnpj_filtro = st.text_input(
                        "📄 CNPJ:",
                        placeholder="Digite CNPJ para filtrar",
                        help="Filtrar por CNPJ específico"
                    )
             
                with col3:
                    dias_filtro = st.selectbox(
                        "📅 Período:",
                        options=["Todos", "Últimos 7 dias", "Últimos 30 dias", "Últimos 90 dias"],
                        help="Filtrar por período"
                    )
             
                # Obter histórico filtrado
                if cnpj_filtro:
                    historico_filtrado = history_manager.buscar_historico_cnpj(cnpj_filtro)
                else:
                    historico_filtrado = history_manager.historico
             
                # Filtrar por tipo
                if tipo_filtro != "Todos":
                    historico_filtrado = [alt for alt in historico_filtrado if alt['tipo'] == tipo_filtro]
             
                # Filtrar por período
                if dias_filtro != "Todos":
                    dias = {"Últimos 7 dias": 7, "Últimos 30 dias": 30, "Últimos 90 dias": 90}[dias_filtro]
                    data_limite = datetime.now() - timedelta(days=dias)
                    historico_filtrado = [alt for alt in historico_filtrado
                                        if datetime.fromisoformat(alt['timestamp']) >= data_limite]
             
                # Mostrar histórico
                if historico_filtrado:
                    st.subheader(f"📋 Histórico Filtrado ({len(historico_filtrado)} registros)")
                 
                    # Ordenar por timestamp (mais recente primeiro)
                    historico_filtrado.sort(key=lambda x: x['timestamp'], reverse=True)
                 
                    for i, alt in enumerate(historico_filtrado[:20]): # Mostrar apenas os 20 mais recentes
                        with st.expander(f"{alt['tipo'].title()} - {alt['cnpj']} - {alt['timestamp'][:19]}"):
                            col1, col2 = st.columns(2)
                         
                            with col1:
                                st.write(f"**Tipo:** {alt['tipo'].title()}")
                                st.write(f"**CNPJ:** {alt['cnpj']}")
                                st.write(f"**Data/Hora:** {alt['timestamp'][:19]}")
                                st.write(f"**Usuário:** {alt.get('usuario', 'N/A')}")
                         
                            with col2:
                                if alt['tipo'] == 'edicao':
                                    st.write(f"**Campo:** {alt.get('campo_alterado', 'N/A')}")
                                    st.write(f"**De:** {alt.get('valor_anterior', 'N/A')}")
                                    st.write(f"**Para:** {alt.get('valor_novo', 'N/A')}")
                             
                                if alt.get('observacoes'):
                                    st.write(f"**Observações:** {alt['observacoes']}")
                 
                    # Botão para exportar histórico
                    if st.button("📥 Exportar Histórico CSV"):
                        try:
                            caminho_export = history_manager.exportar_historico_csv()
                            if caminho_export:
                                st.toast(f"✅ Histórico exportado: {caminho_export}")
                        except Exception as e:
                            st.error(f"❌ Erro ao exportar histórico: {e}")
                else:
                    st.info("ℹ️ Nenhum registro encontrado com os filtros aplicados")
            else:
                st.info("ℹ️ Nenhuma alteração registrada ainda")
    
    # ========================================
    # ANÁLISE DE CONCORRENTE (GRALAB)
    # ========================================
    elif st.session_state.page == "🔍 Análise de Concorrente":
        st.header("🔍 Análise de Concorrente - Gralab (CunhaLab)")
        
        st.markdown("""
        <div style="background: linear-gradient(135deg, #ffd700 0%, #ffed4e 100%);
                    color: #333; padding: 1.5rem; border-radius: 10px;
                    margin-bottom: 2rem; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
            <h3 style="margin: 0; font-size: 1.4rem; color: #b8860b;">📊 Análise Comparativa de Mercado</h3>
            <p style="margin: 0.5rem 0 0 0; font-size: 1rem; color: #666;">
                Compare nossa base de laboratórios com o concorrente Gralab (CunhaLab) para identificar oportunidades e ameaças.
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        # Carregar dados
        loader = show_overlay_loader(
            "Carregando dados do Gralab (CunhaLab)...",
            "Sincronizando informações do concorrente. Isso pode levar alguns segundos."
        )
        try:
            dados_gralab = DataManager.carregar_dados_gralab()
        finally:
            loader.empty()
        
        if not dados_gralab or 'Dados Completos' not in dados_gralab:
            st.error("❌ Não foi possível carregar os dados do Gralab. Verifique a conexão com o SharePoint.")
        else:
            df_gralab = dados_gralab['Dados Completos']
            
            # Normalizar CNPJs da nossa base (usar df completo, não df_filtrado)
            # Isso garante que todos os nossos clientes sejam considerados na comparação
            if 'CNPJ_Normalizado' not in df.columns:
                df['CNPJ_Normalizado'] = df['CNPJ_PCL'].apply(DataManager.normalizar_cnpj)
            
            # Obter conjuntos de CNPJs (usar df completo para ter todos os clientes)
            cnpjs_nossos = set(df['CNPJ_Normalizado'].dropna().unique())
            cnpjs_gralab = set(df_gralab['CNPJ_Normalizado'].dropna().unique())
            
            # Calcular intersecções
            cnpjs_comuns = cnpjs_nossos & cnpjs_gralab
            cnpjs_so_nossos = cnpjs_nossos - cnpjs_gralab
            cnpjs_so_gralab = cnpjs_gralab - cnpjs_nossos
            
            # ========================================
            # KPIs COMPARATIVOS
            # ========================================
            st.subheader("📊 Visão Geral Comparativa")
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                pct_comuns = (len(cnpjs_comuns) / len(cnpjs_nossos) * 100) if len(cnpjs_nossos) > 0 else 0
                st.metric(
                    label="🤝 Labs em Comum",
                    value=f"{len(cnpjs_comuns)}",
                    delta=f"{pct_comuns:.1f}% da nossa base"
                )
            
            with col2:
                pct_exclusivos_nossos = (len(cnpjs_so_nossos) / len(cnpjs_nossos) * 100) if len(cnpjs_nossos) > 0 else 0
                st.metric(
                    label="🔵 Exclusivos Nossos",
                    value=f"{len(cnpjs_so_nossos)}",
                    delta=f"{pct_exclusivos_nossos:.1f}% da nossa base"
                )
            
            with col3:
                pct_exclusivos_gralab = (len(cnpjs_so_gralab) / len(cnpjs_gralab) * 100) if len(cnpjs_gralab) > 0 else 0
                st.metric(
                    label="🟠 Exclusivos Gralab",
                    value=f"{len(cnpjs_so_gralab)}",
                    delta=f"{pct_exclusivos_gralab:.1f}% do Gralab"
                )
            
            with col4:
                st.metric(
                    label="📊 Total Gralab",
                    value=f"{len(cnpjs_gralab)}",
                    delta=f"vs {len(cnpjs_nossos)} nossos"
                )
            
            # ========================================
            # GRÁFICOS COMPARATIVOS
            # ========================================
            st.markdown("---")
            st.subheader("📈 Análise Visual")
            
            col_g1, col_g2 = st.columns(2)
            
            with col_g1:
                # Gráfico de Pizza - Distribuição
                import plotly.graph_objects as go
                
                fig_pizza = go.Figure(data=[go.Pie(
                    labels=['Em Comum', 'Só Nossos', 'Só Gralab'],
                    values=[len(cnpjs_comuns), len(cnpjs_so_nossos), len(cnpjs_so_gralab)],
                    marker=dict(colors=['#6BBF47', '#3B82F6', '#FB923C']),
                    hole=0.4,
                    textinfo='label+percent+value',
                    textposition='outside'
                )])
                
                fig_pizza.update_layout(
                    title="Distribuição de Laboratórios",
                    height=400,
                    showlegend=True
                )
                
                st.plotly_chart(fig_pizza, width='stretch')
            
            with col_g2:
                # Gráfico de Barras - Top UFs em Comum
                if len(cnpjs_comuns) > 0:
                    # Usar df completo para análise geográfica completa
                    df_comuns = df[df['CNPJ_Normalizado'].isin(cnpjs_comuns)]
                    top_ufs = df_comuns['Estado'].value_counts().head(10)
                    
                    fig_ufs = px.bar(
                        x=top_ufs.index,
                        y=top_ufs.values,
                        labels={'x': 'UF', 'y': 'Quantidade'},
                        title="Top 10 UFs com Labs em Comum",
                        color=top_ufs.values,
                        color_continuous_scale='Greens'
                    )
                    
                    fig_ufs.update_layout(
                        height=400,
                        showlegend=False,
                        xaxis_title="Estado",
                        yaxis_title="Quantidade de Laboratórios"
                    )
                    
                    st.plotly_chart(fig_ufs, width='stretch')
                else:
                    st.info("Nenhum laboratório em comum para análise geográfica")
            
            # ========================================
            # TABELAS DETALHADAS COM TABS
            # ========================================
            st.markdown("---")
            st.subheader("📋 Análise Detalhada")
            
            tab1, tab2, tab3, tab4, tab5 = st.tabs([
                "🤝 Labs em Comum",
                "🔵 Exclusivos Nossos",
                "🟠 Exclusivos Gralab",
                "🔄 Movimentações",
                "💰 Análise de Preços"
            ])
            
            with tab1:
                st.markdown("### 🤝 Laboratórios em Ambas as Bases")
                
                if len(cnpjs_comuns) > 0:
                    # Criar DataFrame combinado (usar df completo para pegar todos os labs)
                    df_comuns_nossos = df[df['CNPJ_Normalizado'].isin(cnpjs_comuns)][
                        ['CNPJ_PCL', 'CNPJ_Normalizado', 'Nome_Fantasia_PCL', 'Cidade', 'Estado']
                    ].drop_duplicates('CNPJ_Normalizado')
                    
                    # Selecionar colunas disponíveis do Gralab
                    colunas_gralab_desejadas = ['CNPJ_Normalizado', 'Nome', 'Cidade', 'UF', 'Preço CNH', 'Preço Concurso', 'Preço CLT']
                    colunas_gralab_disponiveis = [col for col in colunas_gralab_desejadas if col in df_gralab.columns]
                    
                    df_comuns_gralab = df_gralab[df_gralab['CNPJ_Normalizado'].isin(cnpjs_comuns)][
                        colunas_gralab_disponiveis
                    ]
                    
                    # Merge
                    df_comparacao = pd.merge(
                        df_comuns_nossos,
                        df_comuns_gralab,
                        on='CNPJ_Normalizado',
                        how='inner',
                        suffixes=('_Nosso', '_Gralab')
                    )
                    
                    # Filtros
                    col_f1, col_f2 = st.columns(2)
                    with col_f1:
                        ufs_disponiveis = ['Todos'] + sorted(df_comparacao['Estado'].dropna().unique().tolist())
                        uf_filtro = st.selectbox("Filtrar por UF:", ufs_disponiveis, key="uf_comuns")
                    
                    with col_f2:
                        if uf_filtro != 'Todos':
                            df_temp = df_comparacao[df_comparacao['Estado'] == uf_filtro]
                            cidades_disponiveis = ['Todas'] + sorted(df_temp['Cidade_Nosso'].dropna().unique().tolist())
                        else:
                            cidades_disponiveis = ['Todas'] + sorted(df_comparacao['Cidade_Nosso'].dropna().unique().tolist())
                        cidade_filtro = st.selectbox("Filtrar por Cidade:", cidades_disponiveis, key="cidade_comuns")
                    
                    # Aplicar filtros
                    df_exibir = df_comparacao.copy()
                    if uf_filtro != 'Todos':
                        df_exibir = df_exibir[df_exibir['Estado'] == uf_filtro]
                    if cidade_filtro != 'Todas':
                        df_exibir = df_exibir[df_exibir['Cidade_Nosso'] == cidade_filtro]
                    
                    # Selecionar colunas disponíveis para exibição
                    colunas_exibir = ['CNPJ_PCL', 'Nome_Fantasia_PCL']
                    if 'Nome' in df_exibir.columns:
                        colunas_exibir.append('Nome')
                    colunas_exibir.extend(['Cidade_Nosso', 'Estado'])
                    
                    # Adicionar colunas de preço se disponíveis
                    for col_preco in ['Preço CNH', 'Preço Concurso', 'Preço CLT']:
                        if col_preco in df_exibir.columns:
                            colunas_exibir.append(col_preco)
                    
                    # Renomear colunas para exibição
                    df_exibir_final = df_exibir[colunas_exibir].copy()
                    
                    rename_map = {
                        'CNPJ_PCL': 'CNPJ',
                        'Nome_Fantasia_PCL': 'Nome (Nossa Base)',
                        'Nome': 'Nome (Gralab/CunhaLab)',
                        'Cidade_Nosso': 'Cidade',
                        'Estado': 'UF',
                        'Preço CNH': 'Preço CNH (Gralab/CunhaLab)',
                        'Preço Concurso': 'Preço Concurso (Gralab/CunhaLab)',
                        'Preço CLT': 'Preço CLT (Gralab/CunhaLab)'
                    }
                    
                    df_exibir_final = df_exibir_final.rename(columns={k: v for k, v in rename_map.items() if k in df_exibir_final.columns})
                    
                    st.dataframe(df_exibir_final, width='stretch', height=400, hide_index=True)
                    
                    # Botões de download
                    col_d1, col_d2 = st.columns(2)
                    
                    with col_d1:
                        csv = df_exibir_final.to_csv(index=False, encoding='utf-8-sig')
                        st.download_button(
                            label="📥 Download CSV",
                            data=csv,
                            file_name=f"labs_em_comum_gralab_{datetime.now().strftime('%Y%m%d')}.csv",
                            mime="text/csv",
                            key="download_comuns_csv"
                        )
                    
                    with col_d2:
                        excel_buffer = BytesIO()
                        df_exibir_final.to_excel(excel_buffer, index=False, engine='openpyxl')
                        excel_data = excel_buffer.getvalue()
                        st.download_button(
                            label="📊 Download Excel",
                            data=excel_data,
                            file_name=f"labs_em_comum_gralab_{datetime.now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="download_comuns_excel"
                        )
                else:
                    st.info("Nenhum laboratório em comum encontrado")
            
            with tab2:
                st.markdown("### 🔵 Laboratórios Exclusivos da Nossa Base")
                st.caption("Laboratórios que temos mas o Gralab (CunhaLab) não tem - potencial para proteção")
                
                if len(cnpjs_so_nossos) > 0:
                    # Usar df completo para pegar todos os labs exclusivos nossos
                    df_exclusivos_nossos = df[df['CNPJ_Normalizado'].isin(cnpjs_so_nossos)][
                        ['CNPJ_PCL', 'Nome_Fantasia_PCL', 'Cidade', 'Estado', 'Vol_Hoje', 'Risco_Diario']
                    ].drop_duplicates('CNPJ_PCL')
                    
                    df_exclusivos_nossos = df_exclusivos_nossos.rename(columns={
                        'CNPJ_PCL': 'CNPJ',
                        'Nome_Fantasia_PCL': 'Nome',
                        'Estado': 'UF',
                        'Vol_Hoje': 'Volume Hoje',
                        'Risco_Diario': 'Risco'
                    })
                    
                    st.dataframe(df_exclusivos_nossos, width='stretch', height=400, hide_index=True)
                    
                    # Botões de download
                    col_d1, col_d2 = st.columns(2)
                    
                    with col_d1:
                        csv = df_exclusivos_nossos.to_csv(index=False, encoding='utf-8-sig')
                        st.download_button(
                            label="📥 Download CSV",
                            data=csv,
                            file_name=f"labs_exclusivos_nossos_{datetime.now().strftime('%Y%m%d')}.csv",
                            mime="text/csv",
                            key="download_exclusivos_nossos_csv"
                        )
                    
                    with col_d2:
                        excel_buffer = BytesIO()
                        df_exclusivos_nossos.to_excel(excel_buffer, index=False, engine='openpyxl')
                        excel_data = excel_buffer.getvalue()
                        st.download_button(
                            label="📊 Download Excel",
                            data=excel_data,
                            file_name=f"labs_exclusivos_nossos_{datetime.now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="download_exclusivos_nossos_excel"
                        )
                else:
                    st.info("Nenhum laboratório exclusivo encontrado")
            
            with tab3:
                st.markdown("### 🟠 Laboratórios Exclusivos do Gralab (CunhaLab)")
                st.caption("Laboratórios que o Gralab (CunhaLab) tem mas não temos - oportunidade de prospecção")
                
                if len(cnpjs_so_gralab) > 0:
                    # Filtrar labs exclusivos do Gralab
                    df_exclusivos_gralab = df_gralab[df_gralab['CNPJ_Normalizado'].isin(cnpjs_so_gralab)].copy()
                    
                    # Selecionar colunas disponíveis - sempre incluir CNPJ_Normalizado
                    colunas_disponiveis = []
                    
                    # Verificar se tem coluna CNPJ ou usar CNPJ_Normalizado
                    if 'CNPJ' in df_exclusivos_gralab.columns:
                        colunas_disponiveis.append('CNPJ')
                    elif 'CNPJ_Normalizado' in df_exclusivos_gralab.columns:
                        colunas_disponiveis.append('CNPJ_Normalizado')
                    
                    # Adicionar outras colunas desejadas
                    colunas_desejadas = ['Nome', 'Cidade', 'UF', 'Telefone', 'Preço CNH', 'Preço Concurso', 'Preço CLT']
                    
                    for col in colunas_desejadas:
                        if col in df_exclusivos_gralab.columns:
                            colunas_disponiveis.append(col)
                    
                    if colunas_disponiveis:
                        df_exclusivos_gralab_filtrado = df_exclusivos_gralab[colunas_disponiveis].copy()
                        
                        # Renomear CNPJ_Normalizado para CNPJ se necessário
                        if 'CNPJ_Normalizado' in df_exclusivos_gralab_filtrado.columns and 'CNPJ' not in df_exclusivos_gralab_filtrado.columns:
                            df_exclusivos_gralab_filtrado = df_exclusivos_gralab_filtrado.rename(columns={'CNPJ_Normalizado': 'CNPJ'})
                        
                        # Usar CNPJ para drop_duplicates
                        if 'CNPJ' in df_exclusivos_gralab_filtrado.columns:
                            df_exclusivos_gralab_filtrado = df_exclusivos_gralab_filtrado.drop_duplicates('CNPJ')
                        
                        st.dataframe(df_exclusivos_gralab_filtrado, width='stretch', height=400, hide_index=True)
                        
                        # Botões de download
                        col_d1, col_d2 = st.columns(2)
                        
                        with col_d1:
                            csv = df_exclusivos_gralab_filtrado.to_csv(index=False, encoding='utf-8-sig')
                            st.download_button(
                                label="📥 Download CSV",
                                data=csv,
                                file_name=f"labs_exclusivos_gralab_{datetime.now().strftime('%Y%m%d')}.csv",
                                mime="text/csv",
                                key="download_exclusivos_gralab_csv"
                            )
                        
                        with col_d2:
                            excel_buffer = BytesIO()
                            df_exclusivos_gralab_filtrado.to_excel(excel_buffer, index=False, engine='openpyxl')
                            excel_data = excel_buffer.getvalue()
                            st.download_button(
                                label="📊 Download Excel",
                                data=excel_data,
                                file_name=f"labs_exclusivos_gralab_{datetime.now().strftime('%Y%m%d')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="download_exclusivos_gralab_excel"
                            )
                    else:
                        st.warning("⚠️ Colunas esperadas não encontradas no arquivo do Gralab")
                else:
                    st.info("Nenhum laboratório exclusivo do Gralab (CunhaLab) encontrado")
            
            with tab4:
                st.markdown("### 🔄 Movimentações do Gralab (CunhaLab)")
                st.caption("Credenciamentos e descredenciamentos registrados")
                
                if 'EntradaSaida' in dados_gralab:
                    df_entrada_saida = dados_gralab['EntradaSaida'].copy()
                    
                    if not df_entrada_saida.empty:
                        # Formatar datas para padrão brasileiro
                        for col_data in ['Data Entrada', 'Data Saída', 'Última Verificação']:
                            if col_data in df_entrada_saida.columns:
                                df_entrada_saida[col_data] = pd.to_datetime(df_entrada_saida[col_data], errors='coerce').dt.strftime('%d/%m/%Y')
                                df_entrada_saida[col_data] = df_entrada_saida[col_data].replace('NaT', '')
                        
                        # Filtros
                        col_mov1, col_mov2 = st.columns(2)
                        
                        with col_mov1:
                            tipo_mov_filtro = st.multiselect(
                                "Tipo de Movimentação:",
                                options=['Todos', 'Credenciamento', 'Descredenciamento'],
                                default=['Todos'],
                                key="tipo_mov_gralab"
                            )
                        
                        with col_mov2:
                            if 'UF' in df_entrada_saida.columns:
                                uf_mov_filtro = st.multiselect(
                                    "UF:",
                                    options=['Todos'] + sorted(df_entrada_saida['UF'].dropna().unique().tolist()),
                                    default=['Todos'],
                                    key="uf_mov_gralab"
                                )
                        
                        # Aplicar filtros
                        df_mov_filtrado = df_entrada_saida.copy()
                        
                        if tipo_mov_filtro and 'Todos' not in tipo_mov_filtro:
                            df_mov_filtrado = df_mov_filtrado[df_mov_filtrado['Tipo Movimentação'].isin(tipo_mov_filtro)]
                        
                        if 'UF' in df_entrada_saida.columns and uf_mov_filtro and 'Todos' not in uf_mov_filtro:
                            df_mov_filtrado = df_mov_filtrado[df_mov_filtrado['UF'].isin(uf_mov_filtro)]
                        
                        # Selecionar colunas para exibição
                        colunas_exibir_mov = []
                        
                        # Sempre tentar incluir CNPJ
                        if 'CNPJ' in df_mov_filtrado.columns:
                            colunas_exibir_mov.append('CNPJ')
                        elif 'CNPJ_Normalizado' in df_mov_filtrado.columns:
                            colunas_exibir_mov.append('CNPJ_Normalizado')
                        
                        # Outras colunas importantes
                        colunas_desejadas_mov = [
                            'Nome', 'Cidade', 'UF', 'Data Entrada', 'Data Saída', 
                            'Tipo Movimentação', 'Última Verificação',
                            'Preço CNH', 'Preço Concurso', 'Preço CLT'
                        ]
                        
                        for col in colunas_desejadas_mov:
                            if col in df_mov_filtrado.columns:
                                colunas_exibir_mov.append(col)
                        
                        if colunas_exibir_mov:
                            df_mov_exibir = df_mov_filtrado[colunas_exibir_mov].copy()
                            
                            # Renomear CNPJ_Normalizado se necessário
                            if 'CNPJ_Normalizado' in df_mov_exibir.columns and 'CNPJ' not in df_mov_exibir.columns:
                                df_mov_exibir = df_mov_exibir.rename(columns={'CNPJ_Normalizado': 'CNPJ'})
                            
                            # Adicionar coluna indicando se é nosso cliente
                            def verificar_nosso_cliente(cnpj):
                                if pd.isna(cnpj) or cnpj == '':
                                    return 'N/A'
                                # Normalizar CNPJ para comparação
                                cnpj_normalizado = DataManager.normalizar_cnpj(str(cnpj))
                                if cnpj_normalizado in cnpjs_nossos:
                                    return '✅ Sim'
                                return '❌ Não'
                            
                            if 'CNPJ' in df_mov_exibir.columns:
                                df_mov_exibir['Nosso Cliente'] = df_mov_exibir['CNPJ'].apply(verificar_nosso_cliente)
                            
                            # Métricas resumidas
                            col_m1, col_m2, col_m3 = st.columns(3)
                            
                            total_movimentacoes = len(df_mov_exibir)
                            credenciamentos = len(df_mov_exibir[df_mov_exibir['Tipo Movimentação'] == 'Credenciamento']) if 'Tipo Movimentação' in df_mov_exibir.columns else 0
                            descredenciamentos = len(df_mov_exibir[df_mov_exibir['Tipo Movimentação'] == 'Descredenciamento']) if 'Tipo Movimentação' in df_mov_exibir.columns else 0
                            
                            with col_m1:
                                st.metric("Total Movimentações", total_movimentacoes)
                            with col_m2:
                                st.metric("✅ Credenciamentos", credenciamentos)
                            with col_m3:
                                st.metric("❌ Descredenciamentos", descredenciamentos)
                            
                            st.markdown("---")
                            
                            # Tabela
                            st.dataframe(df_mov_exibir, width='stretch', height=400, hide_index=True)
                            
                            # Botões de download
                            col_d1, col_d2 = st.columns(2)
                            
                            with col_d1:
                                csv = df_mov_exibir.to_csv(index=False, encoding='utf-8-sig')
                                st.download_button(
                                    label="📥 Download CSV",
                                    data=csv,
                                    file_name=f"movimentacoes_gralab_{datetime.now().strftime('%Y%m%d')}.csv",
                                    mime="text/csv",
                                    key="download_movimentacoes_gralab_csv"
                                )
                            
                            with col_d2:
                                excel_buffer = BytesIO()
                                df_mov_exibir.to_excel(excel_buffer, index=False, engine='openpyxl')
                                excel_data = excel_buffer.getvalue()
                                st.download_button(
                                    label="📊 Download Excel",
                                    data=excel_data,
                                    file_name=f"movimentacoes_gralab_{datetime.now().strftime('%Y%m%d')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key="download_movimentacoes_gralab_excel"
                                )
                        else:
                            st.warning("⚠️ Nenhuma coluna disponível para exibição")
                    else:
                        st.info("Nenhuma movimentação registrada")
                else:
                    st.info("Aba 'EntradaSaida' não encontrada no arquivo do Gralab (CunhaLab)")
            
            with tab5:
                st.markdown("### 💰 Análise de Preços (Labs em Comum)")
                
                if len(cnpjs_comuns) > 0:
                    df_precos = df_gralab[df_gralab['CNPJ_Normalizado'].isin(cnpjs_comuns)].copy()
                    
                    # Converter preços para numérico (limpando strings como 'R$ 120,00')
                    def _parse_preco_gralab(valor):
                        if pd.isna(valor):
                            return np.nan
                        if isinstance(valor, (int, float, np.number)):
                            return float(valor)
                        texto = str(valor).strip()
                        if not texto:
                            return np.nan
                        texto = texto.replace('R$', '').replace('r$', '')
                        texto = texto.replace(' ', '')
                        texto = texto.replace('.', '').replace(',', '.')
                        try:
                            return float(texto)
                        except Exception:
                            import re
                            numeros = re.findall(r'\d+[.,]?\d*', str(valor))
                            if numeros:
                                numero = numeros[0].replace('.', '').replace(',', '.')
                                try:
                                    return float(numero)
                                except Exception:
                                    return np.nan
                            return np.nan

                    for col in ['Preço CNH', 'Preço Concurso', 'Preço CLT']:
                        if col in df_precos.columns:
                            df_precos[col] = df_precos[col].apply(_parse_preco_gralab)
                    
                    # Estatísticas
                    col_s1, col_s2, col_s3 = st.columns(3)
                    
                    with col_s1:
                        st.markdown("#### 🎫 CNH")
                        if 'Preço CNH' in df_precos.columns:
                            precos_cnh = df_precos['Preço CNH'].dropna()
                            if len(precos_cnh) > 0:
                                st.metric("Média", f"R$ {precos_cnh.mean():.2f}")
                                st.metric("Mediana", f"R$ {precos_cnh.median():.2f}")
                                st.metric("Mín / Máx", f"R$ {precos_cnh.min():.2f} / R$ {precos_cnh.max():.2f}")
                            else:
                                st.info("Sem dados")
                    
                    with col_s2:
                        st.markdown("#### 📝 Concurso")
                        if 'Preço Concurso' in df_precos.columns:
                            precos_concurso = df_precos['Preço Concurso'].dropna()
                            if len(precos_concurso) > 0:
                                st.metric("Média", f"R$ {precos_concurso.mean():.2f}")
                                st.metric("Mediana", f"R$ {precos_concurso.median():.2f}")
                                st.metric("Mín / Máx", f"R$ {precos_concurso.min():.2f} / R$ {precos_concurso.max():.2f}")
                            else:
                                st.info("Sem dados")
                    
                    with col_s3:
                        st.markdown("#### 👔 CLT")
                        if 'Preço CLT' in df_precos.columns:
                            precos_clt = df_precos['Preço CLT'].dropna()
                            if len(precos_clt) > 0:
                                st.metric("Média", f"R$ {precos_clt.mean():.2f}")
                                st.metric("Mediana", f"R$ {precos_clt.median():.2f}")
                                st.metric("Mín / Máx", f"R$ {precos_clt.min():.2f} / R$ {precos_clt.max():.2f}")
                            else:
                                st.info("Sem dados")
                    
                    # Boxplot de distribuição
                    st.markdown("---")
                    st.markdown("#### 📊 Distribuição de Preços")
                    
                    # Preparar dados para boxplot
                    dados_boxplot = []
                    for col, nome in [('Preço CNH', 'CNH'), ('Preço Concurso', 'Concurso'), ('Preço CLT', 'CLT')]:
                        if col in df_precos.columns:
                            valores = df_precos[col].dropna()
                            for valor in valores:
                                dados_boxplot.append({'Tipo': nome, 'Preço': valor})
                    
                    if dados_boxplot:
                        df_boxplot = pd.DataFrame(dados_boxplot)
                        
                        fig_box = px.box(
                            df_boxplot,
                            x='Tipo',
                            y='Preço',
                            color='Tipo',
                            title="Distribuição de Preços por Tipo de Exame",
                            labels={'Preço': 'Preço (R$)', 'Tipo': 'Tipo de Exame'}
                        )
                        
                        fig_box.update_layout(height=400, showlegend=False)
                        st.plotly_chart(fig_box, width='stretch')
                    else:
                        st.info("Sem dados de preços disponíveis para análise")
                else:
                    st.info("Nenhum laboratório em comum para análise de preços")
    
    st.markdown("""
    <div class="footer">
        <p>📊 <strong>Syntox Churn</strong> - Dashboard profissional de análise de retenção de laboratórios</p>
        <p>Desenvolvido com ❤️ para otimizar a gestão de relacionamento com PCLs</p>
    </div>
    """, unsafe_allow_html=True)
if __name__ == "__main__":
    main()