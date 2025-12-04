"""Tela inicial com autenticação via Microsoft e navegação principal em 3 páginas via sidebar."""
from __future__ import annotations
import pandas as pd
from pathlib import Path
from typing import Any, Dict
import psycopg2
import streamlit as st

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

# -------------------------------------------------
# Estilos: esconder SOMENTE a navegação padrão do Streamlit
# (mantendo a sidebar visível para nossa navegação)
# -------------------------------------------------
HIDE_SIDEBAR_NAVIGATION = """
<style>
    /* Esconde o seletor de páginas padrão e o botão de "collapse" */
    [data-testid="stSidebarNav"] {display: none;}
    [data-testid="collapsedControl"] {display: none;}
</style>
"""
st.markdown(HIDE_SIDEBAR_NAVIGATION, unsafe_allow_html=True)





def is_admin_in_db(email: str) -> bool:
    """Verifica no banco se o e-mail existe na tabela usuarios."""
    if not email:
        return False

    try:
        conn = get_connection()
        cursor = conn.cursor()

        # ajuste o nome da tabela/coluna se for diferente
        query = """
            SELECT 1
            FROM usuarios
            WHERE LOWER(email) = %s
            LIMIT 1
        """
        cursor.execute(query, (email.lower(),))
        exists = cursor.fetchone() is not None

        cursor.close()
        conn.close()
        return exists

    except Exception as e:
        # Em produção talvez você queira logar isso em vez de mostrar na tela
        st.error(f"Erro ao verificar permissões no banco: {e}")
        return False



# CONEXÃO COM O BANCO VIA secrets.toml
def get_connection():
    db = st.secrets["postgres"]
    return psycopg2.connect(
        host=db["host"],
        port=db["port"],
        dbname=db["database"],
        user=db["user"],
        password=db["password"],
    )


### CARREGANDO INFO DO BANCO DE DADOS DOS DESVIOS ###
def _load_desvios_into_session() -> bool:
    """Carrega os desvios do banco para o session_state, se ainda não estiverem lá."""
    if "desvios_df_full" in st.session_state and "desvios_df_original_full" in st.session_state:
        return True  # já carregado

    try:
        conn = get_connection()
        query = """
            SELECT
                id,
                participante,
                data_ocorrido,
                formulario_status,
                identificacao_desvio,
                centro,
                visita,
                descricao_desvio,
                causa_raiz,
                acao_preventiva,
                acao_corretiva,
                importancia,
                data_identificacao_texto,
                categoria,
                subcategoria,
                codigo,
                escopo,
                avaliacao_gerente_medico,
                avaliacao_investigador,
                formulario_arquivado,
                recorrencia,
                num_ocorrencia_previa,
                prazo_escalonamento,
                data_escalonamento,
                atendeu_prazos_report,
                populacao,
                data_submissao_cep,
                data_finalizacao,
                criado_por_nome,
                criado_por_email,
                atualizado_por,
                xmin AS row_version
            FROM desvios
            ORDER BY id DESC
        """
        df = pd.read_sql_query(query, conn)

        st.session_state["desvios_df_full"] = df
        st.session_state["desvios_df_original_full"] = df.copy()

        return True

    except Exception as e:
        st.error(f"Erro ao carregar desvios do banco: {e}")
        return False

    finally:
        try:
            conn.close()
        except:
            pass

# ajustar o titulo das colunas
def snake_to_title(name: str) -> str:
    """Converte snake_case para 'Title Case' com espaços."""
    parts = name.split("_")
    parts = [p.capitalize() for p in parts]
    return " ".join(parts)





# -------------------------------------------------
# Blocos de "páginas" internas (rascunhos)
# -------------------------------------------------
def render_pagina_cadastro(user_email: str, display_name: str) -> None:

    # Campos do formulário
    with st.form("cadastro_desvio_form"):
        st.subheader("📋 Cadastrar Desvio")
        nome_paciente = st.text_input("Participante")
        data_desvio = st.date_input("Data do Ocorrido", format="DD/MM/YYYY")
        formulario = st.selectbox("Formulário", ["", "Sim", "Pendente", "N/A"])
        identiicacao_desvio = st.text_input("Identificação do Desvio")
        centro = st.text_input(
            "Centro",
            placeholder="Aqui você pode colocar a lógica para selecionar o centro, puxando do Excel",
        )
        visita = st.text_input(
            "Visita",
            placeholder="Aqui você pode colocar a lógica para selecionar o tipo de visita, puxando do Excel",
        )
        desvio = st.text_area("Desvio")
        causa_raiz = st.text_input("Causa Raiz")
        acao_preventiva = st.text_input("Ação Preventiva")
        acao_corretiva = st.text_input("Ação Corretiva")
        importancia = st.selectbox("Importância", ["", "Maior", "Menor"])
        data_identificacao = st.text_input(
            "Data de Identificação",
            placeholder="Ex.: MOV01-2005 (01/01-07/01)",
        )
        categoria = st.selectbox(
            "Categoria",
            [
                "",
                "Avaliações",
                "Consentimento informado",
                "Procedimentos",
                "PSI",
                "Segurança",
                "Outros",
            ],
        )
        subcategorias = st.selectbox(
            "Subcategoria",
            [
                "",
                "Avaliações Perdidas ou não realizadas",
                "Avaliações realizadas fora da janela",
                "Desvios Recorrentes",
                "Visitas perdidas ou fora da janela",
                "Amostras Laboratoriais",
                "Critérios de Inclusão / Exclusão (Elegibilidade)",
                "Outros",
            ],
        )
        codigo = st.selectbox("Código", ["", "A8", "A7", "O4"])
        escopo = st.selectbox("Escopo", ["", "Protocolo", "GCP"])
        avaliacao_medico = st.selectbox("Avaliação Gerente Médico", ["", "Escolha 1", "Escolha 2", "Escolha 3"])
        avaliacao_investigador = st.selectbox(
            "Avaliação Investigador Principal", ["", "Escolha 1", "Escolha 2", "Escolha 3"]
        )
        arquivado = st.selectbox("Formulário Arquivado (ISF e TFM)?", ["", "Sim", "Não", "N/A"])
        recorrencia = st.selectbox("Recorrência", ["", "Recorrente", "Não Recorrente", "Isolado"])
        st.caption("Não-Recorrente (até 3x), Recorrente (>3x)* *Mesmo grau de importância e causa raiz")
        ocorrencia_previa = st.number_input("N° Desvio Ocorrência Prévia", min_value=0, step=1)
        st.caption("Mesmo grau de importância e causa raiz")
        prazo_escalonamento = st.selectbox("Prazo para Escalonamento", ["", "Imediata", "Mensal", "Padrão"])
        data_escalonamento = st.date_input("Data de Escalonamento", format="DD/MM/YYYY")
        prazo_report = st.selectbox("Atendeu os Prazos de Report?", ["", "Sim", "Não"])
        populacao = st.selectbox("População", ["", "Intenção de Tratar (ITT)", "Por Protocolo (PP)"])
        data_cep = st.date_input("Data de Submissão ao CEP", format="DD/MM/YYYY")
        data_finalizacao = st.date_input("Data de Finalização", format="DD/MM/YYYY")
        anexos = st.file_uploader("Attachments", accept_multiple_files=True)
        atualizado_por = None

        submit = st.form_submit_button("Salvar registro")


    if submit:
        anexos_nomes = [uploaded_file.name for uploaded_file in anexos] if anexos else []
        registro = {
            "Participante": nome_paciente,
            "Data do Ocorrido": data_desvio,
            "Formulário": formulario,
            "Identificação do Desvio": identiicacao_desvio,
            "Centro": centro,
            "Visita": visita,
            "Desvio": desvio,
            "Causa Raiz": causa_raiz,
            "Ação Preventiva": acao_preventiva,
            "Ação Corretiva": acao_corretiva,
            "Importância": importancia,
            "Data de Identificação": data_identificacao,
            "Categoria": categoria,
            "Subcategoria": subcategorias,
            "Código": codigo,
            "Escopo": escopo,
            "Avaliação Gerente Médico": avaliacao_medico,
            "Avaliação Investigador Principal": avaliacao_investigador,
            "Formulário Arquivado (ISF e TFM)?": arquivado,
            "Recorrência": recorrencia,
            "N° Desvio Ocorrência Prévia": ocorrencia_previa,
            "Prazo para Escalonamento": prazo_escalonamento,
            "Data de Escalonamento": data_escalonamento,
            "Atendeu os Prazos de Report?": prazo_report,
            "População": populacao,
            "Data de Submissão ao CEP": data_cep,
            "Data de Finalização": data_finalizacao,
            "Attachments": ", ".join(anexos_nomes),
                # 👇 rastreabilidade
            "Criado por (nome)": display_name,
            "Criado por (email)": user_email,
            "Atualizado por": atualizado_por

        }

        df_registro = pd.DataFrame([registro])

        try:
            conn = get_connection()
            cursor = conn.cursor()

            # 🔁 Agora alinhado com os nomes da tabela:
            # id,"participante","data_ocorrido","formulario_status","identificacao_desvio","centro",
            # "visita","descricao_desvio","causa_raiz","acao_preventiva","acao_corretiva",
            # "importancia","data_identificacao_texto","categoria","subcategoria","codigo","escopo",
            # "avaliacao_gerente_medico","avaliacao_investigador","formulario_arquivado","recorrencia",
            # "num_ocorrencia_previa","prazo_escalonamento","data_escalonamento","atendeu_prazos_report",
            # "populacao","data_submissao_cep","data_finalizacao","criado_por_id","criado_em"

            insert_query = """
            INSERT INTO desvios (
                participante,
                data_ocorrido,
                formulario_status,
                identificacao_desvio,
                centro,
                visita,
                descricao_desvio,
                causa_raiz,
                acao_preventiva,
                acao_corretiva,
                importancia,
                data_identificacao_texto,
                categoria,
                subcategoria,
                codigo,
                escopo,
                avaliacao_gerente_medico,
                avaliacao_investigador,
                formulario_arquivado,
                recorrencia,
                num_ocorrencia_previa,
                prazo_escalonamento,
                data_escalonamento,
                atendeu_prazos_report,
                populacao,
                data_submissao_cep,
                data_finalizacao,
                criado_por_nome,
                criado_por_email,
                atualizado_por
 
            )
            VALUES (
                %s, %s, %s, %s,
                %s, %s, %s, %s,
                %s, %s, %s, %s,
                %s, %s, %s, %s,
                %s, %s, %s, %s,
                %s, %s, %s, %s,
                %s, %s, %s, %s, 
                %s, %s
            )
            """

            values = (
                nome_paciente,              # participante
                data_desvio,                # data_ocorrido (DATE)
                formulario,                 # formulario_status
                identiicacao_desvio,        # identificacao_desvio
                centro,                     # centro
                visita,                     # visita
                desvio,                     # descricao_desvio
                causa_raiz,                 # causa_raiz
                acao_preventiva,            # acao_preventiva
                acao_corretiva,             # acao_corretiva
                importancia,                # importancia
                data_identificacao,         # data_identificacao_texto (texto mesmo)
                categoria,                  # categoria
                subcategorias,              # subcategoria
                codigo,                     # codigo
                escopo,                     # escopo
                avaliacao_medico,           # avaliacao_gerente_medico
                avaliacao_investigador,     # avaliacao_investigador
                arquivado,                  # formulario_arquivado
                recorrencia,                # recorrencia
                int(ocorrencia_previa) if ocorrencia_previa is not None else None,  # num_ocorrencia_previa
                prazo_escalonamento,        # prazo_escalonamento
                data_escalonamento,         # data_escalonamento
                prazo_report,               # atendeu_prazos_report
                populacao,                  # populacao
                data_cep,                   # data_submissao_cep
                data_finalizacao,            # data_finalizacao
                display_name,                # criado_por_email
                user_email,
                atualizado_por              # criado_por_nome   
            )

            cursor.execute(insert_query, values)
            conn.commit()

            st.success("Desvio Salvo com Sucesso! ✅")

        except Exception as e:
            st.error(f"Erro ao salvar no banco: {e}")

        finally:
            try:
                cursor.close()
                conn.close()
                st.rerun()
            except:
                pass





def render_pagina_adm_desvios(is_admin: bool) -> None:
    """Rascunho da página de Administração de Desvios (apenas admins)."""
    st.subheader("🛠 Panel ADM")

    if not is_admin:
        st.error("Acesso restrito aos administradores cadastrados.")
        st.info(
            "Se você precisar de acesso administrativo, solicite inclusão do seu e-mail "
            "na whitelist de administradores."
        )
        return

    st.write(
        "Área dedicada para **análises, ajustes e aprovação de desvios**. "
        "Aqui você poderá revisar registros, alterar status, registrar pareceres, etc."
    )

    col1, col2 = st.columns([2, 1])
    with col1:
        st.markdown("**📋 Lista de Desvios**")
        st.dataframe(
            {
                "ID": ["D-001", "D-002"],
                "Título": ["Falha na coleta", "Problema no transporte"],
                "Status": ["Em análise", "Corrigido"],
            },
            use_container_width=True,
        )
    with col2:
        st.markdown("**🔎 Filtros (placeholder):**")
        st.selectbox("Status", ["Todos", "Novo", "Em análise", "Corrigido", "Encerrado"])
        st.selectbox("Criticidade", ["Todas", "Baixa", "Média", "Alta"])

    st.button("Aplicar filtros (placeholder)", disabled=True)


def render_pagina_relatorios() -> None:
    """Rascunho da página de Relatórios & Auditoria."""
    st.subheader("📑 Relatórios & Auditoria")
    st.write(
        "Central para geração de **relatórios, exportações e trilhas de auditoria** "
        "dos desvios registrados no sistema."
    )

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**📅 Relatórios por Período (placeholder):**")
        st.date_input("Data inicial")
        st.date_input("Data final")
        st.multiselect("Status", ["Novo", "Em análise", "Corrigido", "Encerrado"])
        st.button("Gerar Relatório (placeholder)", disabled=True)

    with col2:
        st.markdown("**🧾 Exportações (placeholder):**")
        st.checkbox("Incluir dados sensíveis (apenas usuários autorizados)")
        st.radio("Formato", ["CSV", "XLSX", "PDF"], horizontal=True)
        st.button("Exportar (placeholder)", disabled=True)


def render_pagina_desvios(display_name: str) -> None:
    """
    Página de Lista de Desvios:
    - tabela editável
    - session_state
    - controle de concorrência com xmin
    - colunas criadasPor* não editáveis
    - esconde regex_id
    - camelCase na UI
    - salva atualizado_por com o display_name
    """
    st.subheader("📋 Lista de Desvios")

    # Carrega dados uma vez na sessão
    if not _load_desvios_into_session():
        return

    df_full = st.session_state["desvios_df_full"]

    if df_full.empty:
        st.info("Nenhum desvio cadastrado ainda.")
        if st.button("🔄 Recarregar dados"):
            for k in ["desvios_df_full", "desvios_df_original_full"]:
                st.session_state.pop(k, None)
            st.rerun()
        return

    # Remover da visualização: row_version e regex_id
    df_for_display = df_full.drop(columns=["row_version", "regex_id"], errors="ignore")

    # Montar mapeamento snake_case -> camelCase para as colunas exibidas
    snake_cols = list(df_for_display.columns)
    snake_to_camel_map = {col: snake_to_title(col) for col in snake_cols}
    camel_to_snake_map = {v: k for k, v in snake_to_camel_map.items()}

    display_df = df_for_display.rename(columns=snake_to_camel_map)

    # Nomes camelCase das colunas que queremos travar na UI
    non_editable_cols = [
        "id",
        "criadoPorNome",
        "criadoPorEmail",
        "atualizadoPor",
    ]

    st.caption(
        "Clique nas células para editar. As alterações só são salvas no banco ao clicar em "
        "**Salvar alterações.**"
    )

    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("🔄 Recarregar dados"):
            for k in ["desvios_df_full", "desvios_df_original_full"]:
                st.session_state.pop(k, None)
            st.rerun()

    edited_display_df = st.data_editor(
        display_df,
        num_rows="fixed",
        use_container_width=True,
        key="desvios_editor",
        disabled=[c for c in non_editable_cols if c in display_df.columns],
    )

    with col2:
        salvar = st.button("💾 Salvar", type="primary")

    if not salvar:
        return

    # Converter de volta de camelCase para snake_case para comparar / salvar
    edited_snake_df = edited_display_df.rename(columns=camel_to_snake_map)

    df_original_full = st.session_state["desvios_df_original_full"]

    if "id" not in edited_snake_df.columns:
        st.error("A coluna 'id' é obrigatória para atualizar os registros.")
        return

    # Índice por ID
    edited_idx = edited_snake_df.set_index("id")
    original_full_idx = df_original_full.set_index("id")

    # Colunas de dados que aparecem na tabela (sem id, row_version, regex_id)
    data_columns = [
        c
        for c in edited_snake_df.columns
        if c not in ("id", "row_version", "regex_id")
    ]

    # Para detecção de diff, ignoramos row_version/regex_id
    original_for_compare = original_full_idx[data_columns]

    # Máscara de diferenças
    try:
        diffs_mask = edited_idx[data_columns].ne(original_for_compare).any(axis=1)
    except Exception:
        # se der treta de tipo, assume que tudo mudou
        diffs_mask = pd.Series(True, index=edited_idx.index)

    ids_alterados = edited_idx.index[diffs_mask].tolist()

    if not ids_alterados:
        st.info("Nenhuma alteração detectada para salvar.")
        return

    # Colunas que serão atualizadas no banco:
    # - todas as colunas de dados
    # - exceto criadores
    # - e vamos garantir que atualizado_por SEMPRE seja setado com o display_name
    base_update_columns = [
        c
        for c in data_columns
        if c not in ("criado_por_nome", "criado_por_email", "atualizado_por")
    ]
    columns_to_update = base_update_columns + ["atualizado_por"]

    set_clause = ", ".join(f"{col} = %s" for col in columns_to_update)
    update_sql = f"""
        UPDATE desvios
        SET {set_clause}
        WHERE id = %s
          AND xmin = %s::xid       -- 👈 controle de concorrência
    """

    conflitos = []
    atualizados = 0

    try:
        conn = get_connection()
        cursor = conn.cursor()

        for row_id in ids_alterados:
            row_editada = edited_idx.loc[row_id]
            row_original = original_full_idx.loc[row_id]

            row_version = row_original["row_version"]

            # valores das colunas editáveis normais
            values = [row_editada[col] for col in base_update_columns]

            # atualizado_por sempre com o usuário logado
            values.append(display_name)

            # WHERE id & xmin
            values.append(row_id)
            values.append(row_version)

            cursor.execute(update_sql, values)

            if cursor.rowcount == 0:
                conflitos.append(row_id)
            else:
                atualizados += 1

        conn.commit()

        if atualizados:
            st.success(f"{atualizados} registro(s) atualizado(s) com sucesso! ✅")

        if conflitos:
            st.warning(
                f"{len(conflitos)} registro(s) NÃO foram salvos porque foram alterados "
                "por outra pessoa depois que você carregou a página. "
                "Clique em **Recarregar dados** para ver a versão mais recente."
            )

        # Recarrega do banco para pegar novos row_version e atualizado_por
        for k in ["desvios_df_full", "desvios_df_original_full"]:
            st.session_state.pop(k, None)
        _load_desvios_into_session()

    except Exception as e:
        st.error(f"Erro ao salvar alterações no banco: {e}")
    finally:
        try:
            cursor.close()
            conn.close()
        except:
            pass


# -------------------------------------------------
# Autenticação e contexto do usuário
# -------------------------------------------------
auth = MicrosoftAuth()


logged_in = create_login_page(auth)
if not logged_in:
    st.stop()


# Garantir token válido durante a sessão
AuthManager.check_and_refresh_token(auth)
create_user_header()

user = AuthManager.get_current_user() or {}
display_name = user.get("displayName", "Usuário")
user_email = (user.get("mail") or user.get("userPrincipalName") or "").lower()


st.session_state["display_name"] = display_name
st.session_state["user_email"] = user_email


if "is_admin" not in st.session_state:
    # só consulta no banco UMA vez por sessão
    st.session_state.is_admin = is_admin_in_db(user_email)

is_admin = st.session_state.is_admin

# -------------------------------------------------
# Definição das páginas e estado de navegação
# -------------------------------------------------
PAGES = {
    "📥 Cadastro": "cadastro",
    "📋 Lista Desvios": "lista_desvios",
    "📑 Relatórios & Auditoria": "relatorios",
    "🛠 Painel Adm": "adm_desvios",
}

if "page" not in st.session_state:
    # primeira página como padrão
    st.session_state.page = list(PAGES.values())[0]

# -------------------------------------------------
# Sidebar: SOMENTE navegação entre páginas
# -------------------------------------------------
with st.sidebar:
    st.markdown("### 🧭 Navegação")

    for label, page_key in PAGES.items():
        if st.button(label, key=f"nav_{page_key}", use_container_width=True):
            st.session_state.page = page_key
            st.rerun()

# -------------------------------------------------
# Cabeçalho principal
# -------------------------------------------------
st.title("Portal Pesquisa Clínica")

# -------------------------------------------------
# Renderização da página selecionada
# -------------------------------------------------
current_page = st.session_state.page

if current_page == "cadastro":
    render_pagina_cadastro(
        user_email=user_email,
        display_name=display_name,
    )
elif current_page == "adm_desvios":
    render_pagina_adm_desvios(is_admin=is_admin)
elif current_page == "relatorios":
    render_pagina_relatorios()
elif current_page == "lista_desvios":
    render_pagina_desvios(display_name=display_name)

# Mensagem extra se não for admin (opcional, você pode remover)
if not is_admin:
    st.info(
        "Se precisar de acesso à área administrativa, solicite ao responsável "
        "a inclusão do seu e-mail na lista de administradores."
    )
