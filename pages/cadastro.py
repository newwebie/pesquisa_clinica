import os
import pandas as pd
import streamlit as st
import psycopg2
from psycopg2.extras import RealDictCursor


# =========================
# Config / Conexão com Banco
# =========================

def get_connection():
    """
    Cria uma conexão com o Postgres.
    Ajuste os valores padrão ou use variáveis de ambiente.
    """
    return psycopg2.connect(
        host=os.getenv("DB_HOST", "localhost"),
        port=os.getenv("DB_PORT", "5432"),
        dbname=os.getenv("DB_NAME", "meubanco"),
        user=os.getenv("DB_USER", "postgres"),
        password=os.getenv("DB_PASSWORD", "postgres"),
    )


# =========================
# Helpers
# =========================

def snake_to_camel(name: str) -> str:
    """
    Converte snake_case -> camelCase.
    Ex: criado_por_nome -> criadoPorNome
    """
    if not isinstance(name, str):
        return name
    parts = name.split("_")
    if not parts:
        return name
    return parts[0] + "".join(p.capitalize() for p in parts[1:])


def snake_to_title(name: str) -> str:
    """
    Alias para manter compatível com seu código original.
    Aqui usamos camelCase na UI.
    """
    return snake_to_camel(name)


def _load_desvios_into_session() -> bool:
    """
    Carrega os desvios do banco e guarda em session_state
    como 'desvios_df_full' e 'desvios_df_original_full'.

    Retorna True em caso de sucesso, False em caso de erro.
    """
    if "desvios_df_full" in st.session_state and "desvios_df_original_full" in st.session_state:
        return True

    try:
        conn = get_connection()
        cursor = conn.cursor(cursor_factory=RealDictCursor)

        # SELECT *, incluindo xmin como row_version para controle de concorrência
        cursor.execute("SELECT *, xmin AS row_version FROM desvios ORDER BY id;")
        rows = cursor.fetchall()

        df = pd.DataFrame(rows)

        st.session_state["desvios_df_full"] = df.copy()
        st.session_state["desvios_df_original_full"] = df.copy()
        return True

    except Exception as e:
        st.error(f"Erro ao carregar lista de desvios do banco: {e}")
        return False
    finally:
        try:
            cursor.close()
            conn.close()
        except Exception:
            pass


# =========================
# Autenticação por E-mail
# =========================

def login_screen():
    st.title("🔐 Login - Lista de Desvios")
    st.write("Informe seu e-mail cadastrado na tabela **usuarios** para acessar a lista de desvios.")

    email = st.text_input("E-mail", placeholder="seu.email@empresa.com")

    if st.button("Entrar"):
        if not email:
            st.warning("Por favor, informe um e-mail.")
            return

        try:
            conn = get_connection()
            cursor = conn.cursor(cursor_factory=RealDictCursor)

            # Busca por e-mail (case-insensitive)
            cursor.execute(
                """
                SELECT id, email, display_name
                FROM usuarios
                WHERE LOWER(email) = LOWER(%s)
                """,
                (email.strip(),),
            )
            user = cursor.fetchone()

            if not user:
                st.error("E-mail não encontrado na tabela de usuários.")
                return

            # Guarda informações básicas na sessão
            st.session_state["is_authenticated"] = True
            st.session_state["user_email"] = user["email"]
            st.session_state["user_display_name"] = user.get("display_name") or user["email"]

            st.success(f"Bem-vindo(a), {st.session_state['user_display_name']}!")
            st.rerun()

        except Exception as e:
            st.error(f"Erro ao autenticar: {e}")
        finally:
            try:
                cursor.close()
                conn.close()
            except Exception:
                pass


# =========================
# Página Lista de Desvios
# =========================

def lista_desvios_page():
    display_name = st.session_state.get("user_display_name", "")

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
        "**Salvar**."
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
        except Exception:
            pass


# =========================
# Main
# =========================

def main():
    st.set_page_config(
        page_title="Lista de Desvios",
        page_icon="📋",
        layout="wide",
    )

    # Inicializa flag de autenticação
    if "is_authenticated" not in st.session_state:
        st.session_state["is_authenticated"] = False

    # Se não autenticado, mostra tela de login
    if not st.session_state["is_authenticated"]:
        login_screen()
        return

    # Barra lateral com info do usuário + logout
    with st.sidebar:
        st.markdown("### 👤 Usuário logado")
        st.write(st.session_state.get("user_display_name", ""))
        st.caption(st.session_state.get("user_email", ""))

        if st.button("Sair"):
            # Limpa tudo relacionado à sessão
            for k in [
                "is_authenticated",
                "user_email",
                "user_display_name",
                "desvios_df_full",
                "desvios_df_original_full",
            ]:
                st.session_state.pop(k, None)
            st.rerun()

    # Conteúdo principal
    lista_desvios_page()


if __name__ == "__main__":
    main()
