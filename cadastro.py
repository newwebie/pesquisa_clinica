import streamlit as st
import pandas as pd
from pathlib import Path

st.title("Cadastro de desvio Clínico")


# Campos do formulário
with st.form("cadastro_desvio_form"):
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
    }

    df_registro = pd.DataFrame([registro])

    arquivo_registros = Path("cadastro_registros.csv")
    if arquivo_registros.exists():
        df_registro.to_csv(arquivo_registros, mode="a", header=False, index=False)
    else:
        df_registro.to_csv(arquivo_registros, index=False)

    st.success("Registro salvo com sucesso!")
    st.dataframe(df_registro)



# 1. conferir se os campos abaixo realmente nunca terão novas opções
# Categoria, Visita, (centro já foi falado que muda), Subcategorias, Codigo, Escopo

# 2. Conferir se nenhum selectbox é multiselect

# 3. array de campos obrigatórios [participante, data_desvio, centro]
# 3.1 validar se os campos obrigatórios estão corretos (provavelmente tem que adicionar mais)

