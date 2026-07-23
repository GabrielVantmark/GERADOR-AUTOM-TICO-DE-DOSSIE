import datetime
import io
import re
import zipfile
from docx import Document
import pandas as pd
import streamlit as st

# 1. Configuração da página
st.set_page_config(
    page_title="Gerador Automático de Dossiê PLD-FT",
    page_icon="🤖",
    layout="wide",
)

# 2. Injeção de CSS
st.markdown(
    """
    <style>
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        padding-left: 3rem;
        padding-right: 3rem;
        max-width: 95% !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def gerar_codigo_dossie(indice):
    hoje = datetime.date.today().strftime("%Y%m%d")
    return f"DOS-{hoje}-{str(indice).zfill(3)}"


def formatar_data(valor):
    if pd.isna(valor) or not valor:
        return ""
    val_str = str(valor).strip()
    match = re.search(r"(\d{4})-(\d{2})-(\d{2})", val_str)
    if match:
        ano, mes, dia = match.groups()
        return f"{dia}/{mes}/{ano}"
    return val_str


def substituir_texto(doc_obj, mapa_substituicao):
    for p in doc_obj.paragraphs:
        for chave, valor in mapa_substituicao.items():
            if chave in p.text:
                for run in p.runs:
                    if chave in run.text:
                        run.text = run.text.replace(chave, str(valor))
                if chave in p.text:
                    p.text = p.text.replace(chave, str(valor))

    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                for chave, valor in mapa_substituicao.items():
                    if chave in cell.text:
                        for p in cell.paragraphs:
                            for run in p.runs:
                                if chave in run.text:
                                    run.text = run.text.replace(
                                        chave, str(valor)
                                    )
                            if chave in p.text:
                                p.text = p.text.replace(chave, str(valor))


# --- CABEÇALHO ---
col_logo, col_titulo = st.columns([1, 5], vertical_alignment="center")
with col_logo:
    try:
        st.image("noBgColor.png", width=160)
    except Exception:
        st.write("🤖")

with col_titulo:
    st.title("Gerador de Dossiês PLD-FT (Versão Inteligente)")
    st.subheader(
        "Upload de planilha + Complementação de Análise e Diligências"
    )

st.markdown("---")

# --- PASSO 1: UPLOAD ---
st.markdown("### 1. Selecione a Planilha de Apontamentos")
uploaded_file = st.file_uploader(
    "Arraste e solte o arquivo .xlsx ou .csv aqui:", type=["xlsx", "xls", "csv"]
)

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, dtype=str, skiprows=2)
        else:
            df = pd.read_excel(uploaded_file, dtype=str, skiprows=2)

        df.columns = [str(c).strip() for c in df.columns]
        df = df.fillna("")

        df["CODIGO_DOSSIE"] = [
            gerar_codigo_dossie(i + 1) for i in range(len(df))
        ]
        df["ID_Alerta"] = df.apply(
            lambda r: f"{r['CODIGO_DOSSIE']} | CPF/CNPJ: {r.get('Listas - CPF/CNPJ Pesquisado', 'N/A')}",
            axis=1,
        )

        st.success(
            f"✅ Planilha carregada! **{len(df)} registro(s)** identificados."
        )

        st.markdown("---")
        st.markdown("### 2. Análise e Emissão do Dossiê")

        alerta_selecionado = st.selectbox(
            "Selecione o registro para analisar e gerar:", df["ID_Alerta"].tolist()
        )

        linha = df[df["ID_Alerta"] == alerta_selecionado].iloc[0]

        # --- FORMULÁRIO COMPLEMENTAR DA V1 PARA EDITAR/VALIDAR OS DADOS ---
        with st.expander("📝 Detalhes do Alerta e Campos de Análise", expanded=True):
            col1, col2 = st.columns(2)

            with col1:
                st.markdown("#### 📌 Dados Extraídos da Planilha")
                cod_dossie = linha.get("CODIGO_DOSSIE", "")
                cpf_cnpj_pesquisado = linha.get(
                    "Listas - CPF/CNPJ Pesquisado", ""
                )
                nome_contraparte = linha.get("Listas - Nome Encontrado", "")
                tipologia = linha.get("Listas - Nome da Lista Pesquisada", "")
                data_detecao = formatar_data(
                    linha.get("Listas - Data Detecção", "")
                )

                st.text_input(
                    "Código de Rastreabilidade", cod_dossie, disabled=True
                )
                st.text_input(
                    "CPF/CNPJ Pesquisado", cpf_cnpj_pesquisado, disabled=True
                )
                st.text_input("Nome Encontrado", nome_contraparte, disabled=True)
                st.text_input("Lista / Tipologia", tipologia, disabled=True)

            with col2:
                st.markdown("#### ⚖️ Decisão de Análise e Diligências (V1)")

                analista = st.text_input("Analista Responsável", "Analista PLD")
                data_analise = st.date_input(
                    "Data da Análise", datetime.date.today()
                ).strftime("%d/%m/%Y")

                status_alerta = st.selectbox(
                    "Conclusão / Decisão de Arquivamento:",
                    [
                        "Arquivado - Sem Indício de Irregularidade",
                        "Arquivado - Falso Positivo",
                        "Encaminhado para Comunicação (COAF)",
                        "Em Monitoramento",
                    ],
                )

                diligencias_opcoes = st.multiselect(
                    "Diligências Realizadas:",
                    [
                        "Consulta Mídia Negativa",
                        "Pesquisa de Bens / Cartório",
                        "Consulta Base Pública (Receita / Sanções / CEIS)",
                        "Solicitação de Esclarecimentos ao Cliente",
                        "Verificação de Vínculos / Relacionamento",
                    ],
                    default=[
                        "Consulta Base Pública (Receita / Sanções / CEIS)"
                    ],
                )

                justificativa = st.text_area(
                    "Justificativa da Decisão:",
                    value="Análise realizada com base nos apontamentos identificados. Não foram constatados indícios que justifiquem a comunicação atípica.",
                    height=100,
                )

        # Prepara string das diligências para o Word
        diligencias_str = (
            "\n".join([f"- {d}" for d in diligencias_opcoes])
            if diligencias_opcoes
            else "Nenhuma diligência adicional registrada."
        )

        if st.button("🚀 Gerar Dossiê Completo (.docx)"):
            doc = Document("modelo_dossie.docx")

            dicionario_dados = {
                "{{CODIGO_DOSSIE}}": cod_dossie,
                "{{NUM_ALERTA}}": cpf_cnpj_pesquisado,
                "{{DATA_GERACAO}}": data_detecao,
                "{{ANALISTA}}": analista,
                "{{DATA_ANALISE}}": data_analise,
                "{{SISTEMA}}": "Advice e-Guardian",
                "{{STATUS_ALERTA}}": status_alerta,
                "{{TIPOLOGIA}}": tipologia,
                "{{REGRA}}": "Apontamento em Listas Restritivas",
                "{{NORMATIVA}}": "Lei nº 9.613/1998 e Resolução BCB nº 96/2021",
                "{{NOME_CONTRAPARTE}}": nome_contraparte,
                "{{CPF_CNPJ}}": linha.get("Listas - CPF/CNPJ Encontrado", ""),
                "{{OBS_CONTRAPARTE}}": f"Vinculação: {linha.get('Listas - Parte Relac', '')} | Grupo: {linha.get('Listas - Grupo Atuação', '')}",
                "{{DILIGENCIAS}}": diligencias_str,
                "{{JUSTIFICATIVA}}": justificativa,
            }

            substituir_texto(doc, dicionario_dados)

            nome_arquivo = f"{cod_dossie}_Dossie_PLD.docx"
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)

            st.download_button(
                label=f"📥 Baixar {cod_dossie} (.docx)",
                data=buffer,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

    except Exception as e:
        st.error(f"Erro ao processar a planilha: {e}")
