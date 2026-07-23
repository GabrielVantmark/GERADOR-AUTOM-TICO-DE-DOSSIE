import datetime
import io
import re
import zipfile
from docx import Document
import pandas as pd
import streamlit as st

# 1. Configuração da página
st.set_page_config(
    page_title="Gerador de Dossiê PLD-FT (Automático)",
    page_icon="🤖",
    layout="wide",
)

# 2. Injeção de CSS para expansão da tela e remoção das margens largas
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


# --- FUNÇÕES AUXILIARES ---
def gerar_codigo_dossie(indice):
    """Gera um código único no padrão: DOS-YYYYMMDD-001"""
    hoje = datetime.date.today().strftime("%Y%m%d")
    sequencial = str(indice).zfill(3)
    return f"DOS-{hoje}-{sequencial}"


def formatar_data(valor):
    """Converte datas no formato YYYY-MM-DD para DD/MM/YYYY"""
    if pd.isna(valor) or not valor:
        return ""
    val_str = str(valor).strip()
    match = re.search(r"(\d{4})-(\d{2})-(\d{2})", val_str)
    if match:
        ano, mes, dia = match.groups()
        return f"{dia}/{mes}/{ano}"
    return val_str


def substituir_texto(doc_obj, mapa_substituicao):
    """Substitui placeholders mantendo a formatação original em parágrafos e tabelas."""
    # Substituição em parágrafos
    for p in doc_obj.paragraphs:
        for chave, valor in mapa_substituicao.items():
            if chave in p.text:
                for run in p.runs:
                    if chave in run.text:
                        run.text = run.text.replace(chave, str(valor))
                if chave in p.text:
                    p.text = p.text.replace(chave, str(valor))

    # Substituição dentro de tabelas
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


# --- CABEÇALHO DA PÁGINA ---
col_logo, col_titulo = st.columns([1, 5], vertical_alignment="center")

with col_logo:
    try:
        st.image("noBgColor.png", width=160)
    except Exception:
        st.write("🤖")

with col_titulo:
    st.title("Gerador Automático de Dossiês PLD-FT (V2)")
    st.subheader(
        "Upload de relatórios com geração de código de rastreabilidade"
    )

st.markdown("---")

# --- PASSO 1: UPLOAD DA PLANILHA ---
st.markdown("### 1. Selecione a Planilha de Apontamentos")
uploaded_file = st.file_uploader(
    "Arraste e solte o arquivo .xlsx ou .csv do e-Guardian aqui:",
    type=["xlsx", "xls", "csv"],
)

if uploaded_file is not None:
    try:
        # Carrega a planilha garantindo tipo texto para não perder zeros à esquerda
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, dtype=str, skiprows=2)
        else:
            df = pd.read_excel(uploaded_file, dtype=str, skiprows=2)

        # Limpa espaços e formata cabeçalhos
        df.columns = [str(c).strip() for c in df.columns]
        df = df.fillna("")

        # 1. Gera o Código de Rastreabilidade para cada linha
        df["CODIGO_DOSSIE"] = [
            gerar_codigo_dossie(i + 1) for i in range(len(df))
        ]

        # 2. Cria o ID de seleção combinando o Código Único com o CPF/CNPJ pesquisado
        df["ID_Alerta"] = df.apply(
            lambda r: f"{r['CODIGO_DOSSIE']} | CPF/CNPJ: {r.get('Listas - CPF/CNPJ Pesquisado', 'N/A')}",
            axis=1,
        )

        st.success(
            f"✅ Planilha processada com sucesso! **{len(df)} registro(s)** identificados e codificados."
        )

        # Exibe prévia dos dados com os códigos gerados
        st.markdown("### 2. Prévia dos Registros Codificados")
        colunas_destaque = [
            "CODIGO_DOSSIE",
            "Listas - CPF/CNPJ Pesquisado",
            "Listas - Nome Encontrado",
            "Listas - Nome da Lista Pesquisada",
            "Listas - Data Detecção",
        ]
        cols_existentes = [c for c in colunas_destaque if c in df.columns]
        st.dataframe(
            df[cols_existentes] if cols_existentes else df,
            use_container_width=True,
        )

        st.markdown("---")
        st.markdown("### 3. Emissão dos Dossiês")

        col_acao1, col_acao2 = st.columns(2)

        # --- OPÇÃO A: GERAÇÃO INDIVIDUAL ---
        with col_acao1:
            st.markdown("#### Gerar Alerta Individual")
            alerta_selecionado = st.selectbox(
                "Escolha o registro que deseja emitir:", df["ID_Alerta"].tolist()
            )

            if st.button("🚀 Gerar Dossiê Selecionado"):
                linha = df[df["ID_Alerta"] == alerta_selecionado].iloc[0]
                doc = Document("modelo_dossie.docx")

                data_detecao = formatar_data(
                    linha.get("Listas - Data Detecção", "")
                )
                data_hoje = datetime.date.today().strftime("%d/%m/%Y")

                # Dicionário mapeado para as tags do modelo Word
                dicionario_dados = {
                    "{{CODIGO_DOSSIE}}": linha.get("CODIGO_DOSSIE", ""),
                    "{{NUM_ALERTA}}": linha.get(
                        "Listas - CPF/CNPJ Pesquisado", ""
                    ),
                    "{{DATA_GERACAO}}": data_detecao,
                    "{{ANALISTA}}": "Analista de PLD",
                    "{{DATA_ANALISE}}": data_hoje,
                    "{{SISTEMA}}": "Advice e-Guardian",
                    "{{STATUS_ALERTA}}": "Em análise",
                    "{{TIPOLOGIA}}": linha.get(
                        "Listas - Nome da Lista Pesquisada", ""
                    ),
                    "{{REGRA}}": "Apontamento em Listas Restritivas",
                    "{{NORMATIVA}}": "Lei nº 9.613/1998 e Resolução BCB nº 96/2021",
                    "{{NOME_CONTRAPARTE}}": linha.get(
                        "Listas - Nome Encontrado", ""
                    ),
                    "{{CPF_CNPJ}}": linha.get(
                        "Listas - CPF/CNPJ Encontrado", ""
                    ),
                    "{{OBS_CONTRAPARTE}}": f"Vinculação: {linha.get('Listas - Parte Relac', '')} | Grupo: {linha.get('Listas - Grupo Atuação', '')}",
                    "{{JUSTIFICATIVA}}": "Análise realizada conforme bases públicas/privadas de apontamentos.",
                }

                substituir_texto(doc, dicionario_dados)

                cod_dossie = linha.get("CODIGO_DOSSIE", "DOSSIE")
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

        # --- OPÇÃO B: GERAÇÃO EM LOTE (.ZIP) ---
        with col_acao2:
            st.markdown("#### Gerar Todos em Lote (Zip)")
            st.write(
                "Gere automaticamente todos os dossiês codificados da planilha de uma só vez."
            )

            if st.button("📦 Gerar Todos em Lote (.zip)"):
                zip_buffer = io.BytesIO()
                data_hoje = datetime.date.today().strftime("%d/%m/%Y")

                with zipfile.ZipFile(
                    zip_buffer, "w", zipfile.ZIP_DEFLATED
                ) as zip_file:
                    for idx, linha in df.iterrows():
                        doc = Document("modelo_dossie.docx")
                        cod_dossie = linha.get("CODIGO_DOSSIE", f"DOS_{idx+1}")
                        data_detecao = formatar_data(
                            linha.get("Listas - Data Detecção", "")
                        )

                        dicionario_dados = {
                            "{{CODIGO_DOSSIE}}": cod_dossie,
                            "{{NUM_ALERTA}}": linha.get(
                                "Listas - CPF/CNPJ Pesquisado", ""
                            ),
                            "{{DATA_GERACAO}}": data_detecao,
                            "{{ANALISTA}}": "Analista de PLD",
                            "{{DATA_ANALISE}}": data_hoje,
                            "{{SISTEMA}}": "Advice e-Guardian",
                            "{{STATUS_ALERTA}}": "Em análise",
                            "{{TIPOLOGIA}}": linha.get(
                                "Listas - Nome da Lista Pesquisada", ""
                            ),
                            "{{REGRA}}": "Apontamento em Listas Restritivas",
                            "{{NORMATIVA}}": "Lei nº 9.613/1998 e Resolução BCB nº 96/2021",
                            "{{NOME_CONTRAPARTE}}": linha.get(
                                "Listas - Nome Encontrado", ""
                            ),
                            "{{CPF_CNPJ}}": linha.get(
                                "Listas - CPF/CNPJ Encontrado", ""
                            ),
                            "{{OBS_CONTRAPARTE}}": f"Vinculação: {linha.get('Listas - Parte Relac', '')} | Grupo: {linha.get('Listas - Grupo Atuação', '')}",
                            "{{JUSTIFICATIVA}}": "Análise realizada conforme bases públicas/privadas de apontamentos.",
                        }

                        substituir_texto(doc, dicionario_dados)

                        doc_buffer = io.BytesIO()
                        doc.save(doc_buffer)
                        zip_file.writestr(
                            f"{cod_dossie}_Dossie_PLD.docx",
                            doc_buffer.getvalue(),
                        )

                zip_buffer.seek(0)
                st.download_button(
                    label="📥 Baixar Pacote Completo (.zip)",
                    data=zip_buffer,
                    file_name=f"Dossies_PLD_{datetime.date.today().strftime('%d_%m_%Y')}.zip",
                    mime="application/zip",
                )

    except Exception as e:
        st.error(f"Erro ao processar a planilha: {e}")
