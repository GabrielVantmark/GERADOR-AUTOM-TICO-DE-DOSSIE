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

# 2. Injeção de CSS para expansão da tela
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

# --- CABEÇALHO ---
col_logo, col_titulo = st.columns([1, 5], vertical_alignment="center")

with col_logo:
    try:
        st.image("noBgColor.png", width=160)
    except Exception:
        st.write("🤖")

with col_titulo:
    st.title("Gerador Automático via Planilha (V2)")
    st.subheader("Processamento de relatórios de Apontamentos de Listas Restritivas")

st.markdown("---")

# --- STEP 1: UPLOAD DA PLANILHA ---
st.markdown("### 1. Selecione a Planilha de apontamentos")
uploaded_file = st.file_uploader(
    "Arraste e solte a planilha (.xlsx ou .csv) exportada do sistema:",
    type=["xlsx", "xls", "csv"],
)


def formatar_data(valor):
    """Converte datas em formato YYYY-MM-DD para DD/MM/YYYY"""
    if pd.isna(valor) or not valor:
        return ""
    val_str = str(valor).strip()
    match = re.search(r"(\d{4})-(\d{2})-(\d{2})", val_str)
    if match:
        ano, mes, dia = match.groups()
        return f"{dia}/{mes}/{ano}"
    return val_str


def substituir_texto(doc_obj, mapa_substituicao):
    """Substitui placeholders nos parágrafos e tabelas do Word mantendo a formatação."""
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
                                    run.text = run.text.replace(chave, str(valor))
                            if chave in p.text:
                                p.text = p.text.replace(chave, str(valor))


if uploaded_file is not None:
    try:
        # Carrega os dados garantindo leitura em string para preservar CPF/CNPJ
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, dtype=str, skiprows=2)
        else:
            df = pd.read_excel(uploaded_file, dtype=str, skiprows=2)

        # Trata cabeçalhos e limpa espaços extras
        df.columns = [str(c).strip() for c in df.columns]
        df = df.fillna("")

        # Cria uma coluna de identificação do Alerta
        if "Listas - CPF/CNPJ Pesquisado" in df.columns:
            df["ID_Alerta"] = df.apply(
                lambda row: f"{row['Listas - CPF/CNPJ Pesquisado']}_{row.name + 1}",
                axis=1,
            )
        else:
            df["ID_Alerta"] = [f"Alerta_{i+1}" for i in range(len(df))]

        st.success(
            f"✅ Planilha processada com sucesso! **{len(df)} registro(s)** identificados."
        )

        st.markdown("### 2. Prévia dos Alertas Carregados")
        st.dataframe(df, use_container_width=True)

        st.markdown("---")
        st.markdown("### 3. Geração dos Dossiês")

        col_acao1, col_acao2 = st.columns(2)

        # --- GERAÇÃO INDIVIDUAL ---
        with col_acao1:
            st.markdown("#### Gerar Alerta Individual")
            alerta_selecionado = st.selectbox(
                "Escolha o registro para gerar:", df["ID_Alerta"].tolist()
            )

            if st.button("🚀 Gerar Dossiê Selecionado"):
                linha = df[df["ID_Alerta"] == alerta_selecionado].iloc[0]
                doc = Document("modelo_dossie.docx")

                data_detecao = formatar_data(
                    linha.get("Listas - Data Detecção", "")
                )
                data_hoje = datetime.date.today().strftime("%d/%m/%Y")

                dicionario_dados = {
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

                nome_arquivo = f"Dossie_PLD_{linha.get('Listas - CPF/CNPJ Encontrado', 'Alerta')}.docx"
                buffer = io.BytesIO()
                doc.save(buffer)
                buffer.seek(0)

                st.download_button(
                    label="📥 Baixar Dossiê (.docx)",
                    data=buffer,
                    file_name=nome_arquivo,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

        # --- GERAÇÃO EM LOTE (ZIP) ---
        with col_acao2:
            st.markdown("#### Gerar Todos em Lote (Zip)")
            st.write(
                "Gere automaticamente todos os dossiês da planilha compactados em um arquivo .zip."
            )

            if st.button("📦 Gerar Todos em Lote (.zip)"):
                zip_buffer = io.BytesIO()
                data_hoje = datetime.date.today().strftime("%d/%m/%Y")

                with zipfile.ZipFile(
                    zip_buffer, "w", zipfile.ZIP_DEFLATED
                ) as zip_file:
                    for idx, linha in df.iterrows():
                        doc = Document("modelo_dossie.docx")
                        data_detecao = formatar_data(
                            linha.get("Listas - Data Detecção", "")
                        )
                        cpf_encontrado = linha.get(
                            "Listas - CPF/CNPJ Encontrado", f"Item_{idx+1}"
                        )

                        dicionario_dados = {
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
                            "{{CPF_CNPJ}}": cpf_encontrado,
                            "{{OBS_CONTRAPARTE}}": f"Vinculação: {linha.get('Listas - Parte Relac', '')} | Grupo: {linha.get('Listas - Grupo Atuação', '')}",
                            "{{JUSTIFICATIVA}}": "Análise realizada conforme bases públicas/privadas de apontamentos.",
                        }

                        substituir_texto(doc, dicionario_dados)

                        doc_buffer = io.BytesIO()
                        doc.save(doc_buffer)
                        zip_file.writestr(
                            f"Dossie_PLD_{cpf_encontrado}_{idx+1}.docx",
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
        st.error(f"Erro ao ler/processar a planilha: {e}")