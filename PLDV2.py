import datetime
import io
import re
import zipfile
from docx import Document
import pandas as pd
import streamlit as st

# 1. Configuração da página com layout amplo
st.set_page_config(
    page_title="Gerador de Dossiê PLD-FT",
    page_icon="logo title.png",
    layout="wide",
    initial_sidebar_state="expanded",
)

# 2. Injeção de CSS para identidade visual limpa e profissional
st.markdown(
    """
    <style>
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        padding-left: 2.5rem;
        padding-right: 2.5rem;
        max-width: 95% !important;
    }
    .stButton>button {
        width: 100%;
        background-color: #0F172A;
        color: white;
        font-weight: 600;
        border-radius: 8px;
        padding: 0.6rem 1rem;
        border: none;
    }
    .stButton>button:hover {
        background-color: #1E293B;
        color: white;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def gerar_codigo_dossie(indice):
    """Gera código único de rastreabilidade: DOS-YYYYMMDD-001"""
    hoje = datetime.date.today().strftime("%Y%m%d")
    return f"DOS-{hoje}-{str(indice).zfill(3)}"


def formatar_data(valor):
    """Formata datas YYYY-MM-DD para DD/MM/YYYY"""
    if pd.isna(valor) or not valor:
        return ""
    val_str = str(valor).strip()
    match = re.search(r"(\d{4})-(\d{2})-(\d{2})", val_str)
    if match:
        ano, mes, dia = match.groups()
        return f"{dia}/{mes}/{ano}"
    return val_str


def formatar_moeda(valor):
    """Formata valor numérico para moeda R$ XX,XX"""
    if pd.isna(valor) or not valor:
        return "R$ 0,00"
    try:
        val_float = float(str(valor).replace(",", "."))
        return (
            f"R$ {val_float:,.2f}"
            .replace(",", "v")
            .replace(".", ",")
            .replace("v", ".")
        )
    except Exception:
        return f"R$ {valor}"


def preencher_tabela_diligencias(doc_obj, lista_diligencias, datas_diligencias):
    """Localiza a tabela com {{DILIGENCIAS_NOME}} e adiciona linhas reais na tabela do Word."""
    for table in doc_obj.tables:
        for row in table.rows:
            cell_texts = [c.text for c in row.cells]
            if any("{{DILIGENCIAS_NOME}}" in t for t in cell_texts):
                for dil in lista_diligencias:
                    dt = datas_diligencias.get(
                        dil, datetime.date.today().strftime("%d/%m/%Y")
                    )
                    new_row = table.add_row()
                    if len(new_row.cells) >= 2:
                        new_row.cells[0].text = str(dil)
                        new_row.cells[1].text = str(dt)
                    elif len(new_row.cells) == 1:
                        new_row.cells[0].text = f"{dil} - {dt}"

                tr = row._tr
                table._tbl.remove(tr)
                break


def substituir_texto(doc_obj, mapa_substituicao):
    """Substitui placeholders simples nos parágrafos e tabelas do Word."""
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
        st.image("noBgColor.png", width=150)
    except Exception:
        st.write("🛡️")

with col_titulo:
    st.title("Sistema de Gestão e Emissão de Dossiês PLD-FT")
    st.caption("Plataforma de Automação de Análise de Alertas e Diligências")

st.markdown("---")

# --- PASSO 1: UPLOAD ---
st.markdown("### 1. Carregamento de Alertas (Base de Detectados)")
uploaded_file = st.file_uploader(
    "Arraste e solte o arquivo da retaguarda (.xlsx ou .csv):",
    type=["xlsx", "xls", "csv"],
)

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, dtype=str)
        else:
            df = pd.read_excel(uploaded_file, dtype=str)

        df.columns = [str(c).strip() for c in df.columns]
        df = df.fillna("")

        df["CODIGO_DOSSIE"] = [
            gerar_codigo_dossie(i + 1) for i in range(len(df))
        ]

        col_cpf = (
            "CPF/CNPJ Pesquisado"
            if "CPF/CNPJ Pesquisado" in df.columns
            else df.columns[6]
        )
        col_nome = (
            "Nome Encontrado"
            if "Nome Encontrado" in df.columns
            else df.columns[8]
        )

        df["ID_Alerta"] = df.apply(
            lambda r: f"{r['CODIGO_DOSSIE']} | {r.get(col_nome, '')} (CPF/CNPJ: {r.get(col_cpf, '')})",
            axis=1,
        )

        st.success(
            f"✅ **{len(df)} alerta(s)** carregado(s) com sucesso para análise."
        )

        st.markdown("---")
        st.markdown("### 2. Painel de Análise Humana e Diligências")

        alerta_selecionado = st.selectbox(
            "Selecione o alerta para revisar e emitir:", df["ID_Alerta"].tolist()
        )

        linha = df[df["ID_Alerta"] == alerta_selecionado].iloc[0]

        # Mapeamento DE-PARA
        op_origem = linha.get("Nome do Cliente", "")
        op_data = formatar_data(linha.get("Data da Operação", ""))
        op_valor = formatar_moeda(linha.get("Valor da Operação", ""))
        data_geracao = formatar_data(linha.get("Data da Detecção do Hit", ""))
        cpf_cnpj = linha.get("CPF/CNPJ Pesquisado", "")
        status_ip = linha.get("Parte Relacionada", "")
        nome_contraparte = linha.get("Nome Encontrado", "")
        regra_lista = linha.get("Lista", "")
        obs_complemento = linha.get("Complemento", "")
        op_destino = nome_contraparte

        # Dicionário de Justificativas Dinâmicas
        modelos_justificativas = {
            "Arquivado - Sem Indício de Irregularidade": f"Análise realizada sobre o apontamento na lista '{regra_lista}'. Consultas efetuadas nas fontes abertas e bases públicas não identificaram risco iminente de PLD-FT ou atipicidade financeira.",
            "Arquivado - Falso Positivo / Homônimo": f"Análise efetuada indica tratar-se de Falso Positivo / Homônimo. Os dados cadastrais do pesquisado não coincidem com o indivíduo/entidade registrado na lista restritiva '{regra_lista}'.",
            "Encaminhado para Comunicação (COAF)": f"Constatados indícios de atipicidade e divergência cadastral incompatíveis com a capacidade financeira. Processo encaminhado ao Comitê para deliberação de Comunicação de Atipicidade ao COAF.",
            "Em Monitoramento Contínuo": f"O cliente/contraparte permanece sob monitoramento contínuo reforçado para verificação de novas transações e eventual evolução dos apontamentos na lista '{regra_lista}'.",
            "Outro (Especifique na Justificativa)": "",
        }

        with st.expander(
            "📝 Dados Carregados, Diligências e Formatação do Laudo",
            expanded=True,
        ):
            c1, c2 = st.columns(2)

            with c1:
                st.markdown("#### 📌 Dados do Alerta Detectado")
                st.text_input(
                    "Código de Rastreabilidade / Nº Alerta",
                    linha.get("CODIGO_DOSSIE"),
                    disabled=True,
                )
                st.text_input(
                    "Nome da Contraparte / Destino",
                    nome_contraparte,
                    disabled=True,
                )
                st.text_input("CPF/CNPJ Pesquisado", cpf_cnpj, disabled=True)
                st.text_input(
                    "Regra / Lista Restritiva", regra_lista, disabled=True
                )
                st.text_input("Status na IP", status_ip, disabled=True)
                st.text_input(
                    "Operação - Origem", op_origem, disabled=True
                )
                st.text_input(
                    "Operação - Data", op_data, disabled=True
                )
                st.text_input(
                    "Operação - Valor", op_valor, disabled=True
                )

            with c2:
                st.markdown("#### ⚖️ Decisão do Analista")
                analista = st.text_input("Analista Responsável", "Analista PLD")
                data_analise = st.date_input(
                    "Data da Análise", datetime.date.today()
                ).strftime("%d/%m/%Y")

                risco_cliente = st.selectbox(
                    "Classificação de Risco do Cliente:",
                    ["Baixo", "Médio", "Alto", "Não Classificado"],
                    index=0,
                )

                decisao_arquivamento = st.selectbox(
                    "Conclusão da Análise:",
                    list(modelos_justificativas.keys()),
                )

                st.markdown("---")
                st.markdown("##### 🔎 Diligências Efetuadas")

                diligencias_opcoes = st.multiselect(
                    "Selecione as diligências padrão:",
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

                diligencia_extra = st.text_input(
                    "➕ Adicionar diligência personalizada (opcional):",
                    placeholder="Ex: Análise de Contrato Social na Junta Comercial",
                )

                lista_final_diligencias = list(diligencias_opcoes)
                if diligencia_extra.strip():
                    lista_final_diligencias.append(diligencia_extra.strip())

                datas_diligencias = {}
                if lista_final_diligencias:
                    st.markdown("**Datas da realização:**")
                    for dil in lista_final_diligencias:
                        d_data = st.date_input(
                            f"Data - {dil}:",
                            datetime.date.today(),
                            key=f"data_{dil}",
                        )
                        datas_diligencias[dil] = d_data.strftime("%d/%m/%Y")

                st.markdown("---")

                texto_padrao = modelos_justificativas.get(
                    decisao_arquivamento, ""
                )

                justificativa = st.text_area(
                    "Justificativa da Decisão (Editável):",
                    value=texto_padrao,
                    height=120,
                )

        st.markdown("---")
        st.markdown("### 3. Emissão do Documento")

        if st.button("📄 Gerar Dossiê em Word (.docx)"):
            doc = Document("modelo_dossie.docx")

            if lista_final_diligencias:
                preencher_tabela_diligencias(
                    doc, lista_final_diligencias, datas_diligencias
                )

            dicionario_dados = {
                "{{CODIGO_DOSSIE}}": linha.get("CODIGO_DOSSIE", ""),
                "{{NUM_ALERTA}}": linha.get("CODIGO_DOSSIE", ""),
                "{{SISTEMA}}": "Advice e-Guardian",
                "{{NORMATIVA}}": "Lei nº 9.613/1998 e Carta Circular Nº 3.978/2020",
                "{{DATA_GERACAO}}": data_geracao,
                "{{CPF_CNPJ}}": cpf_cnpj,
                "{{NOME_CONTRAPARTE}}": nome_contraparte,
                "{{REGRA}}": regra_lista,
                "{{TIPOLOGIA}}": regra_lista,
                "{{STATUS_IP}}": status_ip,
                "{{OBS_CONTRAPARTE}}": obs_complemento,
                "{{OPERAÇÃO_ORIGEM}}": op_origem,
                "{{OPERAÇÃO_DESTINO}}": op_destino,
                "{{OPERAÇÃO_DATA}}": op_data,
                "{{OPERAÇÃO_VALOR}}": op_valor,
                "{{RISCO_CLIENTE}}": risco_cliente,
                "{{ANALISTA}}": analista,
                "{{DATA_ANALISE}}": data_analise,
                "{{STATUS_ALERTA}}": decisao_arquivamento,
                "{{DECISAO}}": decisao_arquivamento,
                "{{JUSTIFICATIVA}}": justificativa,
            }

            substituir_texto(doc, dicionario_dados)

            cod_dossie = linha.get("CODIGO_DOSSIE", "DOSSIE")
            
            # ALTERAÇÃO DO NOME DO ARQUIVO CONFORME SOLICITADO:
            nome_arquivo = f"Dossiê de alerta PLD-FT - ({cod_dossie}).docx"

            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)

            st.download_button(
                label=f"📥 Baixar Documento Oficial ({cod_dossie})",
                data=buffer,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

    except Exception as e:
        st.error(f"Erro ao ler/processar a planilha: {e}")
