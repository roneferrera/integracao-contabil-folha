# ============================================================
# app_integracao_dominio.py – Integração Contábil Domínio V1.0
# Dependências: streamlit, pandas, pdfplumber, openpyxl
# pip install streamlit pandas pdfplumber openpyxl
# ============================================================

import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO

VERSAO = "V1.0"

# ==============================
# TEMA TR
# ==============================
def apply_tr_theme():
    st.markdown("""
        <style>
        html, body, [class*="css"] {
            font-family: 'Segoe UI', 'Arial', sans-serif; color: #444444;
        }
        h1, h2, h3 { color: #FF8000; font-weight: 700; }
        section[data-testid="stSidebar"] { background-color: #444444; color: #FFFFFF; }
        section[data-testid="stSidebar"] * { color: #FFFFFF !important; }
        .stButton > button {
            background-color: #FF8000; color: #FFFFFF;
            border: none; border-radius: 4px; font-weight: bold;
        }
        .stButton > button:hover { background-color: #D64001; color: #FFFFFF; }
        .stDownloadButton > button {
            background-color: #FF8000; color: #FFFFFF;
            border: none; border-radius: 4px; font-weight: bold;
        }
        .stDownloadButton > button:hover { background-color: #D64001; color: #FFFFFF; }
        hr { border-color: #FF8000; }
        [data-testid="metric-container"] {
            background-color: #E9E9E9; border-left: 4px solid #FF8000;
            border-radius: 4px; padding: 10px;
        }
        .instrucoes-box {
            background-color: #E9E9E9; border-left: 4px solid #FF8000;
            border-radius: 4px; padding: 16px 20px; margin: 12px 0;
            color: #444444;
        }
        .instrucoes-box h4 { color: #FF8000; margin-top: 14px; margin-bottom: 6px; }
        .instrucoes-box h4:first-child { margin-top: 0; }
        </style>
    """, unsafe_allow_html=True)


# ==============================
# LINHAS A IGNORAR NO PDF
# ==============================
IGNORE_PATTERNS = [
    r"^EMPRESA PADR",
    r"^RUBRICAS",
    r"^Emiss",
    r"^Hora:",
    r"^Pág",
    r"^Cód\.",
    r"^\s*$",
    r"^A\.",
    r"^B\.",
    r"^C\.",
    r"^D\.",
    r"^E\.",
    r"^F\.",
    r"^G\.",
    r"^H\.",
    r"^I\.",
    r"^J\.",
    r"^K\.",
    r"^L\.",
    r"^M\.",
    r"^N\.",
    r"^O\.",
    r"^P\.",
    r"^Q\.",
    r"^R\.",
    r"^S\.",
    r"^T\.",
    r"^U\.",
    r"^V\.",
    r"^W\.",
    r"^X\.",
    r"^Z\.",
    r"^Soma na base",
]

def should_ignore_cad(line: str) -> bool:
    for pat in IGNORE_PATTERNS:
        if re.search(pat, line, re.IGNORECASE):
            return True
    return False


# ==============================
# PARSE DO PDF DE RUBRICAS
# ==============================
def parse_cadastro_eventos_pdf(file_bytes: bytes) -> dict:
    """
    Retorna dict {codigo_evento (str): tipo_rubrica (str)}

    Captura TODAS as variações do tipo no PDF de Rubricas:
      1. Tipo separado por espaço:   "208 REEMBOLSO DE QUILOMETRAGEM Provento Nenhuma..."
      2. Tipo colado na descrição:   "836 INSS DIF FER DESC A MAIORProvento Nenhuma..."
      3. Tipo colado na próx palavra:"813 FGTS FERIAS InformativaNenhuma..."
      4. Tipo com ponto e espaço:    "243 CONVENIO MEDICO Inf. dedutora Nenhuma..."
    """
    catalog = {}

    # Regex abrangente:
    # - código numérico no início
    # - descrição lazy até encontrar o tipo
    # - \s* = zero ou mais espaços antes do tipo (captura colados)
    # - tipo com variações: Provento, Desconto, Inf. ded(utora)?, Informativa
    RE_LINHA = re.compile(
        r"^\s*(\d+)\s+"                                               # código
        r"(.+?)"                                                      # descrição (lazy)
        r"\s*(Provento|Desconto|Inf\.\s*ded(?:utora)?|Informativa)"   # tipo
        r"[\s\w]",                                                    # seguido de espaço ou letra
        re.IGNORECASE,
    )

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for raw in text.splitlines():
                line = raw.strip()
                if not line or should_ignore_cad(line):
                    continue
                m = RE_LINHA.match(line)
                if m:
                    cod      = m.group(1).strip()
                    tipo_raw = m.group(3).strip().lower()

                    if "provento"   in tipo_raw:
                        tipo_norm = "Provento"
                    elif "desconto" in tipo_raw:
                        tipo_norm = "Desconto"
                    elif "inf. ded" in tipo_raw or "inf.ded" in tipo_raw:
                        tipo_norm = "Inf. dedutora"
                    elif "informat" in tipo_raw:
                        tipo_norm = "Informativa"
                    else:
                        tipo_norm = m.group(3).strip()

                    # Primeira ocorrência = mais confiável
                    if cod not in catalog:
                        catalog[cod] = tipo_norm

    return catalog


# ==============================
# LEITURA DO EXCEL DE EVENTOS
# ==============================
def ler_excel_eventos(file_bytes: bytes) -> tuple[pd.DataFrame | None, str]:
    """
    Lê o Excel de eventos configurados.
    Retorna (DataFrame, mensagem_erro).
    Aceita abas 'evento' ou 'Plan1'.
    """
    try:
        xls = pd.ExcelFile(BytesIO(file_bytes))
    except Exception as e:
        return None, f"Erro ao abrir Excel: {e}"

    sheet_name = None
    for candidate in ["evento", "Evento", "EVENTO", "Plan1", "plan1"]:
        if candidate in xls.sheet_names:
            sheet_name = candidate
            break

    if sheet_name is None:
        return None, (
            f"Aba 'evento' não encontrada. "
            f"Abas disponíveis: {xls.sheet_names}"
        )

    try:
        df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, dtype=str)
    except Exception as e:
        return None, f"Erro ao ler aba '{sheet_name}': {e}"

    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")

    return df, ""


def normalizar_colunas(df: pd.DataFrame) -> pd.DataFrame | None:
    """
    Mapeia nomes de colunas variados para nomes padronizados internos.
    Retorna None se colunas obrigatórias não forem encontradas.
    """
    mapa = {
        "cod_empresa":   ["código da empresa", "codigo da empresa", "cod. empresa",
                          "cod_empresa", "empresa"],
        "centro_custo":  ["centro de custo", "centro_custo", "cod centro de custo",
                          "cód centro de custo"],
        "seq":           ["código sequencial da integração",
                          "codigo sequencial da integracao",
                          "código sequencial", "seq", "sequencial"],
        "tipo_integ":    ["tipo da integração (1 - folha mensal; 2 - empresa; "
                          "3 - férias; 4 - rescisao; 5 - prov. férias; 6 - prov. 13)",
                          "tipo da integração", "tipo integracao", "tipo_integ", "tipo"],
        "descricao":     ["descrição", "descricao", "desc", "descrição evento",
                          "descrição do evento"],
        "conta_debito":  ["código da conta débito", "codigo da conta debito",
                          "conta débito", "conta_debito", "débito"],
        "conta_credito": ["código da conta crédito", "codigo da conta credito",
                          "conta crédito", "conta_credito", "crédito"],
        "historico":     ["código do histórico", "codigo do historico",
                          "histórico", "historico"],
    }

    col_lower = {c.lower(): c for c in df.columns}
    rename_map = {}

    for campo, candidatos in mapa.items():
        for cand in candidatos:
            if cand.lower() in col_lower:
                rename_map[col_lower[cand.lower()]] = campo
                break

    df = df.rename(columns=rename_map)

    obrigatorias = ["seq", "tipo_integ", "descricao"]
    faltando = [c for c in obrigatorias if c not in df.columns]
    if faltando:
        return None

    return df


# ==============================
# CRUZAMENTO E GERAÇÃO DO TXT
# ==============================
def gerar_txt_dominio(
    df_eventos: pd.DataFrame,
    catalog_pdf: dict,
    cod_empresa_padrao: str,
    log: list,
) -> tuple[str, pd.DataFrame]:
    """
    Cruza os eventos do Excel com o catálogo do PDF e gera o TXT Domínio.

    Formato de saída (leiaute Domínio Separador):
      Aba 'integra':  cod_empresa|0|seq|tipo_integ|cod_rubrica|
      Aba 'evento':   cod_empresa|centro_custo|seq|tipo_integ|descricao|
                      conta_debito|conta_credito|historico|complemento|

    Na prática, o sistema usa o arquivo de eventos (aba 'evento') para
    configurar as contas contábeis. Aqui geramos o arquivo de integração
    completo no formato esperado pelo Domínio.
    """
    linhas_integra = []
    linhas_evento  = []
    dados_tabela   = []

    nao_identificados  = []
    sem_conta          = []
    gerados_ok         = 0

    for _, row in df_eventos.iterrows():
        seq         = str(row.get("seq",         "")).strip()
        tipo_integ  = str(row.get("tipo_integ",  "")).strip()
        descricao   = str(row.get("descricao",   "")).strip()
        conta_deb   = str(row.get("conta_debito",  "")).strip()
        conta_cred  = str(row.get("conta_credito", "")).strip()
        historico   = str(row.get("historico",     "")).strip()
        centro      = str(row.get("centro_custo",  "")).strip()
        complemento = str(row.get("complemento",   "")).strip()

        # Código da empresa: usa coluna ou padrão informado
        cod_emp_raw = str(row.get("cod_empresa", "")).strip()
        cod_empresa = cod_emp_raw if cod_emp_raw and cod_emp_raw.lower() != "nan" \
                      else cod_empresa_padrao

        # Limpa NaN
        for var_name, var_val in [
            ("conta_deb",  conta_deb),
            ("conta_cred", conta_cred),
            ("historico",  historico),
            ("centro",     centro),
            ("complemento", complemento),
        ]:
            pass  # já são strings; nan vira "nan"

        conta_deb   = "" if conta_deb.lower()   == "nan" else conta_deb
        conta_cred  = "" if conta_cred.lower()  == "nan" else conta_cred
        historico   = "" if historico.lower()   == "nan" else historico
        centro      = "" if centro.lower()      == "nan" else centro
        complemento = "" if complemento.lower() == "nan" else complemento

        if not seq or seq.lower() == "nan":
            continue

        # Busca tipo no catálogo do PDF
        tipo_rubrica = catalog_pdf.get(seq, "")

        status = ""
        if not tipo_rubrica:
            nao_identificados.append(seq)
            status = "⚠️ Não identificado no PDF"
            log.append(
                f"Evento {seq} ({descricao[:40]}) — "
                f"tipo não encontrado no PDF de Rubricas."
            )
        else:
            if not conta_deb and not conta_cred:
                sem_conta.append(seq)
                status = "ℹ️ Sem conta configurada"
            else:
                gerados_ok += 1
                status = f"✅ {tipo_rubrica}"

        # Linha aba 'integra'
        linhas_integra.append(
            f"{cod_empresa}|0|{seq}|{tipo_integ}|{seq}|\n"
        )

        # Linha aba 'evento'
        linhas_evento.append(
            f"{cod_empresa}|{centro}|{seq}|{tipo_integ}|{descricao}|"
            f"{conta_deb}|{conta_cred}|{historico}|{complemento}|\n"
        )

        dados_tabela.append({
            "Seq":            seq,
            "Tipo Integração": tipo_integ,
            "Descrição":      descricao,
            "Tipo Rubrica":   tipo_rubrica or "—",
            "Conta Débito":   conta_deb,
            "Conta Crédito":  conta_cred,
            "Histórico":      historico,
            "Centro Custo":   centro,
            "Status":         status,
        })

    # Monta arquivo final
    conteudo = "".join(linhas_integra) + "".join(linhas_evento)

    log.append(
        f"Geração concluída → "
        f"Gerados OK: {gerados_ok} | "
        f"Sem tipo no PDF: {len(nao_identificados)} | "
        f"Sem conta: {len(sem_conta)}"
    )

    if nao_identificados:
        log.append(
            f"Eventos sem tipo no PDF: {', '.join(nao_identificados[:20])}"
            f"{'...' if len(nao_identificados) > 20 else ''}"
        )

    return conteudo, pd.DataFrame(dados_tabela)


# ==============================
# INTERFACE STREAMLIT
# ==============================
def main():
    st.set_page_config(
        page_title="Domínio Sistemas | Integração Contábil",
        page_icon="🟠",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    apply_tr_theme()

    st.markdown(
        f"""
        <div style="background:#444444; padding:24px 28px 18px 28px;
                    border-radius:8px; border-top:6px solid #FF8000;
                    margin-bottom:28px;">
            <h2 style="color:#FF8000; margin:0;
                       font-family:'Segoe UI',Arial,sans-serif;">
                📊 Integração Contábil — Excel + PDF Rubricas → Domínio
                &nbsp;|&nbsp; {VERSAO}
            </h2>
            <p style="color:#DDDDDD; margin:6px 0 0 0;
                      font-family:'Segoe UI',Arial,sans-serif;">
                Faça upload do <b>Excel de eventos</b> e do
                <b>PDF de Rubricas</b>, depois clique em
                <b>▶ Gerar arquivo Domínio</b>.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── Sidebar ───────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("### ⚙️ Configurações")
        cod_empresa_padrao = st.text_input(
            "Código da empresa (padrão)",
            value="1",
            help="Usado quando o Excel não tiver a coluna 'Código da Empresa'.",
        )
        st.markdown("---")
        st.markdown("### ℹ Sobre")
        st.markdown(f"**Versão:** {VERSAO}")
        st.markdown("**Thomson Reuters**")
        st.markdown("**Domínio Sistemas**")

    # ── Instruções ────────────────────────────────────────────────────
    with st.expander("📖 **Instruções de Uso** — clique para expandir", expanded=False):
        st.markdown(
            """
            <div class="instrucoes-box">

            <h4>🔹 Passo 1 — Prepare os arquivos</h4>
            <ul>
                <li><b>Excel de Eventos</b>: aba <code>evento</code> com colunas
                    Código da Empresa, Centro de Custo, Código Sequencial,
                    Tipo da Integração, Descrição, Conta Débito, Conta Crédito,
                    Histórico.</li>
                <li><b>PDF de Rubricas</b>: relatório "Plano e Acumuladores → Rubricas"
                    exportado do Domínio.</li>
            </ul>

            <h4>🔹 Passo 2 — Faça o upload</h4>
            <p>Selecione os dois arquivos nos campos abaixo.</p>

            <h4>🔹 Passo 3 — Gere e baixe</h4>
            <p>Clique em <b>▶ Gerar arquivo Domínio</b> e depois em
            <b>⬇ Baixar TXT</b>.</p>

            <h4>🔹 Passo 4 — Importe no Domínio</h4>
            <p><b>Utilitários → Importação → Importação Padrão →
            Leiaute Domínio Sistemas com Separador</b>.</p>

            <h4>⚠ Observações</h4>
            <ul>
                <li>Eventos com tipo <b>colado</b> na descrição no PDF
                    (ex: <code>InformativaNenhuma</code>) são detectados
                    automaticamente.</li>
                <li>Eventos marcados como <b>⚠️ Não identificado no PDF</b>
                    precisam ser configurados manualmente.</li>
            </ul>

            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("---")

    # ── Session state ─────────────────────────────────────────────────
    defaults = {
        "log":          [f"Aplicação pronta. Versão: {VERSAO}"],
        "txt_gerado":   None,
        "nome_arquivo": "dominio_integracao.txt",
        "df_resultado": None,
        "catalog_size": 0,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

    # ── Uploads ───────────────────────────────────────────────────────
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        arquivo_excel = st.file_uploader(
            "📄 Excel de Eventos (.xlsx)",
            type=["xlsx", "xls"],
            help="Arquivo com os eventos a serem integrados.",
        )
    with col_up2:
        arquivo_pdf = st.file_uploader(
            "📋 PDF de Rubricas (.pdf)",
            type=["pdf"],
            help="Relatório de Rubricas exportado do Domínio.",
        )

    col_btn1, col_btn2 = st.columns([1, 1])
    with col_btn1:
        gerar = st.button(
            "▶ Gerar arquivo Domínio",
            disabled=(arquivo_excel is None or arquivo_pdf is None),
            use_container_width=True,
            type="primary",
        )
    with col_btn2:
        limpar = st.button("🗑 Limpar", use_container_width=True)

    if limpar:
        for k, v in defaults.items():
            st.session_state[k] = v
        st.session_state.log = ["Campos limpos."]
        st.rerun()

    # ── Processamento ─────────────────────────────────────────────────
    if gerar and arquivo_excel and arquivo_pdf:
        log = ["Iniciando processamento..."]

        # 1. Parse do PDF
        with st.spinner("Lendo PDF de Rubricas..."):
            catalog = parse_cadastro_eventos_pdf(arquivo_pdf.read())
        log.append(
            f"PDF lido: {len(catalog)} evento(s) identificado(s) no catálogo."
        )
        st.session_state.catalog_size = len(catalog)

        # 2. Leitura do Excel
        df_raw, erro = ler_excel_eventos(arquivo_excel.read())
        if erro:
            log.append(f"ERRO: {erro}")
            st.session_state.log = log
            st.rerun()

        df = normalizar_colunas(df_raw)
        if df is None:
            log.append(
                "ERRO: Colunas obrigatórias não encontradas no Excel. "
                "Verifique se as colunas 'Código Sequencial', "
                "'Tipo da Integração' e 'Descrição' existem."
            )
            st.session_state.log = log
            st.rerun()

        log.append(
            f"Excel lido: {len(df)} linha(s) na aba de eventos."
        )

        # 3. Geração do TXT
        conteudo, df_resultado = gerar_txt_dominio(
            df, catalog, cod_empresa_padrao, log
        )

        st.session_state.txt_gerado   = conteudo.encode("latin-1", errors="replace")
        st.session_state.nome_arquivo = "dominio_integracao.txt"
        st.session_state.df_resultado = df_resultado
        st.session_state.log          = log
        st.rerun()

    # ── Resultados ────────────────────────────────────────────────────
    if st.session_state.txt_gerado is not None:
        st.success("✅ Arquivo gerado com sucesso!")
        st.download_button(
            label="⬇ Baixar TXT (Domínio Separador)",
            data=st.session_state.txt_gerado,
            file_name=st.session_state.nome_arquivo,
            mime="text/plain",
            use_container_width=True,
            type="primary",
        )

        df_res = st.session_state.df_resultado
        if df_res is not None and not df_res.empty:
            # Métricas
            total       = len(df_res)
            ok          = len(df_res[df_res["Status"].str.startswith("✅")])
            nao_id      = len(df_res[df_res["Status"].str.startswith("⚠️")])
            sem_conta   = len(df_res[df_res["Status"].str.startswith("ℹ️")])
            catalog_sz  = st.session_state.catalog_size

            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("📋 Catálogo PDF",       catalog_sz)
            m2.metric("📊 Total eventos",      total)
            m3.metric("✅ Com tipo + conta",    ok)
            m4.metric("⚠️ Sem tipo no PDF",    nao_id)
            m5.metric("ℹ️ Sem conta",          sem_conta)

            # Tabela com highlight
            def highlight_row(row):
                s = str(row.get("Status", ""))
                if s.startswith("✅"):  return ["background-color:#d4edda"] * len(row)
                if s.startswith("⚠️"): return ["background-color:#fff3cd"] * len(row)
                if s.startswith("ℹ️"): return ["background-color:#cce5ff"] * len(row)
                return [""] * len(row)

            st.dataframe(
                df_res.style.apply(highlight_row, axis=1),
                use_container_width=True,
            )

            # Expander: eventos não identificados
            nao_id_df = df_res[df_res["Status"].str.startswith("⚠️")]
            if not nao_id_df.empty:
                with st.expander(
                    f"⚠️ Eventos não identificados no PDF ({len(nao_id_df)}) "
                    f"— requerem configuração manual"
                ):
                    st.dataframe(
                        nao_id_df[["Seq", "Tipo Integração",
                                   "Descrição", "Centro Custo"]],
                        use_container_width=True,
                    )

            # Prévia do arquivo
            with st.expander("👁️ Prévia do arquivo gerado (primeiras 30 linhas)"):
                preview = "".join(
                    st.session_state.txt_gerado
                    .decode("latin-1", errors="replace")
                    .splitlines(True)[:30]
                )
                st.code(preview, language="text")

    # ── Log ───────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("**Log de processamento**")
    log_texto = "\n".join(st.session_state.log)
    tem_erro  = any(str(l).startswith("ERRO") for l in st.session_state.log)
    cor_borda = "#D32F2F" if tem_erro else "#388E3C"
    st.markdown(
        f"""
        <div style="background:#FCFCFC; border:1px solid {cor_borda};
                    border-radius:6px; padding:14px;
                    font-family:Consolas,monospace; font-size:13px;
                    white-space:pre-wrap; max-height:340px;
                    overflow-y:auto; color:#1F1F1F;">
{log_texto}
        </div>
        """,
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
