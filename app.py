import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO, StringIO
import re

# ══════════════════════════════════════════════════════════════════════════════
#  CONFIGURAÇÃO DA PÁGINA
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Domínio Sistemas | Thomson Reuters",
    page_icon="🟠",
    layout="wide",
    initial_sidebar_state="expanded",
)

VERSAO = "V2.0"

# ══════════════════════════════════════════════════════════════════════════════
#  TEMA THOMSON REUTERS
# ══════════════════════════════════════════════════════════════════════════════
def apply_tr_theme():
    st.markdown("""
        <style>
        html, body, [class*="css"] {
            font-family: 'Segoe UI', 'Arial', sans-serif;
            color: #444444;
        }
        h1, h2, h3 { color: #FF8000; font-weight: 700; }
        section[data-testid="stSidebar"] {
            background-color: #444444;
            color: #FFFFFF;
        }
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
            background-color: #E9E9E9;
            border-left: 4px solid #FF8000;
            border-radius: 4px;
            padding: 10px;
        }
        [data-testid="stFileUploader"] {
            border: 2px dashed #FF8000;
            border-radius: 6px;
            padding: 10px;
        }
        .instrucoes-box {
            background-color: #E9E9E9;
            border-left: 4px solid #FF8000;
            border-radius: 4px;
            padding: 16px 20px;
            margin: 12px 0;
            color: #444444;
            font-family: 'Segoe UI', Arial, sans-serif;
        }
        .instrucoes-box h4 {
            color: #FF8000;
            margin-top: 14px;
            margin-bottom: 6px;
        }
        .instrucoes-box h4:first-child { margin-top: 0; }
        </style>
    """, unsafe_allow_html=True)

apply_tr_theme()

# ══════════════════════════════════════════════════════════════════════════════
#  MAPEAMENTOS
# ══════════════════════════════════════════════════════════════════════════════
TIPO_MAP = {
    "folha normal":       1,
    "empresa":            2,
    "férias":             3,
    "ferias":             3,
    "rescisão":           4,
    "rescisao":           4,
    "provisão de férias": 5,
    "provisao de ferias": 5,
    "provisão de 13":     6,
    "provisao de 13":     6,
    "provisão de 13º":    6,
}

TIPO_DESC = {
    1: "Folha mensal",
    2: "Empresa",
    3: "Férias",
    4: "Rescisão",
    5: "Prov. Férias",
    6: "Prov. 13",
}

TIPO_ICONE = {1: "📋", 2: "🏢", 3: "🏖️", 4: "📤", 5: "📅", 6: "🎄"}

IGNORE_PATTERNS = [
    re.compile(r"^RELAÇÃO DE RUBRICAS",        re.IGNORECASE),
    re.compile(r"^Página\s*:",                 re.IGNORECASE),
    re.compile(r"^Emissão\s*:",                re.IGNORECASE),
    re.compile(r"^Hora\s*:",                   re.IGNORECASE),
    re.compile(r"^Empresa\s*:",                re.IGNORECASE),
    re.compile(r"^Código\s+Descrição",         re.IGNORECASE),
]

RE_CC    = re.compile(r"^Centro de Custo\s*:\s*(\d+)\s+(.+)$", re.IGNORECASE)
RE_EVENT = re.compile(r"^\s*(\d+)\s+(.+)$")


# ══════════════════════════════════════════════════════════════════════════════
#  PARSER DO PDF
# ══════════════════════════════════════════════════════════════════════════════
def normalize_tipo(line: str) -> int | None:
    key = line.strip().lower()
    if "13" in key:
        key = re.sub(r"[º°].*$", "", key).strip()
        key = re.sub(r"\s+$", "", key)
    return TIPO_MAP.get(key)


def should_ignore(line: str) -> bool:
    return any(p.match(line) for p in IGNORE_PATTERNS)


def parse_pdf(file_bytes: bytes) -> pd.DataFrame:
    rows = []
    current_tipo    = None
    current_cc_cod  = None
    current_cc_desc = None

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for raw in text.splitlines():
                line = raw.strip()
                if not line:
                    continue
                if should_ignore(line):
                    continue

                # Tipo de integração
                tipo_code = normalize_tipo(line)
                if tipo_code is not None:
                    current_tipo = tipo_code
                    continue

                # Centro de Custo
                m_cc = RE_CC.match(line)
                if m_cc:
                    current_cc_cod  = m_cc.group(1).strip()
                    current_cc_desc = m_cc.group(2).strip()
                    continue

                # Evento
                m_ev = RE_EVENT.match(line)
                if m_ev and current_tipo and current_cc_cod:
                    cod_ev  = m_ev.group(1).strip()
                    desc_ev = m_ev.group(2).strip()
                    rows.append({
                        "Cód Centro de Custo":    current_cc_cod,
                        "Desc. Centro de Custo":  current_cc_desc,
                        "Tipo da Integração":     current_tipo,
                        "Desc. Tipo Integração":  TIPO_DESC[current_tipo],
                        "Cod Evento":             cod_ev,
                        "Descrição Evento":       desc_ev,
                        "Cod + Descrição Evento": f"{cod_ev} - {desc_ev}",
                    })

    return pd.DataFrame(rows)


# ══════════════════════════════════════════════════════════════════════════════
#  GERAÇÃO DO ARQUIVO "evento exemplo.txt" (aba "evento")
#
#  Colunas (tabuladas):
#  Código da Empresa | Centro de custo | Código Sequencial da Integração |
#  Tipo da Integração (...) | Descrição | Código da Conta Débito |
#  Código da Conta Crédito | Código do Histórico | Complemento
#
#  Regras observadas no exemplo:
#  - Código da Empresa: preenchido apenas na 1ª linha de cada CC (demais = vazio)
#  - Centro de custo:   preenchido apenas na 1ª linha de cada CC (demais = vazio)
#  - Código Sequencial: sequência global crescente (1, 2, 3 ...)
#  - Conta Débito/Crédito/Histórico/Complemento: vazios (não configurados)
# ══════════════════════════════════════════════════════════════════════════════
def gerar_evento_txt(df: pd.DataFrame, cod_empresa: str) -> bytes:
    """
    Gera o arquivo tabulado no formato da aba 'evento' do evento exemplo.xlsx.
    Tipos incluídos: Folha Normal (1), Empresa (2), Férias (3), Rescisão (4).
    Provisões (5 e 6) vão para o arquivo integra.
    """
    HEADER = "\t".join([
        "Código da Empresa",
        "Centro de custo",
        "Código Sequencial da Integração",
        "Tipo da Integração (1 - Folha mensal; 2 - Empresa; 3 - Férias; 4 - Rescisao; 5 - Prov. Férias; 6 - Prov. 13)",
        "Descrição",
        "Código da Conta Débito",
        "Código da Conta Crédito",
        "Código do Histórico",
        "Complemento",
    ])

    # Filtra tipos 1-4 (evento)
    df_ev = df[df["Tipo da Integração"].isin([1, 2, 3, 4])].copy()

    linhas = [HEADER]
    seq = 1
    prev_cc = None

    for _, row in df_ev.iterrows():
        cc_key = (row["Cód Centro de Custo"], row["Tipo da Integração"])

        if cc_key != prev_cc:
            emp_val = cod_empresa
            cc_val  = row["Cód Centro de Custo"]
            prev_cc = cc_key
        else:
            emp_val = ""
            cc_val  = ""

        linha = "\t".join([
            emp_val,                          # Código da Empresa
            cc_val,                           # Centro de custo
            str(seq),                         # Código Sequencial
            str(row["Tipo da Integração"]),   # Tipo da Integração
            row["Descrição Evento"],          # Descrição
            "",                               # Conta Débito
            "",                               # Conta Crédito
            "",                               # Histórico
            "",                               # Complemento
        ])
        linhas.append(linha)
        seq += 1

    conteudo = "\n".join(linhas)
    return conteudo.encode("utf-8-sig")   # BOM para Excel abrir corretamente


# ══════════════════════════════════════════════════════════════════════════════
#  GERAÇÃO DO ARQUIVO "integra exemplo.txt" (aba "Plan1")
#
#  Colunas (tabuladas):
#  Código da Empresa | Centro de Custo | Código Sequencial da Integração |
#  Tipo da Integração (...) | Descrição | Código da Conta Crédito |
#  Código da Conta Débito | Código do Histórico
#
#  Regras observadas no exemplo:
#  - Centro de Custo: nan (vazio) no exemplo → mantemos vazio
#  - Código Sequencial: sequência global crescente
#  - Contas/Histórico: vazios
#  - Inclui TODOS os tipos (1-6)
# ══════════════════════════════════════════════════════════════════════════════
def gerar_integra_txt(df: pd.DataFrame, cod_empresa: str) -> bytes:
    """
    Gera o arquivo tabulado no formato da aba 'Plan1' do integra exemplo.xls.
    Inclui todos os tipos de integração (1 a 6).
    """
    HEADER = "\t".join([
        "Código da Empresa",
        "Centro de Custo",
        "Código Sequencial da Integração",
        "Tipo da Integração (1 - Folha mensal; 2 - Empresa; 3 - Férias; 4 - Rescisao; 5 - Prov. Férias; 6 - Prov. 13)",
        "Descrição",
        "Código da Conta Crédito",
        "Código da Conta Débito",
        "Código do Histórico",
    ])

    linhas = [HEADER]
    seq = 1

    for _, row in df.iterrows():
        linha = "\t".join([
            cod_empresa,                      # Código da Empresa
            "",                               # Centro de Custo (nan no exemplo)
            str(seq),                         # Código Sequencial
            str(row["Tipo da Integração"]),   # Tipo da Integração
            row["Descrição Evento"],          # Descrição
            "",                               # Conta Crédito
            "",                               # Conta Débito
            "",                               # Histórico
        ])
        linhas.append(linha)
        seq += 1

    conteudo = "\n".join(linhas)
    return conteudo.encode("utf-8-sig")


# ══════════════════════════════════════════════════════════════════════════════
#  GERAÇÃO DO EXCEL DE VISUALIZAÇÃO
# ══════════════════════════════════════════════════════════════════════════════
def to_excel(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Rubricas")
        wb = writer.book
        ws = writer.sheets["Rubricas"]

        hdr = wb.add_format({
            "bold": True, "bg_color": "#FF8000", "font_color": "#FFFFFF",
            "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True,
        })
        cell = wb.add_format({"border": 1, "valign": "vcenter", "align": "left"})
        cnum = wb.add_format({"border": 1, "valign": "vcenter", "align": "center"})
        alt  = wb.add_format({
            "border": 1, "valign": "vcenter", "align": "left", "bg_color": "#FFF0E0"
        })
        anum = wb.add_format({
            "border": 1, "valign": "vcenter", "align": "center", "bg_color": "#FFF0E0"
        })

        for c, col in enumerate(df.columns):
            ws.write(0, c, col, hdr)
        ws.set_row(0, 32)

        widths = [10, 24, 12, 18, 10, 42, 52]
        for c, w in enumerate(widths):
            ws.set_column(c, c, w)

        num_cols_idx = {0, 2, 4}

        for r, row in enumerate(df.itertuples(index=False), start=1):
            even = r % 2 == 0
            for c in range(len(df.columns)):
                val = row[c]
                if c in num_cols_idx:
                    ws.write(r, c, val, anum if even else cnum)
                else:
                    ws.write(r, c, val, alt if even else cell)

        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, len(df), len(df.columns) - 1)

    return output.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
#  HEADER PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(
    f"""
    <div style="background:#444444; padding:24px 28px 18px 28px;
                border-radius:8px; border-top:6px solid #FF8000;
                margin-bottom:28px;">
        <h2 style="color:#FF8000; margin:0;
                   font-family:'Segoe UI',Arial,sans-serif;">
            📊 Rubricas/Itens Não Configurados → TXT Tabulado
            &nbsp;|&nbsp; {VERSAO}
        </h2>
        <p style="color:#DDDDDD; margin:6px 0 0 0;
                  font-family:'Segoe UI',Arial,sans-serif;">
            Converte o PDF em dois arquivos TXT tabulados:
            <strong>evento_exemplo.txt</strong> e <strong>integra_exemplo.txt</strong>
            prontos para importação no Domínio Sistemas.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ══════════════════════════════════════════════════════════════════════════════
#  SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### ⚙️ Configurações")

    cod_empresa = st.text_input(
        "Código da Empresa",
        value="1",
        help="Código da empresa que será preenchido nos arquivos TXT.",
    )

    st.markdown("---")
    st.markdown("### 📋 Tipos de Integração")
    for cod, desc in TIPO_DESC.items():
        icone = TIPO_ICONE.get(cod, "•")
        st.caption(f"{icone} **{cod}** — {desc}")

    st.markdown("---")
    st.markdown("### 📄 Arquivos gerados")
    st.caption("**evento_exemplo.txt**")
    st.caption("→ Tipos 1 (Folha), 2 (Empresa), 3 (Férias), 4 (Rescisão)")
    st.caption("→ Aba: evento")
    st.markdown("")
    st.caption("**integra_exemplo.txt**")
    st.caption("→ Todos os tipos (1 a 6)")
    st.caption("→ Aba: Plan1")

    st.markdown("---")
    st.markdown("### ℹ Sobre")
    st.markdown(f"**Versão:** {VERSAO}")
    st.markdown("**Thomson Reuters**")
    st.markdown("**Domínio Sistemas**")

# ══════════════════════════════════════════════════════════════════════════════
#  INSTRUÇÕES
# ══════════════════════════════════════════════════════════════════════════════
with st.expander("📖 **Instruções de Uso** — clique para expandir", expanded=False):
    st.markdown(
        """
        <div class="instrucoes-box">

        <h4>🔹 Arquivos gerados</h4>
        <table style="width:100%; border-collapse:collapse;">
          <tr style="background:#FF8000; color:white;">
            <th style="padding:6px; text-align:left;">Arquivo</th>
            <th style="padding:6px; text-align:left;">Aba de destino</th>
            <th style="padding:6px; text-align:left;">Tipos incluídos</th>
            <th style="padding:6px; text-align:left;">Colunas</th>
          </tr>
          <tr style="background:#FFF0E0;">
            <td style="padding:6px;"><b>evento_exemplo.txt</b></td>
            <td style="padding:6px;">evento</td>
            <td style="padding:6px;">1, 2, 3, 4</td>
            <td style="padding:6px;">Empresa | CC | Seq | Tipo | Descrição | Déb | Créd | Hist | Complemento</td>
          </tr>
          <tr>
            <td style="padding:6px;"><b>integra_exemplo.txt</b></td>
            <td style="padding:6px;">Plan1</td>
            <td style="padding:6px;">1, 2, 3, 4, 5, 6</td>
            <td style="padding:6px;">Empresa | CC | Seq | Tipo | Descrição | Créd | Déb | Hist</td>
          </tr>
        </table>

        <h4>🔹 Passo a passo</h4>
        <ol>
          <li>Informe o <b>Código da Empresa</b> na sidebar.</li>
          <li>Faça o upload do PDF <b>RubricasItens não Configurados</b>.</li>
          <li>Clique em <b>▶ Processar PDF</b>.</li>
          <li>Baixe os dois arquivos TXT e o Excel de visualização.</li>
          <li>Abra cada TXT no Excel e cole na aba correspondente do template.</li>
        </ol>

        <h4>⚠️ Observações</h4>
        <ul>
          <li>Os campos <b>Conta Débito, Conta Crédito, Histórico e Complemento</b>
              ficam vazios — devem ser preenchidos manualmente no Domínio.</li>
          <li>O <b>Código Sequencial</b> é gerado automaticamente (1, 2, 3...).</li>
          <li>No <b>evento_exemplo</b>, Empresa e CC aparecem apenas na 1ª linha
              de cada bloco (igual ao arquivo de exemplo).</li>
          <li>No <b>integra_exemplo</b>, CC fica vazio (igual ao arquivo de exemplo).</li>
        </ul>

        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown("---")

# ══════════════════════════════════════════════════════════════════════════════
#  SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════
defaults = {
    "df":            None,
    "evento_bytes":  None,
    "integra_bytes": None,
    "excel_bytes":   None,
    "processado":    False,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ══════════════════════════════════════════════════════════════════════════════
#  UPLOAD + BOTÕES
# ══════════════════════════════════════════════════════════════════════════════
uploaded = st.file_uploader(
    "📁 Arquivo PDF — Relação de Rubricas/Itens Não Configurados",
    type=["pdf"],
    help="Relatório exportado do Domínio Sistemas.",
)

col_btn1, col_btn2 = st.columns([1, 1])
with col_btn1:
    gerar = st.button(
        "▶ Processar PDF",
        disabled=(uploaded is None),
        use_container_width=True,
        type="primary",
    )
with col_btn2:
    limpar = st.button("🗑 Limpar", use_container_width=True)

if limpar:
    for k, v in defaults.items():
        st.session_state[k] = v
    st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
#  PROCESSAMENTO
# ══════════════════════════════════════════════════════════════════════════════
if gerar and uploaded is not None:
    with st.spinner("🔄 Processando PDF e gerando arquivos..."):
        df = parse_pdf(uploaded.read())

    if df.empty:
        st.error("⚠️ Nenhum dado extraído. Verifique se o PDF está correto.")
    else:
        st.session_state.df            = df
        st.session_state.evento_bytes  = gerar_evento_txt(df, cod_empresa)
        st.session_state.integra_bytes = gerar_integra_txt(df, cod_empresa)
        st.session_state.excel_bytes   = to_excel(df)
        st.session_state.processado    = True
        st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
#  RESULTADO
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.processado and st.session_state.df is not None:
    df = st.session_state.df

    st.success(f"✅ **{len(df)} registros** extraídos com sucesso!")

    # ── Métricas ──────────────────────────────────────────────────────────────
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("📄 Total de Registros",  len(df))
    m2.metric("🏢 Centros de Custo",     df["Cód Centro de Custo"].nunique())
    m3.metric("🔢 Tipos de Integração",  df["Tipo da Integração"].nunique())
    m4.metric("🎯 Eventos Únicos",       df["Cod Evento"].nunique())

    st.markdown("---")

    # ── Downloads principais ──────────────────────────────────────────────────
    st.markdown("### ⬇️ Downloads")

    d1, d2, d3 = st.columns(3)

    with d1:
        st.markdown("**📄 evento_exemplo.txt**")
        st.caption("Aba: evento | Tipos: 1, 2, 3, 4")
        n_evento = len(df[df["Tipo da Integração"].isin([1, 2, 3, 4])])
        st.download_button(
            label=f"⬇ Baixar evento_exemplo.txt ({n_evento} linhas)",
            data=st.session_state.evento_bytes,
            file_name="evento_exemplo.txt",
            mime="text/plain",
            use_container_width=True,
        )

    with d2:
        st.markdown("**📄 integra_exemplo.txt**")
        st.caption("Aba: Plan1 | Tipos: 1 a 6")
        st.download_button(
            label=f"⬇ Baixar integra_exemplo.txt ({len(df)} linhas)",
            data=st.session_state.integra_bytes,
            file_name="integra_exemplo.txt",
            mime="text/plain",
            use_container_width=True,
        )

    with d3:
        st.markdown("**📊 Visualização Excel**")
        st.caption("Todas as colunas extraídas")
        st.download_button(
            label="⬇ Baixar Excel de visualização",
            data=st.session_state.excel_bytes,
            file_name="rubricas_nao_configuradas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.markdown("---")

    # ── Prévia dos TXTs ───────────────────────────────────────────────────────
    st.markdown("### 👁️ Prévia dos Arquivos Gerados")

    tab1, tab2 = st.tabs(["📄 evento_exemplo.txt", "📄 integra_exemplo.txt"])

    with tab1:
        prev_ev = st.session_state.evento_bytes.decode("utf-8-sig")
        linhas_ev = prev_ev.splitlines()
        st.caption(f"Total: {len(linhas_ev)-1} registros + cabeçalho")
        preview_ev = "\n".join(linhas_ev[:21])
        st.code(preview_ev, language="text")
        if len(linhas_ev) > 21:
            st.caption(f"... e mais {len(linhas_ev)-21} linhas")

    with tab2:
        prev_in = st.session_state.integra_bytes.decode("utf-8-sig")
        linhas_in = prev_in.splitlines()
        st.caption(f"Total: {len(linhas_in)-1} registros + cabeçalho")
        preview_in = "\n".join(linhas_in[:21])
        st.code(preview_in, language="text")
        if len(linhas_in) > 21:
            st.caption(f"... e mais {len(linhas_in)-21} linhas")

    st.markdown("---")

    # ── Filtros + Tabela ──────────────────────────────────────────────────────
    st.markdown("### 🔍 Filtros e Visualização")
    f1, f2, f3 = st.columns(3)

    with f1:
        cc_opts = ["Todos"] + sorted(
            df["Cód Centro de Custo"].unique().tolist(), key=int
        )
        cc_labels = {
            c: f"{c} — {df[df['Cód Centro de Custo']==c]['Desc. Centro de Custo'].iloc[0]}"
            for c in cc_opts if c != "Todos"
        }
        sel_cc = st.selectbox(
            "Centro de Custo", cc_opts,
            format_func=lambda x: cc_labels.get(x, x),
        )

    with f2:
        tipo_opts = ["Todos"] + sorted(df["Tipo da Integração"].unique().tolist())
        sel_tipo = st.selectbox(
            "Tipo de Integração", tipo_opts,
            format_func=lambda x: (
                f"{TIPO_ICONE.get(x,'')} {x} — {TIPO_DESC[x]}"
                if x != "Todos" else x
            ),
        )

    with f3:
        busca = st.text_input("🔎 Buscar evento (código ou descrição)")

    dff = df.copy()
    if sel_cc   != "Todos":
        dff = dff[dff["Cód Centro de Custo"] == sel_cc]
    if sel_tipo != "Todos":
        dff = dff[dff["Tipo da Integração"] == sel_tipo]
    if busca:
        mask = (
            dff["Cod Evento"].str.contains(busca, case=False, na=False)
            | dff["Descrição Evento"].str.contains(busca, case=False, na=False)
        )
        dff = dff[mask]

    st.dataframe(dff, use_container_width=True, height=420)
    st.caption(f"Exibindo **{len(dff)}** de **{len(df)}** registros")

    # ── Resumos ───────────────────────────────────────────────────────────────
    st.markdown("---")
    with st.expander("📊 Resumo por Tipo de Integração"):
        resumo = (
            df.groupby(["Tipo da Integração", "Desc. Tipo Integração"])
            .agg(
                Registros         =("Cod Evento", "count"),
                Centros_de_Custo  =("Cód Centro de Custo", "nunique"),
                Eventos_Únicos    =("Cod Evento", "nunique"),
            )
            .reset_index()
        )
        st.dataframe(resumo, use_container_width=True, hide_index=True)

    with st.expander("📊 Resumo por Centro de Custo"):
        resumo_cc = (
            df.groupby(["Cód Centro de Custo", "Desc. Centro de Custo"])
            .agg(
                Registros        =("Cod Evento", "count"),
                Tipos_Integração =("Tipo da Integração", "nunique"),
                Eventos_Únicos   =("Cod Evento", "nunique"),
            )
            .reset_index()
        )
        st.dataframe(resumo_cc, use_container_width=True, hide_index=True)

else:
    if not st.session_state.processado:
        st.markdown(
            """
            <div class="instrucoes-box">
            <h4>👆 Como começar</h4>
            <p>Informe o <b>Código da Empresa</b> na sidebar, faça o upload do PDF
            <b>Relação de Rubricas/Itens Não Configurados</b> e clique em
            <b>▶ Processar PDF</b>.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        col_l, col_r = st.columns(2)
        with col_l:
            st.markdown("#### 📄 Arquivo: evento_exemplo.txt")
            st.table(pd.DataFrame({
                "Coluna": [
                    "Código da Empresa",
                    "Centro de custo",
                    "Código Sequencial da Integração",
                    "Tipo da Integração",
                    "Descrição",
                    "Código da Conta Débito",
                    "Código da Conta Crédito",
                    "Código do Histórico",
                    "Complemento",
                ],
                "Obs": [
                    "1ª linha do bloco",
                    "1ª linha do bloco",
                    "Sequencial global",
                    "1, 2, 3 ou 4",
                    "Do PDF",
                    "Vazio",
                    "Vazio",
                    "Vazio",
                    "Vazio",
                ]
            }))

        with col_r:
            st.markdown("#### 📄 Arquivo: integra_exemplo.txt")
            st.table(pd.DataFrame({
                "Coluna": [
                    "Código da Empresa",
                    "Centro de Custo",
                    "Código Sequencial da Integração",
                    "Tipo da Integração",
                    "Descrição",
                    "Código da Conta Crédito",
                    "Código da Conta Débito",
                    "Código do Histórico",
                ],
                "Obs": [
                    "Sempre preenchido",
                    "Vazio (nan)",
                    "Sequencial global",
                    "1, 2, 3, 4, 5 ou 6",
                    "Do PDF",
                    "Vazio",
                    "Vazio",
                    "Vazio",
                ]
            }))
