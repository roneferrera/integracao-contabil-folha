import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO
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

VERSAO = "V5.0"

# ══════════════════════════════════════════════════════════════════════════════
#  TEMA THOMSON REUTERS
# ══════════════════════════════════════════════════════════════════════════════
def apply_tr_theme():
    st.markdown("""
        <style>
        html, body, [class*="css"] {
            font-family: 'Segoe UI', 'Arial', sans-serif; color: #444444;
        }
        h1, h2, h3 { color: #FF8000; font-weight: 700; }
        section[data-testid="stSidebar"] { background-color: #444444; }
        section[data-testid="stSidebar"] * { color: #FFFFFF !important; }
        .stButton > button {
            background-color: #FF8000; color: #FFFFFF;
            border: none; border-radius: 4px; font-weight: bold;
        }
        .stButton > button:hover { background-color: #D64001; }
        .stDownloadButton > button {
            background-color: #FF8000; color: #FFFFFF;
            border: none; border-radius: 4px; font-weight: bold;
        }
        .stDownloadButton > button:hover { background-color: #D64001; }
        hr { border-color: #FF8000; }
        [data-testid="metric-container"] {
            background-color: #E9E9E9; border-left: 4px solid #FF8000;
            border-radius: 4px; padding: 10px;
        }
        [data-testid="stFileUploader"] {
            border: 2px dashed #FF8000; border-radius: 6px; padding: 8px;
        }
        .instrucoes-box {
            background-color: #E9E9E9; border-left: 4px solid #FF8000;
            border-radius: 4px; padding: 16px 20px; margin: 12px 0;
            color: #444444;
        }
        .instrucoes-box h4 { color: #FF8000; margin-top: 14px; margin-bottom: 6px; }
        .instrucoes-box h4:first-child { margin-top: 0; }
        .step-box {
            background: #FFF8F0; border: 1px solid #FF8000;
            border-radius: 8px; padding: 14px 18px; margin: 8px 0;
        }
        .step-title {
            color: #FF8000; font-weight: bold; font-size: 15px; margin-bottom: 6px;
        }
        .alert-box {
            background: #FFF3CD; border: 1px solid #FF8000;
            border-radius: 6px; padding: 12px 16px; margin: 8px 0; color: #444;
        }
        </style>
    """, unsafe_allow_html=True)

apply_tr_theme()

# ══════════════════════════════════════════════════════════════════════════════
#  MAPEAMENTOS
# ══════════════════════════════════════════════════════════════════════════════
TIPO_DESC = {
    1: "Folha mensal", 2: "Empresa", 3: "Férias",
    4: "Rescisão", 5: "Prov. Férias", 6: "Prov. 13",
}
TIPO_ICONE = {1: "📋", 2: "🏢", 3: "🏖️", 4: "📤", 5: "📅", 6: "🎄"}

# ══════════════════════════════════════════════════════════════════════════════
#  VALIDAÇÃO DO PDF
# ══════════════════════════════════════════════════════════════════════════════
def detectar_tipo_pdf(file_bytes: bytes) -> str:
    try:
        with pdfplumber.open(BytesIO(file_bytes)) as pdf:
            if not pdf.pages:
                return "desconhecido"
            text = (pdf.pages[0].extract_text() or "").upper()
            if "RELAÇÃO DE RUBRICAS" in text or "RELACAO DE RUBRICAS" in text:
                if "NÃO CONFIGURADOS" in text or "NAO CONFIGURADOS" in text:
                    return "rubricas_nao_config"
            if "PLANO E ACUMULADORES" in text:
                return "cadastro_eventos"
    except Exception:
        pass
    return "desconhecido"

# ══════════════════════════════════════════════════════════════════════════════
#  PARSER — PDF RUBRICAS NÃO CONFIGURADAS
# ══════════════════════════════════════════════════════════════════════════════
RE_TIPO_SECAO = [
    (re.compile(r"^Provisão\s+de\s+Férias\s*$",  re.IGNORECASE), 5),
    (re.compile(r"^Provisao\s+de\s+Ferias\s*$",  re.IGNORECASE), 5),
    (re.compile(r"^Provisão\s+de\s+13",           re.IGNORECASE), 6),
    (re.compile(r"^Provisao\s+de\s+13",           re.IGNORECASE), 6),
    (re.compile(r"^Folha\s+Normal\s*$",           re.IGNORECASE), 1),
    (re.compile(r"^Férias\s*$",                   re.IGNORECASE), 3),
    (re.compile(r"^Ferias\s*$",                   re.IGNORECASE), 3),
    (re.compile(r"^Rescisão\s*$",                 re.IGNORECASE), 4),
    (re.compile(r"^Rescisao\s*$",                 re.IGNORECASE), 4),
    (re.compile(r"^Empresa\s*$",                  re.IGNORECASE), 2),
]
RE_CC    = re.compile(r"^Centro\s+de\s+Custo\s*:\s*(\d+)\s+(.+)$", re.IGNORECASE)
RE_EVENT = re.compile(r"^\s*(\d+)\s{1,}(.+)$")

def should_ignore_rubrica(line: str) -> bool:
    s = line.strip()
    sl = s.lower()
    if not s: return True
    if "relação de rubricas" in sl or "relacao de rubricas" in sl: return True
    if re.match(r"^Empresa\s*:\s*\d+", s, re.IGNORECASE): return True
    if re.match(r"^Página\s*:",  s, re.IGNORECASE): return True
    if re.match(r"^Emissão\s*:", s, re.IGNORECASE): return True
    if re.match(r"^Hora\s*:",    s, re.IGNORECASE): return True
    if re.match(r"^[Cc]ód(?:igo)?\s+[Dd]escrição", s): return True
    if sl in {"código descrição", "codigo descricao",
              "código  descrição", "codigo  descricao"}: return True
    return False

def detect_tipo(line: str) -> int | None:
    s = line.strip()
    for pattern, code in RE_TIPO_SECAO:
        if pattern.match(s):
            return code
    return None

def parse_rubricas_pdf(file_bytes: bytes) -> pd.DataFrame:
    rows = []
    current_tipo = current_cc_cod = current_cc_desc = None
    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for raw in text.splitlines():
                line = raw.strip()
                if should_ignore_rubrica(line):
                    continue
                tipo = detect_tipo(line)
                if tipo is not None:
                    current_tipo = tipo
                    continue
                m_cc = RE_CC.match(line)
                if m_cc:
                    current_cc_cod  = m_cc.group(1).strip()
                    current_cc_desc = m_cc.group(2).strip()
                    continue
                if current_tipo is not None and current_cc_cod is not None:
                    m_ev = RE_EVENT.match(line)
                    if m_ev:
                        rows.append({
                            "Cód Centro de Custo":   current_cc_cod,
                            "Desc. Centro de Custo": current_cc_desc,
                            "Tipo da Integração":    current_tipo,
                            "Desc. Tipo Integração": TIPO_DESC[current_tipo],
                            "Cod Evento":            m_ev.group(1).strip(),
                            "Descrição Evento":      m_ev.group(2).strip(),
                        })
    return pd.DataFrame(rows)

# ══════════════════════════════════════════════════════════════════════════════
#  PARSER — PDF CADASTRO DE EVENTOS (Rubricas.pdf — Plano e Acumuladores)
#
#  Estrutura real do PDF (31 páginas):
#  Cada linha contém: CÓDIGO  DESCRIÇÃO(pode ser truncada)  TIPO  Base  Unidade...
#
#  O TIPO aparece na mesma linha, após a descrição (que pode ser truncada pelo PDF).
#  Valores possíveis: "Provento" | "Desconto" | "Informativa" | "Inf. ded"
#
#  Exemplos reais:
#  "1 HORAS NORMAIS Provento Nenhuma Horas 0,00 ..."
#  "40 HORAS FALTAS Desconto Salário Contratual Horas 0,00 ..."
#  "235 DESC.ADIANT.SALARIAL IRRFInf. ded Formula Automático 0,00 ..."
#  "243 CONVENIO MEDICO - INFORMATIVO Inf. dedutora Nenhuma Valor 0,00 ..."
#  "23 F.G.T.S DE RESCISAO InformativaNenhuma Valor 8,00 ..."
#  "813 FGTS FERIAS InformativaNenhuma Valor 8,00 ..."
#
#  Regex: captura código (numérico) + tudo até encontrar o tipo
# ══════════════════════════════════════════════════════════════════════════════

# Regex principal: código + descrição (lazy) + tipo
# O tipo pode estar colado na descrição (sem espaço) como "InformativaNenhuma"
RE_CAD_EVENTO = re.compile(
    r"^\s*(\d+)\s+"                                      # código
    r"(.+?)\s*"                                          # descrição (lazy, pode ser truncada)
    r"(Provento|Desconto|Informativa|Inf\.\s*ded\w*)"    # tipo
    r"[\s\w]",                                           # seguido de espaço ou letra (Base/Unidade)
    re.IGNORECASE,
)

IGNORE_CAD_PATTERNS = [
    re.compile(r"^EMPRESA PADRÃO",    re.IGNORECASE),
    re.compile(r"^Página\s*:",        re.IGNORECASE),
    re.compile(r"^Emissão\s*:",       re.IGNORECASE),
    re.compile(r"^Hora\s*:",          re.IGNORECASE),
    re.compile(r"^RUBRICAS\s*$",      re.IGNORECASE),
    re.compile(r"^Cód\.\s+Descrição", re.IGNORECASE),
    re.compile(r"^Soma na base",      re.IGNORECASE),
    re.compile(r"^[A-Z]\.\s+[A-Z]",  re.IGNORECASE),
]

def should_ignore_cad(line: str) -> bool:
    return any(p.match(line.strip()) for p in IGNORE_CAD_PATTERNS)

def normalizar_tipo_rubrica(tipo_raw: str) -> str:
    t = tipo_raw.strip().lower()
    if "provento"   in t: return "Provento"
    if "desconto"   in t: return "Desconto"
    if "inf. ded"   in t or "inf.ded" in t: return "Inf. dedutora"
    if "informat"   in t: return "Informativa"
    return tipo_raw.strip()

def parse_cadastro_eventos_pdf(file_bytes: bytes) -> dict:
    """
    Retorna dict {cod_evento (str): tipo_rubrica (str)}
    Identifica o tipo pelo CÓDIGO do evento, lendo a mesma linha.
    """
    catalog = {}
    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for raw in text.splitlines():
                line = raw.strip()
                if not line or should_ignore_cad(line):
                    continue
                m = RE_CAD_EVENTO.match(line)
                if m:
                    cod       = m.group(1).strip()
                    tipo_raw  = m.group(3).strip()
                    tipo_norm = normalizar_tipo_rubrica(tipo_raw)
                    # Só registra se ainda não visto (primeira ocorrência = mais confiável)
                    if cod not in catalog:
                        catalog[cod] = tipo_norm
    return catalog

# ══════════════════════════════════════════════════════════════════════════════
#  COLORAÇÃO DE TIPO RUBRICA
# ══════════════════════════════════════════════════════════════════════════════
def color_tipo_rubrica(val):
    cores = {
        "Provento":      "background-color:#d4edda; color:#155724",
        "Desconto":      "background-color:#f8d7da; color:#721c24",
        "Informativa":   "background-color:#cce5ff; color:#004085",
        "Inf. dedutora": "background-color:#fff3cd; color:#856404",
        "—":             "background-color:#f0f0f0; color:#888888",
    }
    return cores.get(val, "")

# ══════════════════════════════════════════════════════════════════════════════
#  GERAÇÃO DO EXCEL INTERMEDIÁRIO
# ══════════════════════════════════════════════════════════════════════════════
def gerar_excel_intermediario(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    cols_fixas = [
        "Cód Centro de Custo", "Desc. Centro de Custo",
        "Tipo da Integração",  "Desc. Tipo Integração",
        "Cod Evento",          "Descrição Evento",
        "Tipo Rubrica",
    ]
    cols_preencher = [
        "Código da Conta Débito",
        "Código da Conta Crédito",
        "Código do Histórico",
        "Complemento / Histórico",
    ]
    for col in cols_preencher:
        if col not in df.columns:
            df[col] = ""

    df_out = df[cols_fixas + cols_preencher].copy()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_out.to_excel(writer, index=False, sheet_name="Rubricas")
        wb = writer.book
        ws = writer.sheets["Rubricas"]

        hdr = wb.add_format({
            "bold": True, "bg_color": "#FF8000", "font_color": "#FFFFFF",
            "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True,
        })
        hdr_fill = wb.add_format({
            "bold": True, "bg_color": "#444444", "font_color": "#FFFFFF",
            "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True,
        })
        cell  = wb.add_format({"border": 1, "valign": "vcenter", "align": "left"})
        cnum  = wb.add_format({"border": 1, "valign": "vcenter", "align": "center"})
        alt   = wb.add_format({"border": 1, "valign": "vcenter", "align": "left",   "bg_color": "#FFF0E0"})
        anum  = wb.add_format({"border": 1, "valign": "vcenter", "align": "center", "bg_color": "#FFF0E0"})
        fill  = wb.add_format({"border": 1, "valign": "vcenter", "align": "left",   "bg_color": "#FFFACD"})
        falt  = wb.add_format({"border": 1, "valign": "vcenter", "align": "left",   "bg_color": "#FFFDE7"})

        for c, col in enumerate(df_out.columns):
            fmt = hdr_fill if col in cols_preencher else hdr
            ws.write(0, c, col, fmt)
        ws.set_row(0, 36)

        widths = {
            "Cód Centro de Custo":      9,  "Desc. Centro de Custo":   22,
            "Tipo da Integração":      10,  "Desc. Tipo Integração":   16,
            "Cod Evento":               9,  "Descrição Evento":        40,
            "Tipo Rubrica":            14,  "Código da Conta Débito":  16,
            "Código da Conta Crédito": 16,  "Código do Histórico":     14,
            "Complemento / Histórico": 40,
        }
        for c, col in enumerate(df_out.columns):
            ws.set_column(c, c, widths.get(col, 16))

        fill_idx = {df_out.columns.get_loc(c) for c in cols_preencher}
        num_idx  = {
            df_out.columns.get_loc("Cód Centro de Custo"),
            df_out.columns.get_loc("Tipo da Integração"),
            df_out.columns.get_loc("Cod Evento"),
        }

        for r, row in enumerate(df_out.itertuples(index=False), start=1):
            even = r % 2 == 0
            for c in range(len(df_out.columns)):
                val = row[c]
                if c in fill_idx:
                    ws.write(r, c, val if val else "", falt if even else fill)
                elif c in num_idx:
                    ws.write(r, c, val, anum if even else cnum)
                else:
                    ws.write(r, c, val, alt if even else cell)

        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, len(df_out), len(df_out.columns) - 1)

        instr = wb.add_format({
            "bold": True, "font_color": "#D64001", "font_size": 10,
            "align": "left", "valign": "vcenter",
        })
        ws.write(0, len(df_out.columns) + 1,
                 "⬅ Preencha as colunas em AMARELO e faça o upload novamente.", instr)
        ws.set_column(len(df_out.columns) + 1, len(df_out.columns) + 1, 55)

    return output.getvalue()

# ══════════════════════════════════════════════════════════════════════════════
#  GERAÇÃO DOS TXTs FINAIS
# ══════════════════════════════════════════════════════════════════════════════
def _nan_to_str(val) -> str:
    if val is None: return ""
    s = str(val).strip()
    return "" if s.lower() in ("nan", "none") else s

def gerar_evento_txt(df: pd.DataFrame, cod_empresa: str) -> bytes:
    HEADER = "\t".join([
        "Código da Empresa", "Centro de custo",
        "Código Sequencial da Integração",
        "Tipo da Integração (1 - Folha mensal; 2 - Empresa; 3 - Férias; 4 - Rescisao; 5 - Prov. Férias; 6 - Prov. 13)",
        "Descrição", "Código da Conta Débito", "Código da Conta Crédito",
        "Código do Histórico", "Complemento",
    ])
    df_ev   = df[df["Tipo da Integração"].isin([1, 2, 3, 4])].copy()
    linhas  = [HEADER]
    seq     = 1
    prev_cc = None
    for _, row in df_ev.iterrows():
        cc_key  = (row["Cód Centro de Custo"], row["Tipo da Integração"])
        emp_val = cod_empresa if cc_key != prev_cc else ""
        cc_val  = row["Cód Centro de Custo"] if cc_key != prev_cc else ""
        prev_cc = cc_key
        linhas.append("\t".join([
            str(emp_val), str(cc_val), str(seq),
            str(row["Tipo da Integração"]),
            _nan_to_str(row.get("Descrição Evento", "")),
            _nan_to_str(row.get("Código da Conta Débito", "")),
            _nan_to_str(row.get("Código da Conta Crédito", "")),
            _nan_to_str(row.get("Código do Histórico", "")),
            _nan_to_str(row.get("Complemento / Histórico", "")),
        ]))
        seq += 1
    return "\n".join(linhas).encode("utf-8-sig")

def gerar_integra_txt(df: pd.DataFrame, cod_empresa: str) -> bytes:
    HEADER = "\t".join([
        "Código da Empresa", "Centro de Custo",
        "Código Sequencial da Integração",
        "Tipo da Integração (1 - Folha mensal; 2 - Empresa; 3 - Férias; 4 - Rescisao; 5 - Prov. Férias; 6 - Prov. 13)",
        "Descrição", "Código da Conta Crédito", "Código da Conta Débito",
        "Código do Histórico",
    ])
    linhas = [HEADER]
    for seq, (_, row) in enumerate(df.iterrows(), start=1):
        linhas.append("\t".join([
            str(cod_empresa), "",
            str(seq),
            str(row["Tipo da Integração"]),
            _nan_to_str(row.get("Descrição Evento", "")),
            _nan_to_str(row.get("Código da Conta Crédito", "")),
            _nan_to_str(row.get("Código da Conta Débito", "")),
            _nan_to_str(row.get("Código do Histórico", "")),
        ]))
    return "\n".join(linhas).encode("utf-8-sig")

# ══════════════════════════════════════════════════════════════════════════════
#  HEADER PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(
    f"""
    <div style="background:#444444; padding:24px 28px 18px 28px;
                border-radius:8px; border-top:6px solid #FF8000; margin-bottom:28px;">
        <h2 style="color:#FF8000; margin:0; font-family:'Segoe UI',Arial,sans-serif;">
            📊 Rubricas Não Configuradas → Excel + TXT &nbsp;|&nbsp; {VERSAO}
        </h2>
        <p style="color:#DDDDDD; margin:6px 0 0 0; font-family:'Segoe UI',Arial,sans-serif;">
            <strong>① Gerar Excel intermediário</strong> (Tipo da Rubrica identificado pelo código do evento)
            → preencher contabilização →
            <strong>② Gerar TXTs finais</strong>
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
    cod_empresa = st.text_input("Código da Empresa", value="1",
                                help="Preenchido nos arquivos TXT gerados.")
    st.markdown("---")
    st.markdown("### 📋 Tipos de Integração")
    for cod, desc in TIPO_DESC.items():
        st.caption(f"{TIPO_ICONE[cod]} **{cod}** — {desc}")
    st.markdown("---")
    st.markdown("### 🏷️ Tipos de Rubrica")
    st.caption("🟢 Provento")
    st.caption("🔴 Desconto")
    st.caption("🔵 Informativa")
    st.caption("🟡 Inf. dedutora")
    st.markdown("---")
    st.markdown("### 📄 Arquivos gerados")
    st.caption("**evento_exemplo.txt** → aba *evento* | Tipos 1, 2, 3, 4")
    st.caption("**integra_exemplo.txt** → aba *Plan1* | Todos os tipos (1 a 6)")
    st.markdown("---")
    st.markdown(f"**Versão:** {VERSAO}")
    st.markdown("**Thomson Reuters | Domínio Sistemas**")

# ══════════════════════════════════════════════════════════════════════════════
#  INSTRUÇÕES
# ══════════════════════════════════════════════════════════════════════════════
with st.expander("📖 **Instruções de Uso** — clique para expandir", expanded=False):
    st.markdown("""
        <div class="instrucoes-box">

        <h4>🔹 Como funciona a identificação do Tipo de Rubrica</h4>
        <p>O sistema lê o <b>código do evento</b> do PDF de Rubricas Não Configuradas
        e busca esse mesmo código no PDF de Cadastro (Plano e Acumuladores) para
        identificar se é <b>Provento, Desconto, Informativa ou Inf. dedutora</b>.</p>

        <h4>🔹 Qual PDF usar em cada campo</h4>
        <table style="width:100%; border-collapse:collapse;">
          <tr style="background:#FF8000; color:white;">
            <th style="padding:6px;">Campo</th>
            <th style="padding:6px;">PDF correto</th>
          </tr>
          <tr style="background:#FFF0E0;">
            <td style="padding:6px;"><b>PDF 1</b></td>
            <td style="padding:6px;"><b>RubricasItens não Configurados.pdf</b><br>
            Título: <i>RELAÇÃO DE RUBRICAS/ITENS NÃO CONFIGURADOS</i></td>
          </tr>
          <tr>
            <td style="padding:6px;"><b>PDF 2</b></td>
            <td style="padding:6px;"><b>Rubricas.pdf</b><br>
            Título: <i>EMPRESA PADRÃO - PLANO E ACUMULADORES</i></td>
          </tr>
        </table>

        <h4>🔹 Etapa 1 — Gerar Excel intermediário</h4>
        <ol>
          <li>Informe o <b>Código da Empresa</b> na sidebar.</li>
          <li>Faça upload dos dois PDFs.</li>
          <li>Clique em <b>▶ Gerar Excel Intermediário</b>.</li>
          <li>Baixe o Excel e preencha as colunas em <b>amarelo</b>.</li>
        </ol>

        <h4>🔹 Etapa 2 — Gerar TXTs finais</h4>
        <ol>
          <li>Faça upload do Excel preenchido.</li>
          <li>Clique em <b>▶ Gerar TXTs Finais</b>.</li>
          <li>Baixe <b>evento_exemplo.txt</b> e <b>integra_exemplo.txt</b>.</li>
        </ol>

        <h4>⚠️ Colunas a preencher no Excel (em amarelo)</h4>
        <ul>
          <li><b>Código da Conta Débito</b></li>
          <li><b>Código da Conta Crédito</b></li>
          <li><b>Código do Histórico</b></li>
          <li><b>Complemento / Histórico</b></li>
        </ul>
        </div>
    """, unsafe_allow_html=True)

st.markdown("---")

# ══════════════════════════════════════════════════════════════════════════════
#  SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════
for k, v in {
    "df_rubricas":   None,
    "excel_interm":  None,
    "evento_bytes":  None,
    "integra_bytes": None,
    "etapa1_ok":     False,
    "etapa2_ok":     False,
    "log_parse":     [],
    "catalog_size":  0,
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ══════════════════════════════════════════════════════════════════════════════
#  ETAPA 1 — Upload PDFs + Gerar Excel intermediário
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
    <div class="step-box">
        <div class="step-title">① ETAPA 1 — Upload dos PDFs e geração do Excel intermediário</div>
    </div>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="alert-box">
        ⚠️ <b>Atenção aos arquivos:</b><br>
        • <b>Campo 1</b> → <code>RubricasItens não Configurados.pdf</code>
          (título: <i>RELAÇÃO DE RUBRICAS/ITENS NÃO CONFIGURADOS</i>)<br>
        • <b>Campo 2</b> → <code>Rubricas.pdf</code>
          (título: <i>EMPRESA PADRÃO - PLANO E ACUMULADORES</i>)
    </div>
""", unsafe_allow_html=True)

col_up1, col_up2 = st.columns(2)
with col_up1:
    st.markdown("**📄 PDF 1 — Rubricas/Itens Não Configurados**")
    st.caption("Título esperado: *RELAÇÃO DE RUBRICAS/ITENS NÃO CONFIGURADOS*")
    pdf_rubricas = st.file_uploader(
        "Rubricas Não Configuradas", type=["pdf"], key="up_rubricas",
        label_visibility="collapsed",
    )
with col_up2:
    st.markdown("**📄 PDF 2 — Cadastro de Eventos (Plano e Acumuladores)**")
    st.caption("Título esperado: *EMPRESA PADRÃO - PLANO E ACUMULADORES*")
    pdf_cadastro = st.file_uploader(
        "Cadastro de Eventos", type=["pdf"], key="up_cadastro",
        label_visibility="collapsed",
    )

# Validação em tempo real
pdf_rub_ok = pdf_cad_ok = False
bytes_rub = bytes_cad = None

if pdf_rubricas is not None:
    bytes_rub = pdf_rubricas.read()
    tipo_rub  = detectar_tipo_pdf(bytes_rub)
    if tipo_rub == "rubricas_nao_config":
        st.success("✅ PDF 1 identificado: **Relação de Rubricas/Itens Não Configurados**")
        pdf_rub_ok = True
    elif tipo_rub == "cadastro_eventos":
        st.error("❌ **PDF 1 incorreto!** Você enviou o PDF de Plano e Acumuladores no campo errado.")
    else:
        st.warning("⚠️ PDF 1: tipo não identificado automaticamente. Verifique se é o arquivo correto.")
        pdf_rub_ok = True

if pdf_cadastro is not None:
    bytes_cad = pdf_cadastro.read()
    tipo_cad  = detectar_tipo_pdf(bytes_cad)
    if tipo_cad == "cadastro_eventos":
        st.success("✅ PDF 2 identificado: **Empresa Padrão — Plano e Acumuladores**")
        pdf_cad_ok = True
    elif tipo_cad == "rubricas_nao_config":
        st.error("❌ **PDF 2 incorreto!** Você enviou o PDF de Rubricas Não Configuradas no campo errado.")
    else:
        st.warning("⚠️ PDF 2: tipo não identificado automaticamente. Verifique se é o arquivo correto.")
        pdf_cad_ok = True

ambos_pdfs = (pdf_rubricas is not None and pdf_cadastro is not None
              and pdf_rub_ok and pdf_cad_ok)

col_b1, col_b2 = st.columns([1, 1])
with col_b1:
    gerar_excel = st.button(
        "▶ Gerar Excel Intermediário",
        disabled=not ambos_pdfs,
        use_container_width=True,
        type="primary",
    )
with col_b2:
    limpar = st.button("🗑 Limpar tudo", use_container_width=True)

if limpar:
    for k in ["df_rubricas", "excel_interm", "evento_bytes",
              "integra_bytes", "log_parse"]:
        st.session_state[k] = [] if k == "log_parse" else None
    st.session_state.etapa1_ok    = False
    st.session_state.etapa2_ok    = False
    st.session_state.catalog_size = 0
    st.rerun()

if gerar_excel and ambos_pdfs:
    log = []

    with st.spinner("🔄 Lendo PDF de Rubricas Não Configuradas..."):
        df_rub = parse_rubricas_pdf(bytes_rub)
        log.append(f"PDF 1 (Rubricas Não Config.): {len(df_rub)} linhas extraídas.")

    if df_rub.empty:
        st.error("⚠️ Nenhum dado extraído do PDF 1.")
        with pdfplumber.open(BytesIO(bytes_rub)) as pdf:
            txt = pdf.pages[0].extract_text() or ""
        with st.expander("🔍 Debug — primeiras 40 linhas do PDF 1"):
            st.code("\n".join(txt.splitlines()[:40]))
    else:
        with st.spinner("🔄 Lendo PDF de Cadastro (identificando tipo por código)..."):
            catalog = parse_cadastro_eventos_pdf(bytes_cad)
            log.append(f"PDF 2 (Plano e Acumuladores): {len(catalog)} códigos mapeados.")

        # Cruza pelo CÓDIGO DO EVENTO
        df_rub["Tipo Rubrica"] = df_rub["Cod Evento"].map(
            lambda c: catalog.get(str(c), "—")
        )

        encontrados = (df_rub["Tipo Rubrica"] != "—").sum()
        nao_encontrados = (df_rub["Tipo Rubrica"] == "—").sum()
        log.append(f"Cruzamento por código: {encontrados} identificados | {nao_encontrados} sem correspondência (—).")

        if nao_encontrados > 0:
            codigos_nao_encontrados = df_rub[df_rub["Tipo Rubrica"] == "—"]["Cod Evento"].unique()
            log.append(f"Códigos sem tipo: {', '.join(codigos_nao_encontrados[:20])}"
                       + (" ..." if len(codigos_nao_encontrados) > 20 else ""))

        for col in ["Código da Conta Débito", "Código da Conta Crédito",
                    "Código do Histórico", "Complemento / Histórico"]:
            df_rub[col] = ""

        st.session_state.df_rubricas  = df_rub
        st.session_state.excel_interm = gerar_excel_intermediario(df_rub)
        st.session_state.etapa1_ok    = True
        st.session_state.log_parse    = log
        st.session_state.catalog_size = len(catalog)
        st.rerun()

# ── Resultado Etapa 1 ─────────────────────────────────────────────────────────
if st.session_state.etapa1_ok and st.session_state.df_rubricas is not None:
    df_rub = st.session_state.df_rubricas

    st.success(f"✅ **{len(df_rub)} registros** extraídos e cruzados com o cadastro!")

    if st.session_state.log_parse:
        for msg in st.session_state.log_parse:
            st.caption(f"ℹ️ {msg}")

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("📄 Registros",         len(df_rub))
    m2.metric("🏢 Centros de Custo",  df_rub["Cód Centro de Custo"].nunique())
    m3.metric("🔢 Tipos Integração",  df_rub["Tipo da Integração"].nunique())
    m4.metric("🎯 Eventos Únicos",    df_rub["Cod Evento"].nunique())
    m5.metric("🏷️ Tipo Identificado", (df_rub["Tipo Rubrica"] != "—").sum())

    # Prévia com destaque por tipo de rubrica
    with st.expander("👁️ Prévia dos dados (primeiros 30 registros)"):
        df_prev = df_rub.head(30).copy()
        try:
            styled = df_prev.style.map(color_tipo_rubrica, subset=["Tipo Rubrica"])
        except AttributeError:
            styled = df_prev.style.applymap(color_tipo_rubrica, subset=["Tipo Rubrica"])
        st.dataframe(styled, use_container_width=True, hide_index=True)

    # Resumo por Tipo de Rubrica
    with st.expander("📊 Resumo por Tipo de Rubrica"):
        resumo_tr = (
            df_rub.groupby("Tipo Rubrica")
            .agg(Registros=("Cod Evento", "count"),
                 Eventos_Únicos=("Cod Evento", "nunique"))
            .reset_index()
            .sort_values("Registros", ascending=False)
        )
        st.dataframe(resumo_tr, use_container_width=True, hide_index=True)

    # Eventos sem tipo identificado
    sem_tipo = df_rub[df_rub["Tipo Rubrica"] == "—"]
    if not sem_tipo.empty:
        with st.expander(f"⚠️ {len(sem_tipo)} registros sem Tipo Rubrica identificado"):
            st.caption("Esses códigos não foram encontrados no PDF de Cadastro de Eventos.")
            st.dataframe(
                sem_tipo[["Cód Centro de Custo", "Desc. Centro de Custo",
                           "Tipo da Integração", "Cod Evento", "Descrição Evento"]],
                use_container_width=True, hide_index=True
            )

    # Resumo por Tipo de Integração
    with st.expander("📊 Resumo por Tipo de Integração"):
        resumo = (
            df_rub.groupby(["Tipo da Integração", "Desc. Tipo Integração"])
            .agg(Registros=("Cod Evento", "count"),
                 Centros=("Cód Centro de Custo", "nunique"),
                 Eventos_Únicos=("Cod Evento", "nunique"))
            .reset_index()
        )
        st.dataframe(resumo, use_container_width=True, hide_index=True)

    st.markdown("#### ⬇️ Baixe o Excel, preencha as colunas em amarelo e siga para a Etapa 2")
    st.download_button(
        label="📥 Baixar Excel Intermediário (preencher contabilização)",
        data=st.session_state.excel_interm,
        file_name="rubricas_para_preencher.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

st.markdown("---")

# ══════════════════════════════════════════════════════════════════════════════
#  ETAPA 2 — Upload Excel preenchido + Gerar TXTs finais
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
    <div class="step-box">
        <div class="step-title">② ETAPA 2 — Upload do Excel preenchido e geração dos TXTs finais</div>
    </div>
""", unsafe_allow_html=True)

st.markdown("**📊 Excel preenchido (com Débito, Crédito, Histórico e Complemento)**")
excel_preenchido = st.file_uploader(
    "Excel preenchido", type=["xlsx", "xls"], key="up_excel",
    label_visibility="collapsed",
)

gerar_txt = st.button(
    "▶ Gerar TXTs Finais",
    disabled=(excel_preenchido is None),
    use_container_width=True,
    type="primary",
)

if gerar_txt and excel_preenchido is not None:
    with st.spinner("🔄 Lendo Excel preenchido..."):
        try:
            df_filled = pd.read_excel(excel_preenchido, dtype=str).fillna("")
            df_filled["Tipo da Integração"] = pd.to_numeric(
                df_filled["Tipo da Integração"], errors="coerce"
            ).fillna(0).astype(int)

            st.session_state.evento_bytes  = gerar_evento_txt(df_filled, cod_empresa)
            st.session_state.integra_bytes = gerar_integra_txt(df_filled, cod_empresa)
            st.session_state.etapa2_ok     = True
            st.rerun()
        except Exception as e:
            st.error(f"❌ Erro ao processar o Excel: {e}")

# ── Resultado Etapa 2 ─────────────────────────────────────────────────────────
if st.session_state.etapa2_ok:
    st.success("✅ Arquivos TXT gerados com sucesso!")

    ev_lines = st.session_state.evento_bytes.decode("utf-8-sig").splitlines()
    in_lines = st.session_state.integra_bytes.decode("utf-8-sig").splitlines()

    d1, d2 = st.columns(2)
    with d1:
        st.markdown("**📄 evento_exemplo.txt**")
        st.caption(f"Aba: evento | Tipos 1, 2, 3, 4 | {len(ev_lines)-1} registros")
        st.download_button(
            label="⬇ Baixar evento_exemplo.txt",
            data=st.session_state.evento_bytes,
            file_name="evento_exemplo.txt",
            mime="text/plain",
            use_container_width=True,
        )
    with d2:
        st.markdown("**📄 integra_exemplo.txt**")
        st.caption(f"Aba: Plan1 | Tipos 1 a 6 | {len(in_lines)-1} registros")
        st.download_button(
            label="⬇ Baixar integra_exemplo.txt",
            data=st.session_state.integra_bytes,
            file_name="integra_exemplo.txt",
            mime="text/plain",
            use_container_width=True,
        )

    tab1, tab2 = st.tabs(["📄 evento_exemplo.txt", "📄 integra_exemplo.txt"])
    with tab1:
        st.code("\n".join(ev_lines[:21]), language="text")
        if len(ev_lines) > 21:
            st.caption(f"... e mais {len(ev_lines)-21} linhas")
    with tab2:
        st.code("\n".join(in_lines[:21]), language="text")
        if len(in_lines) > 21:
            st.caption(f"... e mais {len(in_lines)-21} linhas")
