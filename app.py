# ============================================================
# app_integracao_dominio.py  –  Integração Contábil Domínio V3.0
# Entradas:
#   1. RubricasItens não Configurados.pdf  → eventos sem config contábil
#   2. rubricas.txt                        → catálogo de tipos de rubrica
# Fluxo:
#   ETAPA 1 → Gera Excel para preenchimento das contas
#   ETAPA 2 → Importa Excel preenchido → gera evento exemplo.xlsx + integra exemplo.xls
# ============================================================

import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

VERSAO = "V3.0"

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
        </style>
    """, unsafe_allow_html=True)


# ==============================
# PARSE DO TXT DE RUBRICAS
# ==============================
def parse_rubricas_txt(file_bytes: bytes, log: list) -> dict:
    catalog = {}
    TIPO_MAP = {"P": "Provento", "D": "Desconto", "I": "Informativa", "ID": "Inf. Dedutora"}
    try:
        texto = file_bytes.decode("latin-1", errors="replace")
    except Exception as e:
        log.append(f"ERRO ao decodificar rubricas.txt: {e}")
        return catalog
    for raw in texto.splitlines():
        raw = raw.strip()
        if not raw:
            continue
        partes = raw.split("\t")
        if len(partes) < 5:
            continue
        cod = partes[2].strip()
        descricao = partes[3].strip()
        tipo_raw = partes[4].strip().upper()
        if not cod:
            continue
        tipo_norm = TIPO_MAP.get(tipo_raw)
        if tipo_norm is None:
            continue
        if cod not in catalog:
            catalog[cod] = {"tipo": tipo_norm, "descricao": descricao}
    log.append(f"rubricas.txt: {len(catalog)} código(s) mapeado(s).")
    return catalog


# ==============================
# PARSE DO PDF DE ITENS NÃO CONFIGURADOS
# ==============================
IGNORE_PATTERNS_NAO_CONFIG = [
    r"^RELAÇÃO DE RUBRICAS",
    r"^Página",
    r"^Emissão",
    r"^Hora:",
    r"^Empresa:",
    r"^Código\s+Descrição",
    r"^\s*$",
]

SECAO_TIPO_FOLHA = {
    "Folha Normal": "1",
    "Empresa": "2",
    "Férias": "3",
    "Rescisão": "4",
    "Provisão de Férias": "5",
    "Provisão de 13º": "6",
    "Provisão de 13o": "6",
}

SECAO_TIPO_FOLHA_DESC = {
    "1": "Folha Normal",
    "2": "Empresa",
    "3": "Férias",
    "4": "Rescisão",
    "5": "Provisão de Férias",
    "6": "Provisão de 13º",
}

RE_SECAO = re.compile(
    r"^(Folha Normal|Empresa|Férias|Rescisão|"
    r"Provisão de Férias|Provisão de 13º|Provisão de 13o)$",
    re.IGNORECASE,
)
RE_CC = re.compile(r"^Centro de Custo:\s*(\d+)\s+(.+)$", re.IGNORECASE)
RE_EVENT = re.compile(r"^\s*(\d+)\s+(.+)$")


def should_ignore(line: str) -> bool:
    for pat in IGNORE_PATTERNS_NAO_CONFIG:
        if re.search(pat, line, re.IGNORECASE):
            return True
    return False


def parse_nao_configurados_pdf(file_bytes: bytes, log: list) -> list:
    eventos = []
    vistos = set()
    tipo_folha_atual = "1"
    cc_cod_atual = ""
    cc_nome_atual = ""

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for raw in text.splitlines():
                line = raw.strip()
                if not line:
                    continue
                m_sec = RE_SECAO.match(line)
                if m_sec:
                    sec = m_sec.group(1).strip()
                    for k, v in SECAO_TIPO_FOLHA.items():
                        if k.lower() in sec.lower():
                            tipo_folha_atual = v
                            break
                    continue
                m_cc = RE_CC.match(line)
                if m_cc:
                    cc_cod_atual = m_cc.group(1).strip()
                    cc_nome_atual = m_cc.group(2).strip()
                    continue
                if should_ignore(line):
                    continue
                m_ev = RE_EVENT.match(line)
                if m_ev:
                    cod = m_ev.group(1).strip()
                    desc = m_ev.group(2).strip()
                    if not cod.isdigit():
                        continue
                    chave = (cod, tipo_folha_atual, cc_cod_atual)
                    if chave not in vistos:
                        vistos.add(chave)
                        eventos.append({
                            "cod": cod,
                            "descricao_pdf": desc,
                            "tipo_folha": tipo_folha_atual,
                            "tipo_folha_desc": SECAO_TIPO_FOLHA_DESC.get(tipo_folha_atual, tipo_folha_atual),
                            "centro_custo_cod": cc_cod_atual,
                            "centro_custo_nome": cc_nome_atual,
                        })

    log.append(f"PDF: {len(eventos)} evento(s) extraído(s).")
    return eventos


# ==============================
# ETAPA 1 — GERA EXCEL PARA PREENCHIMENTO
# ==============================
def gerar_excel_configuracao(eventos: list, catalog: dict, cod_empresa: str, log: list) -> bytes:
    linhas = []
    for ev in eventos:
        cod = ev["cod"]
        info = catalog.get(cod, {})
        tipo = info.get("tipo", "⚠️ Não encontrado")
        desc_rubrica = info.get("descricao", ev["descricao_pdf"])
        linhas.append({
            # Identificação (somente leitura)
            "Cód. Empresa":          cod_empresa,
            "Cód. Evento":           cod,
            "Descrição (PDF)":       ev["descricao_pdf"],
            "Descrição (Rubricas)":  desc_rubrica,
            "Tipo Rubrica":          tipo,
            "Tipo Folha (Nº)":       ev["tipo_folha"],
            "Tipo Folha":            ev["tipo_folha_desc"],
            "Cód. Centro de Custo":  ev["centro_custo_cod"],
            "Centro de Custo":       ev["centro_custo_nome"],
            # Campos a preencher
            "Conta Débito":          "",
            "Conta Crédito":         "",
            "Cód. Histórico":        "",
            "Histórico":             "",
            "Observação":            "",
        })

    df = pd.DataFrame(linhas)
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Configuração", index=False)
        ws = writer.sheets["Configuração"]
        _formatar_planilha_config(ws, df)

    output.seek(0)
    log.append(f"Excel de configuração gerado: {len(linhas)} linha(s).")
    return output.read()


def _formatar_planilha_config(ws, df: pd.DataFrame):
    borda = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    larguras = {
        "A": 12, "B": 12, "C": 38, "D": 38, "E": 16,
        "F": 14, "G": 20, "H": 18, "I": 22,
        "J": 16, "K": 16, "L": 14, "M": 42, "N": 30,
    }
    for col, w in larguras.items():
        ws.column_dimensions[col].width = w

    COR_HEADER_INFO  = "444444"
    COR_HEADER_FILL  = "FF8000"
    COR_FILL_PREENCHER = "FFF8F0"
    COLS_PREENCHER = {10, 11, 12, 13, 14}  # J-N (1-based)

    TIPO_COR = {
        "Provento":      "D4EDDA",
        "Desconto":      "F8D7DA",
        "Informativa":   "CCE5FF",
        "Inf. Dedutora": "FFF3CD",
    }

    for col_idx, cell in enumerate(ws[1], start=1):
        if col_idx in COLS_PREENCHER:
            cell.fill = PatternFill("solid", fgColor=COR_HEADER_FILL)
            cell.font = Font(bold=True, color="FFFFFF", size=10)
        else:
            cell.fill = PatternFill("solid", fgColor=COR_HEADER_INFO)
            cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = borda
    ws.row_dimensions[1].height = 32

    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        tipo_val = ws.cell(row=row_idx, column=5).value or ""
        cor_linha = TIPO_COR.get(tipo_val, "E2E3E5")
        for col_idx, cell in enumerate(row, start=1):
            cell.border = borda
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if col_idx in COLS_PREENCHER:
                cell.fill = PatternFill("solid", fgColor=COR_FILL_PREENCHER)
                cell.font = Font(size=10)
            else:
                cell.fill = PatternFill("solid", fgColor=cor_linha)
                cell.font = Font(size=10)
        ws.row_dimensions[row_idx].height = 18

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


# ==============================
# ETAPA 2 — IMPORTA EXCEL PREENCHIDO → GERA ARQUIVOS FINAIS
# ==============================
def ler_excel_preenchido(file_bytes: bytes, log: list) -> pd.DataFrame | None:
    try:
        xls = pd.ExcelFile(BytesIO(file_bytes))
    except Exception as e:
        log.append(f"ERRO ao abrir Excel preenchido: {e}")
        return None

    sheet = None
    for candidate in ["Configuração", "configuracao", "Plan1", "Sheet1"]:
        if candidate in xls.sheet_names:
            sheet = candidate
            break
    if not sheet:
        sheet = xls.sheet_names[0]

    try:
        df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet, dtype=str)
    except Exception as e:
        log.append(f"ERRO ao ler aba '{sheet}': {e}")
        return None

    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")
    log.append(f"Excel preenchido lido: {len(df)} linha(s) na aba '{sheet}'.")
    return df


def _limpa(val) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s.lower() in ("nan", "none", "") else s


def gerar_arquivos_finais(df: pd.DataFrame, cod_empresa_padrao: str, log: list) -> tuple[bytes, bytes]:
    """
    Retorna (evento_exemplo_xlsx, integra_exemplo_xls)

    Formato evento exemplo.xlsx (aba 'evento'):
      Código da Empresa | Centro de custo | Código Sequencial da Integração |
      Tipo da Integração | Descrição | Código da Conta Débito |
      Código da Conta Crédito | Código do Histórico | Complemento

    Formato evento exemplo.xlsx (aba 'integra'):
      Código da Empresa | Separador | Código Sequencial da Integração |
      Tipo da Integração | Código da Rúbrica Selecionada

    Formato integra exemplo.xls (Plan1):
      Código da Empresa | Centro de Custo | Código Sequencial da Integração |
      Tipo da Integração | Descrição | Código da Conta Crédito |
      Código da Conta Débito | Código do Histórico
    """
    # Mapeamento de colunas do Excel preenchido
    col_map = {}
    for col in df.columns:
        cl = col.lower()
        if "cód. empresa" in cl or "cod. empresa" in cl:
            col_map["empresa"] = col
        elif "cód. evento" in cl or "cod. evento" in cl:
            col_map["seq"] = col
        elif "tipo folha (nº)" in cl or "tipo folha (n" in cl:
            col_map["tipo"] = col
        elif "descrição (rubricas)" in cl:
            col_map["desc"] = col
        elif "descrição (pdf)" in cl and "desc" not in col_map:
            col_map["desc"] = col
        elif "cód. centro de custo" in cl:
            col_map["cc"] = col
        elif "conta débito" in cl or "conta debito" in cl:
            col_map["debito"] = col
        elif "conta crédito" in cl or "conta credito" in cl:
            col_map["credito"] = col
        elif "cód. histórico" in cl or "cod. historico" in cl:
            col_map["historico"] = col
        elif "histórico" in cl and "cód" not in cl and "cod" not in cl:
            col_map["historico_texto"] = col
        elif "observação" in cl:
            col_map["observacao"] = col

    linhas_evento = []
    linhas_integra = []
    linhas_integra_xls = []

    sem_conta = 0
    com_conta = 0

    for _, row in df.iterrows():
        empresa  = _limpa(row.get(col_map.get("empresa", ""), "")) or cod_empresa_padrao
        seq      = _limpa(row.get(col_map.get("seq", ""), ""))
        tipo     = _limpa(row.get(col_map.get("tipo", ""), ""))
        desc     = _limpa(row.get(col_map.get("desc", ""), ""))
        cc       = _limpa(row.get(col_map.get("cc", ""), ""))
        debito   = _limpa(row.get(col_map.get("debito", ""), ""))
        credito  = _limpa(row.get(col_map.get("credito", ""), ""))
        historico = _limpa(row.get(col_map.get("historico", ""), ""))
        complemento = _limpa(row.get(col_map.get("historico_texto", ""), ""))

        if not seq:
            continue

        if debito or credito:
            com_conta += 1
        else:
            sem_conta += 1

        # ── Aba 'evento' do evento exemplo.xlsx ──────────────────────────
        linhas_evento.append({
            "Código da Empresa": empresa,
            "Centro de custo": cc,
            "Código Sequencial da Integração": seq,
            "Tipo da Integração (1 - Folha mensal; 2 - Empresa; 3 - Férias; 4 - Rescisao; 5 - Prov. Férias; 6 - Prov. 13)": tipo,
            "Descrição": desc,
            "Código da Conta Débito": debito,
            "Código da Conta Crédito": credito,
            "Código do Histórico": historico,
            "Complemento": complemento,
        })

        # ── Aba 'integra' do evento exemplo.xlsx ─────────────────────────
        linhas_integra.append({
            "Código da Empresa": empresa,
            "Separador": "0",
            "Código Sequencial da Integração": seq,
            "Tipo da Integração (1 - Folha mensal; 2 - Empresa; 3 - Férias; 4 - Rescisao; 5 - Prov. Férias; 6 - Prov. 13)": tipo,
            "Código da Rúbrica Selecionada": seq,
        })

        # ── Plan1 do integra exemplo.xls ─────────────────────────────────
        linhas_integra_xls.append({
            "Código da Empresa": empresa,
            "Centro de Custo": cc,
            "Código Sequencial da Integração": seq,
            "Tipo da Integração (1 - Folha mensal; 2 - Empresa; 3 - Férias; 4 - Rescisao; 5 - Prov. Férias; 6 - Prov. 13)": tipo,
            "Descrição": desc,
            "Código da Conta Crédito": credito,
            "Código da Conta Débito": debito,
            "Código do Histórico": historico,
        })

    log.append(
        f"Arquivos gerados → Com conta: {com_conta} | Sem conta: {sem_conta}"
    )

    # ── Gera evento exemplo.xlsx (2 abas: integra + evento) ──────────────
    buf_evento = BytesIO()
    with pd.ExcelWriter(buf_evento, engine="openpyxl") as writer:
        pd.DataFrame(linhas_integra).to_excel(writer, sheet_name="integra", index=False)
        pd.DataFrame(linhas_evento).to_excel(writer, sheet_name="evento", index=False)
        # Formata ambas as abas
        for sheet_name in ["integra", "evento"]:
            ws = writer.sheets[sheet_name]
            _formatar_planilha_saida(ws)
    buf_evento.seek(0)

    # ── Gera integra exemplo.xls (Plan1) ─────────────────────────────────
    buf_integra = BytesIO()
    with pd.ExcelWriter(buf_integra, engine="openpyxl") as writer:
        pd.DataFrame(linhas_integra_xls).to_excel(writer, sheet_name="Plan1", index=False)
        ws = writer.sheets["Plan1"]
        _formatar_planilha_saida(ws)
    buf_integra.seek(0)

    return buf_evento.read(), buf_integra.read()


def _formatar_planilha_saida(ws):
    borda = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    # Cabeçalho
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor="444444")
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = borda
    ws.row_dimensions[1].height = 32

    # Dados
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = borda
            cell.alignment = Alignment(vertical="center")
            cell.font = Font(size=10)

    # Larguras automáticas
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 50)

    ws.freeze_panes = "A2"


# ==============================
# INTERFACE STREAMLIT
# ==============================
def main():
    st.set_page_config(
        page_title="Domínio | Integração Contábil",
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
            <h2 style="color:#FF8000; margin:0; font-family:'Segoe UI',Arial,sans-serif;">
                📊 Integração Contábil — Domínio Sistemas &nbsp;|&nbsp; {VERSAO}
            </h2>
            <p style="color:#DDDDDD; margin:6px 0 0 0; font-family:'Segoe UI',Arial,sans-serif;">
                <b>Etapa 1:</b> Gera Excel para preenchimento das contas contábeis.<br>
                <b>Etapa 2:</b> Importa Excel preenchido → gera <b>evento exemplo.xlsx</b>
                e <b>integra exemplo.xls</b>.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── Sidebar ───────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("### ⚙️ Configurações")
        cod_empresa = st.text_input("Código da empresa", value="45")
        st.markdown("---")
        st.markdown("### 🎨 Legenda")
        st.markdown("🟢 Verde → Provento")
        st.markdown("🔴 Vermelho → Desconto")
        st.markdown("🔵 Azul → Informativa")
        st.markdown("🟡 Amarelo → Inf. Dedutora")
        st.markdown("🟠 Laranja → Campos a preencher")
        st.markdown("---")
        st.markdown(f"**Versão:** {VERSAO}")
        st.markdown("**Thomson Reuters | Domínio Sistemas**")

    # ── Session state ─────────────────────────────────────────────────────
    for k, v in {
        "log": [f"Pronto. Versão {VERSAO}"],
        "excel_config": None,
        "evento_xlsx": None,
        "integra_xls": None,
        "df_preview": None,
        "n_eventos": 0,
    }.items():
        if k not in st.session_state:
            st.session_state[k] = v

    # ══════════════════════════════════════════════════════════════════════
    # ETAPA 1
    # ══════════════════════════════════════════════════════════════════════
    st.markdown("## 📋 Etapa 1 — Gerar Excel para Preenchimento")

    col1, col2 = st.columns(2)
    with col1:
        pdf_file = st.file_uploader(
            "1️⃣ PDF — Rubricas/Itens Não Configurados",
            type=["pdf"],
            key="pdf_etapa1",
        )
    with col2:
        txt_file = st.file_uploader(
            "2️⃣ TXT — Rubricas (catálogo de tipos)",
            type=["txt"],
            key="txt_etapa1",
        )

    col_btn1, col_btn2 = st.columns([1, 1])
    with col_btn1:
        gerar_excel = st.button(
            "▶ Gerar Excel de Configuração",
            disabled=(pdf_file is None or txt_file is None),
            use_container_width=True,
            type="primary",
        )
    with col_btn2:
        if st.button("🗑 Limpar tudo", use_container_width=True):
            for k in ["log", "excel_config", "evento_xlsx", "integra_xls", "df_preview", "n_eventos"]:
                st.session_state[k] = [] if k == "log" else None if k != "n_eventos" else 0
            st.session_state.log = ["Campos limpos."]
            st.rerun()

    if gerar_excel and pdf_file and txt_file:
        log = ["[Etapa 1] Iniciando..."]
        with st.spinner("Lendo rubricas.txt..."):
            catalog = parse_rubricas_txt(txt_file.read(), log)
        with st.spinner("Lendo PDF de Itens Não Configurados..."):
            eventos = parse_nao_configurados_pdf(pdf_file.read(), log)
        if not eventos:
            log.append("AVISO: Nenhum evento encontrado no PDF.")
        else:
            with st.spinner("Gerando Excel..."):
                excel_bytes = gerar_excel_configuracao(eventos, catalog, cod_empresa, log)
            st.session_state.excel_config = excel_bytes
            st.session_state.n_eventos = len(eventos)

            # Preview
            linhas_prev = []
            for ev in eventos:
                cod = ev["cod"]
                info = catalog.get(cod, {})
                linhas_prev.append({
                    "Código": cod,
                    "Descrição": ev["descricao_pdf"],
                    "Tipo": info.get("tipo", "⚠️"),
                    "Tipo Folha": ev["tipo_folha_desc"],
                    "Centro Custo": ev["centro_custo_nome"],
                })
            st.session_state.df_preview = pd.DataFrame(linhas_prev)
        st.session_state.log = log
        st.rerun()

    # Resultado Etapa 1
    if st.session_state.excel_config is not None:
        st.success(f"✅ Excel gerado — {st.session_state.n_eventos} evento(s)")
        st.download_button(
            label="⬇ Baixar Excel de Configuração",
            data=st.session_state.excel_config,
            file_name="configuracao_rubricas_dominio.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

        if st.session_state.df_preview is not None:
            df = st.session_state.df_preview
            total = len(df)
            p = len(df[df["Tipo"] == "Provento"])
            d = len(df[df["Tipo"] == "Desconto"])
            i = len(df[df["Tipo"] == "Informativa"])
            id_ = len(df[df["Tipo"] == "Inf. Dedutora"])
            nf = len(df[df["Tipo"].str.startswith("⚠️", na=False)])
            m1, m2, m3, m4, m5, m6 = st.columns(6)
            m1.metric("📋 Total", total)
            m2.metric("🟢 Proventos", p)
            m3.metric("🔴 Descontos", d)
            m4.metric("🔵 Informativas", i)
            m5.metric("🟡 Inf. Ded.", id_)
            m6.metric("⚠️ Não id.", nf)

            def hl(row):
                t = str(row.get("Tipo", ""))
                if t == "Provento":      return ["background-color:#d4edda"] * len(row)
                if t == "Desconto":      return ["background-color:#f8d7da"] * len(row)
                if t == "Informativa":   return ["background-color:#cce5ff"] * len(row)
                if t == "Inf. Dedutora": return ["background-color:#fff3cd"] * len(row)
                return ["background-color:#e2e3e5"] * len(row)

            st.dataframe(df.head(50).style.apply(hl, axis=1), use_container_width=True)

    st.markdown("---")

    # ══════════════════════════════════════════════════════════════════════
    # ETAPA 2
    # ══════════════════════════════════════════════════════════════════════
    st.markdown("## 📥 Etapa 2 — Importar Excel Preenchido → Gerar Arquivos Finais")
    st.markdown("""
    1. Baixe o Excel da Etapa 1
    2. Preencha as colunas **Conta Débito**, **Conta Crédito**, **Cód. Histórico**, **Histórico** e **Observação**
    3. Faça upload do Excel preenchido abaixo e clique em **▶ Gerar Arquivos Finais**
    """)

    excel_preenchido = st.file_uploader(
        "3️⃣ Excel Preenchido (.xlsx)",
        type=["xlsx", "xls"],
        key="excel_etapa2",
    )

    col_btn3, _ = st.columns([1, 1])
    with col_btn3:
        gerar_finais = st.button(
            "▶ Gerar Arquivos Finais",
            disabled=(excel_preenchido is None),
            use_container_width=True,
            type="primary",
        )

    if gerar_finais and excel_preenchido:
        log = st.session_state.log or []
        log.append("[Etapa 2] Iniciando...")
        with st.spinner("Lendo Excel preenchido..."):
            df_preen = ler_excel_preenchido(excel_preenchido.read(), log)
        if df_preen is not None:
            with st.spinner("Gerando arquivos finais..."):
                evento_bytes, integra_bytes = gerar_arquivos_finais(df_preen, cod_empresa, log)
            st.session_state.evento_xlsx = evento_bytes
            st.session_state.integra_xls = integra_bytes
        st.session_state.log = log
        st.rerun()

    # Resultado Etapa 2
    if st.session_state.evento_xlsx is not None:
        st.success("✅ Arquivos finais gerados!")

        col_d1, col_d2 = st.columns(2)
        with col_d1:
            st.download_button(
                label="⬇ Baixar evento exemplo.xlsx",
                data=st.session_state.evento_xlsx,
                file_name="evento exemplo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )
        with col_d2:
            st.download_button(
                label="⬇ Baixar integra exemplo.xlsx",
                data=st.session_state.integra_xls,
                file_name="integra exemplo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )

        st.markdown("---")
        st.markdown("### 📋 Estrutura dos arquivos gerados")
        col_i1, col_i2 = st.columns(2)
        with col_i1:
            st.markdown("""
            **evento exemplo.xlsx**
            - Aba `integra`: Código Empresa | 0 | Seq | Tipo | Rúbrica
            - Aba `evento`: Empresa | CC | Seq | Tipo | Descrição | Débito | Crédito | Histórico | Complemento
            """)
        with col_i2:
            st.markdown("""
            **integra exemplo.xlsx**
            - Aba `Plan1`: Empresa | CC | Seq | Tipo | Descrição | Crédito | Débito | Histórico
            """)

    # ── Log ───────────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("**Log de processamento**")
    log_texto = "\n".join(st.session_state.log)
    tem_erro = any(str(l).upper().startswith("ERRO") for l in st.session_state.log)
    cor_borda = "#D32F2F" if tem_erro else "#388E3C"
    st.markdown(
        f"""
        <div style="background:#FCFCFC; border:1px solid {cor_borda};
                    border-radius:6px; padding:14px;
                    font-family:Consolas,monospace; font-size:13px;
                    white-space:pre-wrap; max-height:300px;
                    overflow-y:auto; color:#1F1F1F;">
{log_texto}
        </div>
        """,
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
