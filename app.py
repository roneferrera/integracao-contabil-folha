# ============================================================
# gerar_excel_configuracao.py
# Gera Excel para preenchimento de contas contábeis
# Entradas:
#   1. RubricasItens não Configurados.pdf
#   2. rubricas.txt
# Saída:
#   Excel com colunas prontas para preenchimento
# ============================================================

import streamlit as st
import pdfplumber
import pandas as pd
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
        </style>
    """, unsafe_allow_html=True)


# ==============================
# PARSE DO TXT DE RUBRICAS
# ==============================
def parse_rubricas_txt(file_bytes: bytes, log: list) -> dict:
    """
    Lê rubricas.txt e retorna:
    { cod (str): {"tipo": str, "descricao": str} }
    Tipos: P=Provento | D=Desconto | I=Informativa | ID=Inf.Dedutora
    """
    catalog = {}
    TIPO_MAP = {
        "P":  "Provento",
        "D":  "Desconto",
        "I":  "Informativa",
        "ID": "Inf. Dedutora",
    }
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
        cod       = partes[2].strip()
        descricao = partes[3].strip()
        tipo_raw  = partes[4].strip().upper()
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
    "Folha Normal":       "1",
    "Empresa":            "2",
    "Férias":             "3",
    "Rescisão":           "4",
    "Provisão de Férias": "5",
    "Provisão de 13º":    "6",
    "Provisão de 13o":    "6",
}

SECAO_TIPO_FOLHA_DESC = {
    "1": "Folha Normal",
    "2": "Empresa",
    "3": "Férias",
    "4": "Rescisão",
    "5": "Provisão de Férias",
    "6": "Provisão de 13º",
}

def should_ignore(line: str) -> bool:
    for pat in IGNORE_PATTERNS_NAO_CONFIG:
        if re.search(pat, line, re.IGNORECASE):
            return True
    return False

RE_SECAO = re.compile(
    r"^(Folha Normal|Empresa|Férias|Rescisão|"
    r"Provisão de Férias|Provisão de 13º|Provisão de 13o)$",
    re.IGNORECASE,
)
RE_CC    = re.compile(r"^Centro de Custo:\s*(\d+)\s+(.+)$", re.IGNORECASE)
RE_EVENT = re.compile(r"^\s*(\d+)\s+(.+)$")


def parse_nao_configurados_pdf(file_bytes: bytes, log: list) -> list:
    """
    Retorna lista de dicts:
    { cod, descricao_pdf, tipo_folha, tipo_folha_desc,
      centro_custo_cod, centro_custo_nome }
    """
    eventos   = []
    vistos    = set()
    tipo_folha_atual   = "1"
    cc_cod_atual       = ""
    cc_nome_atual      = ""

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for raw in text.splitlines():
                line = raw.strip()
                if not line:
                    continue

                # Seção de tipo de folha
                m_sec = RE_SECAO.match(line)
                if m_sec:
                    sec = m_sec.group(1).strip()
                    for k, v in SECAO_TIPO_FOLHA.items():
                        if k.lower() in sec.lower():
                            tipo_folha_atual = v
                            break
                    continue

                # Centro de Custo
                m_cc = RE_CC.match(line)
                if m_cc:
                    cc_cod_atual  = m_cc.group(1).strip()
                    cc_nome_atual = m_cc.group(2).strip()
                    continue

                if should_ignore(line):
                    continue

                # Evento
                m_ev = RE_EVENT.match(line)
                if m_ev:
                    cod  = m_ev.group(1).strip()
                    desc = m_ev.group(2).strip()
                    if not cod.isdigit():
                        continue
                    chave = (cod, tipo_folha_atual, cc_cod_atual)
                    if chave not in vistos:
                        vistos.add(chave)
                        eventos.append({
                            "cod":               cod,
                            "descricao_pdf":     desc,
                            "tipo_folha":        tipo_folha_atual,
                            "tipo_folha_desc":   SECAO_TIPO_FOLHA_DESC.get(tipo_folha_atual, tipo_folha_atual),
                            "centro_custo_cod":  cc_cod_atual,
                            "centro_custo_nome": cc_nome_atual,
                        })

    log.append(f"PDF: {len(eventos)} evento(s) extraído(s) (únicos por código+tipo+CC).")
    return eventos


# ==============================
# GERAÇÃO DO EXCEL
# ==============================
def gerar_excel(eventos: list, catalog: dict, log: list) -> bytes:
    """
    Gera Excel com:
    - Aba "Configuração" → tabela principal para preenchimento
    - Aba "Resumo"       → estatísticas
    """
    linhas = []
    sem_tipo = []

    for ev in eventos:
        cod  = ev["cod"]
        info = catalog.get(cod, {})
        tipo = info.get("tipo", "")
        desc_rubrica = info.get("descricao", ev["descricao_pdf"])

        if not tipo:
            sem_tipo.append(cod)

        linhas.append({
            # ── Identificação ──────────────────────────────────────
            "Código Evento":        cod,
            "Descrição (PDF)":      ev["descricao_pdf"],
            "Descrição (Rubricas)": desc_rubrica,
            "Tipo Rubrica":         tipo or "⚠️ Não encontrado",
            "Tipo Folha (Nº)":      ev["tipo_folha"],
            "Tipo Folha":           ev["tipo_folha_desc"],
            "Cód. Centro de Custo": ev["centro_custo_cod"],
            "Centro de Custo":      ev["centro_custo_nome"],
            # ── Campos para preenchimento ──────────────────────────
            "Conta Débito":         "",
            "Conta Crédito":        "",
            "Cód. Histórico":       "",
            "Histórico":            "",
            "Observação":           "",
        })

    df = pd.DataFrame(linhas)

    # ── Resumo ────────────────────────────────────────────────────
    total   = len(df)
    p_count = len(df[df["Tipo Rubrica"] == "Provento"])
    d_count = len(df[df["Tipo Rubrica"] == "Desconto"])
    i_count = len(df[df["Tipo Rubrica"] == "Informativa"])
    id_count= len(df[df["Tipo Rubrica"] == "Inf. Dedutora"])
    nf_count= len(df[df["Tipo Rubrica"].str.startswith("⚠️")])

    resumo_data = {
        "Tipo":       ["Provento", "Desconto", "Informativa", "Inf. Dedutora", "Não encontrado", "TOTAL"],
        "Quantidade": [p_count, d_count, i_count, id_count, nf_count, total],
    }
    df_resumo = pd.DataFrame(resumo_data)

    # ── Escrita no Excel ──────────────────────────────────────────
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Configuração", index=False)
        df_resumo.to_excel(writer, sheet_name="Resumo", index=False)

        # Formatação aba Configuração
        ws = writer.sheets["Configuração"]
        _formatar_planilha(ws, df)

        # Formatação aba Resumo
        ws_r = writer.sheets["Resumo"]
        _formatar_resumo(ws_r)

    output.seek(0)
    if sem_tipo:
        log.append(
            f"⚠️ {len(sem_tipo)} código(s) não encontrado(s) no rubricas.txt: "
            f"{', '.join(sorted(set(sem_tipo))[:20])}"
            f"{'...' if len(sem_tipo) > 20 else ''}"
        )
    log.append(f"Excel gerado: {total} linha(s) | "
               f"P={p_count} D={d_count} I={i_count} ID={id_count} NF={nf_count}")
    return output.read()


def _formatar_planilha(ws, df: pd.DataFrame):
    """Aplica larguras, cores de cabeçalho e destaque nas colunas a preencher."""
    from openpyxl.styles import (
        PatternFill, Font, Alignment, Border, Side, numbers
    )
    from openpyxl.utils import get_column_letter

    # Larguras das colunas
    larguras = {
        "A": 14,  # Código Evento
        "B": 38,  # Descrição PDF
        "C": 38,  # Descrição Rubricas
        "D": 16,  # Tipo Rubrica
        "E": 14,  # Tipo Folha Nº
        "F": 20,  # Tipo Folha
        "G": 18,  # Cód. CC
        "H": 22,  # Centro de Custo
        "I": 16,  # Conta Débito   ← PREENCHER
        "J": 16,  # Conta Crédito  ← PREENCHER
        "K": 14,  # Cód. Histórico ← PREENCHER
        "L": 42,  # Histórico      ← PREENCHER
        "M": 30,  # Observação     ← PREENCHER
    }
    for col, w in larguras.items():
        ws.column_dimensions[col].width = w

    # Cores
    COR_HEADER_INFO    = "444444"  # cinza escuro  → colunas de identificação
    COR_HEADER_FILL    = "FF8000"  # laranja TR     → colunas a preencher
    COR_PROVENTO       = "D4EDDA"  # verde claro
    COR_DESCONTO       = "F8D7DA"  # vermelho claro
    COR_INFORMATIVA    = "CCE5FF"  # azul claro
    COR_INF_DED        = "FFF3CD"  # amarelo claro
    COR_NAO_ENCONTRADO = "E2E3E5"  # cinza claro
    COR_FILL_PREENCHER = "FFF8F0"  # laranja muito claro → células a preencher

    borda = Border(
        left=Side(style="thin"),  right=Side(style="thin"),
        top=Side(style="thin"),   bottom=Side(style="thin"),
    )

    # Cabeçalho
    COLS_PREENCHER = {9, 10, 11, 12, 13}  # I, J, K, L, M (1-based)
    for col_idx, cell in enumerate(ws[1], start=1):
        if col_idx in COLS_PREENCHER:
            cell.fill      = PatternFill("solid", fgColor=COR_HEADER_FILL)
            cell.font      = Font(bold=True, color="FFFFFF", size=10)
        else:
            cell.fill      = PatternFill("solid", fgColor=COR_HEADER_INFO)
            cell.font      = Font(bold=True, color="FFFFFF", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = borda

    ws.row_dimensions[1].height = 32

    # Mapa tipo → cor de fundo
    TIPO_COR = {
        "Provento":      COR_PROVENTO,
        "Desconto":      COR_DESCONTO,
        "Informativa":   COR_INFORMATIVA,
        "Inf. Dedutora": COR_INF_DED,
    }

    # Dados
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        tipo_val = ws.cell(row=row_idx, column=4).value or ""
        cor_linha = TIPO_COR.get(tipo_val, COR_NAO_ENCONTRADO)

        for col_idx, cell in enumerate(row, start=1):
            cell.border    = borda
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if col_idx in COLS_PREENCHER:
                cell.fill = PatternFill("solid", fgColor=COR_FILL_PREENCHER)
                cell.font = Font(size=10)
            else:
                cell.fill = PatternFill("solid", fgColor=cor_linha)
                cell.font = Font(size=10)

        ws.row_dimensions[row_idx].height = 18

    # Congela cabeçalho
    ws.freeze_panes = "A2"

    # Filtro automático
    ws.auto_filter.ref = ws.dimensions


def _formatar_resumo(ws):
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

    borda = Border(
        left=Side(style="thin"),  right=Side(style="thin"),
        top=Side(style="thin"),   bottom=Side(style="thin"),
    )
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 14

    CORES = {
        "Provento":      "D4EDDA",
        "Desconto":      "F8D7DA",
        "Informativa":   "CCE5FF",
        "Inf. Dedutora": "FFF3CD",
        "Não encontrado":"E2E3E5",
        "TOTAL":         "FF8000",
    }

    for row in ws.iter_rows():
        for cell in row:
            cell.border    = borda
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font      = Font(size=10)

    # Cabeçalho
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor="444444")
        cell.font = Font(bold=True, color="FFFFFF", size=10)

    # Linhas de dados
    for row in ws.iter_rows(min_row=2):
        tipo = row[0].value or ""
        cor  = CORES.get(tipo, "FFFFFF")
        for cell in row:
            cell.fill = PatternFill("solid", fgColor=cor)
            if tipo == "TOTAL":
                cell.font = Font(bold=True, color="FFFFFF", size=11)


# ==============================
# INTERFACE STREAMLIT
# ==============================
def main():
    st.set_page_config(
        page_title="Domínio | Gerador de Excel de Configuração",
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
                📊 Gerador de Excel — Configuração Contábil de Rubricas
                &nbsp;|&nbsp; {VERSAO}
            </h2>
            <p style="color:#DDDDDD; margin:6px 0 0 0;
                      font-family:'Segoe UI',Arial,sans-serif;">
                Gera planilha Excel com todas as rubricas não configuradas,
                já com o tipo (Provento/Desconto/Informativa/Inf.Dedutora),
                pronta para preenchimento das contas contábeis.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.sidebar:
        st.markdown("### ℹ Sobre")
        st.markdown(f"**Versão:** {VERSAO}")
        st.markdown("**Thomson Reuters | Domínio Sistemas**")
        st.markdown("---")
        st.markdown("### 🎨 Legenda de cores")
        st.markdown("🟢 **Verde** → Provento")
        st.markdown("🔴 **Vermelho** → Desconto")
        st.markdown("🔵 **Azul** → Informativa")
        st.markdown("🟡 **Amarelo** → Inf. Dedutora")
        st.markdown("⚪ **Cinza** → Não encontrado")
        st.markdown("🟠 **Laranja** → Colunas a preencher")

    # ── Instruções ────────────────────────────────────────────────────
    with st.expander("📖 **Como usar** — clique para expandir", expanded=True):
        st.markdown("""
        **1.** Faça upload do **PDF de Itens Não Configurados** e do **rubricas.txt**

        **2.** Clique em **▶ Gerar Excel**

        **3.** Baixe o Excel e preencha as colunas destacadas em **laranja**:
        - **Conta Débito** — código da conta contábil a débito
        - **Conta Crédito** — código da conta contábil a crédito
        - **Cód. Histórico** — código do histórico padrão
        - **Histórico** — texto do histórico (ex: `<<Competência>> - <<Descrição>>`)
        - **Observação** — campo livre para anotações

        **4.** Use o Excel preenchido como referência para configurar no Domínio
        ou importe via a ferramenta de integração contábil.
        """)

    st.markdown("---")

    # ── Session state ─────────────────────────────────────────────────
    if "log"          not in st.session_state: st.session_state.log          = [f"Pronto. Versão {VERSAO}"]
    if "excel_gerado" not in st.session_state: st.session_state.excel_gerado = None
    if "nome_arquivo" not in st.session_state: st.session_state.nome_arquivo = "configuracao_rubricas.xlsx"
    if "df_preview"   not in st.session_state: st.session_state.df_preview   = None

    # ── Uploads ───────────────────────────────────────────────────────
    col1, col2 = st.columns(2)
    with col1:
        pdf_file = st.file_uploader(
            "1️⃣ PDF — Rubricas/Itens Não Configurados",
            type=["pdf"],
        )
    with col2:
        txt_file = st.file_uploader(
            "2️⃣ TXT — Rubricas (catálogo de tipos)",
            type=["txt"],
        )

    arquivos_ok = pdf_file is not None and txt_file is not None

    col_btn1, col_btn2 = st.columns([1, 1])
    with col_btn1:
        gerar = st.button(
            "▶ Gerar Excel",
            disabled=not arquivos_ok,
            use_container_width=True,
            type="primary",
        )
    with col_btn2:
        limpar = st.button("🗑 Limpar", use_container_width=True)

    if limpar:
        st.session_state.log          = ["Campos limpos."]
        st.session_state.excel_gerado = None
        st.session_state.df_preview   = None
        st.rerun()

    # ── Processamento ─────────────────────────────────────────────────
    if gerar and arquivos_ok:
        log = ["Iniciando processamento..."]

        with st.spinner("Lendo rubricas.txt..."):
            catalog = parse_rubricas_txt(txt_file.read(), log)

        with st.spinner("Lendo PDF de Itens Não Configurados..."):
            eventos = parse_nao_configurados_pdf(pdf_file.read(), log)

        if not eventos:
            log.append("AVISO: Nenhum evento encontrado no PDF.")
            st.session_state.log = log
            st.rerun()

        with st.spinner("Gerando Excel..."):
            excel_bytes = gerar_excel(eventos, catalog, log)

        st.session_state.excel_gerado = excel_bytes
        st.session_state.nome_arquivo = "configuracao_rubricas_dominio.xlsx"
        st.session_state.log          = log

        # Preview
        linhas_preview = []
        for ev in eventos:
            cod  = ev["cod"]
            info = catalog.get(cod, {})
            linhas_preview.append({
                "Código":        cod,
                "Descrição":     ev["descricao_pdf"],
                "Tipo":          info.get("tipo", "⚠️ Não encontrado"),
                "Tipo Folha":    ev["tipo_folha_desc"],
                "Centro Custo":  ev["centro_custo_nome"],
            })
        st.session_state.df_preview = pd.DataFrame(linhas_preview)
        st.rerun()

    # ── Resultado ─────────────────────────────────────────────────────
    if st.session_state.excel_gerado is not None:
        st.success("✅ Excel gerado com sucesso!")
        st.download_button(
            label="⬇ Baixar Excel — Configuração Contábil",
            data=st.session_state.excel_gerado,
            file_name=st.session_state.nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

        if st.session_state.df_preview is not None:
            df = st.session_state.df_preview
            total     = len(df)
            p_count   = len(df[df["Tipo"] == "Provento"])
            d_count   = len(df[df["Tipo"] == "Desconto"])
            i_count   = len(df[df["Tipo"] == "Informativa"])
            id_count  = len(df[df["Tipo"] == "Inf. Dedutora"])
            nf_count  = len(df[df["Tipo"].str.startswith("⚠️", na=False)])

            m1, m2, m3, m4, m5, m6 = st.columns(6)
            m1.metric("📋 Total",         total)
            m2.metric("🟢 Proventos",     p_count)
            m3.metric("🔴 Descontos",     d_count)
            m4.metric("🔵 Informativas",  i_count)
            m5.metric("🟡 Inf. Dedutora", id_count)
            m6.metric("⚪ Não encontrado",nf_count)

            st.markdown("**Prévia dos dados (primeiras 50 linhas):**")

            def highlight_tipo(row):
                t = str(row.get("Tipo", ""))
                if t == "Provento":      return ["background-color:#d4edda"] * len(row)
                if t == "Desconto":      return ["background-color:#f8d7da"] * len(row)
                if t == "Informativa":   return ["background-color:#cce5ff"] * len(row)
                if t == "Inf. Dedutora": return ["background-color:#fff3cd"] * len(row)
                return ["background-color:#e2e3e5"] * len(row)

            st.dataframe(
                df.head(50).style.apply(highlight_tipo, axis=1),
                use_container_width=True,
            )
            if len(df) > 50:
                st.caption(f"Mostrando 50 de {len(df)} linhas. Baixe o Excel para ver todas.")

    # ── Log ───────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("**Log de processamento**")
    log_texto = "\n".join(st.session_state.log)
    tem_erro  = any(str(l).upper().startswith("ERRO") for l in st.session_state.log)
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
