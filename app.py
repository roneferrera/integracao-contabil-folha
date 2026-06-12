# ============================================================
# app_integracao_dominio.py  –  Integração Contábil Domínio V3.1
# Entradas:
#   1. RubricasItens não Configurados.pdf  → eventos sem config contábil
#   2. rubricas.txt                        → catálogo de tipos de rubrica
#   3. Contas.xls (NOVO)                   → plano de contas para lookup
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

VERSAO = "V3.1"

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
        .conta-card {
            background: #FFF8F0;
            border: 1px solid #FF8000;
            border-radius: 6px;
            padding: 12px 16px;
            margin-bottom: 8px;
        }
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
# PARSE DO PLANO DE CONTAS (XLS/XLSX)
# ==============================
def parse_plano_contas(file_bytes: bytes, log: list) -> pd.DataFrame:
    """
    Lê o arquivo de Plano de Contas e retorna DataFrame normalizado com:
    codigo_conta, classificacao, nome_conta, tipo (S=sintética, A=analítica)
    """
    try:
        xls = pd.ExcelFile(BytesIO(file_bytes))
    except Exception as e:
        log.append(f"ERRO ao abrir Plano de Contas: {e}")
        return pd.DataFrame()

    sheet = xls.sheet_names[0]
    try:
        df_raw = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet, header=None, dtype=str)
    except Exception as e:
        log.append(f"ERRO ao ler Plano de Contas: {e}")
        return pd.DataFrame()

    # O arquivo tem estrutura peculiar - vamos extrair classificação e nome
    # Identificar linhas de dados reais (com classificação numérica)
    registros = []
    
    for _, row in df_raw.iterrows():
        row_vals = [str(v).strip() if pd.notna(v) and str(v).strip() not in ("nan", "None", "") else "" for v in row]
        
        # Buscar classificação (padrão X.X.X.XX.XXXXXX)
        classificacao = ""
        nome_conta = ""
        tipo_conta = ""
        
        for val in row_vals:
            if re.match(r'^\d+(\.\d+)+$', val) and not classificacao:
                classificacao = val
            elif val == "S" or val == "A":
                tipo_conta = val
        
        # Nome da conta: último valor não vazio que não seja classificação nem tipo
        vals_nao_vazios = [v for v in row_vals if v and v not in ("S", "A") 
                          and not re.match(r'^\d+(\.\d+)*\.?$', v)
                          and not re.match(r'^\d+\.\d+$', v)]
        
        if vals_nao_vazios:
            nome_conta = vals_nao_vazios[-1]
        
        if classificacao and nome_conta:
            registros.append({
                "classificacao": classificacao,
                "nome_conta": nome_conta,
                "tipo": tipo_conta,
                "nivel": len(classificacao.split(".")),
            })
    
    df_contas = pd.DataFrame(registros).drop_duplicates(subset=["classificacao"])
    
    # Filtrar apenas contas analíticas (nível mais profundo / sem S)
    # Contas analíticas são as que podem receber lançamentos
    df_analiticas = df_contas[df_contas["tipo"] != "S"].copy() if not df_contas.empty else df_contas
    
    log.append(f"Plano de Contas: {len(df_contas)} conta(s) total, {len(df_analiticas)} analítica(s).")
    return df_contas


def get_contas_para_select(df_contas: pd.DataFrame, filtro: str = "") -> list:
    """Retorna lista formatada 'CLASSIFICACAO - NOME' para selectbox."""
    if df_contas.empty:
        return ["(Sem plano de contas)"]
    
    df_filtrado = df_contas[df_contas["tipo"] != "S"].copy()
    
    if filtro:
        mask = (
            df_filtrado["classificacao"].str.contains(filtro, case=False, na=False) |
            df_filtrado["nome_conta"].str.contains(filtro, case=False, na=False)
        )
        df_filtrado = df_filtrado[mask]
    
    opcoes = [""] + [
        f"{row['classificacao']} - {row['nome_conta']}"
        for _, row in df_filtrado.iterrows()
    ]
    return opcoes


def extrair_classificacao(opcao_str: str) -> str:
    """Extrai só a classificação de uma string 'CLASSIF - NOME'."""
    if not opcao_str or " - " not in opcao_str:
        return opcao_str or ""
    return opcao_str.split(" - ")[0].strip()


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
# CLASSIFICAÇÃO DE GRUPOS POR CENTRO DE CUSTO (COM SEPARADOR)
# ==============================
GRUPOS_DESPESA_PADRAO = [
    "Custo Direto de Produção",
    "Custo Direto de Serviços",
    "Custo Indireto de Produção",
    "Despesa Administrativa",
    "Despesa com Vendas",
    "Despesa Financeira",
    "Despesa Não Operacional",
    "Outro",
]

# Mapeamento de grupo → prefixos de conta sugeridos no plano de contas
GRUPO_CONTA_PREFIXO = {
    "Custo Direto de Produção":   ["4.1.1", "4.1.2"],
    "Custo Direto de Serviços":   ["4.1.3"],
    "Custo Indireto de Produção": ["4.1.2"],
    "Despesa Administrativa":     ["4.2.2"],
    "Despesa com Vendas":         ["4.2.1"],
    "Despesa Financeira":         ["4.2.2.06"],
    "Despesa Não Operacional":    ["4.3"],
    "Outro":                      [],
}


def get_contas_por_grupo(df_contas: pd.DataFrame, grupo: str) -> list:
    """Filtra contas analíticas pelo prefixo do grupo selecionado."""
    if df_contas.empty:
        return [""]
    
    prefixos = GRUPO_CONTA_PREFIXO.get(grupo, [])
    df_anal = df_contas[df_contas["tipo"] != "S"].copy()
    
    if not prefixos:
        # Retorna todas as analíticas
        opcoes = [""] + [
            f"{r['classificacao']} - {r['nome_conta']}"
            for _, r in df_anal.iterrows()
        ]
        return opcoes
    
    mask = df_anal["classificacao"].apply(
        lambda c: any(c.startswith(p) for p in prefixos)
    )
    df_filtrado = df_anal[mask]
    
    if df_filtrado.empty:
        # fallback: todas
        df_filtrado = df_anal
    
    opcoes = [""] + [
        f"{r['classificacao']} - {r['nome_conta']}"
        for _, r in df_filtrado.iterrows()
    ]
    return opcoes


def get_centros_custo_unicos(eventos: list) -> list:
    """Retorna lista de (cod, nome) únicos dos centros de custo."""
    vistos = {}
    for ev in eventos:
        cod = ev["centro_custo_cod"]
        nome = ev["centro_custo_nome"]
        if cod and cod not in vistos:
            vistos[cod] = nome
    return [(cod, nome) for cod, nome in vistos.items()]


# ==============================
# ETAPA 1 — GERA EXCEL PARA PREENCHIMENTO
# ==============================
def gerar_excel_configuracao(
    eventos: list,
    catalog: dict,
    cod_empresa: str,
    log: list,
    usa_separador: bool = False,
    config_cc: dict = None,   # {cc_cod: {"grupo": str, "conta_debito": str, "conta_credito": str, "historico": str}}
    df_contas: pd.DataFrame = None,
) -> bytes:
    """
    config_cc: dict com configurações por centro de custo quando usa_separador=True
    """
    linhas = []
    for ev in eventos:
        cod = ev["cod"]
        info = catalog.get(cod, {})
        tipo = info.get("tipo", "⚠️ Não encontrado")
        desc_rubrica = info.get("descricao", ev["descricao_pdf"])
        cc_cod = ev["centro_custo_cod"]
        
        # Preencher contas automaticamente se usa_separador e há config para este CC
        conta_deb_auto = ""
        conta_cred_auto = ""
        historico_auto = ""
        grupo_auto = ""
        
        if usa_separador and config_cc and cc_cod in config_cc:
            cfg = config_cc[cc_cod]
            conta_deb_auto = cfg.get("conta_debito", "")
            conta_cred_auto = cfg.get("conta_credito", "")
            historico_auto = cfg.get("historico", "")
            grupo_auto = cfg.get("grupo", "")
        
        linhas.append({
            # Identificação (somente leitura)
            "Cód. Empresa":          cod_empresa,
            "Cód. Evento":           cod,
            "Descrição (PDF)":       ev["descricao_pdf"],
            "Descrição (Rubricas)":  desc_rubrica,
            "Tipo Rubrica":          tipo,
            "Tipo Folha (Nº)":       ev["tipo_folha"],
            "Tipo Folha":            ev["tipo_folha_desc"],
            "Cód. Centro de Custo":  cc_cod,
            "Centro de Custo":       ev["centro_custo_nome"],
            "Grupo de Despesa":      grupo_auto,
            "Usa Separador":         "Sim" if usa_separador else "Não",
            # Campos a preencher (pré-preenchidos se usa_separador)
            "Conta Débito":          conta_deb_auto,
            "Conta Crédito":         conta_cred_auto,
            "Cód. Histórico":        "",
            "Histórico":             historico_auto,
            "Observação":            "",
        })

    df = pd.DataFrame(linhas)
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Configuração", index=False)
        ws = writer.sheets["Configuração"]
        _formatar_planilha_config(ws, df)
        
        # Se há plano de contas, exportar como aba de referência
        if df_contas is not None and not df_contas.empty:
            df_contas_exp = df_contas[["classificacao", "nome_conta", "tipo"]].copy()
            df_contas_exp.columns = ["Classificação", "Nome da Conta", "Tipo (S/A)"]
            df_contas_exp.to_excel(writer, sheet_name="Plano de Contas", index=False)
            ws_pc = writer.sheets["Plano de Contas"]
            _formatar_planilha_saida(ws_pc)

    output.seek(0)
    log.append(f"Excel de configuração gerado: {len(linhas)} linha(s). Separador: {'Sim' if usa_separador else 'Não'}.")
    return output.read()


def _formatar_planilha_config(ws, df: pd.DataFrame):
    borda = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    larguras = {
        "A": 12, "B": 12, "C": 38, "D": 38, "E": 16,
        "F": 14, "G": 20, "H": 18, "I": 22, "J": 22,
        "K": 14, "L": 16, "M": 16, "N": 14, "O": 42, "P": 30,
    }
    for col, w in larguras.items():
        ws.column_dimensions[col].width = w

    COR_HEADER_INFO   = "444444"
    COR_HEADER_FILL   = "FF8000"
    COR_FILL_PREENCHER = "FFF8F0"
    # Colunas a preencher: L(12), M(13), N(14), O(15), P(16)
    COLS_PREENCHER = {12, 13, 14, 15, 16}
    # Coluna grupo e separador: J(10), K(11) - informativas
    COLS_INFO_EXTRA = {10, 11}

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
        elif col_idx in COLS_INFO_EXTRA:
            cell.fill = PatternFill("solid", fgColor="6C757D")
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
        elif "usa separador" in cl:
            col_map["usa_separador"] = col

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
        usa_sep  = _limpa(row.get(col_map.get("usa_separador", ""), ""))

        if not seq:
            continue

        # Determinar separador
        separador_val = "1" if usa_sep.lower() == "sim" else "0"

        if debito or credito:
            com_conta += 1
        else:
            sem_conta += 1

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

        linhas_integra.append({
            "Código da Empresa": empresa,
            "Separador": separador_val,
            "Código Sequencial da Integração": seq,
            "Tipo da Integração (1 - Folha mensal; 2 - Empresa; 3 - Férias; 4 - Rescisao; 5 - Prov. Férias; 6 - Prov. 13)": tipo,
            "Código da Rúbrica Selecionada": seq,
        })

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

    log.append(f"Arquivos gerados → Com conta: {com_conta} | Sem conta: {sem_conta}")

    buf_evento = BytesIO()
    with pd.ExcelWriter(buf_evento, engine="openpyxl") as writer:
        pd.DataFrame(linhas_integra).to_excel(writer, sheet_name="integra", index=False)
        pd.DataFrame(linhas_evento).to_excel(writer, sheet_name="evento", index=False)
        for sheet_name in ["integra", "evento"]:
            ws = writer.sheets[sheet_name]
            _formatar_planilha_saida(ws)
    buf_evento.seek(0)

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
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor="444444")
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = borda
    ws.row_dimensions[1].height = 32

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = borda
            cell.alignment = Alignment(vertical="center")
            cell.font = Font(size=10)

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
    defaults = {
        "log": [f"Pronto. Versão {VERSAO}"],
        "excel_config": None,
        "evento_xlsx": None,
        "integra_xls": None,
        "df_preview": None,
        "n_eventos": 0,
        "df_contas": None,
        "eventos_parsed": None,
        "config_cc": {},
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

    # ══════════════════════════════════════════════════════════════════════
    # ETAPA 1
    # ══════════════════════════════════════════════════════════════════════
    st.markdown("## 📋 Etapa 1 — Gerar Excel para Preenchimento")

    col1, col2, col3 = st.columns(3)
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
    with col3:
        contas_file = st.file_uploader(
            "3️⃣ XLS/XLSX — Plano de Contas (opcional)",
            type=["xls", "xlsx"],
            key="contas_etapa1",
        )

    # Carregar plano de contas ao fazer upload
    if contas_file is not None and st.session_state.df_contas is None:
        log_temp = []
        df_c = parse_plano_contas(contas_file.read(), log_temp)
        if not df_c.empty:
            st.session_state.df_contas = df_c
            st.session_state.log.extend(log_temp)
            st.rerun()

    if contas_file is None:
        st.session_state.df_contas = None

    # Mostrar info do plano de contas carregado
    if st.session_state.df_contas is not None:
        df_pc = st.session_state.df_contas
        n_analiticas = len(df_pc[df_pc["tipo"] != "S"])
        st.success(f"✅ Plano de Contas carregado: **{len(df_pc)}** contas ({n_analiticas} analíticas)")

    st.markdown("---")

    # ── Configuração de Separador ─────────────────────────────────────────
    st.markdown("### ⚙️ Configuração de Separador")

    col_sep1, col_sep2 = st.columns([1, 2])
    with col_sep1:
        usa_separador = st.radio(
            "Os lançamentos usam separador por Centro de Custo?",
            options=["Não", "Sim"],
            index=0,
            horizontal=True,
            help="Com separador: cada Centro de Custo recebe um grupo de despesa e contas específicas.",
        )
    usa_sep_bool = (usa_separador == "Sim")

    # ── Configuração por CC quando separador ativo ────────────────────────
    if usa_sep_bool:
        # Precisa ter eventos parseados para mostrar os CCs
        if st.session_state.eventos_parsed:
            ccs = get_centros_custo_unicos(st.session_state.eventos_parsed)
            
            if ccs:
                st.markdown("#### 🏢 Classificação por Centro de Custo")
                st.info(
                    "Para cada Centro de Custo, selecione o **Grupo de Despesa** e as **contas contábeis** "
                    "que serão aplicadas aos eventos desse CC."
                )

                df_contas_atual = st.session_state.df_contas

                for cc_cod, cc_nome in ccs:
                    with st.expander(f"🏢 CC {cc_cod} — {cc_nome}", expanded=True):
                        col_g, col_d, col_c, col_h = st.columns([2, 3, 3, 2])
                        
                        cfg_atual = st.session_state.config_cc.get(cc_cod, {})
                        
                        with col_g:
                            grupo_sel = st.selectbox(
                                "Grupo de Despesa",
                                options=GRUPOS_DESPESA_PADRAO,
                                index=GRUPOS_DESPESA_PADRAO.index(cfg_atual.get("grupo", "Outro"))
                                      if cfg_atual.get("grupo") in GRUPOS_DESPESA_PADRAO else 0,
                                key=f"grupo_{cc_cod}",
                            )
                        
                        # Filtrar contas pelo grupo selecionado
                        if df_contas_atual is not None and not df_contas_atual.empty:
                            opcoes_contas = get_contas_por_grupo(df_contas_atual, grupo_sel)
                        else:
                            opcoes_contas = ["(Carregue o Plano de Contas para sugestões)"]
                        
                        with col_d:
                            deb_atual = cfg_atual.get("conta_debito", "")
                            # Tentar encontrar índice atual
                            idx_deb = 0
                            for i, op in enumerate(opcoes_contas):
                                if deb_atual and (deb_atual in op or op.startswith(deb_atual)):
                                    idx_deb = i
                                    break
                            conta_deb_sel = st.selectbox(
                                "Conta Débito",
                                options=opcoes_contas,
                                index=idx_deb,
                                key=f"debito_{cc_cod}",
                            )
                        
                        with col_c:
                            cred_atual = cfg_atual.get("conta_credito", "")
                            idx_cred = 0
                            for i, op in enumerate(opcoes_contas):
                                if cred_atual and (cred_atual in op or op.startswith(cred_atual)):
                                    idx_cred = i
                                    break
                            conta_cred_sel = st.selectbox(
                                "Conta Crédito",
                                options=opcoes_contas,
                                index=idx_cred,
                                key=f"credito_{cc_cod}",
                            )
                        
                        with col_h:
                            hist_sel = st.text_input(
                                "Histórico",
                                value=cfg_atual.get("historico", ""),
                                key=f"hist_{cc_cod}",
                                placeholder="Ex: 001",
                            )
                        
                        # Salvar config no session_state
                        st.session_state.config_cc[cc_cod] = {
                            "grupo": grupo_sel,
                            "conta_debito": extrair_classificacao(conta_deb_sel),
                            "conta_credito": extrair_classificacao(conta_cred_sel),
                            "historico": hist_sel,
                        }
                        
                        # Preview das contas selecionadas
                        deb_exib = extrair_classificacao(conta_deb_sel)
                        cred_exib = extrair_classificacao(conta_cred_sel)
                        if deb_exib or cred_exib:
                            st.markdown(
                                f"<small>✅ <b>Débito:</b> {deb_exib or '—'} &nbsp;|&nbsp; "
                                f"<b>Crédito:</b> {cred_exib or '—'} &nbsp;|&nbsp; "
                                f"<b>Grupo:</b> {grupo_sel}</small>",
                                unsafe_allow_html=True,
                            )
            else:
                st.warning("Nenhum Centro de Custo encontrado. Processe o PDF primeiro.")
        else:
            st.info("⬆️ Faça upload do PDF e clique em **▶ Processar PDF** para configurar os Centros de Custo.")
            
            # Botão para processar só o PDF (sem gerar Excel ainda)
            if pdf_file and txt_file:
                if st.button("🔍 Processar PDF (pré-visualizar CCs)", use_container_width=False):
                    log = list(st.session_state.log)
                    with st.spinner("Lendo arquivos..."):
                        catalog = parse_rubricas_txt(txt_file.read(), log)
                        eventos = parse_nao_configurados_pdf(pdf_file.read(), log)
                    st.session_state.eventos_parsed = eventos
                    st.session_state.log = log
                    st.rerun()

    st.markdown("---")

    # ── Botões principais ─────────────────────────────────────────────────
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
            for k in ["log", "excel_config", "evento_xlsx", "integra_xls",
                      "df_preview", "n_eventos", "df_contas", "eventos_parsed", "config_cc"]:
                if k == "log":
                    st.session_state[k] = ["Campos limpos."]
                elif k == "n_eventos":
                    st.session_state[k] = 0
                elif k == "config_cc":
                    st.session_state[k] = {}
                else:
                    st.session_state[k] = None
            st.rerun()

    if gerar_excel and pdf_file and txt_file:
        log = ["[Etapa 1] Iniciando..."]
        with st.spinner("Lendo rubricas.txt..."):
            catalog = parse_rubricas_txt(txt_file.read(), log)
        with st.spinner("Lendo PDF de Itens Não Configurados..."):
            eventos = parse_nao_configurados_pdf(pdf_file.read(), log)
        
        st.session_state.eventos_parsed = eventos
        
        if not eventos:
            log.append("AVISO: Nenhum evento encontrado no PDF.")
        else:
            with st.spinner("Gerando Excel..."):
                excel_bytes = gerar_excel_configuracao(
                    eventos,
                    catalog,
                    cod_empresa,
                    log,
                    usa_separador=usa_sep_bool,
                    config_cc=st.session_state.config_cc if usa_sep_bool else None,
                    df_contas=st.session_state.df_contas,
                )
            st.session_state.excel_config = excel_bytes
            st.session_state.n_eventos = len(eventos)

            linhas_prev = []
            for ev in eventos:
                cod_ev = ev["cod"]
                info = catalog.get(cod_ev, {})
                cc_cod = ev["centro_custo_cod"]
                cfg_cc = st.session_state.config_cc.get(cc_cod, {}) if usa_sep_bool else {}
                linhas_prev.append({
                    "Código": cod_ev,
                    "Descrição": ev["descricao_pdf"],
                    "Tipo": info.get("tipo", "⚠️"),
                    "Tipo Folha": ev["tipo_folha_desc"],
                    "Centro Custo": ev["centro_custo_nome"],
                    "Grupo": cfg_cc.get("grupo", "—"),
                    "Conta Débito": cfg_cc.get("conta_debito", ""),
                    "Conta Crédito": cfg_cc.get("conta_credito", ""),
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
            p  = len(df[df["Tipo"] == "Provento"])
            d  = len(df[df["Tipo"] == "Desconto"])
            i  = len(df[df["Tipo"] == "Informativa"])
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
        "4️⃣ Excel Preenchido (.xlsx)",
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
        log = list(st.session_state.log) or []
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
            - Aba `integra`: Empresa | Separador | Seq | Tipo | Rúbrica
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
