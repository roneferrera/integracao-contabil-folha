# ============================================================
# app_integracao_dominio.py  –  Integração Contábil Domínio V4.0
# ============================================================

import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

VERSAO = "V4.0"

# ==============================
# TEMA
# ==============================
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
        div[data-testid="stExpander"] { border: 1px solid #FF8000; border-radius: 6px; }
        </style>
    """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════
# CLASSIFICADOR SEMÂNTICO UNIVERSAL
# Funciona com QUALQUER plano de contas — usa palavras-chave nos nomes
# ══════════════════════════════════════════════════════════════════════════

# Grupos de DÉBITO: contas de despesa/custo
# Cada grupo tem palavras-chave que devem aparecer no nome da conta
GRUPOS_DEBITO_KEYWORDS: dict[str, list[str]] = {
    "Custo Direto de Produção": [
        "mão-de-obra direta", "mao-de-obra direta", "material aplicado",
        "matéria-prima", "materia-prima", "custo direto produção",
        "custo direto de produção", "mão de obra direta produção",
        "salário produção", "inss produção", "fgts produção",
        "férias produção", "13 produção",
    ],
    "Custo Direto de Serviços": [
        "custo direto serviço", "custo direto de serviço",
        "mão-de-obra direta serviço", "mao-de-obra direta serviço",
        "custo da produção de serviço", "custo serviço prestado",
        "salário serviço", "inss serviço", "fgts serviço",
        "férias serviço", "13 serviço",
    ],
    "Custo Indireto de Produção": [
        "custo indireto", "mão-de-obra indireta", "mao-de-obra indireta",
        "material consumo indireto", "utilidade serviço",
        "depreciação custo", "combustível custo",
        "aluguel custo", "energia custo",
    ],
    "Despesa Administrativa": [
        "despesa administrativa", "despesas administrativas",
        "despesa com pessoal admin", "aluguel admin",
        "energia elétrica admin", "telefone admin",
        "material escritório", "material de escritório",
        "serviço contabilidade", "honorário", "serviços tomados",
        "serviços prestados por terceiros", "impostos taxas contribuições",
        "depreciação admin", "amortização admin",
        "salário admin", "inss admin", "fgts admin",
        "férias admin", "13 admin", "pro-labore",
        "pró-labore", "serviços de contabilidade",
        "despesas gerais", "despesa geral",
    ],
    "Despesa com Vendas": [
        "despesa com venda", "despesas com vendas",
        "comissão", "propaganda", "publicidade",
        "frete entrega", "despesa entrega",
        "viagem representação", "perdas recebimento",
        "salário venda", "inss venda", "fgts venda",
        "despesa pessoal venda",
    ],
    "Despesa Financeira": [
        "despesa financeira", "despesas financeiras",
        "juro passivo", "juros passivos",
        "variação monetária passiva", "variação cambial passiva",
        "desconto financeiro concedido", "juro mora",
        "juro empréstimo", "tarifa bancária", "iof",
        "multa passiva", "perda aplicação",
    ],
    "Despesa Não Operacional": [
        "despesa não operacional", "despesas não operacionais",
        "resultado negativo alien", "perda alienação",
        "resultado negativo sinistro", "baixa ativo",
        "provisão irpj", "provisão csll",
        "imposto de renda", "contribuição social",
        "provisão ir", "provisão cs",
    ],
}

# Grupos de CRÉDITO: contas de passivo/obrigação
GRUPOS_CREDITO_KEYWORDS: dict[str, list[str]] = {
    "Custo Direto de Produção": [
        "salário", "salarios", "obrigação pessoal",
        "obrigações com o pessoal", "inss a recolher",
        "fgts a recolher", "provisão férias", "provisão 13",
        "férias a pagar", "13 a pagar", "rescisão",
        "obrigação trabalhista", "obrigações trabalhistas",
    ],
    "Custo Direto de Serviços": [
        "salário", "salarios", "obrigação pessoal",
        "obrigações com o pessoal", "inss a recolher",
        "fgts a recolher", "provisão férias", "provisão 13",
        "férias a pagar", "13 a pagar",
        "obrigação trabalhista", "obrigações trabalhistas",
    ],
    "Custo Indireto de Produção": [
        "salário", "salarios", "obrigação pessoal",
        "obrigações com o pessoal", "inss a recolher",
        "fgts a recolher", "provisão férias", "provisão 13",
        "obrigação trabalhista", "obrigações trabalhistas",
    ],
    "Despesa Administrativa": [
        "salário", "salarios", "obrigação pessoal",
        "obrigações com o pessoal", "inss a recolher",
        "fgts a recolher", "provisão férias", "provisão 13",
        "obrigação trabalhista", "obrigações trabalhistas",
        "impostos contribuições a recolher",
        "contas a pagar", "fornecedores",
    ],
    "Despesa com Vendas": [
        "salário", "salarios", "obrigação pessoal",
        "obrigações com o pessoal", "inss a recolher",
        "fgts a recolher", "provisão férias", "provisão 13",
        "obrigação trabalhista", "obrigações trabalhistas",
        "contas a pagar", "fornecedores",
    ],
    "Despesa Financeira": [
        "contas a pagar", "outras obrigações",
        "empréstimo", "financiamento",
        "impostos contribuições a recolher",
    ],
    "Despesa Não Operacional": [
        "contas a pagar", "outras obrigações",
        "impostos contribuições a recolher",
        "provisão imposto", "obrigação tributária",
    ],
}

GRUPOS_LISTA = list(GRUPOS_DEBITO_KEYWORDS.keys()) + ["Outro"]


def _normalizar(texto: str) -> str:
    """Normaliza para comparação: minúsculas, sem acentos extras."""
    return (
        texto.lower()
        .replace("ã", "a").replace("á", "a").replace("â", "a").replace("à", "a")
        .replace("é", "e").replace("ê", "e").replace("è", "e")
        .replace("í", "i").replace("ï", "i")
        .replace("ó", "o").replace("ô", "o").replace("õ", "o")
        .replace("ú", "u").replace("ü", "u")
        .replace("ç", "c")
        .strip()
    )


def _conta_match(nome_conta: str, keywords: list[str]) -> bool:
    """Verifica se o nome da conta contém alguma das palavras-chave."""
    nome_norm = _normalizar(nome_conta)
    for kw in keywords:
        kw_norm = _normalizar(kw)
        if kw_norm in nome_norm:
            return True
    return False


def classificar_contas_automatico(
    df_contas: pd.DataFrame,
    grupo: str,
) -> tuple[list[str], list[str]]:
    """
    Retorna (opcoes_debito, opcoes_credito) filtradas por keywords do grupo.
    Cada item: 'CLASSIFICACAO - NOME'
    """
    df_a = df_contas[df_contas["tipo"] == "A"].copy()
    if df_a.empty:
        return [""], [""]

    kw_deb  = GRUPOS_DEBITO_KEYWORDS.get(grupo, [])
    kw_cred = GRUPOS_CREDITO_KEYWORDS.get(grupo, [])

    if kw_deb and grupo != "Outro":
        mask_deb = df_a["nome_conta"].apply(lambda n: _conta_match(n, kw_deb))
        df_deb = df_a[mask_deb]
    else:
        df_deb = df_a   # "Outro" → todas

    if kw_cred and grupo != "Outro":
        mask_cred = df_a["nome_conta"].apply(lambda n: _conta_match(n, kw_cred))
        df_cred = df_a[mask_cred]
    else:
        df_cred = df_a

    def _fmt(df):
        return [""] + [f"{r['classificacao']} - {r['nome_conta']}" for _, r in df.iterrows()]

    return _fmt(df_deb), _fmt(df_cred)


def sugerir_grupo_automatico(df_contas: pd.DataFrame, grupo: str) -> dict:
    """
    Retorna {conta_debito_sugerida, conta_credito_sugerida, n_deb, n_cred}
    Pega a PRIMEIRA conta que bater para cada lado.
    """
    ops_deb, ops_cred = classificar_contas_automatico(df_contas, grupo)
    deb  = ops_deb[1]  if len(ops_deb)  > 1 else ""
    cred = ops_cred[1] if len(ops_cred) > 1 else ""
    return {
        "conta_debito":  extrair_codigo(deb),
        "conta_credito": extrair_codigo(cred),
        "n_deb":  len(ops_deb)  - 1,
        "n_cred": len(ops_cred) - 1,
        "ops_deb":  ops_deb,
        "ops_cred": ops_cred,
    }


def extrair_codigo(opcao: str) -> str:
    if not opcao or " - " not in opcao:
        return opcao or ""
    return opcao.split(" - ")[0].strip()


def _idx(opcoes: list, valor: str) -> int:
    if not valor:
        return 0
    for i, op in enumerate(opcoes):
        if op.startswith(valor):
            return i
    return 0


# ==============================
# PARSE DO PLANO DE CONTAS
# Estrutura exportação Domínio:
#   col[0] = Empresa  col[1] = Reduzido
#   col[2] = Classificação  col[3] = Tipo (S/A)  col[4] = Descrição
# ==============================
def parse_plano_contas(file_bytes: bytes, log: list) -> pd.DataFrame:
    try:
        df_raw = pd.read_excel(BytesIO(file_bytes), sheet_name=0, header=0, dtype=str)
    except Exception as e:
        log.append(f"ERRO ao abrir Plano de Contas: {e}")
        return pd.DataFrame()

    if len(df_raw.columns) < 5:
        log.append("ERRO: Plano de Contas com menos de 5 colunas.")
        return pd.DataFrame()

    registros = []
    for _, row in df_raw.iterrows():
        classif = str(row.iloc[2]).strip()
        tipo    = str(row.iloc[3]).strip().upper()
        nome    = str(row.iloc[4]).strip()

        if not re.match(r'^\d+$', classif):
            continue
        if tipo not in ("S", "A"):
            continue
        if not nome or nome.lower() in ("nan", "none", ""):
            continue

        registros.append({
            "classificacao": classif,
            "nome_conta":    nome,
            "tipo":          tipo,
        })

    df_contas = (
        pd.DataFrame(registros)
        .drop_duplicates(subset=["classificacao"])
        .reset_index(drop=True)
    )

    n_a = len(df_contas[df_contas["tipo"] == "A"])
    n_s = len(df_contas[df_contas["tipo"] == "S"])
    log.append(f"Plano de Contas: {len(df_contas)} contas ({n_a} analíticas · {n_s} sintéticas).")
    return df_contas


# ==============================
# PARSE TXT RUBRICAS
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
        cod      = partes[2].strip()
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
# PARSE PDF
# ==============================
IGNORE_PATTERNS = [
    r"^RELAÇÃO DE RUBRICAS", r"^Página", r"^Emissão",
    r"^Hora:", r"^Empresa:", r"^Código\s+Descrição", r"^\s*$",
]
SECAO_TIPO_FOLHA = {
    "Folha Normal": "1", "Empresa": "2", "Férias": "3",
    "Rescisão": "4", "Provisão de Férias": "5",
    "Provisão de 13º": "6", "Provisão de 13o": "6",
}
SECAO_TIPO_FOLHA_DESC = {
    "1": "Folha Normal", "2": "Empresa", "3": "Férias",
    "4": "Rescisão", "5": "Provisão de Férias", "6": "Provisão de 13º",
}
RE_SECAO = re.compile(
    r"^(Folha Normal|Empresa|Férias|Rescisão|"
    r"Provisão de Férias|Provisão de 13º|Provisão de 13o)$", re.IGNORECASE,
)
RE_CC    = re.compile(r"^Centro de Custo:\s*(\d+)\s+(.+)$", re.IGNORECASE)
RE_EVENT = re.compile(r"^\s*(\d+)\s+(.+)$")


def should_ignore(line: str) -> bool:
    return any(re.search(p, line, re.IGNORECASE) for p in IGNORE_PATTERNS)


def parse_nao_configurados_pdf(file_bytes: bytes, log: list) -> list:
    eventos = []
    vistos = set()
    tipo_folha_atual = "1"
    cc_cod_atual = cc_nome_atual = ""

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for raw in text.splitlines():
                line = raw.strip()
                if not line:
                    continue
                m = RE_SECAO.match(line)
                if m:
                    sec = m.group(1).strip()
                    for k, v in SECAO_TIPO_FOLHA.items():
                        if k.lower() in sec.lower():
                            tipo_folha_atual = v
                            break
                    continue
                m = RE_CC.match(line)
                if m:
                    cc_cod_atual  = m.group(1).strip()
                    cc_nome_atual = m.group(2).strip()
                    continue
                if should_ignore(line):
                    continue
                m = RE_EVENT.match(line)
                if m:
                    cod  = m.group(1).strip()
                    desc = m.group(2).strip()
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


def get_centros_custo_unicos(eventos: list) -> list:
    vistos: dict[str, str] = {}
    for ev in eventos:
        cod  = ev["centro_custo_cod"]
        nome = ev["centro_custo_nome"]
        if cod and cod not in vistos:
            vistos[cod] = nome
    return list(vistos.items())


# ==============================
# ETAPA 1 — GERA EXCEL
# ==============================
def gerar_excel_configuracao(
    eventos:       list,
    catalog:       dict,
    cod_empresa:   str,
    log:           list,
    usa_separador: bool = False,
    config_cc:     dict | None = None,
    df_contas:     pd.DataFrame | None = None,
) -> bytes:
    linhas = []
    for ev in eventos:
        cod       = ev["cod"]
        info      = catalog.get(cod, {})
        tipo      = info.get("tipo", "⚠️ Não encontrado")
        desc_rubr = info.get("descricao", ev["descricao_pdf"])
        cc_cod    = ev["centro_custo_cod"]

        conta_deb = conta_cred = historico = grupo = ""
        if usa_separador and config_cc and cc_cod in config_cc:
            cfg        = config_cc[cc_cod]
            conta_deb  = cfg.get("conta_debito",  "")
            conta_cred = cfg.get("conta_credito", "")
            historico  = cfg.get("historico",     "")
            grupo      = cfg.get("grupo",         "")

        linhas.append({
            "Cód. Empresa":         cod_empresa,
            "Cód. Evento":          cod,
            "Descrição (PDF)":      ev["descricao_pdf"],
            "Descrição (Rubricas)": desc_rubr,
            "Tipo Rubrica":         tipo,
            "Tipo Folha (Nº)":      ev["tipo_folha"],
            "Tipo Folha":           ev["tipo_folha_desc"],
            "Cód. Centro de Custo": cc_cod,
            "Centro de Custo":      ev["centro_custo_nome"],
            "Grupo de Despesa":     grupo,
            "Usa Separador":        "Sim" if usa_separador else "Não",
            "Conta Débito":         conta_deb,
            "Conta Crédito":        conta_cred,
            "Cód. Histórico":       "",
            "Histórico":            historico,
            "Observação":           "",
        })

    df = pd.DataFrame(linhas)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Configuração", index=False)
        _formatar_planilha_config(writer.sheets["Configuração"], df)

        if df_contas is not None and not df_contas.empty:
            df_exp = df_contas[["classificacao", "nome_conta", "tipo"]].copy()
            df_exp.columns = ["Classificação", "Nome da Conta", "Tipo (S/A)"]
            df_exp.to_excel(writer, sheet_name="Plano de Contas", index=False)
            _formatar_planilha_saida(writer.sheets["Plano de Contas"])

    output.seek(0)
    log.append(f"Excel gerado: {len(linhas)} linha(s). Separador: {'Sim' if usa_separador else 'Não'}.")
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

    COLS_PREENCHER  = {12, 13, 14, 15, 16}
    COLS_INFO_EXTRA = {10, 11}
    TIPO_COR = {
        "Provento": "D4EDDA", "Desconto": "F8D7DA",
        "Informativa": "CCE5FF", "Inf. Dedutora": "FFF3CD",
    }

    for col_idx, cell in enumerate(ws[1], start=1):
        if col_idx in COLS_PREENCHER:
            cell.fill = PatternFill("solid", fgColor="FF8000")
            cell.font = Font(bold=True, color="FFFFFF", size=10)
        elif col_idx in COLS_INFO_EXTRA:
            cell.fill = PatternFill("solid", fgColor="6C757D")
            cell.font = Font(bold=True, color="FFFFFF", size=10)
        else:
            cell.fill = PatternFill("solid", fgColor="444444")
            cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = borda
    ws.row_dimensions[1].height = 32

    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        tipo_val  = ws.cell(row=row_idx, column=5).value or ""
        cor_linha = TIPO_COR.get(tipo_val, "E2E3E5")
        for col_idx, cell in enumerate(row, start=1):
            cell.border = borda
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if col_idx in COLS_PREENCHER:
                cell.fill = PatternFill("solid", fgColor="FFF8F0")
                cell.font = Font(size=10)
            else:
                cell.fill = PatternFill("solid", fgColor=cor_linha)
                cell.font = Font(size=10)
        ws.row_dimensions[row_idx].height = 18

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


# ==============================
# ETAPA 2 — GERA ARQUIVOS FINAIS
# ==============================
def ler_excel_preenchido(file_bytes: bytes, log: list) -> pd.DataFrame | None:
    try:
        xls = pd.ExcelFile(BytesIO(file_bytes))
    except Exception as e:
        log.append(f"ERRO ao abrir Excel preenchido: {e}")
        return None
    sheet = None
    for c in ["Configuração", "configuracao", "Plan1", "Sheet1"]:
        if c in xls.sheet_names:
            sheet = c
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
    log.append(f"Excel preenchido: {len(df)} linha(s) na aba '{sheet}'.")
    return df


def _limpa(val) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s.lower() in ("nan", "none", "") else s


def gerar_arquivos_finais(df: pd.DataFrame, cod_empresa_padrao: str, log: list) -> tuple[bytes, bytes]:
    col_map: dict[str, str] = {}
    for col in df.columns:
        cl = col.lower()
        if "cód. empresa"    in cl or "cod. empresa"    in cl: col_map["empresa"]        = col
        elif "cód. evento"   in cl or "cod. evento"     in cl: col_map["seq"]            = col
        elif "tipo folha (nº)" in cl or "tipo folha (n" in cl: col_map["tipo"]           = col
        elif "descrição (rubricas)" in cl:                      col_map["desc"]           = col
        elif "descrição (pdf)" in cl and "desc" not in col_map: col_map["desc"]          = col
        elif "cód. centro de custo" in cl:                      col_map["cc"]             = col
        elif "conta débito"  in cl or "conta debito"   in cl:  col_map["debito"]         = col
        elif "conta crédito" in cl or "conta credito"  in cl:  col_map["credito"]        = col
        elif "cód. histórico" in cl or "cod. historico" in cl: col_map["historico"]      = col
        elif "histórico" in cl and "cód" not in cl and "cod" not in cl:
            col_map["historico_texto"] = col
        elif "observação" in cl:                                col_map["observacao"]     = col
        elif "usa separador" in cl:                             col_map["usa_separador"]  = col

    linhas_evento, linhas_integra, linhas_integra_xls = [], [], []
    sem_conta = com_conta = 0

    for _, row in df.iterrows():
        empresa     = _limpa(row.get(col_map.get("empresa",         ""), "")) or cod_empresa_padrao
        seq         = _limpa(row.get(col_map.get("seq",             ""), ""))
        tipo        = _limpa(row.get(col_map.get("tipo",            ""), ""))
        desc        = _limpa(row.get(col_map.get("desc",            ""), ""))
        cc          = _limpa(row.get(col_map.get("cc",              ""), ""))
        debito      = _limpa(row.get(col_map.get("debito",          ""), ""))
        credito     = _limpa(row.get(col_map.get("credito",         ""), ""))
        historico   = _limpa(row.get(col_map.get("historico",       ""), ""))
        complemento = _limpa(row.get(col_map.get("historico_texto", ""), ""))
        usa_sep     = _limpa(row.get(col_map.get("usa_separador",   ""), ""))

        if not seq:
            continue

        sep_val = "1" if usa_sep.lower() == "sim" else "0"
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
            "Separador": sep_val,
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

    log.append(f"Arquivos → Com conta: {com_conta} | Sem conta: {sem_conta}")

    buf_evento = BytesIO()
    with pd.ExcelWriter(buf_evento, engine="openpyxl") as writer:
        pd.DataFrame(linhas_integra).to_excel(writer, sheet_name="integra", index=False)
        pd.DataFrame(linhas_evento).to_excel(writer,  sheet_name="evento",  index=False)
        for sn in ["integra", "evento"]:
            _formatar_planilha_saida(writer.sheets[sn])
    buf_evento.seek(0)

    buf_integra = BytesIO()
    with pd.ExcelWriter(buf_integra, engine="openpyxl") as writer:
        pd.DataFrame(linhas_integra_xls).to_excel(writer, sheet_name="Plan1", index=False)
        _formatar_planilha_saida(writer.sheets["Plan1"])
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
        max_len    = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 50)
    ws.freeze_panes = "A2"


# ══════════════════════════════════════════════════════════════════════════
# INTERFACE STREAMLIT
# ══════════════════════════════════════════════════════════════════════════
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
                    border-radius:8px; border-top:6px solid #FF8000; margin-bottom:28px;">
            <h2 style="color:#FF8000; margin:0;">
                📊 Integração Contábil — Domínio Sistemas &nbsp;|&nbsp; {VERSAO}
            </h2>
            <p style="color:#DDDDDD; margin:6px 0 0 0;">
                <b>Etapa 1:</b> Importa PDF + TXT + Plano de Contas → classifica automaticamente → gera Excel.<br>
                <b>Etapa 2:</b> Importa Excel preenchido → gera <b>evento exemplo.xlsx</b> e <b>integra exemplo.xlsx</b>.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── Sidebar ────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("### ⚙️ Configurações")
        cod_empresa = st.text_input("Código da empresa", value="45")
        st.markdown("---")
        st.markdown("### 🎨 Legenda de Tipos")
        st.markdown("🟢 Verde → Provento")
        st.markdown("🔴 Vermelho → Desconto")
        st.markdown("🔵 Azul → Informativa")
        st.markdown("🟡 Amarelo → Inf. Dedutora")
        st.markdown("🟠 Laranja → Campos a preencher")
        st.markdown("---")
        st.markdown(f"**Versão:** {VERSAO}")

    # ── Session state ──────────────────────────────────────────────────────
    _defaults = {
        "log":            [f"Pronto. Versão {VERSAO}"],
        "excel_config":   None,
        "evento_xlsx":    None,
        "integra_xls":    None,
        "df_preview":     None,
        "n_eventos":      0,
        "df_contas":      None,
        "eventos_parsed": None,
        "config_cc":      {},
        "_contas_fid":    None,
    }
    for k, v in _defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

    # ══════════════════════════════════════════════════════════════════════
    # ETAPA 1
    # ══════════════════════════════════════════════════════════════════════
    st.markdown("## 📋 Etapa 1 — Gerar Excel para Preenchimento")

    col1, col2, col3 = st.columns(3)
    with col1:
        pdf_file = st.file_uploader("1️⃣ PDF — Rubricas/Itens Não Configurados",
                                    type=["pdf"], key="pdf_etapa1")
    with col2:
        txt_file = st.file_uploader("2️⃣ TXT — Rubricas (catálogo de tipos)",
                                    type=["txt"], key="txt_etapa1")
    with col3:
        contas_file = st.file_uploader("3️⃣ XLS/XLSX — Plano de Contas (opcional)",
                                       type=["xls", "xlsx"], key="contas_etapa1")

    # Carrega plano de contas
    if contas_file is not None:
        fid = getattr(contas_file, "file_id", id(contas_file))
        if st.session_state._contas_fid != fid:
            log_tmp: list[str] = []
            df_c = parse_plano_contas(contas_file.read(), log_tmp)
            st.session_state.df_contas   = df_c if not df_c.empty else None
            st.session_state._contas_fid = fid
            st.session_state.log.extend(log_tmp)
    else:
        st.session_state.df_contas   = None
        st.session_state._contas_fid = None

    df_pc = st.session_state.df_contas
    if df_pc is not None:
        n_a = len(df_pc[df_pc["tipo"] == "A"])
        n_s = len(df_pc[df_pc["tipo"] == "S"])
        st.success(
            f"✅ Plano de Contas carregado: **{len(df_pc)}** contas "
            f"({n_a} analíticas · {n_s} sintéticas)"
        )

    st.markdown("---")

    # ── Separador ─────────────────────────────────────────────────────────
    st.markdown("### ⚙️ Configuração de Separador")
    usa_separador = st.radio(
        "Os lançamentos usam separador por Centro de Custo?",
        ["Não", "Sim"], index=0, horizontal=True,
        help="Com separador: cada CC recebe contas contábeis específicas.",
    )
    usa_sep_bool = (usa_separador == "Sim")

    # ── Configuração por CC ────────────────────────────────────────────────
    if usa_sep_bool:
        if st.session_state.eventos_parsed:
            ccs = get_centros_custo_unicos(st.session_state.eventos_parsed)
            if ccs:
                st.markdown("#### 🏢 Classificação por Centro de Custo")

                # ── Painel de não classificados ───────────────────────────
                nao_classif = []
                for cc_cod, cc_nome in ccs:
                    cfg = st.session_state.config_cc.get(cc_cod, {})
                    if not cfg.get("conta_debito") or not cfg.get("conta_credito"):
                        nao_classif.append(f"CC {cc_cod} — {cc_nome}")

                if nao_classif:
                    with st.expander(
                        f"⚠️ {len(nao_classif)} Centro(s) de Custo sem classificação completa",
                        expanded=True,
                    ):
                        st.warning(
                            "Os seguintes CCs ainda não têm Conta Débito **e** "
                            "Conta Crédito definidas:"
                        )
                        for item in nao_classif:
                            st.markdown(f"- {item}")
                        st.info(
                            "💡 Selecione o **Grupo de Despesa** abaixo — o sistema "
                            "classificará automaticamente as contas do seu plano."
                        )
                else:
                    st.success("✅ Todos os Centros de Custo estão classificados!")

                st.markdown("---")

                # ── Expanders por CC ──────────────────────────────────────
                for cc_cod, cc_nome in ccs:
                    cfg_atual = st.session_state.config_cc.get(cc_cod, {})
                    deb_ok    = bool(cfg_atual.get("conta_debito"))
                    cred_ok   = bool(cfg_atual.get("conta_credito"))
                    status    = "✅" if (deb_ok and cred_ok) else "⚠️"

                    with st.expander(
                        f"{status} CC {cc_cod} — {cc_nome}",
                        expanded=not (deb_ok and cred_ok),
                    ):
                        # ── Seleção de grupo ──────────────────────────────
                        grupo_idx = GRUPOS_LISTA.index(cfg_atual.get("grupo", "Outro")) \
                                    if cfg_atual.get("grupo") in GRUPOS_LISTA \
                                    else len(GRUPOS_LISTA) - 1

                        grupo_sel = st.selectbox(
                            "📂 Grupo de Despesa",
                            options=GRUPOS_LISTA,
                            index=grupo_idx,
                            key=f"grupo_{cc_cod}",
                        )

                        # ── Classificação automática ──────────────────────
                        if df_pc is not None:
                            auto = sugerir_grupo_automatico(df_pc, grupo_sel)
                            ops_deb  = auto["ops_deb"]
                            ops_cred = auto["ops_cred"]

                            # Métricas de cobertura
                            col_m1, col_m2, col_m3 = st.columns(3)
                            col_m1.metric("Contas Débito encontradas",  auto["n_deb"])
                            col_m2.metric("Contas Crédito encontradas", auto["n_cred"])
                            col_m3.metric(
                                "Status classificação",
                                "✅ OK" if (auto["n_deb"] > 0 and auto["n_cred"] > 0)
                                else "⚠️ Verificar",
                            )

                            # Alerta se não encontrou contas
                            if auto["n_deb"] == 0:
                                st.warning(
                                    "⚠️ Nenhuma conta de **Débito** encontrada para "
                                    f"o grupo **{grupo_sel}** neste plano de contas. "
                                    "Selecione manualmente ou escolha outro grupo."
                                )
                            if auto["n_cred"] == 0:
                                st.warning(
                                    "⚠️ Nenhuma conta de **Crédito** encontrada para "
                                    f"o grupo **{grupo_sel}** neste plano de contas. "
                                    "Selecione manualmente ou escolha outro grupo."
                                )
                        else:
                            ops_deb  = ["(Carregue o Plano de Contas)"]
                            ops_cred = ["(Carregue o Plano de Contas)"]

                        # ── Selectboxes Débito / Crédito ──────────────────
                        col_d, col_c, col_h = st.columns([3, 3, 2])

                        with col_d:
                            deb_sel = st.selectbox(
                                f"💸 Conta Débito ({len(ops_deb)-1} opções)",
                                options=ops_deb,
                                index=_idx(ops_deb, cfg_atual.get("conta_debito", "")),
                                key=f"deb_{cc_cod}",
                            )

                        with col_c:
                            cred_sel = st.selectbox(
                                f"💰 Conta Crédito ({len(ops_cred)-1} opções)",
                                options=ops_cred,
                                index=_idx(ops_cred, cfg_atual.get("conta_credito", "")),
                                key=f"cred_{cc_cod}",
                            )

                        with col_h:
                            hist_sel = st.text_input(
                                "📋 Histórico",
                                value=cfg_atual.get("historico", ""),
                                key=f"hist_{cc_cod}",
                                placeholder="Ex: 001",
                            )

                        # ── Preview ───────────────────────────────────────
                        deb_cod  = extrair_codigo(deb_sel)
                        cred_cod = extrair_codigo(cred_sel)

                        def _nome_conta(cod):
                            if not cod or df_pc is None:
                                return "—"
                            r = df_pc[df_pc["classificacao"] == cod]
                            return r.iloc[0]["nome_conta"] if not r.empty else cod

                        if deb_cod or cred_cod:
                            cor_preview = "#e8f5e9" if (deb_cod and cred_cod) else "#fff3e0"
                            borda_preview = "#4caf50" if (deb_cod and cred_cod) else "#FF8000"
                            st.markdown(
                                f"""
                                <div style="background:{cor_preview};
                                            border-left:4px solid {borda_preview};
                                            padding:8px 12px; border-radius:4px; margin-top:6px;
                                            font-size:13px;">
                                    <b>D:</b> <code>{deb_cod or '—'}</code>
                                    {_nome_conta(deb_cod)}
                                    &nbsp;&nbsp;&nbsp;
                                    <b>C:</b> <code>{cred_cod or '—'}</code>
                                    {_nome_conta(cred_cod)}
                                </div>
                                """,
                                unsafe_allow_html=True,
                            )

                        # Persiste
                        st.session_state.config_cc[cc_cod] = {
                            "grupo":         grupo_sel,
                            "conta_debito":  deb_cod,
                            "conta_credito": cred_cod,
                            "historico":     hist_sel,
                        }
            else:
                st.warning("Nenhum Centro de Custo encontrado. Processe o PDF primeiro.")
        else:
            st.info("⬆️ Faça upload do PDF e clique em **🔍 Pré-visualizar CCs**.")
            if pdf_file and txt_file:
                if st.button("🔍 Pré-visualizar CCs", use_container_width=False):
                    log_tmp: list[str] = []
                    parse_rubricas_txt(txt_file.read(), log_tmp)
                    evs = parse_nao_configurados_pdf(pdf_file.read(), log_tmp)
                    st.session_state.eventos_parsed = evs
                    st.session_state.log.extend(log_tmp)

                    # Classificação automática imediata se há plano de contas
                    if df_pc is not None:
                        ccs_tmp = get_centros_custo_unicos(evs)
                        for cc_cod, _ in ccs_tmp:
                            if cc_cod not in st.session_state.config_cc:
                                # Grupo padrão: Despesa Administrativa
                                auto = sugerir_grupo_automatico(df_pc, "Despesa Administrativa")
                                st.session_state.config_cc[cc_cod] = {
                                    "grupo":         "Despesa Administrativa",
                                    "conta_debito":  auto["conta_debito"],
                                    "conta_credito": auto["conta_credito"],
                                    "historico":     "",
                                }
                    st.rerun()

    st.markdown("---")

    # ── Botões principais ──────────────────────────────────────────────────
    col_btn1, col_btn2 = st.columns([1, 1])
    with col_btn1:
        gerar_excel = st.button(
            "▶ Gerar Excel de Configuração",
            disabled=(pdf_file is None or txt_file is None),
            use_container_width=True, type="primary",
        )
    with col_btn2:
        if st.button("🗑 Limpar tudo", use_container_width=True):
            for k in ["log", "excel_config", "evento_xlsx", "integra_xls",
                      "df_preview", "n_eventos", "df_contas", "eventos_parsed",
                      "config_cc", "_contas_fid"]:
                if k == "log":         st.session_state[k] = ["Campos limpos."]
                elif k == "n_eventos": st.session_state[k] = 0
                elif k == "config_cc": st.session_state[k] = {}
                else:                  st.session_state[k] = None
            st.rerun()

    if gerar_excel and pdf_file and txt_file:
        log: list[str] = ["[Etapa 1] Iniciando..."]
        with st.spinner("Lendo rubricas.txt..."):
            catalog = parse_rubricas_txt(txt_file.read(), log)
        with st.spinner("Lendo PDF..."):
            eventos = parse_nao_configurados_pdf(pdf_file.read(), log)
        st.session_state.eventos_parsed = eventos

        # Classificação automática para CCs novos
        if df_pc is not None and usa_sep_bool:
            ccs_novos = get_centros_custo_unicos(eventos)
            for cc_cod, _ in ccs_novos:
                if cc_cod not in st.session_state.config_cc or \
                   not st.session_state.config_cc[cc_cod].get("conta_debito"):
                    grupo_default = "Despesa Administrativa"
                    auto = sugerir_grupo_automatico(df_pc, grupo_default)
                    st.session_state.config_cc[cc_cod] = {
                        "grupo":         grupo_default,
                        "conta_debito":  auto["conta_debito"],
                        "conta_credito": auto["conta_credito"],
                        "historico":     "",
                    }
                    log.append(
                        f"CC {cc_cod}: classificado automaticamente → "
                        f"D:{auto['conta_debito']} / C:{auto['conta_credito']}"
                    )

        if not eventos:
            log.append("AVISO: Nenhum evento encontrado no PDF.")
        else:
            with st.spinner("Gerando Excel..."):
                excel_bytes = gerar_excel_configuracao(
                    eventos, catalog, cod_empresa, log,
                    usa_separador=usa_sep_bool,
                    config_cc=st.session_state.config_cc if usa_sep_bool else None,
                    df_contas=df_pc,
                )
            st.session_state.excel_config = excel_bytes
            st.session_state.n_eventos    = len(eventos)

            # Preview
            linhas_prev = []
            for ev in eventos:
                cod_ev = ev["cod"]
                info   = catalog.get(cod_ev, {})
                cc_cod = ev["centro_custo_cod"]
                cfg_cc = st.session_state.config_cc.get(cc_cod, {}) if usa_sep_bool else {}
                classif_ok = bool(cfg_cc.get("conta_debito") and cfg_cc.get("conta_credito"))
                linhas_prev.append({
                    "Código":         cod_ev,
                    "Descrição":      ev["descricao_pdf"],
                    "Tipo":           info.get("tipo", "⚠️"),
                    "Tipo Folha":     ev["tipo_folha_desc"],
                    "Centro Custo":   ev["centro_custo_nome"],
                    "Grupo":          cfg_cc.get("grupo", "—"),
                    "Conta Débito":   cfg_cc.get("conta_debito",  ""),
                    "Conta Crédito":  cfg_cc.get("conta_credito", ""),
                    "Classificado":   "✅" if classif_ok else "⚠️",
                })
            st.session_state.df_preview = pd.DataFrame(linhas_prev)

        st.session_state.log = log
        st.rerun()

    # ── Resultado Etapa 1 ──────────────────────────────────────────────────
    if st.session_state.excel_config is not None:
        st.success(f"✅ Excel gerado — {st.session_state.n_eventos} evento(s)")
        st.download_button(
            label="⬇ Baixar Excel de Configuração",
            data=st.session_state.excel_config,
            file_name="configuracao_rubricas_dominio.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, type="primary",
        )

        if st.session_state.df_preview is not None:
            df = st.session_state.df_preview
            total = len(df)
            p    = len(df[df["Tipo"] == "Provento"])
            d    = len(df[df["Tipo"] == "Desconto"])
            i    = len(df[df["Tipo"] == "Informativa"])
            id_  = len(df[df["Tipo"] == "Inf. Dedutora"])
            nf   = len(df[df["Tipo"].str.startswith("⚠️", na=False)])
            ok   = len(df[df.get("Classificado", pd.Series(dtype=str)) == "✅"]) \
                   if "Classificado" in df.columns else 0
            nok  = len(df[df.get("Classificado", pd.Series(dtype=str)) == "⚠️"]) \
                   if "Classificado" in df.columns else 0

            m1, m2, m3, m4, m5, m6, m7, m8 = st.columns(8)
            m1.metric("📋 Total",       total)
            m2.metric("🟢 Proventos",   p)
            m3.metric("🔴 Descontos",   d)
            m4.metric("🔵 Informativas",i)
            m5.metric("🟡 Inf. Ded.",   id_)
            m6.metric("⚠️ Tipo n/id",   nf)
            m7.metric("✅ Classif.",     ok)
            m8.metric("⚠️ Sem conta",   nok)

            # Alerta de não classificados
            if nok > 0 and usa_sep_bool:
                df_nok = df[df["Classificado"] == "⚠️"][
                    ["Código", "Descrição", "Centro Custo", "Conta Débito", "Conta Crédito"]
                ]
                with st.expander(f"⚠️ {nok} evento(s) sem classificação completa", expanded=True):
                    st.warning(
                        "Estes eventos não têm Conta Débito **e/ou** Conta Crédito preenchidas. "
                        "Ajuste os Centros de Custo acima ou preencha manualmente no Excel."
                    )
                    st.dataframe(df_nok, use_container_width=True)

            def hl(row):
                t = str(row.get("Tipo", ""))
                if t == "Provento":      return ["background-color:#d4edda"] * len(row)
                if t == "Desconto":      return ["background-color:#f8d7da"] * len(row)
                if t == "Informativa":   return ["background-color:#cce5ff"] * len(row)
                if t == "Inf. Dedutora": return ["background-color:#fff3cd"] * len(row)
                return ["background-color:#e2e3e5"] * len(row)

            st.dataframe(df.head(100).style.apply(hl, axis=1), use_container_width=True)

    st.markdown("---")

    # ══════════════════════════════════════════════════════════════════════
    # ETAPA 2
    # ══════════════════════════════════════════════════════════════════════
    st.markdown("## 📥 Etapa 2 — Importar Excel Preenchido → Gerar Arquivos Finais")
    st.markdown("""
    1. Baixe o Excel da Etapa 1 · 2. Ajuste as contas se necessário · 3. Faça upload e clique em **▶ Gerar Arquivos Finais**
    """)

    excel_preenchido = st.file_uploader(
        "4️⃣ Excel Preenchido (.xlsx)", type=["xlsx", "xls"], key="excel_etapa2",
    )
    col_btn3, _ = st.columns([1, 1])
    with col_btn3:
        gerar_finais = st.button(
            "▶ Gerar Arquivos Finais",
            disabled=(excel_preenchido is None),
            use_container_width=True, type="primary",
        )

    if gerar_finais and excel_preenchido:
        log = list(st.session_state.log)
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
                use_container_width=True, type="primary",
            )
        with col_d2:
            st.download_button(
                label="⬇ Baixar integra exemplo.xlsx",
                data=st.session_state.integra_xls,
                file_name="integra exemplo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, type="primary",
            )

    # ── Log ────────────────────────────────────────────────────────────────
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
                    overflow-y:auto; color:#1F1F1F;">{log_texto}</div>
        """,
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
