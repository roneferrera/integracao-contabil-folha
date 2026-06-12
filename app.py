# ============================================================
# app_integracao_dominio.py  –  Integração Contábil Domínio V4.2
# ============================================================

import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

VERSAO = "V4.2"

# ══════════════════════════════════════════════════════════════════════════
# TEMA
# ══════════════════════════════════════════════════════════════════════════
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
# PARSE DO PLANO DE CONTAS
# Estrutura real do Contas.xls (Domínio):
#   Linha 0 = cabeçalho: "Plano de Contas - Completo", "Unnamed:1"(Reduzido),
#             "Unnamed:2"(Classificação), "Unnamed:3"(Tipo), "Unnamed:4"(Descrição)...
#   col[0] = Empresa  |  col[1] = Reduzido  |  col[2] = Classificação
#   col[3] = Tipo S/A |  col[4] = Descrição
#   Última linha pode ser "Total de : NNN" → ignorar
# ══════════════════════════════════════════════════════════════════════════
def parse_plano_contas(file_bytes: bytes, log: list) -> pd.DataFrame:
    # Tenta ler como XLS (xlrd) e como XLSX (openpyxl)
    df_raw = None
    for engine in [None, "xlrd", "openpyxl"]:
        try:
            kwargs = {"sheet_name": 0, "header": 0, "dtype": str}
            if engine:
                kwargs["engine"] = engine
            df_raw = pd.read_excel(BytesIO(file_bytes), **kwargs)
            break
        except Exception:
            continue

    if df_raw is None:
        log.append("ERRO: Não foi possível abrir o Plano de Contas. "
                   "Tente salvar como .xlsx no Excel e importe novamente.")
        return pd.DataFrame()

    log.append(f"Plano de Contas: arquivo aberto — {len(df_raw)} linhas brutas, "
               f"{len(df_raw.columns)} colunas.")

    if len(df_raw.columns) < 5:
        log.append("ERRO: Plano de Contas com menos de 5 colunas. "
                   "Verifique se o arquivo está correto.")
        return pd.DataFrame()

    registros = []
    ignorados = 0

    for _, row in df_raw.iterrows():
        # Lê os campos pelas posições fixas
        empresa  = str(row.iloc[0]).strip()
        classif  = str(row.iloc[2]).strip()
        tipo_raw = str(row.iloc[3]).strip().upper()
        nome     = str(row.iloc[4]).strip()

        # Ignora linhas de rodapé / totalizador / vazias
        if empresa.lower().startswith("total") or classif.lower().startswith("total"):
            ignorados += 1
            continue

        # Classificação deve ser puramente numérica
        if not re.match(r'^\d+$', classif):
            ignorados += 1
            continue

        # Tipo deve ser S ou A
        if tipo_raw not in ("S", "A"):
            ignorados += 1
            continue

        # Nome não pode ser vazio / nan
        if not nome or nome.lower() in ("nan", "none", ""):
            ignorados += 1
            continue

        registros.append({
            "classificacao": classif,
            "nome_conta":    nome.upper(),   # MAIÚSCULAS para comparação uniforme
            "nome_original": nome,           # preserva original para exibição
            "tipo":          tipo_raw,
        })

    df = (
        pd.DataFrame(registros)
        .drop_duplicates(subset=["classificacao"])
        .reset_index(drop=True)
    )

    n_a = len(df[df["tipo"] == "A"])
    n_s = len(df[df["tipo"] == "S"])

    log.append(
        f"Plano de Contas: {len(df)} contas válidas "
        f"({n_a} analíticas · {n_s} sintéticas · {ignorados} linhas ignoradas)."
    )

    if n_a == 0:
        log.append(
            "AVISO: Nenhuma conta analítica encontrada. "
            "Verifique se a coluna 'Tipo' contém os valores 'A' e 'S'."
        )

    return df


# ══════════════════════════════════════════════════════════════════════════
# CLASSIFICADOR SEMÂNTICO UNIVERSAL
#
# Compara palavras-chave com os NOMES das contas (em MAIÚSCULAS, sem acento).
# Funciona com qualquer plano de contas — não depende de prefixos numéricos.
# ══════════════════════════════════════════════════════════════════════════

def _norm(texto: str) -> str:
    """Maiúsculas + remove acentos para comparação."""
    t = texto.upper()
    for orig, sub in [
        ("Ã","A"),("Á","A"),("Â","A"),("À","A"),("Ä","A"),
        ("É","E"),("Ê","E"),("È","E"),("Ë","E"),
        ("Í","I"),("Î","I"),("Ï","I"),
        ("Ó","O"),("Ô","O"),("Õ","O"),("Ö","O"),
        ("Ú","U"),("Û","U"),("Ü","U"),
        ("Ç","C"),("Ñ","N"),
    ]:
        t = t.replace(orig, sub)
    return t


# ── Palavras-chave DÉBITO por grupo ───────────────────────────────────────
# Baseadas nos nomes REAIS do Contas.xls (já em maiúsculas / sem acento)
KWORDS_DEBITO: dict[str, list[str]] = {
    "Custo Direto de Produção": [
        "MATERIA-PRIMA", "MATERIAL APLICADO", "MAO-DE-OBRA DIRETA",
        "SALARIOS E ORDENADOS", "PRO-LABORE", "PREMIOS DE GRATIFICACOES",
        "13 SALARIO", "FERIAS", "INSS", "FGTS", "INDENIZACOES",
        "ASSISTENCIA MEDICA", "VALE TRANSPORTE", "PIS S/ FOLHA",
        "ALIMENTACAO", "VALE REFEICAO", "HORAS EXTRAS", "SEGURO DE VIDA",
        "TREINAMENTO", "BOLSA AUXILIO", "CONTRIBUICAO ASSISTENCIAL",
        "SERVICOS PESSOAL PJ", "INDUSTRIALIZACAO",
        "CUSTOS DIRETOS DE PRODUCAO",
    ],
    "Custo Direto de Serviços": [
        "CUSTOS DIRETOS DA PRODUCAO DE SERVICOS",
        "MAO-DE-OBRA DIRETA", "SALARIOS E ORDENADOS", "PRO-LABORE",
        "PREMIOS DE GRATIFICACOES", "13 SALARIO", "FERIAS", "INSS", "FGTS",
        "INDENIZACOES", "ASSISTENCIA MEDICA", "VALE TRANSPORTE",
        "PIS S/ FOLHA", "ALIMENTACAO", "VALE REFEICAO", "HORAS EXTRAS",
        "SEGURO DE VIDA", "TREINAMENTO", "BOLSA AUXILIO",
        "CONTRIBUICAO ASSISTENCIAL", "SERVICOS PESSOAL PJ",
        "ROYALTIES", "DESPESAS - KM OP", "ESTACIONAMENTOS E PEDAGIOS",
        "ASSISTENCIA ODONTOLOGICA", "SEGUROS DE ACIDENTES",
        "BENEFICIOS CONCEDIDOS",
    ],
    "Custo Indireto de Produção": [
        "MAO-DE-OBRA INDIRETA", "MATERIAIS DE CONSUMO INDIRETO",
        "MATERIAIS DE MANUTENCAO", "UTILIDADES E SERVICOS",
        "ALUGUEIS E ARRENDAMENTOS", "DEPRECIACOES", "AMORTIZACOES",
        "COMBUSTIVEIS", "ENERGIA ELETRICA", "AUDITORIA E CONSULTORIA",
        "LOCACAO DE MAQUINAS", "OUTROS SERVICOS TOMADOS",
        "ASSESSORIA EM INFORMATICA", "CONDOMINIO",
        "CUSTOS INDIRETOS DE PRODUCAO",
    ],
    "Despesa Administrativa": [
        "DESPESAS ADMINISTRATIVAS", "DESPESAS COM PESSOAL",
        "SALARIOS E ORDENADOS", "PRO-LABORE", "PREMIOS E GRATIFICACOES",
        "13 SALARIO", "FERIAS", "INSS", "FGTS", "INDENIZACOES",
        "ASSISTENCIA MEDICA", "VALE TRANSPORTE", "PIS S/ FOLHA",
        "ALIMENTACAO", "HORAS EXTRAS", "SEGURO DE VIDA", "TREINAMENTO",
        "BOLSA AUXILIO", "CONTRIBUICAO ASSISTENCIAL", "SERVICOS PESSOAL PJ",
        "ALUGUEIS E ARRENDAMENTOS", "ALUGUEIS DE IMOVEIS",
        "ALUGUEIS DE MAQUINAS", "ARRENDAMENTO", "LEASING",
        "IMPOSTOS, TAXAS E CONTRIBUICOES", "PIS", "COFINS", "IPTU", "IPVA",
        "TAXAS DIVERSAS", "MULTAS DE MORA",
        "ENERGIA ELETRICA", "AGUA E ESGOTO", "TELEFONE",
        "DESPESAS POSTAIS", "SEGUROS", "MATERIAL DE ESCRITORIO",
        "MATERIAL DE HIGIENE", "DEPRECIACOES E AMORTIZACOES",
        "REPRODUCOES", "DESPESAS LEGAIS", "LIVROS, JORNAIS",
        "COMBUSTIVEIS E LUBRIFICANTES", "MATERIAIS DE CONSUMO",
        "CONDOMINIOS", "CELULAR", "CONSELHOS DE CLASSE",
        "ESTACIONAMENTOS E PEDAGIOS", "CARTORIO", "GAS", "CONDUCOES",
        "REFEICOES", "MANUTENCAO E REPARO", "VIAGENS",
        "MANUTENCAO DE VEICULOS", "FRETES E CARRETOS",
        "SERVICOS TOMADOS DE PJ", "SERVS. DE PUBLICIDADE",
        "SERVS. MEDICINAIS", "SERVS. SEGURANCA DO TRABALHO",
        "SERVS. ASSIST. TECNICA", "SERVS. DE MANUTENCAO",
        "SERVS. ADVOCATICIOS", "SERVS. DE CONTABILIDADE",
        "SERVS. DE TRANSPORTE", "SERVS. SISTEMAS E MONITORAMENTO",
        "SERVS. ADMINISTRATIVOS", "SERVS. MANUTENCAO DE INFORMATICA",
        "SERVICOS DE LIMPEZA", "SERVICOS PRESTADOS POR TERCEIROS",
        "SEGURANCA PATRIMONIAL", "DESPESAS PLATAFORMAS", "CORREIOS",
        "ASSESSORIA DE IMPRENSA", "PROVEDOR DE INTERNET",
        "LICENCA DE USO", "EVENTOS INTERNOS", "FEIRAS E EVENTOS",
        "LOCACAO DE MAQUINAS E EQUIPAMENTOS", "LOCACAO DE VEICULOS",
        "REEMBOLSO DE DESPESAS", "DESPESAS - KM ADM",
        "ASSISTENCIA ODONTOLOGICA", "SEGUROS DE ACIDENTES",
        "BENEFICIOS CONCEDIDOS", "COMISSOES",
        "DESPESAS GERAIS",
    ],
    "Despesa com Vendas": [
        "DESPESAS COM VENDAS", "COMISSOES SOBRE VENDAS", "COMISSOES",
        "PROPAGANDA E PUBLICIDADE", "AMOSTRAS GRATIS",
        "DESPESAS COM ENTREGA", "FRETES E CARRETOS",
        "MANUTENCAO DE VEICULOS", "DESPESAS COM VIAGENS",
        "VIAGENS TERRESTRES", "VIAGENS AEREAS", "HOSPEDAGEM",
        "REFEICOES", "DESPESAS GERAIS", "ALUGUEIS",
        "MANUTENCAO E REPARO", "TELEFONE", "DESPESAS POSTAIS",
        "DEPRECIACOES E AMORTIZACOES", "SERVICOS PRESTADOS POR TERCEIROS",
        "SEGUROS", "PERDAS NO RECEBIMENTO",
        "CREDITOS VENCIDOS E NAO LIQUIDADOS",
        "SALARIOS E ORDENADOS", "PRO-LABORE", "13 SALARIO",
        "FERIAS", "INSS", "FGTS", "INDENIZACOES",
        "ASSISTENCIA MEDICA", "VALE TRANSPORTE", "PIS S/ FOLHA",
        "HORAS EXTRAS", "VALE REFEICAO", "SEGURO DE VIDA",
        "TREINAMENTO", "BOLSA AUXILIO", "CONTRIBUICAO ASSISTENCIAL",
        "SERVICOS PESSOAL PJ",
    ],
    "Despesa Financeira": [
        "DESPESAS FINANCEIRAS", "JUROS PASSIVOS",
        "VARIACOES MONETARIAS PASSIVAS", "VARIACOES CAMBIAIS PASSIVAS",
        "DESCONTOS FINANCEIROS CONCEDIDOS", "JUROS DE MORA",
        "JUROS E COMISSOES BANCARIAS",
        "JUROS SOBRE EMPRESTIMOS E FINANCIAMENTOS",
        "MULTAS PASSIVAS", "MULTAS DE MORA",
        "TARIFA BANCARIA", "EMPRESTIMO / FINANCIAMENTO",
        "PERDAS DE APLICACOES FINANCEIRAS", "IOF",
    ],
    "Despesa Não Operacional": [
        "DESPESAS NAO OPERACIONAIS", "RESULTADOS NAO OPERACIONAIS",
        "RESULTADOS NEGATIVOS", "PERDAS NA ALIENACAO",
        "RESULTADO NEGATIVO NA ALIENACAO",
        "RESULTADO NEGATIVO DE SINISTRO",
        "OUTRAS BAIXAS DO ATIVO", "BAIXAS DE INVESTIMENTOS",
        "BAIXAS DE IMOBILIZADO", "BAIXAS DE ATIVO DIFERIDO",
        "PROVISOES PARA PERDAS PERMANENTE",
        "PROVISAO DE IRPJ", "PROVISAO DE CSLL",
        "PROVISAO IRPJ", "PROVISAO CSLL",
        "IMPOSTO DE RENDA", "CONTRIBUICAO SOCIAL",
        "PERDAS POR FALTA NO INVENTARIO",
    ],
}

# ── Palavras-chave CRÉDITO por grupo ──────────────────────────────────────
KWORDS_CREDITO: dict[str, list[str]] = {
    "Custo Direto de Produção": [
        "SALARIOS E ORDENADOS A PAGAR", "PRO-LABORE A PAGAR",
        "GRATIFICACOES A PAGAR", "FERIAS A PAGAR", "RESCISOES A PAGAR",
        "13 SALARIO A PAGAR", "PENSAO ALIMENTICIA", "INDENIZACOES A PAGAR",
        "INSS A RECOLHER", "FGTS A RECOLHER", "PIS S/ FOLHA A RECOLHER",
        "PROVISOES PARA FERIAS", "PROVISOES PARA 13",
        "INSS SOBRE PROVISOES", "FGTS SOBRE PROVISOES",
        "OBRIGACOES COM O PESSOAL", "OBRIGACOES SOCIAIS", "PROVISOES",
        "OBRIGACOES TRABALHISTA",
    ],
    "Custo Direto de Serviços": [
        "SALARIOS E ORDENADOS A PAGAR", "PRO-LABORE A PAGAR",
        "GRATIFICACOES A PAGAR", "FERIAS A PAGAR", "RESCISOES A PAGAR",
        "13 SALARIO A PAGAR", "INDENIZACOES A PAGAR",
        "INSS A RECOLHER", "FGTS A RECOLHER", "PIS S/ FOLHA A RECOLHER",
        "PROVISOES PARA FERIAS", "PROVISOES PARA 13",
        "INSS SOBRE PROVISOES", "FGTS SOBRE PROVISOES",
        "OBRIGACOES COM O PESSOAL", "OBRIGACOES SOCIAIS", "PROVISOES",
        "OBRIGACOES TRABALHISTA", "FORNECEDORES", "CONTAS A PAGAR",
    ],
    "Custo Indireto de Produção": [
        "SALARIOS E ORDENADOS A PAGAR", "PRO-LABORE A PAGAR",
        "FERIAS A PAGAR", "13 SALARIO A PAGAR", "INDENIZACOES A PAGAR",
        "INSS A RECOLHER", "FGTS A RECOLHER",
        "PROVISOES PARA FERIAS", "PROVISOES PARA 13",
        "OBRIGACOES COM O PESSOAL", "OBRIGACOES SOCIAIS", "PROVISOES",
        "OBRIGACOES TRABALHISTA", "FORNECEDORES", "CONTAS A PAGAR",
        "ALUGUEIS A PAGAR",
    ],
    "Despesa Administrativa": [
        "SALARIOS E ORDENADOS A PAGAR", "PRO-LABORE A PAGAR",
        "GRATIFICACOES A PAGAR", "FERIAS A PAGAR", "RESCISOES A PAGAR",
        "13 SALARIO A PAGAR", "PENSAO ALIMENTICIA", "INDENIZACOES A PAGAR",
        "INSS A RECOLHER", "FGTS A RECOLHER", "PIS S/ FOLHA A RECOLHER",
        "PROVISOES PARA FERIAS", "PROVISOES PARA 13",
        "INSS SOBRE PROVISOES", "FGTS SOBRE PROVISOES",
        "OBRIGACOES COM O PESSOAL", "OBRIGACOES SOCIAIS", "PROVISOES",
        "OBRIGACOES TRABALHISTA",
        "IMPOSTOS E CONTRIBUICOES A RECOLHER", "ISS A RECOLHER",
        "IRRF A RECOLHER", "INSS RETIDO A RECOLHER",
        "FORNECEDORES", "CONTAS A PAGAR",
        "HONORARIOS CONTABEIS", "ENERGIA ELETRICA A PAGAR",
        "TELEFONE A PAGAR", "ALUGUEIS A PAGAR",
        "CARTAO DE CREDITO A PAGAR", "SEGUROS A PAGAR",
        "OUTRAS OBRIGACOES",
    ],
    "Despesa com Vendas": [
        "SALARIOS E ORDENADOS A PAGAR", "PRO-LABORE A PAGAR",
        "FERIAS A PAGAR", "13 SALARIO A PAGAR", "INDENIZACOES A PAGAR",
        "INSS A RECOLHER", "FGTS A RECOLHER",
        "PROVISOES PARA FERIAS", "PROVISOES PARA 13",
        "OBRIGACOES COM O PESSOAL", "OBRIGACOES SOCIAIS", "PROVISOES",
        "OBRIGACOES TRABALHISTA", "FORNECEDORES", "CONTAS A PAGAR",
        "OUTRAS OBRIGACOES",
    ],
    "Despesa Financeira": [
        "CONTAS A PAGAR", "OUTRAS OBRIGACOES",
        "EMPRESTIMO BANCO", "EMPRESTIMOS PAGAR", "FINANCIAMENTO",
        "IMPOSTOS E CONTRIBUICOES A RECOLHER", "HONORARIOS CONTABEIS",
    ],
    "Despesa Não Operacional": [
        "CONTAS A PAGAR", "OUTRAS OBRIGACOES",
        "IMPOSTOS E CONTRIBUICOES A RECOLHER",
        "PROVISAO PARA IMPOSTO DE RENDA",
        "PROVISAO P/ CONTRIBUICAO SOCIAL",
        "IMPOSTO DE RENDA A RECOLHER",
        "CONTRIBUICAO SOCIAL A RECOLHER",
    ],
}

GRUPOS_LISTA = list(KWORDS_DEBITO.keys()) + ["Outro"]


def _conta_bate(nome_conta_upper: str, keywords: list[str]) -> bool:
    """Verifica se o nome (já em MAIÚSCULAS normalizadas) contém alguma keyword."""
    nome_norm = _norm(nome_conta_upper)
    for kw in keywords:
        if _norm(kw) in nome_norm:
            return True
    return False


def _analiticas(df: pd.DataFrame) -> pd.DataFrame:
    return df[df["tipo"] == "A"].copy() if not df.empty else df


def _fmt_opcoes(df_f: pd.DataFrame) -> list[str]:
    """Formata lista de opções para selectbox: 'CLASSIF - nome_original'."""
    col_nome = "nome_original" if "nome_original" in df_f.columns else "nome_conta"
    return [""] + [
        f"{r['classificacao']} - {r[col_nome]}"
        for _, r in df_f.iterrows()
    ]


def classificar_contas(df_contas: pd.DataFrame, grupo: str) -> tuple[list[str], list[str]]:
    """
    Retorna (opcoes_debito, opcoes_credito) filtradas por keywords do grupo.
    Fallback para todas as analíticas se nenhuma conta for encontrada.
    """
    df_a = _analiticas(df_contas)
    if df_a.empty:
        return [""], [""]

    kw_d = KWORDS_DEBITO.get(grupo, [])
    kw_c = KWORDS_CREDITO.get(grupo, [])

    if kw_d and grupo != "Outro":
        mask_d = df_a["nome_conta"].apply(lambda n: _conta_bate(n, kw_d))
        df_d   = df_a[mask_d]
        if df_d.empty:
            df_d = df_a   # fallback
    else:
        df_d = df_a

    if kw_c and grupo != "Outro":
        mask_c = df_a["nome_conta"].apply(lambda n: _conta_bate(n, kw_c))
        df_c   = df_a[mask_c]
        if df_c.empty:
            df_c = df_a   # fallback
    else:
        df_c = df_a

    return _fmt_opcoes(df_d), _fmt_opcoes(df_c)


def sugerir_contas(df_contas: pd.DataFrame, grupo: str) -> dict:
    """Sugestão automática: primeira conta de cada lado."""
    ops_d, ops_c = classificar_contas(df_contas, grupo)
    return {
        "ops_deb":       ops_d,
        "ops_cred":      ops_c,
        "conta_debito":  extrair_codigo(ops_d[1])  if len(ops_d)  > 1 else "",
        "conta_credito": extrair_codigo(ops_c[1]) if len(ops_c) > 1 else "",
        "n_deb":  len(ops_d)  - 1,
        "n_cred": len(ops_c) - 1,
    }


def extrair_codigo(opcao: str) -> str:
    if not opcao or " - " not in opcao:
        return opcao or ""
    return opcao.split(" - ")[0].strip()


def _idx(opcoes: list[str], valor: str) -> int:
    if not valor:
        return 0
    for i, op in enumerate(opcoes):
        if op.startswith(valor):
            return i
    return 0


# ══════════════════════════════════════════════════════════════════════════
# PARSE TXT RUBRICAS
# ══════════════════════════════════════════════════════════════════════════
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


# ══════════════════════════════════════════════════════════════════════════
# PARSE PDF
# ══════════════════════════════════════════════════════════════════════════
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


def get_centros_custo_unicos(eventos: list) -> list[tuple[str, str]]:
    vistos: dict[str, str] = {}
    for ev in eventos:
        cod  = ev["centro_custo_cod"]
        nome = ev["centro_custo_nome"]
        if cod and cod not in vistos:
            vistos[cod] = nome
    return list(vistos.items())


# ══════════════════════════════════════════════════════════════════════════
# ETAPA 1 — GERA EXCEL
# ══════════════════════════════════════════════════════════════════════════
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
            col_nome = "nome_original" if "nome_original" in df_contas.columns else "nome_conta"
            df_exp = df_contas[["classificacao", col_nome, "tipo"]].copy()
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


# ══════════════════════════════════════════════════════════════════════════
# ETAPA 2 — GERA ARQUIVOS FINAIS
# ══════════════════════════════════════════════════════════════════════════
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
        if   "cód. empresa"    in cl or "cod. empresa"    in cl: col_map["empresa"]        = col
        elif "cód. evento"     in cl or "cod. evento"     in cl: col_map["seq"]            = col
        elif "tipo folha (nº)" in cl or "tipo folha (n"   in cl: col_map["tipo"]           = col
        elif "descrição (rubricas)" in cl:                        col_map["desc"]           = col
        elif "descrição (pdf)" in cl and "desc" not in col_map:  col_map["desc"]           = col
        elif "cód. centro de custo" in cl:                        col_map["cc"]             = col
        elif "conta débito"    in cl or "conta debito"    in cl: col_map["debito"]         = col
        elif "conta crédito"   in cl or "conta credito"   in cl: col_map["credito"]        = col
        elif "cód. histórico"  in cl or "cod. historico"  in cl: col_map["historico"]      = col
        elif "histórico"       in cl and "cód" not in cl and "cod" not in cl:
            col_map["historico_texto"] = col
        elif "observação"      in cl:                             col_map["observacao"]     = col
        elif "usa separador"   in cl:                             col_map["usa_separador"]  = col

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

    # ── Carrega plano de contas ────────────────────────────────────────────
    if contas_file is not None:
        fid = getattr(contas_file, "file_id", id(contas_file))
        if st.session_state._contas_fid != fid:
            log_tmp: list[str] = []
            df_c = parse_plano_contas(contas_file.read(), log_tmp)
            st.session_state.df_contas   = df_c if not df_c.empty else None
            st.session_state._contas_fid = fid
            st.session_state.config_cc   = {}   # reset ao trocar plano
            st.session_state.log.extend(log_tmp)
    else:
        if st.session_state._contas_fid is not None:
            st.session_state.df_contas   = None
            st.session_state._contas_fid = None
            st.session_state.config_cc   = {}

    df_pc = st.session_state.df_contas

    if df_pc is not None and not df_pc.empty:
        n_a = len(df_pc[df_pc["tipo"] == "A"])
        n_s = len(df_pc[df_pc["tipo"] == "S"])
        st.success(
            f"✅ Plano de Contas carregado: **{len(df_pc)}** contas "
            f"({n_a} analíticas · {n_s} sintéticas)"
        )
        with st.expander("🔍 Ver amostra das contas analíticas carregadas", expanded=False):
            col_nome = "nome_original" if "nome_original" in df_pc.columns else "nome_conta"
            df_amostra = df_pc[df_pc["tipo"] == "A"][["classificacao", col_nome]].head(30)
            df_amostra.columns = ["Classificação", "Nome da Conta"]
            st.dataframe(df_amostra, use_container_width=True)
            st.caption(f"Exibindo 30 de {n_a} contas analíticas.")
    elif contas_file is not None:
        st.error(
            "❌ Plano de Contas não pôde ser carregado. "
            "Verifique o log abaixo para detalhes. "
            "Se o arquivo for .xls antigo, abra no Excel e salve como .xlsx."
        )

    st.markdown("---")

    # ── Separador ─────────────────────────────────────────────────────────
    st.markdown("### ⚙️ Configuração de Separador")
    usa_separador = st.radio(
        "Os lançamentos usam separador por Centro de Custo?",
        ["Não", "Sim"], index=0, horizontal=True,
    )
    usa_sep_bool = (usa_separador == "Sim")

    # ── Configuração por CC ────────────────────────────────────────────────
    if usa_sep_bool:
        if st.session_state.eventos_parsed:
            ccs = get_centros_custo_unicos(st.session_state.eventos_parsed)
            if ccs:
                st.markdown("#### 🏢 Classificação por Centro de Custo")

                nao_classif = [
                    f"CC {cc} — {nm}"
                    for cc, nm in ccs
                    if not st.session_state.config_cc.get(cc, {}).get("conta_debito")
                    or not st.session_state.config_cc.get(cc, {}).get("conta_credito")
                ]
                if nao_classif:
                    with st.expander(
                        f"⚠️ {len(nao_classif)} CC(s) sem classificação completa",
                        expanded=True,
                    ):
                        for item in nao_classif:
                            st.markdown(f"- {item}")
                        if df_pc is None:
                            st.info("💡 Carregue o Plano de Contas para classificação automática.")
                else:
                    st.success("✅ Todos os Centros de Custo estão classificados!")

                st.markdown("---")

                for cc_cod, cc_nome in ccs:
                    cfg_atual = st.session_state.config_cc.get(cc_cod, {})
                    deb_ok    = bool(cfg_atual.get("conta_debito"))
                    cred_ok   = bool(cfg_atual.get("conta_credito"))
                    status    = "✅" if (deb_ok and cred_ok) else "⚠️"

                    with st.expander(
                        f"{status} CC {cc_cod} — {cc_nome}",
                        expanded=not (deb_ok and cred_ok),
                    ):
                        grupo_idx = (
                            GRUPOS_LISTA.index(cfg_atual.get("grupo", "Outro"))
                            if cfg_atual.get("grupo") in GRUPOS_LISTA
                            else len(GRUPOS_LISTA) - 1
                        )
                        grupo_sel = st.selectbox(
                            "📂 Grupo de Despesa",
                            options=GRUPOS_LISTA,
                            index=grupo_idx,
                            key=f"grupo_{cc_cod}",
                        )

                        if df_pc is not None and not df_pc.empty:
                            ops_deb, ops_cred = classificar_contas(df_pc, grupo_sel)
                            n_deb  = len(ops_deb)  - 1
                            n_cred = len(ops_cred) - 1
                        else:
                            ops_deb  = [""]
                            ops_cred = [""]
                            n_deb = n_cred = 0

                        col_m1, col_m2, col_m3 = st.columns(3)
                        col_m1.metric("Contas Débito encontradas",  n_deb)
                        col_m2.metric("Contas Crédito encontradas", n_cred)
                        col_m3.metric(
                            "Status",
                            "✅ OK" if (n_deb > 0 and n_cred > 0) else "⚠️ Verificar",
                        )

                        if df_pc is not None and n_deb == 0:
                            st.warning(f"⚠️ Nenhuma conta de Débito para **{grupo_sel}**. "
                                       "Selecione manualmente ou mude o grupo.")
                        if df_pc is not None and n_cred == 0:
                            st.warning(f"⚠️ Nenhuma conta de Crédito para **{grupo_sel}**. "
                                       "Selecione manualmente ou mude o grupo.")
                        if df_pc is None:
                            st.info("💡 Carregue o Plano de Contas para sugestões automáticas.")

                        col_d, col_c, col_h = st.columns([3, 3, 2])
                        with col_d:
                            deb_sel = st.selectbox(
                                f"💸 Conta Débito ({n_deb} opções)",
                                options=ops_deb,
                                index=_idx(ops_deb, cfg_atual.get("conta_debito", "")),
                                key=f"deb_{cc_cod}",
                            )
                        with col_c:
                            cred_sel = st.selectbox(
                                f"💰 Conta Crédito ({n_cred} opções)",
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

                        deb_cod  = extrair_codigo(deb_sel)
                        cred_cod = extrair_codigo(cred_sel)

                        def _nome_conta(cod):
                            if not cod or df_pc is None:
                                return "—"
                            col_n = "nome_original" if "nome_original" in df_pc.columns else "nome_conta"
                            r = df_pc[df_pc["classificacao"] == cod]
                            return r.iloc[0][col_n] if not r.empty else cod

                        if deb_cod or cred_cod:
                            cor = "#e8f5e9" if (deb_cod and cred_cod) else "#fff3e0"
                            brd = "#4caf50" if (deb_cod and cred_cod) else "#FF8000"
                            st.markdown(
                                f"""
                                <div style="background:{cor}; border-left:4px solid {brd};
                                            padding:8px 12px; border-radius:4px; margin-top:6px;
                                            font-size:13px;">
                                    <b>D:</b> <code>{deb_cod or '—'}</code> {_nome_conta(deb_cod)}
                                    &nbsp;&nbsp;
                                    <b>C:</b> <code>{cred_cod or '—'}</code> {_nome_conta(cred_cod)}
                                </div>
                                """,
                                unsafe_allow_html=True,
                            )

                        st.session_state.config_cc[cc_cod] = {
                            "grupo":         grupo_sel,
                            "conta_debito":  deb_cod,
                            "conta_credito": cred_cod,
                            "historico":     hist_sel,
                        }
            else:
                st.warning("Nenhum Centro de Custo encontrado. Processe o PDF primeiro.")
        else:
            st.info("⬆️ Faça upload do PDF e clique em **▶ Gerar Excel** para configurar os CCs.")

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
        if df_pc is not None and not df_pc.empty and usa_sep_bool:
            ccs_novos = get_centros_custo_unicos(eventos)
            for cc_cod, _ in ccs_novos:
                if cc_cod not in st.session_state.config_cc or \
                   not st.session_state.config_cc[cc_cod].get("conta_debito"):
                    grupo_default = "Despesa Administrativa"
                    auto = sugerir_contas(df_pc, grupo_default)
                    st.session_state.config_cc[cc_cod] = {
                        "grupo":         grupo_default,
                        "conta_debito":  auto["conta_debito"],
                        "conta_credito": auto["conta_credito"],
                        "historico":     "",
                    }
                    log.append(
                        f"CC {cc_cod}: sugestão automática → "
                        f"D:{auto['conta_debito']} ({auto['n_deb']} opções) | "
                        f"C:{auto['conta_credito']} ({auto['n_cred']} opções)"
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

            linhas_prev = []
            for ev in eventos:
                cod_ev = ev["cod"]
                info   = catalog.get(cod_ev, {})
                cc_cod = ev["centro_custo_cod"]
                cfg_cc = st.session_state.config_cc.get(cc_cod, {}) if usa_sep_bool else {}
                ok     = bool(cfg_cc.get("conta_debito") and cfg_cc.get("conta_credito"))
                linhas_prev.append({
                    "Código":        cod_ev,
                    "Descrição":     ev["descricao_pdf"],
                    "Tipo":          info.get("tipo", "⚠️"),
                    "Tipo Folha":    ev["tipo_folha_desc"],
                    "Centro Custo":  ev["centro_custo_nome"],
                    "Grupo":         cfg_cc.get("grupo", "—"),
                    "Conta Débito":  cfg_cc.get("conta_debito",  ""),
                    "Conta Crédito": cfg_cc.get("conta_credito", ""),
                    "Classif.":      "✅" if ok else "⚠️",
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
            p   = len(df[df["Tipo"] == "Provento"])
            d   = len(df[df["Tipo"] == "Desconto"])
            i   = len(df[df["Tipo"] == "Informativa"])
            id_ = len(df[df["Tipo"] == "Inf. Dedutora"])
            nf  = len(df[df["Tipo"].str.startswith("⚠️", na=False)])
            ok  = len(df[df.get("Classif.", pd.Series(dtype=str)) == "✅"]) \
                  if "Classif." in df.columns else 0
            nok = len(df[df.get("Classif.", pd.Series(dtype=str)) == "⚠️"]) \
                  if "Classif." in df.columns else 0

            cols_m = st.columns(8)
            for col_m, lbl, val in zip(cols_m, [
                "📋 Total","🟢 Proventos","🔴 Descontos","🔵 Informativas",
                "🟡 Inf.Ded.","⚠️ Tipo n/id","✅ Classif.","⚠️ Sem conta"
            ], [total, p, d, i, id_, nf, ok, nok]):
                col_m.metric(lbl, val)

            if nok > 0 and usa_sep_bool and "Classif." in df.columns:
                df_nok = df[df["Classif."] == "⚠️"][
                    ["Código","Descrição","Centro Custo","Conta Débito","Conta Crédito"]
                ]
                with st.expander(f"⚠️ {nok} evento(s) sem classificação completa", expanded=True):
                    st.warning("Ajuste os CCs acima ou preencha manualmente no Excel.")
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
    st.markdown("1. Baixe o Excel da Etapa 1 · 2. Ajuste se necessário · 3. Faça upload e clique em **▶ Gerar**")

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
