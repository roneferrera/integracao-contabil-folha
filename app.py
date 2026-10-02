"""
Integrador Contábil da Folha — Domínio Sistemas
Motor 100% determinístico (regex + regras de substring). Sem IA, sem APIs.
Nenhum código, classificação ou nome de conta de um plano específico fica fixo
no programa: a estrutura é lida do plano importado e confirmada pelo usuário.
Executar: pip install -r requirements.txt && streamlit run app.py
"""
import io
import json
import re
import unicodedata
from difflib import SequenceMatcher
from pathlib import Path

import pandas as pd
import streamlit as st

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

# =====================================================================
# 0. PARÂMETROS
# =====================================================================
NOME_PADRAO_SEM_SEPARADOR = "Geral"
LIMIAR_SIMILARIDADE = 0.60
PASTA_PERFIS = Path("perfis")

# Confirme os cabeçalhos com o layout de importação da sua versão da Domínio
COLS_EVENTO = ["Código da Empresa", "Separador", "Código Sequencial", "Tipo de Integração",
               "Código da Rubrica", "Descrição da Rubrica", "Conta Débito", "Conta Crédito",
               "Histórico Padrão"]

SECOES = {  # título normalizado no PDF -> nome interno
    "FOLHA NORMAL": "Folha Normal", "FOLHA MENSAL": "Folha Normal",
    "ADIANTAMENTO": "Adiantamento", "13O SALARIO": "13º Salário",
    "FERIAS": "Férias", "RESCISAO": "Rescisão", "EMPRESA": "Empresa",
    "PROVISAO DE FERIAS": "Provisão de Férias",
    "PROVISAO DE 13O": "Provisão de 13º", "PROVISAO DE 13O SALARIO": "Provisão de 13º",
    "INFORMACOES INSS": "Informações INSS",
}
SECOES_DE_ITENS = {"Empresa", "Provisão de Férias", "Provisão de 13º", "Informações INSS"}
TIPO_INTEGRACAO = {s: s for s in set(SECOES.values())}

# Catálogo fixo da Domínio: Configurar Integração > aba Empresa (não depende do plano).
# (descrição, natureza, competência)  competência: "M" mensal | "F" férias | "13" 13º
ITENS_EMPRESA = {
    1: ("INSS Empresa", "INSS", "M"), 2: ("INSS Terceiros", "INSS", "M"),
    3: ("INSS Acid. Trabalho", "INSS", "M"), 4: ("INSS Pro-Lab", "INSS_SOCIO", "M"),
    5: ("PIS", "PIS", "M"), 6: ("INSS Empresa 13o.", "INSS", "13"),
    7: ("INSS Terceiros 13o.", "INSS", "13"), 8: ("INSS Acid. Trabalho 13o.", "INSS", "13"),
    9: ("INSS Pro-Lab/Aut. 13o.", "INSS_SOCIO", "M"), 10: ("PIS 13o.", "PIS", "13"),
    11: ("INSS Empresa Férias", "INSS", "F"), 12: ("INSS Terceiros Férias", "INSS", "F"),
    13: ("INSS Acid. Trabalho Férias", "INSS", "F"), 14: ("PIS Férias", "PIS", "F"),
    15: ("Dedução 13o. lic. maternidade", "DEDUCAO", "M"),
    16: ("INSS Empresa CCT", "INSS", "M"), 17: ("INSS Terceiros CCT", "INSS", "M"),
    18: ("INSS Acid. Trabalho CCT", "INSS", "M"), 19: ("INSS Empresa 13° CCT", "INSS", "13"),
    20: ("INSS Terceiros 13° CCT", "INSS", "13"), 21: ("INSS Acid. Trabalho 13° CCT", "INSS", "13"),
    22: ("INSS Empresa 13º Indenizado", "INSS", "13"),
    23: ("INSS Terceiros 13º Indenizado", "INSS", "13"),
    24: ("INSS Acid. Trabalho 13º Indenizado", "INSS", "13"),
    25: ("Isenção Filantropia INSS Mensal Empresa", "ISENCAO", "M"),
    26: ("Isenção Filantropia INSS Mensal RAT", "ISENCAO", "M"),
    27: ("Isenção Filantropia INSS Mensal Pró-Labore", "ISENCAO", "M"),
    28: ("Isenção Filantropia INSS Férias Empresa", "ISENCAO", "F"),
    29: ("Isenção Filantropia INSS Férias RAT", "ISENCAO", "F"),
    30: ("Isenção Filantropia INSS 13º Empresa", "ISENCAO", "13"),
    31: ("Isenção Filantropia INSS 13º RAT", "ISENCAO", "13"),
    32: ("Isenção Filantropia INSS 13º Pró-Labore", "ISENCAO", "13"),
    33: ("INSS Receita Bruta", "CPRB", "M"),
    34: ("Isenção Filantropia INSS Mensal Terceiros", "ISENCAO", "M"),
    35: ("Isenção Filantropia INSS Férias Terceiros", "ISENCAO", "F"),
    36: ("Isenção Filantropia 13º Terceiros", "ISENCAO", "13"),
    38: ("GRCS Patronal", "SIND_PATRONAL", "M"),
    39: ("INSS Empresa Mensal Dirigente sindical", "INSS", "M"),
    40: ("INSS Terceiros Mensal Dirigente sindical", "INSS", "M"),
    41: ("INSS Acid. Trabalho Mensal Dirigente sindical", "INSS", "M"),
    42: ("INSS Empresa Férias Dirigente sindical", "INSS", "F"),
    43: ("INSS Terceiros Férias Dirigente sindical", "INSS", "F"),
    44: ("INSS Acid. Trabalho Férias Dirigente sindical", "INSS", "F"),
    45: ("INSS Empresa 13ª Dirigente sindical", "INSS", "13"),
    46: ("INSS Terceiros 13ª Dirigente sindical", "INSS", "13"),
    47: ("INSS Acid. Trabalho 13ª Dirigente sindical", "INSS", "13"),
    48: ("INSS Aut.", "AUTONOMO", "M"), 49: ("Adicional ao SENAI", "SENAI", "M"),
    50: ("Adicional ao SENAI 13º", "SENAI", "13"),
}

# =====================================================================
# 1. NORMALIZAÇÃO E BUSCA DE TERMOS
# =====================================================================
def norm(txt) -> str:
    if txt is None:
        return ""
    s = unicodedata.normalize("NFKD", str(txt))
    s = "".join(c for c in s if not unicodedata.combining(c)).upper()
    s = re.sub(r"\b(?:[A-Z]\.){2,}[A-Z]?\.?", lambda m: m.group(0).replace(".", ""), s)
    s = re.sub(r"[^A-Z0-9%]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def tem(d: str, *termos) -> bool:
    """Palavra inteira; termo terminado em '*' = prefixo."""
    for t in termos:
        if t.endswith("*"):
            pat = r"(?<![A-Z0-9])" + re.escape(t[:-1])
        else:
            pat = r"(?<![A-Z0-9])" + re.escape(t) + r"(?![A-Z0-9%])"
        if re.search(pat, d):
            return True
    return False


RE_13 = re.compile(r"(?<![A-Z0-9])13O?(?![A-Z0-9])")
def e13(d): return bool(RE_13.search(d)) or tem(d, "DECIMO TERCEIRO", "DECIMO*")

# =====================================================================
# 2. CONCEITOS CONTÁBEIS (vocabulário geral — não depende de nenhum plano)
# =====================================================================
CONCEITOS = {
    "@SAL": ["SALARIO*", "ORDENADO*", "REMUNERAC*", "VENCIMENTOS"],
    "@PAGAR": ["PAGAR"],
    "@RECOLHER": ["RECOLHER", "PAGAR"],
    "@PROV": ["PROVIS*"],
    "@FER": ["FERIAS"],
    "@13": e13,
    "@INSS": ["INSS", "PREVIDENCIA*", "CPP"],
    "@FGTS": ["FGTS", "FUNDO DE GARANTIA"],
    "@PIS": ["PIS"],
    "@IRRF": ["IRRF", "IR RETIDO", "IR FONTE", "IMPOSTO DE RENDA RETIDO", "IR S FOLHA"],
    "@PROLAB": ["PRO LAB*", "PROLAB*"],
    "@SINDIC": ["SINDICA*"],
    "@ADIANT": ["ADIANT*", "ADTO*"],
}


def casa(d, termo):
    c = CONCEITOS.get(termo)
    if c is None:
        return tem(d, termo)
    return c(d) if callable(c) else tem(d, *c)


# Escopos: G = grupo do separador | GD = grupo, senão qualquer custo/despesa
#          A = Ativo | P = Passivo/PL | R = Receita | D = Custo/Despesa
G, GD = "G", "GD"
ESC_ROT = {G: "grupo do separador", GD: "grupo ou outra despesa", "A": "Ativo",
           "P": "Passivo/PL", "R": "Receita", "D": "Custo/Despesa"}
_ENC = ["@PROV", "@INSS", "@FGTS", "@PIS"]
_NAO_IRRF = ["ALUGUE*", "APLICAC*", "JUROS", "PJ", "SERVICO*", "TERCEIRO*", "CAPITAL", "@PROV"]

ALVOS = {
    # ---- DRE: dentro do grupo do separador
    "SALARIOS": (G, "Salários e Ordenados", [(["@SAL"], ["@PAGAR", "@PROV", "@13", "@FER", "@ADIANT",
                 "FAMILIA", "MATERN*", "@INSS", "@FGTS", "@PIS", "@PROLAB"])]),
    "PRO_LABORE": (G, "Pró-labore", [(["@PROLAB"], ["@PAGAR", "@INSS"]),
                   (["HONORARIOS", "DIRETORIA"], ["@PAGAR"])]),
    "HE": (G, "Horas Extras", [(["HORAS EXTRA*"], []), (["HORA EXTRA*"], []), (["EXTRAORDINAR*"], [])]),
    "PREMIOS": (G, "Prêmios e Gratificações", [(["PREMIO*"], ["SEGURO*"]),
                (["GRATIFICAC*"], ["@PAGAR", "@13"])]),
    "COMISSOES": (G, "Comissões", [(["COMISS*"], ["BANCARI*", "@PAGAR"])]),
    "DECIMO": (G, "13º Salário", [(["@13"], _ENC + ["@PAGAR"])]),
    "FERIAS": (G, "Férias", [(["@FER"], _ENC + ["@PAGAR"])]),
    "INSS": (G, "INSS (encargo)", [(["@INSS"], ["@PROV", "@RECOLHER", "@13", "@FER", "RECEITA BRUTA",
             "RETIDO", "PRIVADA"]), (["ENCARGOS SOCIAIS"], ["@PROV"])]),
    "FGTS": (G, "FGTS (encargo)", [(["@FGTS"], ["@PROV", "@RECOLHER", "@13", "@FER"]),
             (["ENCARGOS SOCIAIS"], ["@PROV"])]),
    "PIS": (G, "PIS s/ Folha", [(["@PIS"], ["@PROV", "@RECOLHER", "RETIDO", "COFINS"])]),
    "INDENIZ": (G, "Indenizações e Aviso Prévio", [(["INDENIZ*"], []), (["AVISO PREVIO"], []),
                (["RESCIS*"], ["@PAGAR"])]),
    "ASSIST": (G, "Assistência Médica", [(["ASSISTENCIA MEDICA"], []), (["PLANO DE SAUDE"], []),
               (["SAUDE"], []), (["ASSISTENCIA*"], ["CONTRIB*"])]),
    "VT": (G, "Vale-Transporte", [(["VALE TRANSP*"], []), (["TRANSPORTE"], ["FRETE*", "CARRETO*", "SERV*"])]),
    "ALIM": (G, "Alimentação / VR", [(["ALIMENTAC*"], []), (["REFEIC*"], []), (["CESTA*"], [])]),
    "BOLSA": (G, "Bolsa-Auxílio / Estágio", [(["BOLSA*"], []), (["ESTAGI*"], [])]),
    "SEGURO": (G, "Seguro de Vida", [(["SEGURO DE VIDA"], []), (["SEGURO*"], [])]),
    "TREIN": (G, "Treinamento", [(["TREINAMENT*"], []), (["CAPACITAC*"], []), (["CURSO*"], [])]),
    "D_PROV_FER": (G, "Férias - Provisão (DRE)", [(["@FER", "@PROV"], ["@INSS", "@FGTS", "@PIS"])]),
    "D_INSS_FER": (G, "INSS s/ Férias - Provisão (DRE)", [(["@INSS", "@FER", "@PROV"], [])]),
    "D_FGTS_FER": (G, "FGTS s/ Férias - Provisão (DRE)", [(["@FGTS", "@FER", "@PROV"], [])]),
    "D_PROV_13": (G, "13º - Provisão (DRE)", [(["@13", "@PROV"], ["@INSS", "@FGTS", "@PIS"])]),
    "D_INSS_13": (G, "INSS s/ 13º - Provisão (DRE)", [(["@INSS", "@13", "@PROV"], [])]),
    "D_FGTS_13": (G, "FGTS s/ 13º - Provisão (DRE)", [(["@FGTS", "@13", "@PROV"], [])]),
    # ---- Passivo
    "SAL_PAGAR": ("P", "Salários a Pagar", [(["@SAL", "@PAGAR"], ["@PROLAB", "@FER", "@13", "RESCIS*",
                  "FAMILIA", "MATERN*", "@PROV"]), (["FOLHA DE PAGAMENTO"], ["@PROV"])]),
    "FER_PAGAR": ("P", "Férias a Pagar", [(["@FER", "@PAGAR"], _ENC)]),
    "RESC_PAGAR": ("P", "Rescisões a Pagar", [(["RESCIS*", "@PAGAR"], [])]),
    "PROLAB_PAGAR": ("P", "Pró-labore a Pagar", [(["@PROLAB", "@PAGAR"], []),
                     (["HONORARIOS", "DIRETORIA", "@PAGAR"], [])]),
    "INSS_REC": ("P", "INSS a Recolher", [(["@INSS", "@RECOLHER"], ["RETIDO", "RECEITA BRUTA", "@PROV",
                 "AUTONOMO*", "PARCELAMENT*"])]),
    "FGTS_REC": ("P", "FGTS a Recolher", [(["@FGTS", "@RECOLHER"], ["@PROV", "PARCELAMENT*"])]),
    "IRRF_REC": ("P", "IRRF s/ Folha a Recolher", [(["@IRRF", "FOLHA"], ["@PROV"]),
                 (["@IRRF", "@RECOLHER"], _NAO_IRRF), (["@IRRF"], _NAO_IRRF + ["RECUPERAR", "COMPENSAR"])]),
    "PIS_REC": ("P", "PIS s/ Folha a Recolher", [(["@PIS", "FOLHA", "@RECOLHER"], ["@PROV"]),
                (["@PIS", "FOLHA"], ["@PROV"])]),
    "SIND_REC": ("P", "Contribuição Sindical a Recolher", [(["@SINDIC", "@RECOLHER"], ["PATRONAL"]),
                 (["@SINDIC"], ["PATRONAL"])]),
    "EMPREST": ("P", "Empréstimos Consignados a Repassar", [(["CONSIGNAD*"], []),
                (["EMPRESTIMO*", "EMPREGADO*"], []), (["CREDITO DO TRABALHADOR"], [])]),
    "PENSAO": ("P", "Pensão Alimentícia a Repassar", [(["PENSA*"], [])]),
    "P_PROV_FER": ("P", "Provisão de Férias", [(["@PROV", "@FER"], ["@INSS", "@FGTS", "@PIS"])]),
    "P_INSS_FER": ("P", "INSS s/ Provisão de Férias", [(["@INSS", "@PROV", "@FER"], [])]),
    "P_FGTS_FER": ("P", "FGTS s/ Provisão de Férias", [(["@FGTS", "@PROV", "@FER"], [])]),
    "P_PIS_FER": ("P", "PIS s/ Provisão de Férias", [(["@PIS", "@PROV", "@FER"], [])]),
    "P_PROV_13": ("P", "Provisão de 13º", [(["@PROV", "@13"], ["@INSS", "@FGTS", "@PIS"])]),
    "P_INSS_13": ("P", "INSS s/ Provisão de 13º", [(["@INSS", "@PROV", "@13"], [])]),
    "P_FGTS_13": ("P", "FGTS s/ Provisão de 13º", [(["@FGTS", "@PROV", "@13"], [])]),
    "P_PIS_13": ("P", "PIS s/ Provisão de 13º", [(["@PIS", "@PROV", "@13"], [])]),
    "CPRB_REC": ("P", "CPRB a Recolher", [(["@INSS", "RECEITA BRUTA"], []), (["CPRB"], []),
                 (["CONTRIBUICAO PREVIDENCIARIA", "RECEITA"], [])]),
    # ---- Ativo
    "ADIANT_SAL": ("A", "Adiantamento de Salário", [(["@ADIANT", "@SAL"], ["@13", "@FER", "FORNECEDOR*",
                   "CLIENTE*"]), (["@ADIANT", "EMPREGADO*"], ["@13", "@FER"]),
                   (["@ADIANT", "FUNCIONARIO*"], ["@13", "@FER"])]),
    "ADIANT_13": ("A", "Adiantamento de 13º", [(["@ADIANT", "@13"], [])]),
    "ADIANT_FER": ("A", "Adiantamento de Férias", [(["@ADIANT", "@FER"], [])]),
    "INSS_COMP": ("A", "INSS a Compensar", [(["@INSS", "COMPENSAR"], ["RECEITA BRUTA"]),
                  (["@INSS", "RECUPERAR"], [])]),
    "BENEF_INSS": ("A", "Sal.-Família/Maternidade a Compensar", [(["MATERN*", "COMPENSAR"], []),
                   (["FAMILIA", "COMPENSAR"], []), (["MATERN*", "RECUPERAR"], []),
                   (["FAMILIA", "RECUPERAR"], []), (["@INSS", "COMPENSAR"], ["RECEITA BRUTA", "RETIDO"]),
                   (["@INSS", "RECUPERAR"], [])]),
    # ---- Dedução da receita
    "CPRB_DED": ("R", "(-) CPRB / INSS s/ Receita Bruta", [(["@INSS", "RECEITA BRUTA"], ["@RECOLHER"]),
                 (["CPRB"], ["@RECOLHER"]), (["CONTRIBUICAO PREVIDENCIARIA", "RECEITA"], ["@RECOLHER"])]),
    # ---- Despesas fora do grupo de pessoal
    "SIND_PAT": (GD, "Contribuição Sindical Patronal", [(["@SINDIC", "PATRONAL"], []),
                 (["CONTRIBUICAO SINDICAL"], ["@RECOLHER"])]),
    "TAXAS_DIV": ("D", "Taxas e Contribuições Diversas", [(["TAXAS DIVERSAS"], []),
                  (["CONTRIBUICOES DIVERSAS"], []), (["TAXAS*"], [])]),
}

# Contas de obrigação/direito com o colaborador: itens patronais nunca podem usá-las
ALVOS_COLABORADOR = ("SAL_PAGAR", "PROLAB_PAGAR", "FER_PAGAR", "RESC_PAGAR",
                     "ADIANT_SAL", "ADIANT_13", "ADIANT_FER", "EMPREST", "PENSAO")
OBRIGACOES_GUIA = ("INSS_REC", "FGTS_REC", "IRRF_REC", "PIS_REC", "SIND_REC", "CPRB_REC")


def casa_alvo(d, alvo):
    return any(all(casa(d, x) for x in m) and not any(casa(d, x) for x in n)
               for m, n in ALVOS[alvo][2])

# =====================================================================
# 3. PLANO DE CONTAS GENÉRICO
# =====================================================================
def _segs(c) -> tuple:
    return tuple(p for p in re.split(r"\D+", str(c)) if p)


def dentro(segs, grupo) -> bool:
    """segs é o próprio grupo ou descendente dele (com ou sem pontos na máscara)."""
    if not grupo:
        return False
    if segs == grupo:
        return True
    if len(grupo) == 1:
        alvo = segs[0] if len(segs) == 1 else "".join(segs)
        return alvo.startswith(grupo[0]) and alvo != grupo[0]
    return len(segs) > len(grupo) and segs[:len(grupo)] == grupo


def _prefixos(segs):
    if len(segs) == 1:
        s = segs[0]
        return [(s[:k],) for k in range(1, len(s))]
    return [segs[:k] for k in range(1, len(segs))]


@st.cache_data(show_spinner=False)
def carregar_plano(dados: bytes, nome: str):
    if nome.lower().endswith(".csv"):
        df = pd.read_csv(io.BytesIO(dados), sep=None, engine="python", dtype=str)
    else:
        df = pd.read_excel(io.BytesIO(dados), dtype=str)
    ren = {}
    for c in df.columns:
        n = norm(c)
        if "REDUZ" in n: ren[c] = "reduzido"
        elif "CLASSIF" in n: ren[c] = "classificacao"
        elif n.startswith("TIPO"): ren[c] = "tipo"
        elif "DESCRI" in n or n in ("NOME", "NOME DA CONTA"): ren[c] = "descricao"
    df = df.rename(columns=ren)
    falta = {"reduzido", "classificacao", "descricao"} - set(df.columns)
    if falta:
        raise ValueError(f"Colunas não encontradas no plano: {falta}")
    cols = ["reduzido", "classificacao", "descricao"] + (["tipo"] if "tipo" in df.columns else [])
    df = df[cols].dropna(subset=["reduzido", "classificacao"]).copy()
    for c in cols:
        df[c] = df[c].fillna("").astype(str).str.strip()
    df["reduzido"] = df["reduzido"].str.replace(r"\.0$", "", regex=True)
    df["classificacao"] = df["classificacao"].map(
        lambda c: c[:-2] if re.fullmatch(r"\d+\.0", c) else c)
    df = df.drop_duplicates(subset=["reduzido", "classificacao"])
    df["segs"] = df["classificacao"].map(_segs)
    df = df[df["segs"].map(len) > 0].reset_index(drop=True)

    pos = {s: i for i, s in enumerate(df["segs"])}
    df["anc"] = [[pos[p] for p in _prefixos(s) if p in pos] for s in df["segs"]]  # raiz primeiro
    pais = {j for a in df["anc"] for j in a}
    inferido = ["S" if i in pais else "A" for i in range(len(df))]
    if "tipo" in df.columns:
        t = df["tipo"].str.upper().str[:1]
        df["tipo"] = [x if x in ("A", "S") else inf for x, inf in zip(t, inferido)]
    else:
        df["tipo"] = inferido
    df["desc_norm"] = df["descricao"].map(norm)
    return df


NAT_ROT = {"A": "Ativo", "P": "Passivo/PL", "R": "Receita", "D": "Custo/Despesa",
           "X": "Ignorar (apuração/compensação)", "": "(decidir no nível abaixo)"}
ROT_NAT = {v: k for k, v in NAT_ROT.items()}


def natureza_desc(d) -> str:
    if tem(d, "APURAC*", "ENCERRAMENTO*", "CONTAS DE COMPENSACAO", "COMPENSACAO ATIVA",
           "COMPENSACAO PASSIVA"):
        return "X"
    if tem(d, "PASSIVO", "PATRIMONIO LIQUIDO"):
        return "P"
    if tem(d, "ATIVO"):
        return "A"
    r, dsp = tem(d, "RECEITA*"), tem(d, "CUSTO*", "DESPESA*")
    if r and not dsp: return "R"
    if dsp and not r: return "D"
    return ""


def tabela_naturezas(plano, salvo):
    """Grupos de 1º e 2º nível com a natureza sugerida (ou a salva no perfil)."""
    auto = plano["desc_norm"].map(natureza_desc)
    linhas = []
    for r in [i for i, a in enumerate(plano["anc"]) if not a]:
        cr = plano.at[r, "classificacao"]
        vr = salvo.get(cr, auto[r])
        linhas.append({"Classificação": cr, "Descrição": plano.at[r, "descricao"], "Nível": 1,
                       "Natureza": NAT_ROT.get(vr, NAT_ROT[""])})
        for i, a in enumerate(plano["anc"]):
            if len(a) == 1 and a[0] == r and plano.at[i, "tipo"] == "S":
                ci = plano.at[i, "classificacao"]
                vi = salvo.get(ci, "" if auto[r] else auto[i])
                linhas.append({"Classificação": ci, "Descrição": plano.at[i, "descricao"], "Nível": 2,
                               "Natureza": NAT_ROT.get(vi, NAT_ROT[""])})
    return pd.DataFrame(linhas)


def aplicar_naturezas(plano, ov):
    """Níveis 1-2: vence o mais profundo definido na tabela. Abaixo disso: detecção automática."""
    auto = plano["desc_norm"].map(natureza_desc).tolist()
    cls = plano["classificacao"].tolist()
    out = []
    for i, anc in enumerate(plano["anc"]):
        cadeia = list(anc) + [i]
        v = next((ov[cls[j]] for j in reversed(cadeia[:2]) if ov.get(cls[j])), "")
        if not v:
            v = next((auto[j] for j in cadeia[2:] if auto[j]), "")
        out.append(v or "?")
    return out


class Resolvedor:
    def __init__(self, plano):
        self.an = plano[plano["tipo"] == "A"]
        self.idx = plano.drop_duplicates("reduzido").set_index("reduzido")
        self.ov = {}
        self.cache = {}

    @staticmethod
    def chave(alvo, grupo):
        return f"{alvo}@{grupo if ALVOS[alvo][0] in (G, GD) else '*'}"

    def candidatos(self, alvo, grupo):
        escopo, _, alts = ALVOS[alvo]
        bases = []
        if escopo in (G, GD) and grupo:
            gs = _segs(grupo)
            bases.append(self.an[self.an["segs"].apply(lambda s: dentro(s, gs))])
        if escopo == GD:
            bases.append(self.an[self.an["nat"] == "D"])
        if escopo in ("A", "P", "R", "D"):
            bases.append(self.an[self.an["nat"] == escopo])
        for base in bases:
            for must, nots in alts:
                ok = base["desc_norm"].apply(
                    lambda d: all(casa(d, x) for x in must) and not any(casa(d, x) for x in nots))
                c = base[ok]
                if not c.empty:
                    return c.assign(_n=c["desc_norm"].str.len()).sort_values(["_n", "classificacao"])
        return self.an.iloc[0:0]

    def sugestao(self, alvo, grupo):
        k = (alvo, grupo)
        if k not in self.cache:
            c = self.candidatos(alvo, grupo)
            self.cache[k] = ((c.iloc[0]["reduzido"], c.iloc[0]["descricao"], len(c))
                             if not c.empty else ("", "", 0))
        return self.cache[k]

    def valida(self, cod):
        return bool(cod) and cod in self.idx.index and self.idx.loc[cod, "tipo"] == "A"

    def conta(self, alvo, grupo):
        k = self.chave(alvo, grupo)
        if k in self.ov:  # a tabela de contas-chave governa o que é usado
            v = self.ov[k]
            return (v, self.idx.loc[v, "descricao"]) if self.valida(v) else None
        cod, desc, _ = self.sugestao(alvo, grupo)
        return (cod, desc) if cod else None

    def no_escopo(self, cod, alvo, grupo):
        esc, row = ALVOS[alvo][0], self.idx.loc[cod]
        if esc in (G, GD):
            return dentro(row["segs"], _segs(grupo or "")) or (esc == GD and row["nat"] == "D")
        return row["nat"] == esc

    def resolver(self, alvos, grupo):
        if not alvos:
            return "", "", ""
        for a in alvos:
            r = self.conta(a, grupo)
            if r:
                aviso = "" if a == alvos[0] else f"alternativa: {ALVOS[alvos[0]][1]} → {ALVOS[a][1]}"
                return r[0], r[1], aviso
        return "", "", f"❌ sem conta '{ALVOS[alvos[0]][1]}'"


def grupos_resultado(plano):
    """Grupo imediato de cada conta analítica de salários com natureza Custo/Despesa."""
    out, vistos = [], set()
    desc = plano["desc_norm"].tolist()
    for i in range(len(plano)):
        if plano.at[i, "tipo"] != "A" or plano.at[i, "nat"] != "D" or not casa_alvo(desc[i], "SALARIOS"):
            continue
        anc = plano.at[i, "anc"]
        if not anc or anc[-1] in vistos:
            continue
        g = anc[-1]
        vistos.add(g)
        out.append({
            "mascara": plano.at[g, "classificacao"], "descricao": plano.at[g, "descricao"],
            "pai": plano.at[anc[-2], "descricao"] if len(anc) >= 2 else "",
            "cadeia": " ".join(desc[j] for j in list(plano.at[g, "anc"]) + [g]),
        })
    return sorted(out, key=lambda c: c["mascara"])


def sugerir_grupo(nome_sep, cands):
    n = norm(nome_sep)
    if tem(n, "VENDA*", "COMERCI*", "LOJA*", "MARKETING"): chave = "VENDA*"
    elif tem(n, "PRODU*", "FABRI*", "INDUSTR*", "OBRA*", "OPERAC*", "MANUTENC*"): chave = "CUSTO*"
    else: chave = "ADMINISTRATIV*"
    return next((i for i, c in enumerate(cands) if tem(c["cadeia"], chave)), 0)


def existe_analitica(plano, mascara):
    gs = _segs(mascara)
    return bool(gs) and plano[plano["tipo"] == "A"]["segs"].apply(lambda s: dentro(s, gs)).any()

# =====================================================================
# 4. LEITURA DOS RELATÓRIOS DA FOLHA
# =====================================================================
def extrair_linhas(dados: bytes, nome: str):
    if nome.lower().endswith(".pdf"):
        if pdfplumber is None:
            raise RuntimeError("Instale o pdfplumber: pip install pdfplumber")
        linhas = []
        with pdfplumber.open(io.BytesIO(dados)) as pdf:
            for p in pdf.pages:
                linhas.extend((p.extract_text() or "").splitlines())
        return linhas
    for enc in ("utf-8", "cp1252", "latin-1"):
        try:
            return dados.decode(enc).splitlines()
        except UnicodeDecodeError:
            continue


RE_RUB = re.compile(r"^\s*(\d{1,5})\s+(.*?)\s*"
                    r"(Provento|Desconto|Informativa|Informat|Inf\.\s*dedutora|Inf\.\s*ded)")

def _tipo_cad(t):
    t = t.lower()
    if t.startswith("prov"): return "Provento"
    if t.startswith("desc"): return "Desconto"
    if t.startswith("inf."): return "Inf. dedutora"
    return "Informativa"


@st.cache_data(show_spinner=False)
def parse_cadastro(dados: bytes, nome: str):
    linhas = extrair_linhas(dados, nome)
    cab = next((re.split(r"P[áa]gina", l, flags=re.I)[0].strip() for l in linhas if l.strip()), "")
    cad = {}
    for l in linhas:
        m = RE_RUB.match(l)
        if m and int(m.group(1)) not in cad:
            desc = m.group(2).strip()
            cad[int(m.group(1))] = {"descricao": desc, "desc_norm": norm(desc), "tipo": _tipo_cad(m.group(3))}
    return cab, cad


RE_EMPRESA = re.compile(r"^\s*Empresa\s*:\s*(\d+)\s*-\s*(.+)$", re.I)
RE_SEP = re.compile(r"^\s*(Centro de Custo|Filial|Servi[çc]o)\s*:\s*(\d*)\s*(.*)$", re.I)
RE_ITEM = re.compile(r"^\s*(\d{1,5})\s+(\S.*)$")
IGNORAR = ("PAGINA", "EMISSAO", "HORA ", "RELACAO DE RUBRICAS", "CODIGO DESCRICAO")


@st.cache_data(show_spinner=False)
def parse_pendencias(dados: bytes, nome: str):
    cab = {"codigo": "", "nome": ""}
    itens, vistos = [], set()
    secao, sep_cod, sep_nome, sep_tipo = None, "", "", ""
    for bruta in extrair_linhas(dados, nome):
        txt = bruta.strip()
        if not txt:
            continue
        n = norm(txt)
        m = RE_EMPRESA.match(txt)
        if m:
            cab["codigo"], cab["nome"] = m.group(1), m.group(2).strip()
            continue
        if n in SECOES:
            if SECOES[n] != secao:
                secao, sep_cod, sep_nome = SECOES[n], "", ""
            continue
        m = RE_SEP.match(txt)
        if m:
            sep_tipo, sep_cod, sep_nome = m.group(1), m.group(2).strip(), m.group(3).strip()
            continue
        if n.startswith(IGNORAR):
            continue
        m = RE_ITEM.match(txt)
        if m and secao:
            chave = (secao, sep_cod, int(m.group(1)))
            if chave not in vistos:
                vistos.add(chave)
                itens.append({"secao": secao, "sep_tipo": sep_tipo, "sep_cod": sep_cod,
                              "sep_nome": sep_nome, "codigo": int(m.group(1)),
                              "descricao": m.group(2).strip()})
    return cab, itens

# =====================================================================
# 5. REGRAS DETERMINÍSTICAS (devolvem conceitos, nunca contas)
# =====================================================================
def inferir_tipo(d):
    if tem(d, "MATERN*") and not tem(d, "DESC*"): return "Provento"
    if tem(d, "INSS") and tem(d, "A MAIOR") and not tem(d, "DESCONTO"): return "Provento"
    if tem(d, "REEMBOLSO", "DEV", "DEVOLUCAO", "RESTITUI*"): return "Provento"
    if tem(d, "FGTS", "CONTRIBUICAO SOCIAL", "CONTRIB SOCIAL", "BASE"): return "Informativa"
    if tem(d, "DESC*") or (tem(d, "ESTOURO", "TROCO") and tem(d, "ANTERIOR")): return "Desconto"
    if tem(d, "INSS", "IRRF", "IMPOSTO DE RENDA", "PENSAO", "EMPREST*", "CONTRIB*",
           "MENSALIDADE", "FALTA*", "ATRASO*"): return "Desconto"
    if tem(d, "VALE TRANSPORTE") and "%" in d: return "Desconto"
    return "Provento"


def tipo_rubrica(codigo, d, cad):
    reg = cad.get(codigo)
    if reg:
        a, b = d, reg["desc_norm"]
        if (SequenceMatcher(None, a, b).ratio() >= LIMIAR_SIMILARIDADE
                or (min(len(a), len(b)) >= 8 and (a.startswith(b) or b.startswith(a)))):
            return reg["tipo"], "Cadastro"
        return inferir_tipo(d), f"Inferido — no cadastro o cód. {codigo} é '{reg['descricao']}'"
    return inferir_tipo(d), "Inferido — código ausente no cadastro"


def e_socio(d):
    """Pró-labore e rubricas 'EMPREGADOR' (INSS/IRRF/troco do sócio)."""
    if tem(d, "PRO LAB*", "PROLAB*"):
        return True
    return tem(d, "EMPREGADOR") and tem(d, "INSS", "IRRF", "TROCO", "ESTOURO", "ADTO", "ADIANT*")


def passivo_secao(secao, d):
    if secao == "Férias": return "FER_PAGAR"
    if secao == "Rescisão": return "RESC_PAGAR"
    if e_socio(d): return "PROLAB_PAGAR"
    return "SAL_PAGAR"


def alvo_adiant(d):
    if e13(d): return "ADIANT_13"
    if tem(d, "FERIAS"): return "ADIANT_FER"
    return "ADIANT_SAL"


def e_licenca_remunerada(d):
    return tem(d, "LICENC*", "LIC") and tem(d, "REMUN*", "REM")


def alvo_dre_provento(d):
    if tem(d, "PRO LAB*", "PROLAB*"): return ["PRO_LABORE"]
    if tem(d, "BOLSA", "ESTAGI*", "RECESSO"): return ["BOLSA", "SALARIOS"]
    if e_licenca_remunerada(d) and not tem(d, "GOZ*"): return ["SALARIOS"]
    if tem(d, "HORAS EXTRA*", "HORA EXTRA*", "EXTRAS", "BANCO DE HORAS"): return ["HE", "SALARIOS"]
    if e13(d): return ["DECIMO"]
    if tem(d, "FERIAS", "ABONO"): return ["FERIAS"]
    if tem(d, "AVISO PREVIO", "INDENIZ*", "MULTA", "ESTABILIDADE"): return ["INDENIZ", "SALARIOS"]
    if tem(d, "PREMIO*", "GRATIFIC*", "BONUS", "PLR", "PARTICIPACAO LUCRO*", "PARTIC LUCRO*"):
        return ["PREMIOS", "SALARIOS"]
    if tem(d, "COMISS*"): return ["COMISSOES", "PREMIOS", "SALARIOS"]
    if tem(d, "VALE TRANSPORTE", "VT"): return ["VT"]
    if tem(d, "PLANO DE SAUDE", "PLANO SAUDE", "ODONTO*", "ASSISTENCIA MEDICA"): return ["ASSIST"]
    if tem(d, "REFEICAO", "ALIMENTACAO", "CESTA", "VR"): return ["ALIM"]
    if tem(d, "SEGURO*"): return ["SEGURO", "ASSIST"]
    if tem(d, "TREINAMENTO", "CURSO*"): return ["TREIN"]
    return ["SALARIOS"]


def classificar_provento(d, secao, baixa_prov):
    L = passivo_secao(secao, d)
    if tem(d, "ESTOURO", "TROCO"):
        return ["ADIANT_SAL"], [L], "Estouro/troco: crédito a receber do empregado", False
    if tem(d, "ADIANT*", "ADTO*"):
        return [alvo_adiant(d)], [L], "Adiantamento pago: Ativo × obrigação", False
    if tem(d, "DEV", "DEVOLUCAO", "ESTORNO") and tem(d, "EMPREST*", "CONSIG*"):
        return ["EMPREST"], [L], "Devolução/estorno de consignado", True
    if tem(d, "SALARIO FAMILIA", "SAL FAM*"):
        if tem(d, "SEM COMPENSAC*"):
            return ["SALARIOS"], [L], "Salário-família sem compensação na guia: custo", True
        return ["BENEF_INSS"], [L], "Salário-família → conta-ponte até a compensação na guia", False
    if tem(d, "MATERN*") and not tem(d, "DESC*", "DEDUC*"):
        if tem(d, "INSS"):
            return [], [], "Maternidade paga direto pelo INSS — não integrar", True
        if not tem(d, "PRORROG*", "EMPREGADOR"):
            return ["BENEF_INSS"], [L], "Salário-maternidade reembolsável → conta-ponte", False
    if tem(d, "INSS", "IRRF") and tem(d, "A MAIOR", "DEVOL*", "RESTITUI*"):
        return ["IRRF_REC" if tem(d, "IRRF") else "INSS_REC"], [L], "Restituição de retenção a maior", False
    dre = alvo_dre_provento(d)
    if baixa_prov and secao in ("Férias", "Rescisão", "13º Salário") and dre[0] in ("FERIAS", "DECIMO"):
        return (["P_PROV_FER"] if dre[0] == "FERIAS" else ["P_PROV_13"]), [L], "Baixa contra a provisão", False
    return dre, [L], "", False


def classificar_desconto(d, secao):
    L = passivo_secao(secao, d)
    def r(c, obs="", rev=False): return [L], c, obs, rev
    if tem(d, "ESTOURO", "TROCO"): return r(["ADIANT_SAL"], "Baixa de estouro/troco anterior")
    if tem(d, "PENSAO"): return r(["PENSAO"], "Repasse ao beneficiário", True)
    if tem(d, "INSS"): return r(["INSS_REC"])
    if tem(d, "IRRF", "IMPOSTO DE RENDA"): return r(["IRRF_REC"])
    if tem(d, "SAL FAM*", "SALARIO FAMILIA"):
        return r(["BENEF_INSS"], "Estorno de salário-família pago a maior", True)
    if tem(d, "ADIANT*", "ADTO*", "FERIAS PAGAS"): return r([alvo_adiant(d)], "Baixa do adiantamento")
    if tem(d, "SINDICA*", "ASSISTENCIAL", "CONFEDERATIVA", "NEGOC*"): return r(["SIND_REC"])
    if tem(d, "EMPREST*", "EMP", "CONSIG*", "CRED TRAB"): return r(["EMPREST"], "Consignado a repassar", True)
    if tem(d, "PLANO", "ODONTO*", "COPARTICIP*", "ASSISTENCIA MEDICA", "FARMACIA", "SAUDE"):
        return r(["ASSIST"], "Recuperação do custo do benefício")
    if tem(d, "VALE TRANSPORTE", "VT"): return r(["VT"], "Recuperação dos 6% do VT")
    if tem(d, "REFEICAO", "ALIMENTACAO", "VR", "CESTA", "SUPERMERCADO"):
        return r(["ALIM"], "Recuperação do custo do benefício")
    if tem(d, "SEGURO*"): return r(["SEGURO", "ASSIST"], "Recuperação do custo do benefício")
    if tem(d, "AVISO PREVIO", "MULTA", "ESTABILIDADE", "INDENIZ*"):
        return r(["INDENIZ"], "Indenização devida pelo empregado", True)
    if e13(d) and tem(d, "AFAST*", "MATERN*"): return r(["DECIMO"], "Redução do custo de 13º")
    if tem(d, "FALTA*", "ATRASO*", "DSR", "HORAS", "PAGO A MAIOR", "INATIVAS"):
        return r(["SALARIOS"], "Redução do custo de salários")
    return r([], "Desconto sem regra — definir conta manualmente", True)


def classificar_informativa(d, tipo):
    if tem(d, "BASE", "E SOCIAL", "INFORMATIVO", "HORAS CREDITO", "HORAS COMPENSADA",
           "BANCO DE HORAS", "ADIANT*", "DEMONSTR*", "DEPENDENTE"):
        return [], [], "Informativa de base/controle — não integrar", True
    if tem(d, "FGTS", "CONTRIBUICAO SOCIAL", "CONTRIB SOCIAL"):
        dre = ["INDENIZ", "FGTS"] if tem(d, "40%", "20%") else ["FGTS"]
        if tem(d, "A MAIOR") or tipo == "Inf. dedutora":
            return ["FGTS_REC"], dre, "Estorno de FGTS a maior", True
        return dre, ["FGTS_REC"], "Encargo FGTS", False
    if tem(d, "INSS"): return ["INSS"], ["INSS_REC"], "Encargo INSS", False
    if tem(d, "PIS"): return ["PIS"], ["PIS_REC"], "Encargo PIS s/ folha", False
    return [], [], "Informativa não reconhecida", True


_BASE_PROV = {"D_PROV_FER": "FERIAS", "D_INSS_FER": "INSS", "D_FGTS_FER": "FGTS",
              "D_PROV_13": "DECIMO", "D_INSS_13": "INSS", "D_FGTS_13": "FGTS"}


def classificar_item(d, secao):
    """Informações INSS e Provisões."""
    if secao == "Informações INSS":
        if tem(d, "RETENC*", "RETIDO"):
            return ["INSS_REC"], ["INSS_COMP"], "Retenção s/ NF compensada na guia", True
        return (["INSS_REC"], ["BENEF_INSS"], "Compensação na guia — baixa da conta-ponte",
                not tem(d, "FAMILIA", "MATERN*"))
    sfx = "FER" if "Férias" in secao else "13"
    if tem(d, "FGTS"): dre, pas = f"D_FGTS_{sfx}", f"P_FGTS_{sfx}"
    elif tem(d, "INSS", "RAT", "TERCEIROS", "FAP"): dre, pas = f"D_INSS_{sfx}", f"P_INSS_{sfx}"
    elif tem(d, "PIS"): dre, pas = "PIS", f"P_PIS_{sfx}"
    else: dre, pas = f"D_PROV_{sfx}", f"P_PROV_{sfx}"
    dres = [dre] + ([_BASE_PROV[dre]] if dre in _BASE_PROV else [])
    if tem(d, "ESTORNO*", "BAIXA*", "REVERS*"):
        return [pas], dres, "Estorno da provisão", False
    return dres, [pas], "Constituição da provisão", False


OBS_EMPRESA = {
    "INSS": ("Encargo patronal INSS", False),
    "INSS_SOCIO": ("INSS patronal s/ pró-labore", False),
    "AUTONOMO": ("INSS patronal s/ autônomo — confirme a conta de despesa", True),
    "SENAI": ("Adicional ao SENAI — confirme a forma de recolhimento", True),
    "PIS": ("PIS s/ folha", False),
    "FGTS": ("Encargo FGTS", False),
    "CPRB": ("CPRB — dedução da receita bruta, fora do grupo de pessoal", False),
    "SIND_PATRONAL": ("Contribuição sindical patronal (GRCS)", True),
    "DEDUCAO": ("Compensação na guia — baixa da conta-ponte", False),
    "ISENCAO": ("Isenção filantropia: valor não devido — não integrar", True),
    None: ("Item patronal sem regra — definir manualmente", True),
}


def inferir_item_empresa(d):
    if tem(d, "ISENCAO*", "FILANTROP*"): nat = "ISENCAO"
    elif tem(d, "DEDUCAO*", "COMPENSACAO*"): nat = "DEDUCAO"
    elif tem(d, "RECEITA BRUTA"): nat = "CPRB"
    elif tem(d, "GRCS", "SINDICA*"): nat = "SIND_PATRONAL"
    elif tem(d, "SENAI"): nat = "SENAI"
    elif tem(d, "PIS"): nat = "PIS"
    elif tem(d, "FGTS"): nat = "FGTS"
    elif tem(d, "PRO LAB*"): nat = "INSS_SOCIO"
    elif tem(d, "AUT", "AUTONOMO*"): nat = "AUTONOMO"
    elif tem(d, "INSS", "RAT", "SAT", "FAP", "TERCEIROS", "ACID*"): nat = "INSS"
    else: nat = None
    return nat, ("13" if e13(d) else "F" if tem(d, "FERIAS") else "M")


def classificar_item_empresa(codigo, d, baixa_prov):
    cat = ITENS_EMPRESA.get(codigo)
    if cat and SequenceMatcher(None, d, norm(cat[0])).ratio() >= LIMIAR_SIMILARIDADE:
        _, nat, comp = cat
        origem = "Catálogo Empresa"
    else:
        nat, comp = inferir_item_empresa(d)
        origem = (f"Inferido — no catálogo o item {codigo} é '{cat[0]}'" if cat
                  else f"Inferido — item {codigo} fora do catálogo")
    sfx = {"F": "FER", "13": "13"}.get(comp)
    pela_prov = bool(baixa_prov and sfx and nat in ("INSS", "SENAI", "PIS"))
    if nat in ("INSS", "INSS_SOCIO", "AUTONOMO", "SENAI"):
        deb, cred = ([f"P_INSS_{sfx}"] if pela_prov else ["INSS"]), ["INSS_REC"]
    elif nat == "PIS":
        deb, cred = ([f"P_PIS_{sfx}"] if pela_prov else ["PIS"]), ["PIS_REC"]
    elif nat == "FGTS":
        deb, cred = ["FGTS"], ["FGTS_REC"]
    elif nat == "CPRB":
        deb, cred = ["CPRB_DED"], ["CPRB_REC"]
    elif nat == "SIND_PATRONAL":
        deb, cred = ["SIND_PAT", "TAXAS_DIV"], ["SIND_REC"]
    elif nat == "DEDUCAO":
        deb, cred = ["INSS_REC"], ["BENEF_INSS"]
    else:
        deb, cred = [], []
    obs, rev = OBS_EMPRESA.get(nat, OBS_EMPRESA[None])
    if pela_prov:
        obs += " — baixa contra a provisão"
    return deb, cred, obs, rev, origem, nat


def classificar(it, cad, baixa_prov):
    d, sec, nat_item = norm(it["descricao"]), it["secao"], None
    if sec == "Empresa":
        tipo = "Item (Empresa)"
        deb, cred, obs, rev, origem, nat_item = classificar_item_empresa(it["codigo"], d, baixa_prov)
    elif sec in SECOES_DE_ITENS:
        tipo, origem = f"Item ({sec})", "Seção"
        deb, cred, obs, rev = classificar_item(d, sec)
    else:
        tipo, origem = tipo_rubrica(it["codigo"], d, cad)
        if tipo in ("Informativa", "Inf. dedutora"):
            deb, cred, obs, rev = classificar_informativa(d, tipo)
        elif tipo == "Desconto":
            deb, cred, obs, rev = classificar_desconto(d, sec)
        else:
            deb, cred, obs, rev = classificar_provento(d, sec, baixa_prov)
    return {"d": d, "tipo": tipo, "origem": origem, "deb": deb, "cred": cred,
            "obs": obs, "rev": rev, "nat_item": nat_item}

# =====================================================================
# 6. PERFIL DA EMPRESA (confirmações salvas em JSON)
# =====================================================================
def caminho_perfil(cod):
    nome = re.sub(r"\W", "_", str(cod)) or "sem_codigo"
    return PASTA_PERFIS / f"empresa_{nome}.json"


def carregar_perfil(cod):
    p = caminho_perfil(cod)
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def salvar_perfil(cod, dados):
    PASTA_PERFIS.mkdir(exist_ok=True)
    caminho_perfil(cod).write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")


def limpa(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return re.sub(r"\.0$", "", str(v).strip())

# =====================================================================
# 7. INTERFACE
# =====================================================================
st.set_page_config(page_title="Integrador Contábil da Folha", layout="wide")
st.title("📒 Integrador Contábil da Folha — Domínio")
st.caption("Motor 100% determinístico. Nenhum plano de contas fica fixo no programa: "
           "a estrutura é lida do arquivo importado e confirmada por você.")

with st.sidebar:
    st.header("1. Arquivos")
    f_pend = st.file_uploader("Rubricas/Itens não configurados", type=["pdf", "txt"])
    f_cad = st.file_uploader("Cadastro geral de rubricas", type=["pdf", "txt"])
    f_plano = st.file_uploader("Plano de contas", type=["xlsx", "xls", "csv"])
    st.header("2. Parâmetros")
    historico = st.text_input("Código do histórico padrão", "")
    baixa_prov = st.checkbox("Baixar férias/13º pagos contra a provisão", False,
                             help="Deixe desmarcado se a Domínio já gera o 'Valor Estorno Provisão'.")
    socio_adm = st.checkbox("Pró-labore e encargos do sócio sempre em Despesas Administrativas", True)
    st.header("3. Perfil da empresa")
    f_perfil = st.file_uploader("Importar perfil (.json)", type=["json"])

if not (f_pend and f_cad and f_plano):
    st.info("Envie os três arquivos para começar.")
    st.stop()

try:
    cab, itens = parse_pendencias(f_pend.getvalue(), f_pend.name)
    nome_cad, cad = parse_cadastro(f_cad.getvalue(), f_cad.name)
    plano = carregar_plano(f_plano.getvalue(), f_plano.name).copy()
except Exception as e:
    st.error(f"Erro na leitura: {e}")
    st.stop()

if not itens:
    st.error("Nenhuma rubrica/item encontrado no relatório de pendências.")
    st.stop()

cod_empresa = st.sidebar.text_input("Código da empresa na Domínio", cab["codigo"])
perfil = carregar_perfil(cod_empresa)
if f_perfil is not None:
    try:
        perfil = json.loads(f_perfil.getvalue().decode("utf-8"))
    except Exception:
        st.sidebar.error("Arquivo de perfil inválido.")
if perfil:
    st.sidebar.success("Perfil da empresa carregado.")

m1, m2, m3, m4 = st.columns(4)
m1.metric("Itens pendentes", len(itens))
m2.metric("Rubricas no cadastro", len(cad))
m3.metric("Contas no plano", len(plano))
m4.metric("Empresa", f'{cab["codigo"]} - {cab["nome"][:25]}')

dup = plano[plano.duplicated("reduzido", keep=False)]
if not dup.empty:
    st.warning(f"O plano tem {dup['reduzido'].nunique()} código(s) reduzido(s) repetido(s) "
               "com classificações diferentes. Corrija na origem antes de integrar.")

# ---------- 1. Estrutura do plano ----------
st.subheader("1. Estrutura do plano de contas")
tab = tabela_naturezas(plano, perfil.get("naturezas", {}))
with st.expander("Natureza dos grupos — detectada pela descrição, confirme",
                 expanded=not perfil.get("naturezas")):
    tab_ed = st.data_editor(
        tab, key=f"nat_{cod_empresa}", hide_index=True, use_container_width=True,
        disabled=["Classificação", "Descrição", "Nível"],
        column_config={"Natureza": st.column_config.SelectboxColumn(
            "Natureza", options=list(NAT_ROT.values()), required=True)})
nat_ov = dict(zip(tab_ed["Classificação"], tab_ed["Natureza"].map(ROT_NAT).fillna("")))
plano["nat"] = aplicar_naturezas(plano, nat_ov)

cont = plano[plano["tipo"] == "A"]["nat"].value_counts()
cs = st.columns(5)
for col, (k, rot) in zip(cs, [("A", "Ativo"), ("P", "Passivo/PL"), ("R", "Receita"),
                               ("D", "Custo/Despesa"), ("?", "Sem natureza")]):
    col.metric(rot, int(cont.get(k, 0)))
if cont.get("?", 0):
    st.warning("Há contas analíticas sem natureza. Defina-as na tabela acima.")
if not cont.get("D", 0) or not cont.get("P", 0):
    st.error("O plano precisa ter contas de Passivo e de Custo/Despesa. Revise a tabela de natureza.")
    st.stop()

# ---------- 2. Separador → grupo ----------
st.subheader("2. Separador → grupo de resultado")
usa_sep = any(i["sep_cod"] for i in itens)
precisa_geral = (not usa_sep) or any(not i["sep_cod"] for i in itens)
nome_geral = NOME_PADRAO_SEM_SEPARADOR
if precisa_geral:
    nome_geral = (st.text_input("Nome do lote sem separador",
                                perfil.get("nome_geral", NOME_PADRAO_SEM_SEPARADOR))
                  or NOME_PADRAO_SEM_SEPARADOR)

cands = grupos_resultado(plano)
rotulos = [f"{c['mascara']} — {c['descricao']}" + (f" ({c['pai']})" if c["pai"] else "") for c in cands]
mascaras = [c["mascara"] for c in cands]
if usa_sep:
    st.success("Folha com separador detectada (Centro de Custo / Filial / Serviço).")
else:
    st.warning("Nenhuma quebra por Centro de Custo / Filial / Serviço → folha centralizada. "
               "Escolha o grupo contábil em que a folha inteira será classificada.")
if not cands:
    st.info("Nenhum grupo de custo/despesa com conta de salários foi encontrado: informe a máscara manualmente.")

seps = {}
for it in itens:
    k = it["sep_cod"] or nome_geral
    seps.setdefault(k, it["sep_nome"] if it["sep_cod"] else "itens sem separador")

sep_salvo = perfil.get("separadores", {})
mapa_grupos = {}
for k, nome in seps.items():
    c1, c2 = st.columns([3, 1])
    salvo = sep_salvo.get(k, "")
    if salvo in mascaras: idx = mascaras.index(salvo)
    elif usa_sep and cands: idx = sugerir_grupo(nome, cands)
    else: idx = None
    escolha = c1.selectbox(f"Separador {k} — {nome}", rotulos, index=idx, key=f"g_{k}",
                           placeholder="Escolha o grupo contábil")
    manual = c2.text_input("ou máscara manual", value="" if salvo in mascaras else salvo,
                           key=f"m_{k}", placeholder="classificação do grupo")
    pref = manual.strip() or (mascaras[rotulos.index(escolha)] if escolha else "")
    if pref and not existe_analitica(plano, pref):
        c2.error("Máscara sem contas analíticas")
        pref = ""
    mapa_grupos[k] = pref

if any(not v for v in mapa_grupos.values()):
    st.info("Defina o grupo de todos os separadores para continuar.")
    st.stop()

grupo_adm = next((c["mascara"] for c in cands if tem(c["cadeia"], "ADMINISTRATIV*")), None)

# ---------- Classificação (conceitos, ainda sem contas) ----------
classif = []
for it in itens:
    r = classificar(it, cad, baixa_prov)
    k = it["sep_cod"] or nome_geral
    grupo = mapa_grupos[k]
    if socio_adm and grupo_adm and e_socio(r["d"]):
        grupo = grupo_adm
    r.update(it=it, sep=k, grupo=grupo)
    classif.append(r)

ne, nc = norm(cab["nome"]), norm(nome_cad)
n_div = sum(r["origem"].startswith("Inferido") for r in classif if r["it"]["secao"] not in SECOES_DE_ITENS)
if ne and nc and ne not in nc and nc not in ne:
    if n_div:
        st.warning(f"⚠️ Cadastro de outra empresa (**{nome_cad}**): {n_div} rubrica(s) com "
                   "código divergente/ausente — Tipo inferido pela descrição.")
    else:
        st.info(f"Cadastro do modelo **{nome_cad}**, mas todas as rubricas conferem por código + descrição.")

# ---------- 3. Contas-chave ----------
st.subheader("3. Contas-chave do plano")
st.caption("Sugestões tiradas da descrição das contas do plano importado. Confirme, digite outro "
           "código reduzido ou deixe em branco se a empresa não tiver a conta.")
res = Resolvedor(plano)
pedidos = {}
def pedir(alvo, grupo):
    g = grupo if ALVOS[alvo][0] in (G, GD) else None
    pedidos.setdefault(Resolvedor.chave(alvo, g), (alvo, g))
for r in classif:
    for a in r["deb"] + r["cred"]:
        pedir(a, r["grupo"])
for a in ALVOS_COLABORADOR:
    pedir(a, None)

contas_salvas = perfil.get("contas", {})
linhas_map = []
for k, (alvo, grupo) in pedidos.items():
    cod, desc, n = res.sugestao(alvo, grupo)
    salvo = contas_salvas.get(k)
    usa_salvo = salvo is not None and (salvo == "" or res.valida(salvo))
    linhas_map.append({"Chave": k, "Conta-chave": ALVOS[alvo][1], "Grupo": grupo or "—",
                       "Sugestão": f"{cod} — {desc}" if cod else "(nenhuma)", "Candidatos": n,
                       "Origem": "Perfil" if usa_salvo else "Sugestão",
                       "Conta": salvo if usa_salvo else cod})
df_map = pd.DataFrame(linhas_map).sort_values(["Grupo", "Conta-chave"]).reset_index(drop=True)
with st.expander("Contas-chave usadas nesta folha", expanded=not contas_salvas):
    ed_map = st.data_editor(
        df_map, key=f"contas_{cod_empresa}_{abs(hash(tuple(df_map['Chave'])))}",
        hide_index=True, use_container_width
