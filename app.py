"""
Integrador Contábil da Folha — Domínio Sistemas
Motor 100% determinístico (regex + regras de substring). Sem IA, sem APIs, sem chaves.

NENHUMA conta, código reduzido ou classificação de plano é fixada no código:
- as raízes (Ativo / Passivo / Custos-Despesas / Receitas) são detectadas no plano importado;
- as contas são localizadas por descrição dentro do escopo correto;
- tudo pode ser sobrescrito no "Mapa de contas" e salvo por empresa (.json).

Executar:
    pip install -r requirements.txt
    streamlit run app.py
"""
import csv
import hashlib
import io
import json
import re
import unicodedata
from difflib import SequenceMatcher
from functools import lru_cache

import pandas as pd
import streamlit as st

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

# =====================================================================
# 0. PARÂMETROS (layout Domínio — não dependem do plano de contas)
# =====================================================================
NOME_PADRAO_SEM_SEPARADOR = "Geral"
LIMIAR_SIMILARIDADE = 0.60

ABA_INTEGRA, ABA_EVENTO = "integra", "evento"
COLS_INTEGRA = ["Código da Empresa", "Separador"]
COLS_EVENTO = ["Código da Empresa", "Separador", "Código Sequencial", "Tipo de Integração",
               "Código da Rubrica", "Descrição da Rubrica", "Conta Débito", "Conta Crédito",
               "Histórico Padrão"]

SECOES = {  # título normalizado no PDF -> nome interno
    "FOLHA NORMAL": "Folha Normal", "FOLHA MENSAL": "Folha Normal",
    "ADIANTAMENTO": "Adiantamento",
    "13O SALARIO": "13º Salário", "13 SALARIO": "13º Salário", "DECIMO TERCEIRO SALARIO": "13º Salário",
    "FERIAS": "Férias", "RESCISAO": "Rescisão", "EMPRESA": "Empresa",
    "PROVISAO DE FERIAS": "Provisão de Férias", "PROVISAO FERIAS": "Provisão de Férias",
    "PROVISAO DE 13O": "Provisão de 13º", "PROVISAO DE 13O SALARIO": "Provisão de 13º",
    "PROVISAO DE 13 SALARIO": "Provisão de 13º", "PROVISAO 13O SALARIO": "Provisão de 13º",
    "INFORMACOES INSS": "Informações INSS", "INFORMACOES DO INSS": "Informações INSS",
}
SECOES_DE_ITENS = {"Empresa", "Provisão de Férias", "Provisão de 13º", "Informações INSS"}
TIPO_INTEGRACAO = {s: s for s in set(SECOES.values())}  # troque por códigos do layout, se houver

# Catálogo do sistema Domínio: Configurar Integração > aba Empresa (não tem o item 37).
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
    if txt is None or (isinstance(txt, float) and pd.isna(txt)):
        return ""
    s = unicodedata.normalize("NFKD", str(txt))
    s = "".join(c for c in s if not unicodedata.combining(c)).upper()
    s = re.sub(r"\b(?:[A-Z]\.){2,}[A-Z]?\.?", lambda m: m.group(0).replace(".", ""), s)  # I.N.S.S -> INSS
    s = re.sub(r"[^A-Z0-9%]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


@lru_cache(maxsize=4096)
def _pat(t):
    if t.endswith("*"):  # prefixo
        return re.compile(r"(?<![A-Z0-9])" + re.escape(t[:-1]))
    return re.compile(r"(?<![A-Z0-9])" + re.escape(t) + r"(?![A-Z0-9%])")


def tem(d: str, *termos) -> bool:
    """Palavra/expressão inteira; termo terminado em '*' = prefixo."""
    return any(_pat(t).search(d) for t in termos)


RE_13 = re.compile(r"(?<![A-Z0-9])13[OA]?(?![A-Z0-9])")
def e13(d): return bool(RE_13.search(d)) or tem(d, "DECIMO*")

# =====================================================================
# 2. PLANO DE CONTAS GENÉRICO
# =====================================================================
CAMPOS_PLANO = {"reduzido": "Código reduzido", "classificacao": "Classificação",
                "descricao": "Descrição", "tipo": "Tipo A/S (opcional)"}
ROTULOS_RAIZ = {"ATIVO": "Raiz(es) do Ativo", "PASSIVO": "Raiz(es) do Passivo",
                "RESULTADO": "Raiz(es) de Custos/Despesas", "RECEITA": "Raiz(es) de Receitas"}


def norm_classif(v) -> str:
    s = str(v).strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    s = re.sub(r"[\s\-/]+", ".", s)
    s = re.sub(r"[^0-9.]", "", s)
    return re.sub(r"\.+", ".", s).strip(".")


def norm_mascara(m, pontuado):
    s = norm_classif(m or "")
    return s if pontuado else s.replace(".", "")


def ancestrais(c, pontuado):
    if pontuado:
        seg = c.split(".")
        return [".".join(seg[:i]) for i in range(1, len(seg))]
    return [c[:i] for i in range(1, len(c))]


def sob(c, p, pontuado) -> bool:
    if not p:
        return False
    if pontuado:
        return c == p or c.startswith(p + ".")
    return c.startswith(p)


@st.cache_data(show_spinner=False)
def ler_tabela_bruta(dados: bytes, nome: str) -> pd.DataFrame:
    if nome.lower().endswith((".csv", ".txt")):
        txt = None
        for enc in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                txt = dados.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        try:
            delim = csv.Sniffer().sniff(txt[:20000], delimiters=";,\t|").delimiter
        except csv.Error:
            delim = ";"
        rows = list(csv.reader(io.StringIO(txt), delimiter=delim))
        larg = max((len(r) for r in rows), default=0)
        df = pd.DataFrame([r + [None] * (larg - len(r)) for r in rows])
    else:
        df = pd.read_excel(io.BytesIO(dados), dtype=str, header=None)
    hdr = 0
    for i in range(min(40, len(df))):
        vals = [norm(x) for x in df.iloc[i].tolist() if x is not None and not pd.isna(x)]
        if any(v.startswith("CLASSIF") for v in vals) and any(v.startswith(("DESCRI", "NOME")) for v in vals):
            hdr = i
            break
    cols = []
    for j, x in enumerate(df.iloc[hdr].tolist()):
        c = str(x).strip() if x is not None and not pd.isna(x) and str(x).strip() else f"Coluna {j + 1}"
        while c in cols:
            c += "_"
        cols.append(c)
    out = df.iloc[hdr + 1:].copy()
    out.columns = cols
    return out.dropna(how="all").reset_index(drop=True)


def auto_colunas(cols):
    m = {}
    for c in cols:
        n = norm(c)
        if "reduzido" not in m and "REDUZ" in n:
            m["reduzido"] = c
        elif "classificacao" not in m and n.startswith("CLASSIF"):
            m["classificacao"] = c
        elif "descricao" not in m and n.startswith(("DESCRI", "NOME")):
            m["descricao"] = c
        elif "tipo" not in m and (n.startswith("TIPO") or n in ("T", "A S")):
            m["tipo"] = c
    if "reduzido" not in m:
        m["reduzido"] = next((c for c in cols if norm(c) in ("CODIGO", "COD", "CONTA", "COD CONTA")), None)
    return m


@st.cache_data(show_spinner=False)
def preparar_plano(raw: pd.DataFrame, c_red, c_cla, c_des, c_tip):
    df = pd.DataFrame({
        "reduzido": raw[c_red], "classificacao": raw[c_cla], "descricao": raw[c_des],
        "tipo": raw[c_tip] if c_tip else "",
    }).dropna(subset=["reduzido", "classificacao"])
    df["reduzido"] = df["reduzido"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    df["classificacao"] = df["classificacao"].map(norm_classif)
    df["descricao"] = df["descricao"].fillna("").astype(str).str.strip()
    df = df[df["reduzido"].str.fullmatch(r"\d+") & (df["classificacao"] != "")]
    df = df.drop_duplicates(subset=["reduzido", "classificacao"]).reset_index(drop=True)
    pontuado = bool(df["classificacao"].str.contains(".", regex=False).mean() > 0.5)
    if not pontuado:
        df["classificacao"] = df["classificacao"].str.replace(".", "", regex=False)
    existentes = set(df["classificacao"])
    df["pai"] = df["classificacao"].map(
        lambda c: next((p for p in reversed(ancestrais(c, pontuado)) if p in existentes), ""))
    pais = set(df["pai"])
    t = df["tipo"].fillna("").astype(str).map(norm).str[:1]
    if not t.isin(["A", "S"]).mean() >= 0.9:  # coluna ausente/irregular -> deduz pela hierarquia
        t = df["classificacao"].map(lambda c: "S" if c in pais else "A")
    df["tipo"] = t
    df["desc_norm"] = df["descricao"].map(norm)
    return df, pontuado


def detectar_raizes(plano):
    raizes = plano[plano["pai"] == ""]
    out = {k: [] for k in ROTULOS_RAIZ}
    for _, r in raizes.iterrows():
        d, c = r["desc_norm"], r["classificacao"]
        if tem(d, "ATIVO") and not tem(d, "PASSIVO"):
            out["ATIVO"].append(c)
        elif tem(d, "PASSIVO"):
            out["PASSIVO"].append(c)
        elif tem(d, "RECEITA*") and not tem(d, "CUSTO*", "DESPESA*"):
            out["RECEITA"].append(c)
        elif tem(d, "CUSTO*", "DESPESA*", "RESULTADO*"):
            out["RESULTADO"].append(c)
    return out, raizes


def caminho_de(plano, pontuado):
    dsc = dict(zip(plano["classificacao"], plano["descricao"]))
    return lambda c: " > ".join(dsc[a] for a in ancestrais(c, pontuado) if a in dsc)


def grupos_resultado(plano, raizes_res, pontuado):
    """Grupos sintéticos de custos/despesas que contêm conta analítica de Salários."""
    cam = caminho_de(plano, pontuado)
    na_raiz = plano["classificacao"].map(lambda c: any(sob(c, r, pontuado) for r in raizes_res))
    an = plano[(plano["tipo"] == "A") & na_raiz]
    sint = plano[(plano["tipo"] == "S") & na_raiz & ~plano["classificacao"].isin(raizes_res)]
    sal = an[an["desc_norm"].map(lambda d: tem(d, "SALARIO*", "ORDENADO*")
                                 and not tem(d, "PAGAR", "FAMILIA", "MATERN*", "PROVIS*", "EDUCACAO")
                                 and not e13(d))]["classificacao"].tolist()
    out = [{"classif": g.classificacao, "descricao": g.descricao, "caminho": cam(g.classificacao)}
           for g in sint.itertuples() if any(sob(c, g.classificacao, pontuado) for c in sal)]
    if not out:  # plano sem "Salários": oferece os pais imediatos de contas analíticas
        pais = set(an["pai"])
        out = [{"classif": g.classificacao, "descricao": g.descricao, "caminho": cam(g.classificacao)}
               for g in sint.itertuples() if g.classificacao in pais]
    return [g for g in out
            if not any(o is not g and sob(o["classif"], g["classif"], pontuado)
                       and o["classif"] != g["classif"] for o in out)]


def sugerir_grupo(nome_sep, cands):
    if not cands:
        return None
    n = norm(nome_sep)
    if tem(n, "VENDA*", "COMERCI*", "LOJA*", "MARKETING"):
        chave = "VENDA*"
    elif tem(n, "PRODU*", "FABRI*", "INDUSTR*", "OBRA*", "OPERAC*", "MANUTENC*"):
        chave = "CUSTO*"
    else:
        chave = "ADMINISTRATIV*"
    for i, g in enumerate(cands):
        if tem(norm(f"{g['caminho']} {g['descricao']}"), chave):
            return i
    return 0

# =====================================================================
# 3. ALVOS CONTÁBEIS (localizados por descrição; nada fixo ao plano)
#    escopo: G = grupo do separador | ATIVO | PASSIVO | RESULTADO | RECEITA_RES
# =====================================================================
G, AT, PA, RS, RR = "G", "ATIVO", "PASSIVO", "RESULTADO", "RECEITA_RES"
ESCOPO_ROTULO = {AT: "Ativo", PA: "Passivo", RS: "Custos/Despesas", RR: "Receitas + Custos/Despesas"}


def _13(extra=(), nots=()):
    return [(["13O", *extra], list(nots)), (["13A", *extra], list(nots)),
            (["13", *extra], list(nots)), (["DECIMO*", *extra], list(nots))]


NAO_TERC = ["SERVICO*", "TERCEIROS", "ALUGUE*", "APLICAC*", "PJ", "JUROS", "PROVIS*"]

ALVOS = {
    # ---- DRE (restrito ao grupo do separador)
    "SALARIOS": (G, "Salários e Ordenados", [
        (["SALARIO*"], ["PAGAR", "FAMILIA", "MATERN*", "PROVIS*", "13O", "13A", "13", "DECIMO*",
                        "ADIANT*", "INSS", "FGTS", "EDUCACAO"]),
        (["ORDENADO*"], ["PAGAR"])]),
    "PRO_LABORE": (G, "Pró-labore", [(["PRO LABORE"], ["PAGAR", "INSS"]),
                                     (["HONORARIO*", "DIRETOR*"], []), (["RETIRADA*", "SOCIO*"], [])]),
    "HE": (G, "Horas Extras", [(["HORAS EXTRA*"], []), (["HORA EXTRA*"], []), (["EXTRAORDINAR*"], [])]),
    "PREMIOS": (G, "Prêmios e Gratificações", [(["PREMIO*"], []), (["GRATIFICAC*"], []),
                                               (["BONIFICAC*"], [])]),
    "COMISSOES": (G, "Comissões", [(["COMISS*"], ["BANCARI*"])]),
    "DECIMO": (G, "13º Salário", _13((), ["PROVIS*", "INSS", "FGTS", "PIS", "PAGAR", "ADIANT*"])),
    "FERIAS": (G, "Férias", [(["FERIAS"], ["PROVIS*", "PAGAR", "INSS", "FGTS", "PIS", "ADIANT*"])]),
    "INSS": (G, "INSS (encargo)", [
        (["INSS"], ["PROVIS*", "RECOLHER", "PAGAR", "RECEITA BRUTA", "RETIDO", "COMPENSAR"]),
        (["PREVIDENCIA*"], ["RECOLHER", "PAGAR", "PRIVADA", "COMPLEMENTAR"]),
        (["CONTRIBUIC*", "PREVIDENCIARIA*"], ["RECOLHER", "PAGAR", "RECEITA BRUTA"])]),
    "FGTS": (G, "FGTS (encargo)", [(["FGTS"], ["PROVIS*", "RECOLHER", "PAGAR"])]),
    "PIS": (G, "PIS s/ Folha", [(["PIS"], ["PROVIS*", "RECOLHER", "PAGAR", "RETIDO", "FATURAMENTO",
                                           "RECEITA*", "COFINS"])]),
    "INDENIZ": (G, "Indenizações e Aviso Prévio", [(["INDENIZAC*"], []), (["AVISO PREVIO"], []),
                                                   (["RESCIS*"], ["PAGAR"])]),
    "ASSIST": (G, "Assistência Médica", [(["ASSISTENCIA MEDICA*"], []), (["PLANO*", "SAUDE"], []),
                                         (["CONVENIO*", "MEDIC*"], []), (["ODONTO*"], [])]),
    "VT": (G, "Vale-Transporte", [(["VALE TRANSP*"], []), (["TRANSPORTE*", "EMPREGADO*"], []),
                                  (["TRANSPORTE*", "FUNCIONARIO*"], [])]),
    "ALIM": (G, "Alimentação / VR", [(["ALIMENTACAO"], []), (["VALE REFEICAO"], []),
                                     (["REFEICAO*"], []), (["CESTA*", "BASICA*"], [])]),
    "BOLSA": (G, "Bolsa-Auxílio / Estágio", [(["BOLSA*"], []), (["ESTAGI*"], [])]),
    "SEGURO": (G, "Seguro de Vida", [(["SEGURO*", "VIDA"], [])]),
    "TREIN": (G, "Treinamento", [(["TREINAMENTO*"], []), (["CURSO*"], []), (["CAPACITAC*"], [])]),
    "D_PROV_FER": (G, "Férias - Provisão (despesa)", [(["FERIAS", "PROVIS*"], ["INSS", "FGTS", "PIS"])]),
    "D_INSS_FER": (G, "INSS s/ Provisão Férias (despesa)", [(["INSS", "FERIAS", "PROVIS*"], [])]),
    "D_FGTS_FER": (G, "FGTS s/ Provisão Férias (despesa)", [(["FGTS", "FERIAS", "PROVIS*"], [])]),
    "D_PIS_FER": (G, "PIS s/ Provisão Férias (despesa)", [(["PIS", "FERIAS", "PROVIS*"], [])]),
    "D_PROV_13": (G, "13º - Provisão (despesa)", _13(["PROVIS*"], ["INSS", "FGTS", "PIS"])),
    "D_INSS_13": (G, "INSS s/ Provisão 13º (despesa)", _13(["INSS", "PROVIS*"])),
    "D_FGTS_13": (G, "FGTS s/ Provisão 13º (despesa)", _13(["FGTS", "PROVIS*"])),
    "D_PIS_13": (G, "PIS s/ Provisão 13º (despesa)", _13(["PIS", "PROVIS*"])),
    # ---- Passivo
    "SAL_PAGAR": (PA, "Salários a Pagar", [
        (["SALARIO*", "PAGAR"], ["FERIAS", "13O", "13A", "13", "DECIMO*", "PRO LABORE", "RESCIS*"]),
        (["ORDENADO*", "PAGAR"], []), (["FOLHA*", "PAGAR"], [])]),
    "FER_PAGAR": (PA, "Férias a Pagar", [(["FERIAS", "PAGAR"], ["PROVIS*"])]),
    "RESC_PAGAR": (PA, "Rescisões a Pagar", [(["RESCIS*", "PAGAR"], [])]),
    "PROLAB_PAGAR": (PA, "Pró-labore a Pagar", [(["PRO LABORE", "PAGAR"], []), (["PRO LABORE"], ["INSS"])]),
    "INSS_REC": (PA, "INSS a Recolher", [
        (["INSS", "RECOLHER"], ["RECEITA BRUTA", "PROVIS*", "RETIDO"]),
        (["INSS", "PAGAR"], ["RECEITA BRUTA", "PROVIS*", "RETIDO"]),
        (["PREVIDENCIA SOCIAL"], ["PROVIS*"])]),
    "FGTS_REC": (PA, "FGTS a Recolher", [(["FGTS", "RECOLHER"], ["PROVIS*"]),
                                         (["FGTS", "PAGAR"], ["PROVIS*"]), (["FGTS"], ["PROVIS*"])]),
    "IRRF_REC": (PA, "IRRF s/ Folha a Recolher", [
        (["IRRF", "FOLHA*"], ["PROVIS*"]), (["IRRF", "SALARIO*"], []), (["IRRF", "TRABALH*"], []),
        (["IRRF", "RECOLHER"], NAO_TERC), (["IRRF", "PAGAR"], NAO_TERC),
        (["IMPOSTO DE RENDA", "FONTE"], NAO_TERC), (["IR", "FONTE"], NAO_TERC), (["IRRF"], NAO_TERC)]),
    "PIS_REC": (PA, "PIS s/ Folha a Recolher", [(["PIS", "FOLHA*"], ["PROVIS*"]),
                                                (["PIS", "RECOLHER"], ["PROVIS*", "RETIDO", "FATURAMENTO"])]),
    "SIND_REC": (PA, "Contribuição Sindical a Recolher", [
        (["CONTRIBUIC*", "SINDICA*"], ["PATRONAL"]), (["SINDICA*"], ["PATRONAL"])]),
    "SIND_PAT_REC": (PA, "Contribuição Sindical Patronal a Recolher", [(["SINDICA*", "PATRONAL"], [])]),
    "EMPREST": (PA, "Empréstimos Consignados a Repassar", [
        (["CONSIGNA*"], []), (["EMPRESTIMO*", "EMPREGADO*"], []),
        (["EMPRESTIMO*", "FUNCIONARIO*"], []), (["EMPRESTIMO*", "FOLHA"], [])]),
    "PENSAO": (PA, "Pensão Alimentícia a Repassar", [(["PENSAO*"], [])]),
    "CPRB_REC": (PA, "INSS Receita Bruta (CPRB) a Recolher", [(["RECEITA BRUTA"], ["PROVIS*"]),
                                                               (["CPRB"], [])]),
    "P_PROV_FER": (PA, "Provisão para Férias", [(["PROVIS*", "FERIAS"], ["INSS", "FGTS", "PIS"])]),
    "P_INSS_FER": (PA, "INSS s/ Provisão Férias", [(["INSS", "PROVIS*", "FERIAS"], [])]),
    "P_FGTS_FER": (PA, "FGTS s/ Provisão Férias", [(["FGTS", "PROVIS*", "FERIAS"], [])]),
    "P_PIS_FER": (PA, "PIS s/ Provisão Férias", [(["PIS", "PROVIS*", "FERIAS"], [])]),
    "P_PROV_13": (PA, "Provisão para 13º", _13(["PROVIS*"], ["INSS", "FGTS", "PIS"])),
    "P_INSS_13": (PA, "INSS s/ Provisão 13º", _13(["INSS", "PROVIS*"])),
    "P_FGTS_13": (PA, "FGTS s/ Provisão 13º", _13(["FGTS", "PROVIS*"])),
    "P_PIS_13": (PA, "PIS s/ Provisão 13º", _13(["PIS", "PROVIS*"])),
    # ---- Ativo
    "ADIANT_SAL": (AT, "Adiantamento de Salário", [
        (["ADIANTAMENTO*", "SALARIO*"], ["13O", "13A", "13", "DECIMO*", "FERIAS"]),
        (["ADIANTAMENTO*", "EMPREGADO*"], []), (["ADIANTAMENTO*", "FUNCIONARIO*"], []),
        (["ADIANTAMENTO*", "PESSOAL"], [])]),
    "ADIANT_13": (AT, "Adiantamento de 13º", _13(["ADIANTAMENTO*"])),
    "ADIANT_FER": (AT, "Adiantamento de Férias", [(["ADIANTAMENTO*", "FERIAS"], [])]),
    "INSS_COMP": (AT, "INSS a Compensar (retenções)", [(["INSS", "COMPENSAR"], ["MATERN*", "FAMILIA"]),
                                                       (["INSS", "RECUPERAR"], [])]),
    "BENEF_INSS": (AT, "Sal.-Família/Maternidade a Compensar", [
        (["MATERN*", "COMPENSAR"], []), (["FAMILIA", "COMPENSAR"], []),
        (["INSS", "COMPENSAR"], []), (["INSS", "RECUPERAR"], [])]),
    # ---- Resultado (todas as raízes de custos/despesas) e receita
    "SIND_PAT": (RS, "Contribuição Sindical Patronal (despesa)", [(["SINDICA*", "PATRONAL"], [])]),
    "TAXAS_DIV": (RS, "Taxas Diversas", [(["TAXAS DIVERSAS"], []), (["TAXAS", "CONTRIBUIC*"], [])]),
    "CPRB_DED": (RR, "(-) INSS Receita Bruta (CPRB)", [
        (["INSS", "RECEITA BRUTA"], ["RECOLHER", "PAGAR"]), (["CPRB"], ["RECOLHER", "PAGAR"]),
        (["CONTRIBUIC*", "PREVIDENCIARIA*", "RECEITA*"], ["RECOLHER", "PAGAR"])]),
}

# Contas de obrigação/direito com o colaborador: itens patronais nunca podem usá-las
COLAB_ALVOS = ["SAL_PAGAR", "FER_PAGAR", "RESC_PAGAR", "PROLAB_PAGAR",
               "ADIANT_SAL", "ADIANT_13", "ADIANT_FER", "EMPREST", "PENSAO"]


class Resolvedor:
    def __init__(self, plano, pontuado, raizes, overrides):
        self.an = plano[plano["tipo"] == "A"]
        self.p, self.raizes, self.ov = pontuado, raizes, overrides
        self.desc = dict(zip(plano["reduzido"], plano["descricao"]))
        self.cache, self.bases = {}, {}

    @staticmethod
    def chave(alvo, grupo):
        return f"{alvo}|{grupo if ALVOS[alvo][0] == G else '*'}"

    def _prefixos(self, alvo, grupo):
        esc = ALVOS[alvo][0]
        if esc == G:
            return (grupo,) if grupo else ()
        if esc == RR:
            return tuple(self.raizes["RECEITA"] + self.raizes["RESULTADO"])
        return tuple(self.raizes[esc])

    def _base(self, prefs):
        if prefs not in self.bases:
            self.bases[prefs] = self.an[self.an["classificacao"].map(
                lambda c: any(sob(c, p, self.p) for p in prefs))]
        return self.bases[prefs]

    def sugestao(self, alvo, grupo):
        k = self.chave(alvo, grupo)
        if k not in self.cache:
            res, prefs = None, self._prefixos(alvo, grupo)
            base = self._base(prefs) if prefs else self.an.iloc[0:0]
            if not base.empty:
                for must, nots in ALVOS[alvo][2]:
                    ok = base["desc_norm"].map(
                        lambda d: all(tem(d, x) for x in must) and not any(tem(d, x) for x in nots))
                    c = base[ok]
                    if not c.empty:
                        c = c.assign(_n=c["desc_norm"].str.len()).sort_values(["_n", "classificacao"])
                        res = (c.iloc[0]["reduzido"], c.iloc[0]["descricao"])
                        break
            self.cache[k] = res
        return self.cache[k]

    def conta(self, alvo, grupo):
        k = self.chave(alvo, grupo)
        if self.ov.get(k):
            return self.ov[k], self.desc.get(self.ov[k], "?")
        return self.sugestao(alvo, grupo)

    def resolver(self, alvos, grupo):
        if not alvos:
            return "", "", ""
        for a in alvos:
            r = self.conta(a, grupo)
            if r:
                aviso = "" if a == alvos[0] else f"fallback {ALVOS[alvos[0]][1]} → {ALVOS[a][1]}"
                return r[0], r[1], aviso
        return "", "", f"❌ sem conta '{ALVOS[alvos[0]][1]}' (defina no Mapa de contas)"

# =====================================================================
# 4. LEITURA DOS RELATÓRIOS DA DOMÍNIO
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
    return []


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
            if SECOES[n] != secao:  # cabeçalho repetido na quebra de página não zera o separador
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
# 5. REGRAS DETERMINÍSTICAS
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
    if tem(d, "PRO LAB*"):
        return True
    return tem(d, "EMPREGADOR") and tem(d, "INSS", "IRRF", "TROCO", "ESTOURO", "ADTO", "ADIANT*")


def passivo_secao(secao, d):
    if secao == "Férias": return ["FER_PAGAR", "SAL_PAGAR"]
    if secao == "Rescisão": return ["RESC_PAGAR", "SAL_PAGAR"]
    if e_socio(d): return ["PROLAB_PAGAR", "SAL_PAGAR"]
    return ["SAL_PAGAR"]


def alvo_adiant(d):
    if e13(d): return ["ADIANT_13", "ADIANT_SAL"]
    if tem(d, "FERIAS"): return ["ADIANT_FER", "ADIANT_SAL"]
    return ["ADIANT_SAL"]


def e_licenca_remunerada(d):
    """LICENCA REMUNERADA / LICENC REMUN / LIC REMUN / LIC REM"""
    return tem(d, "LICENC*", "LIC") and tem(d, "REMUN*", "REM")


def alvo_dre_provento(d):
    if tem(d, "PRO LABORE"): return ["PRO_LABORE", "SALARIOS"]
    if tem(d, "BOLSA", "ESTAGI*", "RECESSO"): return ["BOLSA", "SALARIOS"]
    # Licença remunerada (férias coletivas s/ período aquisitivo) = natureza salarial;
    # a versão "GOZADA(S)" é férias de fato.
    if e_licenca_remunerada(d) and not tem(d, "GOZ*"): return ["SALARIOS"]
    if tem(d, "HORAS EXTRA*", "EXTRAS", "BANCO DE HORAS"): return ["HE", "SALARIOS"]
    if e13(d): return ["DECIMO", "SALARIOS"]
    if tem(d, "FERIAS", "ABONO"): return ["FERIAS", "SALARIOS"]
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
        return ["ADIANT_SAL"], L, "Estouro/troco: crédito a receber do empregado", False
    if tem(d, "ADIANT*", "ADTO"):
        return alvo_adiant(d), L, "Adiantamento pago: Ativo × obrigação", False
    if tem(d, "DEV", "DEVOLUCAO") and tem(d, "EMPREST*", "CONSIG*"):
        return ["EMPREST"], L, "Devolução de consignado", True
    if tem(d, "INSS", "IRRF") and tem(d, "A MAIOR", "DIF*", "DEVOL*", "RESTITUI*"):
        return (["IRRF_REC"] if tem(d, "IRRF") else ["INSS_REC"]), L, "Restituição de retenção a maior", False
    if tem(d, "SALARIO FAMILIA", "SAL FAM*"):
        return ["BENEF_INSS"], L, "Salário-família → conta-ponte até a compensação na guia", False
    if tem(d, "MATERN*") and not tem(d, "DESC*", "DEDUC*"):
        if tem(d, "INSS"):
            return [], [], "Maternidade paga direto pelo INSS — não integrar", True
        if not tem(d, "PRORROG*", "EMPREGADOR"):
            return ["BENEF_INSS"], L, "Salário-maternidade reembolsável → conta-ponte", False
    dre = alvo_dre_provento(d)
    if baixa_prov and secao in ("Férias", "Rescisão", "13º Salário") and dre[0] in ("FERIAS", "DECIMO"):
        return (["P_PROV_FER"] if dre[0] == "FERIAS" else ["P_PROV_13"]), L, "Baixa contra a provisão", False
    return dre, L, "", False


def classificar_desconto(d, secao):
    L = passivo_secao(secao, d)
    def r(c, obs="", rev=False): return L, c, obs, rev
    if tem(d, "ESTOURO", "TROCO"): return r(["ADIANT_SAL"], "Baixa de estouro/troco anterior")
    if tem(d, "PENSAO*"): return r(["PENSAO"], "Repasse ao beneficiário", True)
    if tem(d, "SAL FAM*", "SALARIO FAMILIA"):
        return r(["BENEF_INSS"], "Estorno de salário-família pago a maior", True)
    if tem(d, "INSS"): return r(["INSS_REC"])
    if tem(d, "IRRF", "IMPOSTO DE RENDA", "IR"): return r(["IRRF_REC"])
    if tem(d, "ADIANT*", "ADTO", "FERIAS PAGAS"): return r(alvo_adiant(d), "Baixa do adiantamento")
    if tem(d, "SINDICA*", "ASSISTENCIAL", "CONFEDERATIVA", "NEGOC*"): return r(["SIND_REC"])
    if tem(d, "EMPREST*", "EMP", "CONSIG*", "CRED TRAB"): return r(["EMPREST"], "Consignado a repassar", True)
    if tem(d, "PLANO", "ODONTO*", "COPARTICIP*", "ASSISTENCIA MEDICA", "FARMACIA", "SAUDE"):
        return r(["ASSIST"], "Recuperação do custo do benefício")
    if tem(d, "VALE TRANSPORTE", "VT"): return r(["VT"], "Recuperação dos 6% do VT")
    if tem(d, "REFEICAO", "ALIMENTACAO", "VR", "CESTA", "SUPERMERCADO"):
        return r(["ALIM"], "Recuperação do custo do benefício")
    if tem(d, "SEGURO*"): return r(["SEGURO", "ASSIST"], "Recuperação do custo do benefício")
    if tem(d, "AVISO PREVIO", "MULTA", "ESTABILIDADE", "INDENIZ*"):
        return r(["INDENIZ", "SALARIOS"], "Indenização devida pelo empregado", True)
    if e13(d) and tem(d, "AFAST*", "MATERN*"): return r(["DECIMO", "SALARIOS"], "Redução do custo de 13º")
    if tem(d, "FALTA*", "ATRASO*", "DSR", "HORAS", "PAGO A MAIOR", "INATIVAS"):
        return r(["SALARIOS"], "Redução do custo de salários")
    return r([], "Desconto sem regra — definir conta manualmente", True)


def classificar_informativa(d, tipo):
    if tem(d, "BASE", "E SOCIAL", "INFORMATIVO", "HORAS CREDITO", "HORAS COMPENSADA",
           "BANCO DE HORAS", "ADIANT*"):
        return [], [], "Informativa de base/controle — não integrar", True
    if tem(d, "FGTS", "CONTRIBUICAO SOCIAL", "CONTRIB SOCIAL"):
        dre = ["INDENIZ", "FGTS"] if tem(d, "40%", "20%") else ["FGTS"]
        if tem(d, "A MAIOR") or tipo == "Inf. dedutora":
            return ["FGTS_REC"], dre, "Estorno de FGTS a maior", True
        return dre, ["FGTS_REC"], "Encargo FGTS", False
    if tem(d, "INSS"): return ["INSS"], ["INSS_REC"], "Encargo INSS", False
    if tem(d, "PIS"): return ["PIS"], ["PIS_REC"], "PIS s/ folha — só entidades sem fins lucrativos", True
    return [], [], "Informativa não reconhecida", True


def classificar_item(d, secao):
    """Provisões e Informações INSS (itens da seção, não rubricas)."""
    if secao == "Informações INSS":
        if tem(d, "RETENC*", "RETIDO", "RETIDA*"):
            return ["INSS_REC"], ["INSS_COMP", "BENEF_INSS"], "Retenção s/ NF compensada na guia", True
        reconhecido = tem(d, "FAMILIA", "MATERN*")
        return ["INSS_REC"], ["BENEF_INSS"], "Compensação na guia — baixa da conta-ponte", not reconhecido
    sfx = "FER" if "Férias" in secao else "13"
    base = "FERIAS" if sfx == "FER" else "DECIMO"
    if tem(d, "FGTS"):
        dre, pas = [f"D_FGTS_{sfx}", "FGTS"], [f"P_FGTS_{sfx}", f"P_PROV_{sfx}"]
    elif tem(d, "INSS", "RAT", "SAT", "TERCEIROS", "FAP", "ACID*"):
        dre, pas = [f"D_INSS_{sfx}", "INSS"], [f"P_INSS_{sfx}", f"P_PROV_{sfx}"]
    elif tem(d, "PIS"):
        dre, pas = [f"D_PIS_{sfx}", "PIS"], [f"P_PIS_{sfx}", f"P_PROV_{sfx}"]
    else:
        dre, pas = [f"D_PROV_{sfx}", base], [f"P_PROV_{sfx}"]
    if tem(d, "ESTORNO*", "BAIXA*", "REVERS*"):
        return pas, dre, "Estorno da provisão", False
    return dre, pas, "Constituição da provisão", False


OBS_EMPRESA = {
    "INSS": ("Encargo patronal INSS", False),
    "INSS_SOCIO": ("INSS patronal s/ pró-labore", False),
    "AUTONOMO": ("INSS patronal s/ autônomo — confirme a conta de despesa", True),
    "SENAI": ("Adicional ao SENAI — confirme a forma de recolhimento", True),
    "PIS": ("PIS s/ folha — só entidades sem fins lucrativos", True),
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
    prov = bool(baixa_prov and sfx and nat in ("INSS", "SENAI", "PIS"))
    if nat in ("INSS", "INSS_SOCIO", "AUTONOMO", "SENAI"):
        deb, cred = ([f"P_INSS_{sfx}", f"P_PROV_{sfx}"] if prov else ["INSS"]), ["INSS_REC"]
    elif nat == "PIS":
        deb, cred = ([f"P_PIS_{sfx}", f"P_PROV_{sfx}"] if prov else ["PIS"]), ["PIS_REC"]
    elif nat == "FGTS":
        deb, cred = ["FGTS"], ["FGTS_REC"]
    elif nat == "CPRB":
        deb, cred = ["CPRB_DED"], ["CPRB_REC", "INSS_REC"]
    elif nat == "SIND_PATRONAL":
        deb, cred = ["SIND_PAT", "TAXAS_DIV"], ["SIND_PAT_REC", "SIND_REC"]
    elif nat == "DEDUCAO":
        deb, cred = ["INSS_REC"], ["BENEF_INSS"]
    else:
        deb, cred = [], []
    obs, rev = OBS_EMPRESA.get(nat, OBS_EMPRESA[None])
    if prov:
        obs += " — baixa contra a provisão"
    return deb, cred, obs, rev, origem, nat

# =====================================================================
# 6. CONFIGURAÇÃO POR EMPRESA (.json)
# =====================================================================
def aplicar_config(dados: bytes):
    fid = hashlib.md5(dados).hexdigest()
    if st.session_state.get("cfg_id") == fid:
        return
    try:
        cfg = json.loads(dados.decode("utf-8"))
    except Exception as e:
        st.sidebar.error(f"Configuração inválida: {e}")
        return
    st.session_state["cfg_id"], st.session_state["cfg"] = fid, cfg
    for k in list(st.session_state.keys()):
        if k.startswith(("g_", "m_", "r_", "c_")) or k == "w_grupo_socio":
            del st.session_state[k]
    st.session_state["cfg_contas"] = {str(k): str(v) for k, v in cfg.get("contas", {}).items()}
    for chave, wk in (("historico", "w_hist"), ("baixa_prov", "w_baixa"), ("socio_adm", "w_socio"),
                      ("bloqueio_extra", "w_bloq"), ("nome_geral", "w_nome_geral"), ("empresa", "w_cod")):
        if chave in cfg:
            st.session_state[wk] = cfg[chave]


def md5(txt: str) -> str:
    return hashlib.md5(txt.encode("utf-8")).hexdigest()[:10]

# =====================================================================
# 7. INTERFACE
# =====================================================================
st.set_page_config(page_title="Integrador Contábil da Folha", layout="wide")
st.title("📒 Integrador Contábil da Folha — Domínio")
st.caption("Motor 100% determinístico (regex + regras de substring). Nenhuma IA, API ou chave paga. "
           "Nenhuma conta fixa: tudo é lido do plano importado.")

with st.sidebar:
    st.header("0. Configuração da empresa")
    f_cfg = st.file_uploader("Carregar configuração salva (.json)", type=["json"])
    if f_cfg is not None:
        aplicar_config(f_cfg.getvalue())
    st.header("1. Arquivos")
    f_pend = st.file_uploader("Rubricas/Itens não configurados", type=["pdf", "txt"])
    f_cad = st.file_uploader("Cadastro geral de rubricas", type=["pdf", "txt"])
    f_plano = st.file_uploader("Plano de contas", type=["xlsx", "xls", "csv", "txt"])
    st.header("2. Parâmetros")
    for k, v in (("w_hist", ""), ("w_baixa", False), ("w_socio", True), ("w_bloq", "")):
        st.session_state.setdefault(k, v)
    historico = st.text_input("Código do histórico padrão", key="w_hist")
    baixa_prov = st.checkbox("Baixar férias/13º pagos contra a provisão", key="w_baixa",
                             help="Deixe desmarcado se a Domínio já gera o 'Valor Estorno Provisão'.")
    socio_adm = st.checkbox("Pró-labore e encargos do sócio sempre em Despesas Administrativas", key="w_socio")
    bloq_extra = st.text_input("Contas extras de colaborador (reduzidos, separados por vírgula)", key="w_bloq",
                               help="Além das detectadas automaticamente; itens patronais nunca poderão usá-las.")

if not (f_pend and f_cad and f_plano):
    st.info("Envie os três arquivos para começar. Opcional: carregue a configuração salva da empresa.")
    st.stop()

cfg = st.session_state.get("cfg", {})
try:
    cab, itens = parse_pendencias(f_pend.getvalue(), f_pend.name)
    nome_cad, cad = parse_cadastro(f_cad.getvalue(), f_cad.name)
    raw = ler_tabela_bruta(f_plano.getvalue(), f_plano.name)
except Exception as e:
    st.error(f"Erro na leitura: {e}")
    st.stop()

if not itens:
    st.error("Nenhuma rubrica/item encontrado no relatório de pendências.")
    st.stop()
if not cad:
    st.warning("Nenhuma rubrica lida no cadastro geral — o Tipo de todas será inferido pela descrição.")

pid = hashlib.md5(f_pend.getvalue()).hexdigest()
if st.session_state.get("pend_id") != pid:
    st.session_state["pend_id"] = pid
    if not cfg.get("empresa"):
        st.session_state["w_cod"] = cab["codigo"]
st.session_state.setdefault("w_cod", cab["codigo"])
cod_empresa = st.sidebar.text_input("Código da empresa na Domínio", key="w_cod")
if cfg.get("empresa") and cab["codigo"] and str(cfg["empresa"]) != cab["codigo"]:
    st.warning(f"⚠️ A configuração carregada é da empresa {cfg['empresa']}, "
               f"mas o relatório é da empresa {cab['codigo']}.")

# ---------- 1. Estrutura do plano ----------
st.subheader("1. Estrutura do plano de contas")
cols = list(raw.columns)
auto = auto_colunas(cols)
cfg_cols = cfg.get("colunas", {})
exp_plano = st.expander("⚙️ Colunas e raízes do plano",
                        expanded=not all(auto.get(c) for c in ("reduzido", "classificacao", "descricao")))
with exp_plano:
    cc = st.columns(4)
    sel = {}
    for i, (campo, rot) in enumerate(CAMPOS_PLANO.items()):
        opts = (["(nenhuma)"] if campo == "tipo" else []) + cols
        wk = f"c_{campo}"
        if st.session_state.get(wk) not in opts:
            pref = cfg_cols.get(campo) if cfg_cols.get(campo) in opts else auto.get(campo)
            st.session_state[wk] = pref if pref in opts else opts[0]
        sel[campo] = cc[i].selectbox(rot, opts, key=wk, index=None)

try:
    plano, pontuado = preparar_plano(raw, sel["reduzido"], sel["classificacao"], sel["descricao"],
                                     None if sel["tipo"] in (None, "(nenhuma)") else sel["tipo"])
except Exception as e:
    st.error(f"Não foi possível montar o plano com essas colunas: {e}")
    st.stop()
if plano.empty:
    st.error("Plano vazio — confira as colunas de Código reduzido e Classificação.")
    st.stop()

raizes_det, df_raizes = detectar_raizes(plano)
opts_r = df_raizes["classificacao"].tolist()
desc_r = dict(zip(df_raizes["classificacao"], df_raizes["descricao"]))
cfg_r = cfg.get("raizes", {})
raizes = {}
with exp_plano:
    rc = st.columns(4)
    for i, (nome_r, rot) in enumerate(ROTULOS_RAIZ.items()):
        wk = f"r_{nome_r}"
        cur = st.session_state.get(wk)
        if cur is None or any(x not in opts_r for x in cur):
            st.session_state[wk] = [x for x in cfg_r.get(nome_r, raizes_det[nome_r]) if x in opts_r]
        raizes[nome_r] = rc[i].multiselect(rot, opts_r, key=wk,
                                           format_func=lambda c: f"{c} — {desc_r.get(c, '')}")
    st.caption(f"{len(plano)} contas ({int((plano['tipo'] == 'A').sum())} analíticas) · classificação "
               f"{'com pontos' if pontuado else 'sem pontos'} · raízes encontradas: "
               + ", ".join(f"{c} {desc_r[c]}" for c in opts_r[:10]))

if not raizes["PASSIVO"] or not raizes["RESULTADO"]:
    st.error("Defina ao menos as raízes do Passivo e de Custos/Despesas em 'Colunas e raízes do plano'.")
    st.stop()
if not raizes["ATIVO"]:
    st.warning("Raiz do Ativo não definida — adiantamentos e compensações ficarão pendentes.")

m1, m2, m3, m4 = st.columns(4)
m1.metric("Itens pendentes", len(itens))
m2.metric("Rubricas no cadastro", len(cad))
m3.metric("Contas no plano", len(plano))
m4.metric("Empresa", f'{cab["codigo"]} - {cab["nome"][:25]}')

# ---------- 2. Separadores ----------
st.subheader("2. Separador → grupo de resultado")
usa_sep = any(i["sep_cod"] for i in itens)
precisa_geral = (not usa_sep) or any(not i["sep_cod"] for i in itens)
st.session_state.setdefault("w_nome_geral", NOME_PADRAO_SEM_SEPARADOR)
if precisa_geral:
    st.text_input("Nome do lote sem separador", key="w_nome_geral")
nome_geral = (st.session_state.get("w_nome_geral") or NOME_PADRAO_SEM_SEPARADOR).strip()

cands = grupos_resultado(plano, raizes["RESULTADO"], pontuado)
rotulos = [f"{g['classif']} — {g['descricao']}" + (f"  ({g['caminho']})" if g["caminho"] else "")
           for g in cands]
rot2pref = {r: g["classif"] for r, g in zip(rotulos, cands)}
pref2rot = {v: k for k, v in rot2pref.items()}
if not cands:
    st.warning("Nenhum grupo de resultado com conta de Salários foi detectado — informe a máscara manualmente.")

if usa_sep:
    tipos = sorted({i["sep_tipo"] for i in itens if i["sep_tipo"]})
    st.success(f"Folha com separador detectada ({', '.join(tipos)}).")
else:
    st.warning("Nenhuma quebra por Centro de Custo / Filial / Serviço → folha centralizada. "
               "Escolha o grupo contábil em que a folha inteira será classificada.")

seps = {}
for it in itens:
    seps.setdefault(it["sep_cod"] or nome_geral, it["sep_nome"] if it["sep_cod"] else "itens sem separador")

an_cla = plano.loc[plano["tipo"] == "A", "classificacao"].tolist()
cfg_grupos = cfg.get("grupos", {})
mapa_grupos = {}
for k, nome in seps.items():
    gk, mk = f"g_{k}", f"m_{k}"
    if gk not in st.session_state or (st.session_state[gk] is not None and st.session_state[gk] not in rotulos):
        pc = cfg_grupos.get(k)
        if pc and pc in pref2rot:
            st.session_state[gk] = pref2rot[pc]
        else:
            if pc:
                st.session_state[mk] = pc
            idx = sugerir_grupo(nome, cands) if usa_sep else None
            st.session_state[gk] = rotulos[idx] if idx is not None else None
    st.session_state.setdefault(mk, "")
    c1, c2 = st.columns([3, 1])
    escolha = c1.selectbox(f"Separador {k} — {nome}", rotulos, key=gk, index=None,
                           placeholder="Escolha o grupo contábil")
    c2.text_input("ou máscara manual", key=mk, placeholder="ex.: 4.2.01")
    pref = norm_mascara(st.session_state[mk], pontuado) or (rot2pref.get(escolha, "") if escolha else "")
    if pref and not any(sob(c, pref, pontuado) for c in an_cla):
        c2.error("Máscara sem contas analíticas")
    mapa_grupos[k] = pref

if any(not v for v in mapa_grupos.values()):
    st.info("Defina o grupo de todos os separadores para continuar.")
    st.stop()

grupo_socio = None
if socio_adm:
    adm = next((g["classif"] for g in cands
                if tem(norm(f"{g['caminho']} {g['descricao']}"), "ADMINISTRATIV*")), None)
    op_socio = ["(mesmo grupo do separador)"] + rotulos
    if st.session_state.get("w_grupo_socio") not in op_socio:
        pc = cfg.get("grupo_socio")
        st.session_state["w_grupo_socio"] = pref2rot.get(pc) or pref2rot.get(adm) or op_socio[0]
    esc_socio = st.selectbox("Grupo do pró-labore e encargos do sócio", op_socio, key="w_grupo_socio", index=None)
    grupo_socio = rot2pref.get(esc_socio)

# ---------- Classificação (regras) ----------
regras = []
for it in itens:
    d, sec = norm(it["descricao"]), it["secao"]
    k = it["sep_cod"] or nome_geral
    prefixo, nat = mapa_grupos[k], None
    if sec == "Empresa":
        tipo = "Item (Empresa)"
        deb, cred, obs, rev, origem, nat = classificar_item_empresa(it["codigo"], d, baixa_prov)
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
    if grupo_socio and (e_socio(d) or nat == "INSS_SOCIO"):
        prefixo = grupo_socio
    regras.append(dict(it=it, d=d, sec=sec, k=k, prefixo=prefixo, tipo=tipo, origem=origem,
                       deb=deb, cred=cred, obs=obs, rev=rev, nat=nat))

# ---------- 3. Mapa de contas ----------
st.subheader("3. Mapa de contas")
base_res = Resolvedor(plano, pontuado, raizes, {})
pares = {Resolvedor.chave(a, ""): (a, "") for a in COLAB_ALVOS}
for r in regras:
    for a in r["deb"] + r["cred"]:
        pares.setdefault(Resolvedor.chave(a, r["prefixo"]), (a, r["prefixo"]))
cfg_contas = st.session_state.setdefault("cfg_contas", {})
linhas_map = []
for chave, (a, g) in sorted(pares.items()):
    sug = base_res.sugestao(a, g)
    linhas_map.append({
        "Chave": chave, "Alvo": ALVOS[a][1],
        "Escopo": f"grupo {g}" if ALVOS[a][0] == G else ESCOPO_ROTULO[ALVOS[a][0]],
        "Sugerida": sug[0] if sug else "", "Descrição sugerida": sug[1] if sug else "❌ não encontrada",
        "Conta definida": cfg_contas.get(chave, ""),
    })
df_map = pd.DataFrame(linhas_map)
n_nf = int(((df_map["Sugerida"] == "") & (df_map["Conta definida"] == "")).sum())
with st.expander(f"Contas localizadas por descrição — {n_nf} alvo(s) sem conta", expanded=n_nf > 0):
    st.caption("Preencha 'Conta definida' com o código reduzido para substituir a sugestão. "
               "Os valores entram na configuração salva da empresa.")
    ed_map = st.data_editor(df_map, key=f"map_{md5(df_map.to_json())}", hide_index=True,
                            disabled=[c for c in df_map.columns if c != "Conta definida"],
                            column_config={"Conta definida": st.column_config.TextColumn()})

analiticas = set(plano.loc[plano["tipo"] == "A", "reduzido"])
overrides, erros_ov = {}, []
for row in ed_map.to_dict("records"):
    v = row.get("Conta definida")
    v = "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()
    if v:
        cfg_contas[row["Chave"]] = v
        if v in analiticas:
            overrides[row["Chave"]] = v
        else:
            erros_ov.append(f"{row['Alvo']} ({row['Escopo']}): {v} inexistente ou sintética")
    else:
        cfg_contas.pop(row["Chave"], None)
if erros_ov:
    st.error("Contas definidas ignoradas: " + "; ".join(erros_ov))

res = Resolvedor(plano, pontuado, raizes, overrides)
colab = {r[0] for r in (res.conta(a, "") for a in COLAB_ALVOS) if r}
colab |= {x for x in re.split(r"[,;\s]+", bloq_extra or "") if x}
st.caption("🔒 Contas de colaborador protegidas contra itens patronais: " + (", ".join(sorted(colab)) or "—"))

# ---------- Resolução ----------
linhas = []
for r in regras:
    it, sec = r["it"], r["sec"]
    dc, dd, ad = res.resolver(r["deb"], r["prefixo"])
    cc_, cd, ac = res.resolver(r["cred"], r["prefixo"])
    alertas = [a for a in (ad, ac) if a]
    if dc and dc == cc_:
        dc, dd = "", ""
        alertas.append("⚡ CURTO-CIRCUITO EVITADO")
    if sec in SECOES_DE_ITENS:
        if dc in colab:
            dc, dd = "", ""
            alertas.append("🚫 débito em conta de colaborador bloqueado")
        if cc_ in colab:
            cc_, cd = "", ""
            alertas.append("🚫 crédito em conta de colaborador bloqueado")
    if not r["deb"] and not r["cred"]:
        status = "❌ Pendente" if (sec == "Empresa" and r["nat"] != "ISENCAO") else "⏭️ Não integrar"
    elif not dc or not cc_:
        status = "❌ Pendente"
    elif r["rev"] or alertas or r["origem"].startswith("Inferido"):
        status = "⚠️ Revisar"
    else:
        status = "✅ OK"
    linhas.append({
        "Status": status, "Seção": sec, "Separador": r["k"],
        "Nome do separador": it["sep_nome"] or nome_geral, "Grupo": r["prefixo"],
        "Código": it["codigo"], "Descrição": it["descricao"], "Tipo": r["tipo"],
        "Origem do tipo": r["origem"], "Débito": dc, "Desc. Débito": dd,
        "Crédito": cc_, "Desc. Crédito": cd,
        "Observação": " | ".join([r["obs"]] + alertas).strip(" |"),
        "Exportar": bool(dc and cc_),
    })
df = pd.DataFrame(linhas)

n_div = int(df["Origem do tipo"].str.startswith("Inferido").sum())
ne, nc = norm(cab["nome"]), norm(nome_cad)
if ne and nc and ne not in nc and nc not in ne:
    if n_div:
        st.warning(f"⚠️ Cadastro de outra empresa (**{nome_cad}**): {n_div} item(ns) com código "
                   "divergente/ausente — Tipo inferido pela descrição.")
    else:
        st.info(f"Cadastro do modelo **{nome_cad}**, mas todas as rubricas conferem por código + descrição.")

# ---------- 4. Conferência ----------
st.subheader("4. Conferência")
st.caption("Edições aqui valem só para este lote. Para correções permanentes, use o Mapa de contas.")
f1, f2, f3, f4 = st.columns(4)
for col, s in zip((f1, f2, f3, f4), ("✅ OK", "⚠️ Revisar", "❌ Pendente", "⏭️ Não integrar")):
    col.metric(s, int((df["Status"] == s).sum()))

ed = st.data_editor(
    df, key=f"editor_{md5(df.to_json())}", hide_index=True,
    disabled=[c for c in df.columns if c not in ("Débito", "Crédito", "Exportar")],
    column_config={"Exportar": st.column_config.CheckboxColumn(),
                   "Débito": st.column_config.TextColumn(), "Crédito": st.column_config.TextColumn()},
)

idx_plano = plano.drop_duplicates("reduzido").set_index("reduzido")


def validar(row):
    msgs = []
    vals = {}
    for lado in ("Débito", "Crédito"):
        v = row[lado]
        c = "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()
        vals[lado] = c
        if not c:
            msgs.append(f"{lado} vazio")
        elif c not in idx_plano.index:
            msgs.append(f"{lado} {c} inexistente")
        elif idx_plano.loc[c, "tipo"] != "A":
            msgs.append(f"{lado} {c} é sintética")
        elif row["Seção"] in SECOES_DE_ITENS and c in colab:
            msgs.append(f"{lado} {c} é conta de colaborador — proibido em item patronal")
    if vals["Débito"] and vals["Débito"] == vals["Crédito"]:
        msgs.append("⚡ débito = crédito")
    return "; ".join(msgs)


ed = ed.copy()
ed["Validação"] = ed.apply(validar, axis=1)
exportar = ed[ed["Exportar"] & (ed["Validação"] == "")].reset_index(drop=True)
bloqueadas = ed[ed["Exportar"] & (ed["Validação"] != "")]
if not bloqueadas.empty:
    st.error(f"{len(bloqueadas)} linha(s) marcadas para exportar com erro — ficarão fora do lote.")
    st.dataframe(bloqueadas[["Seção", "Separador", "Código", "Descrição", "Validação"]], hide_index=True)

# ---------- 5. Exportação ----------
st.subheader("5. Arquivos")


def excel_bytes(abas: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for nome, d in abas.items():
            d.to_excel(w, sheet_name=nome, index=False)
    return buf.getvalue()


integra = pd.DataFrame([{"Código da Empresa": cod_empresa, "Separador": 1 if usa_sep else 0}])[COLS_INTEGRA]
evento = pd.DataFrame({
    "Código da Empresa": [cod_empresa] * len(exportar),
    "Separador": exportar["Separador"].astype(str).tolist(),
    "Código Sequencial": list(range(1, len(exportar) + 1)),
    "Tipo de Integração": exportar["Seção"].map(TIPO_INTEGRACAO).tolist(),
    "Código da Rubrica": exportar["Código"].tolist(),
    "Descrição da Rubrica": exportar["Descrição"].tolist(),
    "Conta Débito": exportar["Débito"].astype(str).str.strip().tolist(),
    "Conta Crédito": exportar["Crédito"].astype(str).str.strip().tolist(),
    "Histórico Padrão": [historico] * len(exportar),
})[COLS_EVENTO]

cfg_out = {
    "versao": 1, "empresa": cod_empresa, "nome_empresa": cab["nome"], "nome_geral": nome_geral,
    "grupos": mapa_grupos, "grupo_socio": grupo_socio, "contas": dict(cfg_contas),
    "historico": historico, "baixa_prov": baixa_prov, "socio_adm": socio_adm,
    "bloqueio_extra": bloq_extra, "raizes": raizes,
    "colunas": {k: v for k, v in sel.items() if v and v != "(nenhuma)"},
}

c1, c2, c3 = st.columns(3)
c1.download_button(f"📥 Importação Domínio ({len(evento)} linhas)",
                   excel_bytes({ABA_INTEGRA: integra, ABA_EVENTO: evento}),
                   file_name=f"integracao_folha_emp{cod_empresa}.xlsx")
c2.download_button("📋 Planilha de conferência completa", excel_bytes({"conferencia": ed}),
                   file_name=f"conferencia_folha_emp{cod_empresa}.xlsx")
c3.download_button("💾 Salvar configuração da empresa (.json)",
                   json.dumps(cfg_out, ensure_ascii=False, indent=2).encode("utf-8"),
                   file_name=f"config_folha_emp{cod_empresa}.json", mime="application/json")
