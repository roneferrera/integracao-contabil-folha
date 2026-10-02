"""
Integrador Contábil da Folha — Domínio Sistemas
Motor 100% determinístico (regex + regras de substring). Sem IA, sem APIs, sem chaves.

- Nenhuma conta, código reduzido ou classificação de plano é fixada no código.
- Exportação no layout de importação da Domínio (1 lançamento por rubrica/item):
    aba "evento"  -> lançamento: Código Sequencial = código da rubrica/item,
                     Descrição = "código - descrição" (caixa mista, máx. 40 caracteres)
    aba "integra" -> vínculo da rubrica/item ao seu lançamento
- Tipos: 1 Folha mensal | 2 Empresa | 3 Férias | 4 Rescisão | 5 Prov. Férias | 6 Prov. 13 | 10 Outras Informações
- Dois modos:
    * Pendências: lê o relatório "Rubricas/Itens não configurados"
    * Todos os eventos cadastrados: usa só o Cadastro geral de rubricas + Plano de contas
- "Usar a mesma configuração da folha normal para Férias/Rescisão" (igual à Domínio):
    marcado    = o Tipo 3/4 não é gerado; os itens usam o lançamento da Folha (Tipo 1)
    desmarcado = Tipo 3/4 próprio: passivo próprio (Férias/Rescisões a Pagar) E despesa própria
                 (Férias / Rescisão), com exceções automáticas e por código

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
# 0. LAYOUT DE IMPORTAÇÃO DA DOMÍNIO (não depende do plano de contas)
# =====================================================================
NOME_PADRAO_SEM_SEPARADOR = "Geral"      # só para exibição na tela
SEP_SEM = "0"                             # valor gravado no arquivo quando não há separador
LIMIAR_SIMILARIDADE = 0.60
LIM_DESC_EVENTO = 40                      # tamanho máximo do campo "Descrição" do lançamento
COMPLEMENTO_PADRAO = "<<Competencia>> - <<Descrição do Lançamento>>"

ABA_INTEGRA, ABA_EVENTO = "integra", "evento"
COL_TIPO = ("Tipo da Integração (1 - Folha mensal; 2 - Empresa; 3 - Férias; 4 - Rescisao; "
            "5 - Prov. Férias; 6 - Prov. 13)")
COLS_INTEGRA = ["Código da Empresa", "Separador", "Código Sequencial da Integração", COL_TIPO,
                "Código da Rúbrica Selecionada"]
COLS_EVENTO = ["Código da Empresa", "Separador (0 quando é sem separador)",
               "Código Sequencial da Integração", COL_TIPO, "Descrição",
               "Código da Conta Débito", "Código da Conta Crédito", "Código do Histórico", "Complemento"]

TIPOS_LAYOUT = {1: "Folha mensal", 2: "Empresa", 3: "Férias", 4: "Rescisão",
                5: "Prov. Férias", 6: "Prov. 13", 10: "Outras Informações"}
TIPO_INTEGRACAO = {  # seção do relatório -> código do layout (None = não importável)
    "Folha Normal": 1, "Adiantamento": 1, "13º Salário": 1,
    "Empresa": 2, "Férias": 3, "Rescisão": 4,
    "Provisão de Férias": 5, "Provisão de 13º": 6,
    "Outras Informações": 10,
}

# Modo "todos os eventos do cadastro": Tipo do layout -> seção interna usada pelas regras
SECAO_POR_TIPO = {1: "Folha Normal", 2: "Empresa", 3: "Férias", 4: "Rescisão", 10: "Outras Informações"}
TIPOS_CADASTRO = list(SECAO_POR_TIPO)          # Prov. Férias/13 (5 e 6) não têm catálogo de itens

SECOES = {  # título normalizado no PDF -> nome interno
    "FOLHA NORMAL": "Folha Normal", "FOLHA MENSAL": "Folha Normal",
    "ADIANTAMENTO": "Adiantamento",
    "13O SALARIO": "13º Salário", "13 SALARIO": "13º Salário", "DECIMO TERCEIRO SALARIO": "13º Salário",
    "FERIAS": "Férias", "RESCISAO": "Rescisão", "EMPRESA": "Empresa",
    "PROVISAO DE FERIAS": "Provisão de Férias", "PROVISAO FERIAS": "Provisão de Férias",
    "PROVISAO DE 13O": "Provisão de 13º", "PROVISAO DE 13O SALARIO": "Provisão de 13º",
    "PROVISAO DE 13 SALARIO": "Provisão de 13º", "PROVISAO 13O SALARIO": "Provisão de 13º",
    "INFORMACOES INSS": "Outras Informações", "INFORMACOES DO INSS": "Outras Informações",
    "OUTRAS INFORMACOES": "Outras Informações",
}
SECOES_DE_ITENS = {"Empresa", "Provisão de Férias", "Provisão de 13º", "Outras Informações"}
SECOES_CATALOGO = {"Empresa", "Outras Informações"}

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

# Catálogo do sistema Domínio: Configurar Integração > aba Outras Informações (Tipo 10)
# Cada item gera o SEU PRÓPRIO lançamento (Código Sequencial = código do item).
ITENS_OUTRAS = {
    1: ("Compensação Ded. FPAS Saldo anterior", "COMP_INSS"),
    2: ("Compensação da Ded. Sal. Família", "COMP_BENEF"),
    3: ("Compensação da Ded. Sal. Maternidade", "COMP_BENEF"),
    4: ("Compensação da Retenção", "COMP_RETENCAO"),
    5: ("Outras Compensações INSS", "COMP_INSS"),
    6: ("Contrib. valor pago coop.trab.", "COOP"),
    7: ("Compensação IRRF Cooperativa", "COMP_IRRF"),
    8: ("Compensação IRRF Pagamento indevido ou a maior", "COMP_IRRF"),
    9: ("Outras Compensações IRRF", "COMP_IRRF"),
    10: ("Compensação PIS Pagamento indevido ou a maior", "COMP_PIS"),
    11: ("Outras Compensações PIS", "COMP_PIS"),
    12: ("Compensação INSS Valor pago a maior", "COMP_INSS"),
    13: ("Compensação diferença INSS Empresa Receita Bruta", "COMP_INSS"),
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


# FGTS com ou sem ponto: "FGTS", "F.G.T.S", "F.G.T.S.", "F. G. T. S." (após norm)
RE_FGTS = re.compile(r"(?<![A-Z0-9])F ?G ?T ?S(?![A-Z0-9])")
def e_fgts(d): return bool(RE_FGTS.search(d))


def e_base(d):
    """Base de cálculo ou dedução de base (BASE INSS DE FÉRIAS, DEDUÇÃO BASE FGTS...)."""
    return tem(d, "BASE", "BASES")


def txt_cel(v) -> str:
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()


def para_int(v):
    s = txt_cel(v)
    return int(s) if s.isdigit() else s


def codigos(txt) -> set:
    """'1, 150; 8781' -> {1, 150, 8781}"""
    return {int(x) for x in re.findall(r"\d+", txt or "")}


def confere(a: str, b: str) -> bool:
    """Descrição do relatório confere com a de referência (cadastro/catálogo)?"""
    return (SequenceMatcher(None, a, b).ratio() >= LIMIAR_SIMILARIDADE
            or (min(len(a), len(b)) >= 8 and (a.startswith(b) or b.startswith(a))))

# =====================================================================
# 1b. CAIXA MISTA E LIMITE DA DESCRIÇÃO DO LANÇAMENTO
# =====================================================================
SIGLAS = {"INSS", "IRRF", "IR", "FGTS", "PIS", "COFINS", "CSLL", "DSR", "RSR", "CCT", "ACT", "RAT", "SAT",
          "FAP", "FPAS", "CPRB", "GRCS", "SENAI", "SESI", "SENAC", "SESC", "SEBRAE", "INCRA", "SENAR",
          "SEST", "SENAT", "VT", "VR", "VA", "PLR", "PPR", "CLT", "CIPA", "NF", "PJ", "EPI", "CPF",
          "CNPJ", "DCTF", "MP", "LC", "CTPS", "ISS", "RPA", "II", "III", "IV", "VI", "VII", "VIII", "IX",
          "XI", "XII"}
MINUSCULAS = {"A", "O", "AS", "OS", "E", "OU", "DE", "DA", "DO", "DAS", "DOS", "EM", "NO", "NA", "NOS",
              "NAS", "POR", "PARA", "COM", "SEM", "AO", "AOS"}
ACENTOS = {
    "FERIAS": "Férias", "SALARIO": "Salário", "SALARIOS": "Salários", "LICENCA": "Licença",
    "LICENCAS": "Licenças", "FAMILIA": "Família", "ALIMENTICIA": "Alimentícia",
    "ASSISTENCIA": "Assistência", "MEDICA": "Médica", "MEDICO": "Médico", "PREVIO": "Prévio",
    "AUXILIO": "Auxílio", "MES": "Mês", "EMPRESTIMO": "Empréstimo", "EMPRESTIMOS": "Empréstimos",
    "ESTAGIO": "Estágio", "ESTAGIARIO": "Estagiário", "PREMIO": "Prêmio", "PREMIOS": "Prêmios",
    "SAUDE": "Saúde", "ODONTOLOGICO": "Odontológico", "ODONTOLOGICA": "Odontológica",
    "PREVIDENCIA": "Previdência", "PREVIDENCIARIA": "Previdenciária", "DIFERENCA": "Diferença",
    "DIFERENCAS": "Diferenças", "PECUNIARIO": "Pecuniário", "MEDIA": "Média", "MEDIAS": "Médias",
    "ULTIMO": "Último", "ACUMULO": "Acúmulo", "HORARIO": "Horário", "EXTRAORDINARIO": "Extraordinário",
    "BENEFICIO": "Benefício", "BENEFICIOS": "Benefícios", "PERIODO": "Período", "ANUENIO": "Anuênio",
    "QUINQUENIO": "Quinquênio", "TRIENIO": "Triênio", "BIENIO": "Biênio",
    "INDENIZATORIO": "Indenizatório", "SAO": "São", "FARMACIA": "Farmácia", "SEGURANCA": "Segurança",
    "CREDITO": "Crédito", "DEBITO": "Débito", "DEPOSITO": "Depósito", "FUNCIONARIO": "Funcionário",
    "MAXIMO": "Máximo", "MINIMO": "Mínimo", "NUMERO": "Número", "TECNICO": "Técnico",
    "CONVENIO": "Convênio", "MAE": "Mãe", "CONJUGE": "Cônjuge", "UTEIS": "Úteis", "UTIL": "Útil",
    "PROPRIO": "Próprio", "AGUA": "Água", "ENERGIA": "Energia", "VALIDO": "Válido",
    "LIQUIDO": "Líquido", "BASICA": "Básica", "AUTONOMO": "Autônomo", "VARIAVEL": "Variável",
}
SUFIXOS = (("coes", "ções"), ("cao", "ção"), ("soes", "sões"), ("sao", "são"))
RE_PALAVRA = re.compile(r"[0-9A-Za-zÀ-ÖØ-öø-ÿ]+")


def _sufixo(w: str) -> str:
    for a, b in SUFIXOS:
        if w.endswith(a) and len(w) > len(a) + 1:
            return w[:-len(a)] + b
    return w


def caixa_mista(txt) -> str:
    """'LICENCA REMUNERADA FERIAS' -> 'Licença Remunerada Férias'; siglas em maiúsculas."""
    s = re.sub(r"(?i)\bpro[\s\-]?labore\b", "Pró-Labore", txt_cel(txt))
    out, pos, primeira = [], 0, True
    for m in RE_PALAVRA.finditer(s):
        out.append(s[pos:m.start()])
        w, up, prox = m.group(0), norm(m.group(0)), s[m.end():m.end() + 1]
        if w.isdigit():
            r = w
        elif re.fullmatch(r"\d+[OoAa]", w):                      # 13O -> 13º / 13A -> 13ª
            r = w[:-1] + ("º" if w[-1] in "Oo" else "ª")
        elif any(ch.isdigit() for ch in w):
            r = w.upper()
        elif up in SIGLAS:
            r = up
        elif len(w) == 1 and up in "SCP" and prox == "/":         # s/ c/ p/
            r = w.lower()
        elif up in MINUSCULAS and not primeira:
            r = w.lower()
        elif w.isascii():
            r = ACENTOS.get(up) or _sufixo(w.lower()).capitalize()
        else:
            r = w.capitalize()
        out.append(r)
        pos, primeira = m.end(), False
    out.append(s[pos:])
    return re.sub(r"\s+", " ", "".join(out)).strip()


def cortar(txt: str, lim: int = LIM_DESC_EVENTO) -> str:
    txt = re.sub(r"\s+", " ", txt).strip()
    if len(txt) <= lim:
        return txt
    c = txt[:lim]
    if txt[lim] != " ":                       # evita cortar no meio da palavra, se der
        esp = c.rfind(" ")
        if esp >= lim - 8:
            c = c[:esp]
    return c.rstrip(" .,-/")


def descricao_lancamento(codigo, descricao) -> str:
    return cortar(f"{codigo} - {caixa_mista(descricao)}")

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
    # Despesa do cálculo de Rescisão (Tipo 4 próprio). Conta "Rescisões" se existir; senão Indenizações.
    "D_RESC": (G, "Rescisões (despesa)", [
        (["RESCIS*"], ["PAGAR", "PROVIS*", "FGTS", "INSS", "MULTA"]),
        (["INDENIZAC*"], []), (["AVISO PREVIO"], [])]),
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
    "ISS_REC": (PA, "ISS Retido a Recolher", [
        (["ISS", "RECOLHER"], ["PROVIS*"]), (["ISS", "RETIDO"], []), (["ISSQN"], []), (["ISS"], ["PROVIS*"])]),
    "REPASSES": (PA, "Outros Descontos a Repassar", [
        (["ASSOCIAC*"], []), (["REPASS*"], []), (["OUTRAS", "CONSIGNAC*"], []), (["OUTRAS", "PAGAR"], [])]),
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
    "IRRF_COMP": (AT, "IRRF a Compensar", [
        (["IRRF", "COMPENSAR"], []), (["IRRF", "RECUPERAR"], []),
        (["IMPOSTO DE RENDA", "COMPENSAR"], []), (["IMPOSTO DE RENDA", "RECUPERAR"], []),
        (["IR", "COMPENSAR"], []), (["IR", "RECUPERAR"], [])]),
    "PIS_COMP": (AT, "PIS a Compensar", [(["PIS", "COMPENSAR"], []), (["PIS", "RECUPERAR"], [])]),
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

# Despesa própria do cálculo de Férias/Rescisão (quando a flag "mesma configuração" está desmarcada).
# Só as despesas de natureza salarial são trocadas; o resto é exceção automática
# (13º, férias, aviso/indenizações, FGTS, benefícios, pró-labore e contas patrimoniais).
DESPESA_TROCAVEL = ("SALARIOS", "HE", "PREMIOS", "COMISSOES")
DESPESA_SECAO = {"Férias": ["FERIAS"], "Rescisão": ["D_RESC", "INDENIZ"]}


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


# ---- Modo "todos os eventos cadastrados" (sem relatório de pendências)
RE_CAB_EMP = re.compile(r"^\s*(?:Empresa\s*:\s*)?(\d{1,6})\s*-\s*(.+?)\s*$", re.I)


@st.cache_data(show_spinner=False)
def empresa_do_cadastro(dados: bytes, nome: str):
    """Código e nome da empresa lidos do cabeçalho do cadastro geral de rubricas."""
    linhas = [l for l in extrair_linhas(dados, nome) if l.strip()]
    for l in linhas[:15]:
        t = re.split(r"P[áa]gina", l, flags=re.I)[0].strip()
        if RE_RUB.match(t):                       # já chegou nas rubricas
            break
        m = RE_EMPRESA.match(t) or RE_CAB_EMP.match(t)
        if m:
            return {"codigo": m.group(1), "nome": m.group(2).strip()}
    primeira = re.split(r"P[áa]gina", linhas[0], flags=re.I)[0].strip() if linhas else ""
    return {"codigo": "", "nome": primeira}


def parse_separadores(txt: str):
    """'1=Administrativo; 2=Produção' -> [('1','Administrativo'), ('2','Produção')]"""
    out = []
    for parte in re.split(r"[;\n]+", txt or ""):
        m = re.match(r"\s*(\d+)\s*(?:[=:\-]\s*(.*))?$", parte)
        if m:
            out.append((m.group(1), (m.group(2) or "").strip() or f"Separador {m.group(1)}"))
    return out


def itens_do_cadastro(cad: dict, tipos, separadores):
    """Itens a contabilizar sem o relatório de pendências:
    Tipos 1/3/4 = todas as rubricas do cadastro; Tipo 2 = catálogo Empresa;
    Tipo 10 = catálogo Outras Informações (itens 1 a 13), CADA ITEM EM LANÇAMENTO PRÓPRIO."""
    fontes = {
        "Empresa": {c: v[0] for c, v in ITENS_EMPRESA.items()},
        "Outras Informações": {c: v[0] for c, v in ITENS_OUTRAS.items()},
    }
    rubricas = {c: r["descricao"] for c, r in cad.items()}
    itens = []
    for sep_cod, sep_nome in (separadores or [("", "")]):
        for t in sorted(tipos):
            secao = SECAO_POR_TIPO[t]
            for cod, desc in sorted(fontes.get(secao, rubricas).items()):
                itens.append({"secao": secao, "sep_tipo": "Informado" if sep_cod else "",
                              "sep_cod": sep_cod, "sep_nome": sep_nome,
                              "codigo": int(cod), "descricao": desc})
    return itens


def unificar_com_folha(itens, mesma_ferias: bool, mesma_rescisao: bool):
    """'Usar a mesma configuração da folha normal para Férias/Rescisão' (Domínio):
    os itens dessas seções passam a usar o lançamento da Folha (Tipo 1). Não duplica rubrica
    que já está no Tipo 1 do mesmo separador. Retorna (itens, quantidade unificada)."""
    unir = {s for s, f in (("Férias", mesma_ferias), ("Rescisão", mesma_rescisao)) if f}
    if not unir:
        return itens, 0
    out, vistos, movidos = [], set(), 0
    for it in itens:
        if it["secao"] in unir:
            continue
        out.append(it)
        if TIPO_INTEGRACAO.get(it["secao"]) == 1:
            vistos.add((it["sep_cod"], it["codigo"]))
    for it in itens:
        if it["secao"] not in unir:
            continue
        movidos += 1
        k = (it["sep_cod"], it["codigo"])
        if k not in vistos:
            vistos.add(k)
            out.append({**it, "secao": "Folha Normal"})
    return out, movidos

# =====================================================================
# 5. REGRAS DETERMINÍSTICAS
# =====================================================================
def inferir_tipo(d):
    if tem(d, "MATERN*") and not tem(d, "DESC*"): return "Provento"
    if tem(d, "INSS") and tem(d, "A MAIOR") and not tem(d, "DESCONTO"): return "Provento"
    if tem(d, "REEMBOLSO", "DEV", "DEVOLUCAO", "RESTITUI*"): return "Provento"
    if e_fgts(d) or tem(d, "CONTRIBUICAO SOCIAL", "CONTRIB SOCIAL", "BASE"): return "Informativa"
    if tem(d, "DESC*") or (tem(d, "ESTOURO", "TROCO") and tem(d, "ANTERIOR")): return "Desconto"
    if tem(d, "INSS", "IRRF", "IMPOSTO DE RENDA", "PENSAO", "EMPREST*", "CONTRIB*",
           "MENSALIDADE", "FALTA*", "ATRASO*"): return "Desconto"
    if tem(d, "VALE TRANSPORTE") and "%" in d: return "Desconto"
    return "Provento"


def tipo_rubrica(codigo, d, cad):
    reg = cad.get(codigo)
    if reg:
        if confere(d, reg["desc_norm"]):
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


def e_maternidade(d):
    """SALARIO MATERNIDADE / LIC.MATERN / LIC.MAT.INSS / SAL MAT"""
    return tem(d, "MATERN*") or (tem(d, "LIC", "LICENC*", "SAL") and tem(d, "MAT"))


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
    # Maternidade "INSS": a empresa paga e abate o valor do INSS a recolher
    if e_maternidade(d) and tem(d, "INSS") and not tem(d, "DESC*", "DEDUC*"):
        return ["INSS_REC"], L, "Maternidade paga pela empresa e deduzida do INSS a recolher", False
    if tem(d, "INSS", "IRRF") and tem(d, "A MAIOR", "DIF*", "DEVOL*", "RESTITUI*"):
        return (["IRRF_REC"] if tem(d, "IRRF") else ["INSS_REC"]), L, "Restituição de retenção a maior", False
    if tem(d, "SALARIO FAMILIA", "SAL FAM*"):
        return ["BENEF_INSS"], L, "Salário-família → conta-ponte até a compensação na guia", False
    if e_maternidade(d) and not tem(d, "DESC*", "DEDUC*", "PRORROG*", "EMPREGADOR"):
        return ["BENEF_INSS"], L, "Salário-maternidade reembolsável → conta-ponte", False
    dre = alvo_dre_provento(d)
    if baixa_prov and secao in ("Férias", "Rescisão", "13º Salário") and dre[0] in ("FERIAS", "DECIMO"):
        return (["P_PROV_FER"] if dre[0] == "FERIAS" else ["P_PROV_13"]), L, "Baixa contra a provisão", False
    return dre, L, "", False


def classificar_desconto(d, secao):
    L = passivo_secao(secao, d)
    def r(c, obs="", rev=False): return L, c, obs, rev
    if tem(d, "ESTOURO", "TROCO"): return r(["ADIANT_SAL"], "Baixa de estouro/troco anterior")
    if tem(d, "LIQUIDO") and tem(d, "RESCIS*"):
        return r(["RESC_PAGAR"], "Líquido da rescisão pago à parte — transfere p/ Rescisões a Pagar", True)
    if tem(d, "PENSAO*"): return r(["PENSAO"], "Repasse ao beneficiário", True)
    if tem(d, "SAL FAM*", "SALARIO FAMILIA"):
        return r(["BENEF_INSS"], "Estorno de salário-família pago a maior", True)
    if tem(d, "INSS"): return r(["INSS_REC"])
    if tem(d, "IRRF", "IMPOSTO DE RENDA", "IR"): return r(["IRRF_REC"])
    if tem(d, "ISS", "ISSQN"): return r(["ISS_REC"], "ISS retido do autônomo", True)
    if tem(d, "ADIANT*", "ADTO", "FERIAS PAGAS") or d in ("VALE", "VALES", "DESC VALE"):
        return r(alvo_adiant(d), "Baixa do adiantamento")
    if tem(d, "SINDICA*", "ASSISTENCIAL", "CONFEDERATIVA", "NEGOC*"): return r(["SIND_REC"])
    if tem(d, "EMPREST*", "EMP", "CONSIG*", "CRED TRAB"): return r(["EMPREST"], "Consignado a repassar", True)
    if tem(d, "ASSOCIAC*"): return r(["REPASSES"], "Mensalidade descontada a repassar", True)
    if tem(d, "PLANO", "ODONTO*", "COPARTICIP*", "ASSISTENCIA MEDICA", "FARMACIA", "SAUDE"):
        return r(["ASSIST"], "Recuperação do custo do benefício")
    if tem(d, "VALE TRANSPORTE", "VT"): return r(["VT"], "Recuperação dos 6% do VT")
    if tem(d, "REFEICAO", "ALIMENTACAO", "VR", "CESTA", "SUPERMERCADO"):
        return r(["ALIM"], "Recuperação do custo do benefício")
    if tem(d, "SEGURO*"): return r(["SEGURO", "ASSIST"], "Recuperação do custo do benefício")
    if tem(d, "AVISO PREVIO", "MULTA", "ESTABILIDADE", "INDENIZ*"):
        return r(["INDENIZ", "SALARIOS"], "Indenização devida pelo empregado", True)
    if e13(d): return r(["DECIMO", "SALARIOS"], "Redução do custo de 13º")
    if e_maternidade(d): return r(["BENEF_INSS"], "Estorno de salário-maternidade pago a maior", True)
    if tem(d, "FERIAS", "ABONO"): return r(["FERIAS", "SALARIOS"], "Redução do custo de férias")
    if tem(d, "COMISS*"): return r(["COMISSOES", "PREMIOS", "SALARIOS"], "Estorno de comissões")
    if tem(d, "PREMIO*", "GRATIFIC*"): return r(["PREMIOS", "SALARIOS"], "Estorno de prêmio/gratificação")
    if tem(d, "FALTA*", "ATRASO*", "DSR", "HORAS", "DIAS", "PAGO A MAIOR", "INATIV*", "SUSPENS*", "AFASTAD*"):
        return r(["SALARIOS"], "Redução do custo de salários")
    return r([], "Desconto sem regra — definir conta manualmente", True)


def classificar_informativa(d, tipo):
    """Regra: informativa só integra se a descrição tiver FGTS (com ou sem ponto)
    e não for base/dedução de base. Todo o resto: não integrar."""
    if e_base(d):
        return [], [], "Base/dedução de base — não gera lançamento", True
    if not e_fgts(d):
        return [], [], "Informativa sem FGTS na descrição — não integrar", True
    dre = ["INDENIZ", "FGTS"] if tem(d, "40%", "20%") else ["FGTS"]
    if tem(d, "A MAIOR") or tipo == "Inf. dedutora":
        return ["FGTS_REC"], dre, "Estorno de FGTS a maior", True
    return dre, ["FGTS_REC"], "Encargo FGTS", False


def despesa_da_secao(alvos, secao, codigo, excecoes):
    """Cálculo de Férias/Rescisão com configuração própria (flag desmarcada):
    a despesa salarial (Salários, Horas Extras, Prêmios, Comissões) vira a despesa do cálculo
    — Férias → despesa de Férias | Rescisão → despesa de Rescisão/Indenizações.
    Exceções: códigos informados pelo usuário e tudo que não for despesa salarial
    (13º, férias, aviso, FGTS, benefícios, pró-labore, contas patrimoniais)."""
    if secao not in DESPESA_SECAO or not alvos or alvos[0] not in DESPESA_TROCAVEL:
        return alvos, ""
    if codigo in excecoes:
        return alvos, f"Exceção: mantém a despesa da folha no cálculo de {secao}"
    novo = DESPESA_SECAO[secao]
    return novo + [a for a in alvos if a not in novo], f"Despesa do cálculo de {secao}"


def classificar_provisao(d, secao):
    """Itens das seções Provisão de Férias / Provisão de 13º."""
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


# ---- Aba Empresa (Tipo 2)
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
    if cat and confere(d, norm(cat[0])):
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


# ---- Aba Outras Informações (Tipo 10) — cada item = 1 lançamento próprio
OBS_OUTRAS = {
    "COMP_BENEF": ("Compensação sal.-família/maternidade — baixa da conta-ponte", False),
    "COMP_INSS": ("Compensação de INSS na guia", False),
    "COMP_RETENCAO": ("Compensação da retenção de INSS (NF)", False),
    "COOP": ("Contribuição s/ cooperativa de trabalho — confirme a conta", True),
    "COMP_IRRF": ("Compensação de IRRF", False),
    "COMP_PIS": ("Compensação de PIS s/ folha — confirme a conta", True),
    None: ("Item de Outras Informações sem regra — definir manualmente", True),
}
CONTAS_OUTRAS = {  # D INSS a Recolher × C INSS a Compensar (padrão da tela Configurar Integração)
    "COMP_BENEF": (["INSS_REC"], ["BENEF_INSS"]),
    "COMP_INSS": (["INSS_REC"], ["INSS_COMP", "BENEF_INSS"]),
    "COMP_RETENCAO": (["INSS_REC"], ["INSS_COMP", "BENEF_INSS"]),
    "COOP": (["INSS"], ["INSS_REC"]),
    "COMP_IRRF": (["IRRF_REC"], ["IRRF_COMP"]),
    "COMP_PIS": (["PIS_REC"], ["PIS_COMP"]),
}


def inferir_item_outras(d):
    if tem(d, "FAMILIA", "MATERN*"): return "COMP_BENEF"
    if tem(d, "COOP*") and not tem(d, "IRRF", "IR"): return "COOP"
    if tem(d, "IRRF", "IR"): return "COMP_IRRF"
    if tem(d, "PIS"): return "COMP_PIS"
    if tem(d, "RETENC*", "RETIDO", "RETIDA*"): return "COMP_RETENCAO"
    if tem(d, "INSS", "FPAS", "COMPENSAC*"): return "COMP_INSS"
    return None


def classificar_item_outras(codigo, d):
    cat = ITENS_OUTRAS.get(codigo)
    if cat and confere(d, norm(cat[0])):
        nat, origem = cat[1], "Catálogo Outras Informações"
    else:
        nat = inferir_item_outras(d)
        origem = (f"Inferido — no catálogo o item {codigo} é '{cat[0]}'" if cat
                  else f"Inferido — item {codigo} fora do catálogo")
    deb, cred = CONTAS_OUTRAS.get(nat, ([], []))
    obs, rev = OBS_OUTRAS.get(nat, OBS_OUTRAS[None])
    return list(deb), list(cred), obs, rev, origem, nat

# =====================================================================
# 6. CONFIGURAÇÃO POR EMPRESA (.json)
# =====================================================================
CFG_WIDGETS = (("historico", "w_hist"), ("baixa_prov", "w_baixa"), ("socio_adm", "w_socio"),
               ("bloqueio_extra", "w_bloq"), ("empresa", "w_cod"), ("complemento", "w_compl"),
               ("tipos_cadastro", "w_tipos_cad"), ("separadores_cadastro", "w_seps_cad"),
               ("mesma_rescisao", "w_mesma_resc"), ("mesma_ferias", "w_mesma_fer"),
               ("excecoes_ferias", "w_exc_fer"), ("excecoes_rescisao", "w_exc_resc"))


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
    for chave, wk in CFG_WIDGETS:
        if chave in cfg:
            st.session_state[wk] = cfg[chave]
    for t, v in cfg.get("hist_tipo", {}).items():
        st.session_state[f"w_hist_{t}"] = str(v)


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
    st.session_state.setdefault("modo_cad", False)
    if st.session_state["modo_cad"]:
        st.success("Modo: **todos os eventos cadastrados** (cadastro geral + plano de contas)")
        if st.button("↩️ Voltar ao modo pendências", use_container_width=True):
            st.session_state["modo_cad"] = False
            st.rerun()
    else:
        if st.button("🗂️ Contabilizar todos os eventos cadastrados", use_container_width=True,
                     help="Ignora o relatório de pendências e gera o lançamento de todas as rubricas "
                          "do cadastro geral + itens das abas Empresa e Outras Informações."):
            st.session_state["modo_cad"] = True
            st.rerun()
    modo_cad = st.session_state["modo_cad"]

    f_pend = None
    if not modo_cad:
        f_pend = st.file_uploader("Rubricas/Itens não configurados", type=["pdf", "txt"])
    f_cad = st.file_uploader("Cadastro geral de rubricas", type=["pdf", "txt"])
    f_plano = st.file_uploader("Plano de contas", type=["xlsx", "xls", "csv", "txt"])

    tipos_cad, seps_cad = [], ""
    if modo_cad:
        st.session_state.setdefault("w_tipos_cad", TIPOS_CADASTRO)
        st.session_state.setdefault("w_seps_cad", "")
        st.session_state["w_tipos_cad"] = [int(t) for t in st.session_state["w_tipos_cad"]
                                           if str(t).isdigit() and int(t) in TIPOS_CADASTRO]
        tipos_cad = st.multiselect("Tipos da Integração a gerar", TIPOS_CADASTRO, key="w_tipos_cad",
                                   format_func=lambda t: f"{t} - {TIPOS_LAYOUT[t]}")
        seps_cad = st.text_area("Separadores (opcional)", key="w_seps_cad",
                                placeholder="1=Administrativo; 2=Produção",
                                help="Sem separador = folha centralizada (Separador 0). "
                                     "Informando, todos os eventos são replicados para cada separador.")
        st.caption("Tipo 10: cada item da aba Outras Informações (1 a 13) gera o seu próprio lançamento.")

    st.header("2. Regras")
    for k, v in (("w_hist", ""), ("w_baixa", False), ("w_socio", True), ("w_bloq", ""),
                 ("w_compl", COMPLEMENTO_PADRAO), ("w_mesma_resc", False), ("w_mesma_fer", False),
                 ("w_exc_fer", ""), ("w_exc_resc", "")):
        st.session_state.setdefault(k, v)
    st.markdown("**Usar a mesma configuração da folha normal para**")
    mesma_resc = st.checkbox("Rescisão", key="w_mesma_resc",
                             help="Igual à Domínio. Marcado: a Rescisão usa os lançamentos da Folha "
                                  "(Tipo 1) e o Tipo 4 não é gerado. Desmarcado: Tipo 4 próprio, com "
                                  "crédito em Rescisões a Pagar e despesa de Rescisão.")
    mesma_fer = st.checkbox("Férias", key="w_mesma_fer",
                            help="Igual à Domínio. Marcado: as Férias usam os lançamentos da Folha "
                                 "(Tipo 1) e o Tipo 3 não é gerado. Desmarcado: Tipo 3 próprio, com "
                                 "crédito em Férias a Pagar e despesa de Férias.")
    with st.expander("Despesa própria de Férias/Rescisão — exceções",
                     expanded=not (mesma_resc and mesma_fer)):
        st.caption("Vale para o cálculo **desmarcado** acima. Salários, horas extras, adicionais, "
                   "prêmios e comissões passam a debitar a despesa de Férias (Tipo 3) ou de Rescisão "
                   "(Tipo 4); descontos de faltas/horas creditam a mesma despesa. Exceções automáticas: "
                   "13º, férias, aviso prévio/indenizações, FGTS, benefícios, pró-labore e contas "
                   "patrimoniais continuam como na folha.")
        exc_fer = st.text_input("Exceções Férias (códigos que mantêm a despesa da folha)",
                                key="w_exc_fer", disabled=mesma_fer, placeholder="ex.: 1, 150, 8781")
        exc_resc = st.text_input("Exceções Rescisão (códigos que mantêm a despesa da folha)",
                                 key="w_exc_resc", disabled=mesma_resc, placeholder="ex.: 9179, 9180")
    excecoes = {"Férias": codigos(exc_fer), "Rescisão": codigos(exc_resc)}
    baixa_prov = st.checkbox("Baixar férias/13º pagos contra a provisão", key="w_baixa",
                             help="Deixe desmarcado se a Domínio já gera o 'Valor Estorno Provisão'.")
    socio_adm = st.checkbox("Pró-labore e encargos do sócio sempre em Despesas Administrativas", key="w_socio")
    bloq_extra = st.text_input("Contas extras de colaborador (reduzidos, separados por vírgula)", key="w_bloq",
                               help="Além das detectadas automaticamente; itens patronais nunca poderão usá-las.")

    st.header("3. Layout de importação")
    historico = st.text_input("Código do Histórico (padrão)", key="w_hist")
    hist_tipo = {}
    with st.expander("Histórico por Tipo da Integração (opcional)"):
        for t, n in TIPOS_LAYOUT.items():
            st.session_state.setdefault(f"w_hist_{t}", "")
            hist_tipo[t] = st.text_input(f"{t} - {n}", key=f"w_hist_{t}",
                                         placeholder=historico or "usa o padrão")
    complemento = st.text_input("Complemento", key="w_compl")
    st.caption(f"1 lançamento por rubrica/item · Código Sequencial = código da rubrica/item · "
               f"Descrição = 'código - descrição' em caixa mista, máx. {LIM_DESC_EVENTO} caracteres.")

if not (f_cad and f_plano and (f_pend or modo_cad)):
    st.info("Envie o cadastro geral de rubricas e o plano de contas"
            + ("." if modo_cad else " e o relatório de pendências — ou use o botão "
               "'Contabilizar todos os eventos cadastrados'.")
            + " Opcional: carregue a configuração salva da empresa.")
    st.stop()

cfg = st.session_state.get("cfg", {})
try:
    nome_cad, cad = parse_cadastro(f_cad.getvalue(), f_cad.name)
    raw = ler_tabela_bruta(f_plano.getvalue(), f_plano.name)
    if modo_cad:
        if not tipos_cad:
            st.info("Escolha ao menos um Tipo da Integração na barra lateral.")
            st.stop()
        cab = empresa_do_cadastro(f_cad.getvalue(), f_cad.name)
        itens = itens_do_cadastro(cad, tipos_cad, parse_separadores(seps_cad))
    else:
        cab, itens = parse_pendencias(f_pend.getvalue(), f_pend.name)
except Exception as e:
    st.error(f"Erro na leitura: {e}")
    st.stop()

itens, n_unif = unificar_com_folha(itens, mesma_fer, mesma_resc)
if n_unif:
    nomes_unif = " e ".join(n for n, f in (("Férias", mesma_fer), ("Rescisão", mesma_resc)) if f)
    st.info(f"🔗 {nomes_unif}: mesma configuração da folha normal — {n_unif} item(ns) usam o "
            "lançamento da Folha (Tipo 1); o Tipo 3/4 correspondente não será gerado.")

if modo_cad and not cad and any(t in (1, 3, 4) for t in tipos_cad):
    st.error("Nenhuma rubrica lida no cadastro geral — não há eventos para os Tipos 1, 3 e 4.")
    st.stop()
if not itens:
    st.error("Nenhuma rubrica/item encontrado" + (" no cadastro." if modo_cad else " no relatório de pendências."))
    st.stop()
if not cad and not modo_cad:
    st.warning("Nenhuma rubrica lida no cadastro geral — o Tipo de todas será inferido pela descrição.")

origem_bytes = f_cad.getvalue() if modo_cad else f_pend.getvalue()
pid = hashlib.md5(origem_bytes + (b"|cad" if modo_cad else b"|pend")).hexdigest()
if st.session_state.get("pend_id") != pid:
    st.session_state["pend_id"] = pid
    if not cfg.get("empresa"):
        st.session_state["w_cod"] = cab["codigo"]
st.session_state.setdefault("w_cod", cab["codigo"])
cod_empresa = st.sidebar.text_input("Código da empresa na Domínio", key="w_cod")
if cfg.get("empresa") and cab["codigo"] and str(cfg["empresa"]) != cab["codigo"]:
    st.warning(f"⚠️ A configuração carregada é da empresa {cfg['empresa']}, "
               f"mas o {'cadastro' if modo_cad else 'relatório'} é da empresa {cab['codigo']}.")

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
m1.metric("Eventos no lote" if modo_cad else "Itens pendentes", len(itens))
m2.metric("Rubricas no cadastro", len(cad))
m3.metric("Contas no plano", len(plano))
m4.metric("Empresa", f'{cab["codigo"]} - {cab["nome"][:25]}')
if modo_cad:
    cont = pd.Series([TIPO_INTEGRACAO[i["secao"]] for i in itens]).value_counts().sort_index()
    st.caption("Eventos por Tipo da Integração: "
               + " · ".join(f"{t} - {TIPOS_LAYOUT[t]}: {n}" for t, n in cont.items()))

# ---------- 2. Separadores ----------
st.subheader("2. Separador → grupo de resultado")
usa_sep = any(i["sep_cod"] for i in itens)

cands = grupos_resultado(plano, raizes["RESULTADO"], pontuado)
rotulos = [f"{g['classif']} — {g['descricao']}" + (f"  ({g['caminho']})" if g["caminho"] else "")
           for g in cands]
rot2pref = {r: g["classif"] for r, g in zip(rotulos, cands)}
pref2rot = {v: k for k, v in rot2pref.items()}
if not cands:
    st.warning("Nenhum grupo de resultado com conta de Salários foi detectado — informe a máscara manualmente.")

if usa_sep:
    tipos = sorted({i["sep_tipo"] for i in itens if i["sep_tipo"]})
    st.success(f"Folha com separador ({', '.join(tipos)}).")
else:
    st.warning(f"Nenhuma quebra por Centro de Custo / Filial / Serviço → folha centralizada "
               f"({NOME_PADRAO_SEM_SEPARADOR}, Separador = {SEP_SEM} no arquivo). "
               "Escolha o grupo contábil em que a folha inteira será classificada.")

seps = {}
for it in itens:
    seps.setdefault(it["sep_cod"] or SEP_SEM,
                    it["sep_nome"] if it["sep_cod"] else f"sem separador ({NOME_PADRAO_SEM_SEPARADOR})")

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
    k = it["sep_cod"] or SEP_SEM
    prefixo, nat = mapa_grupos[k], None
    if sec == "Empresa":
        tipo = "Item (Empresa)"
        deb, cred, obs, rev, origem, nat = classificar_item_empresa(it["codigo"], d, baixa_prov)
    elif sec == "Outras Informações":
        tipo = "Item (Outras Informações)"
        deb, cred, obs, rev, origem, nat = classificar_item_outras(it["codigo"], d)
    elif sec in SECOES_DE_ITENS:
        tipo, origem = f"Item ({sec})", "Seção"
        deb, cred, obs, rev = classificar_provisao(d, sec)
    else:
        tipo, origem = tipo_rubrica(it["codigo"], d, cad)
        if tipo in ("Informativa", "Inf. dedutora"):
            deb, cred, obs, rev = classificar_informativa(d, tipo)
        elif tipo == "Desconto":
            deb, cred, obs, rev = classificar_desconto(d, sec)
        else:
            deb, cred, obs, rev = classificar_provento(d, sec, baixa_prov)
        # Férias/Rescisão com configuração própria: a despesa também muda (com exceções)
        if sec in DESPESA_SECAO:
            exc = excecoes.get(sec, set())
            deb, o1 = despesa_da_secao(deb, sec, it["codigo"], exc)
            cred, o2 = despesa_da_secao(cred, sec, it["codigo"], exc)
            obs = " | ".join(x for x in (obs, o1, o2) if x)
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
    v = txt_cel(row.get("Conta definida"))
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
ST_OK, ST_REV, ST_PEND, ST_NAO, ST_MAN = "✅ OK", "⚠️ Revisar", "❌ Pendente", "⏭️ Não integrar", "📝 Manual"
linhas = []
for r in regras:
    it, sec = r["it"], r["sec"]
    tipo_int = TIPO_INTEGRACAO.get(sec)
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
    if usa_sep and r["k"] == SEP_SEM and sec not in SECOES_DE_ITENS:
        alertas.append(f"separador {SEP_SEM} numa folha com separador — confirme")
    if not r["deb"] and not r["cred"]:
        status = ST_PEND if (sec in SECOES_CATALOGO and r["nat"] != "ISENCAO") else ST_NAO
    elif tipo_int is None:
        status = ST_MAN
        alertas.append("seção sem Tipo da Integração no layout — configurar na Domínio")
    elif not dc or not cc_:
        status = ST_PEND
    elif r["rev"] or alertas or r["origem"].startswith("Inferido"):
        status = ST_REV
    else:
        status = ST_OK
    linhas.append({
        "Status": status, "Seção": sec, "Tipo Integração": str(tipo_int) if tipo_int else "—",
        "Separador": r["k"], "Nome do separador": it["sep_nome"] or NOME_PADRAO_SEM_SEPARADOR,
        "Grupo": r["prefixo"], "Código": it["codigo"], "Descrição": it["descricao"],
        "Descrição Lançamento": descricao_lancamento(it["codigo"], it["descricao"]),
        "Tipo": r["tipo"], "Origem do tipo": r["origem"], "Débito": dc, "Desc. Débito": dd,
        "Crédito": cc_, "Desc. Crédito": cd,
        "Observação": " | ".join([r["obs"]] + alertas).strip(" |"),
        "Exportar": bool(dc and cc_ and tipo_int),
    })
df = pd.DataFrame(linhas)

n_div = int(df["Origem do tipo"].str.startswith("Inferido").sum())
ne, nc = norm(cab["nome"]), norm(nome_cad)
if not modo_cad and ne and nc and ne not in nc and nc not in ne:
    if n_div:
        st.warning(f"⚠️ Cadastro de outra empresa (**{nome_cad}**): {n_div} item(ns) com código "
                   "divergente/ausente — Tipo inferido pela descrição.")
    else:
        st.info(f"Cadastro do modelo **{nome_cad}**, mas todas as rubricas conferem por código + descrição.")

# ---------- 4. Conferência ----------
st.subheader("4. Conferência")
st.caption("Edições aqui valem só para este lote. Para correções permanentes de contas, use o Mapa de contas.")
mcols = st.columns(5)
for col, s in zip(mcols, (ST_OK, ST_REV, ST_PEND, ST_NAO, ST_MAN)):
    col.metric(s, int((df["Status"] == s).sum()))

EDITAVEIS = ("Débito", "Crédito", "Descrição Lançamento", "Exportar")
ed = st.data_editor(
    df, key=f"editor_{md5(df.to_json())}", hide_index=True,
    disabled=[c for c in df.columns if c not in EDITAVEIS],
    column_config={"Exportar": st.column_config.CheckboxColumn(),
                   "Débito": st.column_config.TextColumn(), "Crédito": st.column_config.TextColumn(),
                   "Descrição Lançamento": st.column_config.TextColumn(
                       help=f"Máximo {LIM_DESC_EVENTO} caracteres")},
)

idx_plano = plano.drop_duplicates("reduzido").set_index("reduzido")


def validar(row):
    msgs, vals = [], {}
    if TIPO_INTEGRACAO.get(row["Seção"]) is None:
        msgs.append("seção sem Tipo da Integração no layout")
    for lado in ("Débito", "Crédito"):
        c = txt_cel(row[lado])
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
    dl = txt_cel(row["Descrição Lançamento"])
    if not dl:
        msgs.append("Descrição do lançamento vazia")
    elif len(dl) > LIM_DESC_EVENTO:
        msgs.append(f"Descrição com {len(dl)} caracteres (máx. {LIM_DESC_EVENTO})")
    return "; ".join(msgs)


ed = ed.copy()
ed["Validação"] = ed.apply(validar, axis=1)
exportar = ed[ed["Exportar"] & (ed["Validação"] == "")].reset_index(drop=True)
bloqueadas = ed[ed["Exportar"] & (ed["Validação"] != "")]
if not bloqueadas.empty:
    st.error(f"{len(bloqueadas)} linha(s) marcadas para exportar com erro — ficarão fora do lote.")
    st.dataframe(bloqueadas[["Seção", "Separador", "Código", "Descrição", "Validação"]], hide_index=True)

manuais = ed[ed["Status"] == ST_MAN]
if not manuais.empty:
    st.warning(f"📝 {len(manuais)} item(ns) de seção sem código no layout — configure-os manualmente na Domínio.")
    st.dataframe(manuais[["Seção", "Separador", "Código", "Descrição", "Débito", "Desc. Débito",
                          "Crédito", "Desc. Crédito"]], hide_index=True)

# ---------- 5. Montagem do layout (1 lançamento por rubrica/item) ----------
st.subheader("5. Arquivo de importação")
emp = para_int(cod_empresa)

exp = exportar.copy()
exp["_sep"] = [para_int(v) for v in exp["Separador"]]
exp["_s"] = [str(v) for v in exp["_sep"]]
exp["_tipo"] = [int(TIPO_INTEGRACAO[s]) for s in exp["Seção"]]
exp["_cod"] = [int(c) for c in exp["Código"]]
exp["_deb"] = [txt_cel(v) for v in exp["Débito"]]
exp["_cred"] = [txt_cel(v) for v in exp["Crédito"]]
exp["_hist"] = [txt_cel(hist_tipo.get(t)) or txt_cel(historico) for t in exp["_tipo"]]
exp["_desc"] = [txt_cel(v) for v in exp["Descrição Lançamento"]]

# Mesmo código em Tipos diferentes e em seções do mesmo Tipo é permitido.
# Só remove linhas 100% idênticas (ex.: a mesma rubrica em Folha Normal e Adiantamento com as mesmas contas).
chave_ev = ["_s", "_tipo", "_cod", "_deb", "_cred", "_hist", "_desc"]
exp = (exp.drop_duplicates(chave_ev)
       .sort_values(["_s", "_tipo", "_cod"], kind="stable").reset_index(drop=True))
repetidas = exp[exp.duplicated(["_s", "_tipo", "_cod"], keep=False)]
if not repetidas.empty:
    st.info(f"ℹ️ {len(repetidas)} lançamento(s) com o mesmo código no mesmo Tipo/Separador e contas "
            "diferentes — exportados conforme permitido.")
    st.dataframe(repetidas[["Seção", "Separador", "Código", "Descrição Lançamento", "Débito", "Crédito"]],
                 hide_index=True)
n = len(exp)

evento = pd.DataFrame({
    COLS_EVENTO[0]: [emp] * n,
    COLS_EVENTO[1]: exp["_sep"].tolist(),
    COLS_EVENTO[2]: exp["_cod"].tolist(),           # Código Sequencial = código da rubrica/item
    COLS_EVENTO[3]: exp["_tipo"].tolist(),
    COLS_EVENTO[4]: exp["_desc"].tolist(),          # "1 - Horas Normais" (máx. 40)
    COLS_EVENTO[5]: [para_int(v) for v in exp["_deb"]],
    COLS_EVENTO[6]: [para_int(v) for v in exp["_cred"]],
    COLS_EVENTO[7]: [para_int(v) for v in exp["_hist"]],
    COLS_EVENTO[8]: [complemento] * n,
}, columns=COLS_EVENTO)

integra = pd.DataFrame({
    COLS_INTEGRA[0]: [emp] * n,
    COLS_INTEGRA[1]: exp["_sep"].tolist(),
    COLS_INTEGRA[2]: exp["_cod"].tolist(),          # aponta para o lançamento da própria rubrica/item
    COLS_INTEGRA[3]: exp["_tipo"].tolist(),
    COLS_INTEGRA[4]: exp["_cod"].tolist(),
}, columns=COLS_INTEGRA).drop_duplicates().reset_index(drop=True)

sem_hist = sum(1 for h in exp["_hist"] if not h)
if sem_hist:
    st.warning(f"{sem_hist} lançamento(s) sem Código do Histórico — preencha na barra lateral.")
if not isinstance(emp, int):
    st.warning("Código da empresa não numérico — confira na barra lateral.")
cortadas = sum(1 for c, d in zip(exp["_cod"], exp["Descrição"])
               if len(f"{c} - {caixa_mista(d)}") > LIM_DESC_EVENTO)
if cortadas:
    st.info(f"{cortadas} descrição(ões) cortada(s) em {LIM_DESC_EVENTO} caracteres — "
            "ajuste na coluna 'Descrição Lançamento' se quiser abreviar.")

e1, e2, e3 = st.columns(3)
e1.metric("Lançamentos (aba evento)", len(evento))
e2.metric("Vínculos (aba integra)", len(integra))
e3.metric("Configurar manualmente", len(manuais))
t1, t2 = st.tabs([ABA_EVENTO, ABA_INTEGRA])
t1.dataframe(evento, hide_index=True, use_container_width=True)
t2.dataframe(integra, hide_index=True, use_container_width=True)


def excel_bytes(abas: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for nome, d in abas.items():
            d.to_excel(w, sheet_name=nome[:31], index=False)
    return buf.getvalue()


cfg_out = {
    "versao": 7, "empresa": cod_empresa, "nome_empresa": cab["nome"],
    "grupos": mapa_grupos, "grupo_socio": grupo_socio, "contas": dict(cfg_contas),
    "historico": historico, "hist_tipo": {str(t): v for t, v in hist_tipo.items() if v},
    "complemento": complemento, "baixa_prov": baixa_prov, "socio_adm": socio_adm,
    "bloqueio_extra": bloq_extra, "raizes": raizes,
    "mesma_rescisao": mesma_resc, "mesma_ferias": mesma_fer,
    "excecoes_ferias": exc_fer, "excecoes_rescisao": exc_resc,
    "colunas": {k: v for k, v in sel.items() if v and v != "(nenhuma)"},
}
if modo_cad:
    cfg_out.update({"tipos_cadastro": tipos_cad, "separadores_cadastro": seps_cad})
else:
    for chave in ("tipos_cadastro", "separadores_cadastro"):   # preserva o que veio do .json
        if chave in cfg:
            cfg_out[chave] = cfg[chave]

conf_abas = {"conferencia": ed}
if not manuais.empty:
    conf_abas["configurar_manual"] = manuais

sufixo = "_completo" if modo_cad else ""
c1, c2, c3 = st.columns(3)
c1.download_button(f"📥 Importação Domínio ({len(evento)} lançamentos)",
                   excel_bytes({ABA_INTEGRA: integra, ABA_EVENTO: evento}),
                   file_name=f"integracao_folha_emp{cod_empresa}{sufixo}.xlsx")
c2.download_button("📋 Planilha de conferência completa", excel_bytes(conf_abas),
                   file_name=f"conferencia_folha_emp{cod_empresa}{sufixo}.xlsx")
c3.download_button("💾 Salvar configuração da empresa (.json)",
                   json.dumps(cfg_out, ensure_ascii=False, indent=2).encode("utf-8"),
                   file_name=f"config_folha_emp{cod_empresa}.json", mime="application/json")
