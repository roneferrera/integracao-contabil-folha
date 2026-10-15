"""
Integrador Contábil da Folha — Domínio Sistemas
Motor 100% determinístico (regex + regras de substring). Sem IA, sem APIs, sem chaves.

- Nenhuma conta, código reduzido ou classificação de plano é fixada no código. As contas são achadas em
  qualquer plano: pelo nome (com as abreviações usuais expandidas), pelo nome dos grupos acima da conta
  ("Provisão de Férias > INSS", "Obrigações Trabalhistas > Salários"), pelo lado do balanço (no Passivo o
  "a pagar/a recolher" fica implícito) e pelo código do referencial da RFB (ECF), quando o plano tem.
  Conta achada só pelo grupo, pelo referencial ou empatada com outra vai para "Revisar".
- Exportação no layout de importação da Domínio (1 lançamento por rubrica/item):
    evento  -> lançamento (tabela FOINTEGCONT): Código Sequencial = código da rubrica/item,
               Descrição = "código - descrição" (caixa mista, máx. 40 caracteres)
    integra -> vínculo da rubrica/item ao seu lançamento (tabela FOINTEGCONTEVE)
  Arquivos TXT da importação via tabela (TAB, sem cabeçalho, Windows-1252). A importação grava 0 em campo
  numérico vazio lido do arquivo; o lado sem conta (partidas múltiplas) só fica em branco com o campo em
  "Valor fixo = NULO" no leiaute. Por isso o .zip separa os lançamentos em partes: duas contas, só débito
  (sem a coluna CONTA_CRED) e só crédito (sem a coluna CONTA_DEB), com o passo a passo e os scripts do leiaute.
- Tipos: 1 Folha mensal | 2 Empresa | 3 Férias | 4 Rescisão | 5 Prov. Férias | 6 Prov. 13 | 10 Outras Informações
- Dois modos:
    * Pendências: lê o relatório "Rubricas/Itens não configurados"
    * Todos os eventos cadastrados: usa só o Cadastro geral de rubricas + Plano de contas
- "Usar a mesma configuração da folha normal para Férias/Rescisão" (igual à Domínio):
    marcado    = o Tipo 3/4 não é gerado; os itens usam o lançamento da Folha (Tipo 1).
                 Rubricas exclusivas de Férias/Rescisão (que não existem na Folha) podem permanecer
                 no próprio Tipo 3/4 (opção "Manter exclusivas"), como a Domínio grava.
    desmarcado = Tipo 3/4 próprio com a mesma classificação por natureza da Folha. Como na integração real
                 da Domínio, o cálculo passa por Salários a Pagar: a 937 leva o líquido das férias para
                 Adiantamento de Férias e a 51 leva o líquido da rescisão para Rescisões a Pagar.
- Provisão (Tipos 5/6): férias e 13º pagos (e o FGTS/INSS sobre eles) baixam a provisão; o "Valor Estorno
  Provisão" da Domínio é ajuste do saldo da provisão, não a baixa das férias pagas.
- Entidade filantrópica: gera também os itens de isenção (25 a 36) como estorno do INSS patronal.
- Conferências: chave única do lançamento (empresa+separador+código+tipo), uma rubrica por lançamento
  em cada empresa+separador+tipo, verificação do líquido e conformidade com FOINTEGCONT/FOINTEGCONTEVE.

Executar:
    pip install -r requirements.txt
    streamlit run app_v9.py
"""
import csv
import datetime as dt
import hashlib
import io
import json
import numbers
import re
import struct
import unicodedata
import zipfile
from difflib import SequenceMatcher
from functools import lru_cache

import pandas as pd
import streamlit as st

try:
    import pypdfium2 as pdfium
except ImportError:
    pdfium = None
try:
    import pdfplumber
except ImportError:
    pdfplumber = None

# =====================================================================
# 0. LAYOUT DE IMPORTAÇÃO DA DOMÍNIO (não depende do plano de contas)
# =====================================================================
NOME_PADRAO_SEM_SEPARADOR = "Geral"      # só para exibição na tela
SEP_SEM = "0"                             # valor gravado no arquivo quando não há separador
LIMIAR_SIMILARIDADE = 0.75
LIM_DESC_EVENTO = 40                      # FOINTEGCONT.descricao = char(40)
LIM_COMP_HIST = 512                       # FOINTEGCONT.comp_hist = varchar(512)
HISTORICO_FIXO = "186"                   # Código do Histórico fixo em todos os lançamentos
COMPLEMENTO_PADRAO = "<<Competencia>> - <<Descrição do Lançamento>>"
NULO = "NULO"                            # lado sem conta (partidas múltiplas). Na importação via tabela o NULO
                                         # vale só como "Valor fixo" do campo no leiaute: no TXT, vazio ou texto
                                         # em campo numérico viram 0. O lado NULO sai do TXT (ver PARTES)
SECOES_TIPO1 = ("Folha Normal", "Adiantamento", "13º Salário")
# Itens de Líquido do Tipo 1 (FOINTEGCONTITENS): código -> (descrição, contas candidatas em ordem)
LIQUIDOS = {
    -1: ("Líquido Folha", ["SAL_PAGAR"]), -2: ("Líquido Adiantamento", ["SAL_PAGAR"]),
    -3: ("Líquido Adiantamento 13º", ["SAL_PAGAR"]), -4: ("Líquido 13º", ["SAL_PAGAR"]),
    -5: ("Líquido Empregador", ["PROLAB_PAGAR", "SAL_PAGAR"]), -6: ("Líquido Autônomo", ["SAL_PAGAR"]),
    -7: ("Líquido Férias", ["FER_PAGAR", "SAL_PAGAR"]),
    -15: ("Líquido Participação Lucro", ["SAL_PAGAR"]), -16: ("Líquido RPA", ["SAL_PAGAR"]),
    -17: ("Líquido Resilição Professor", ["SAL_PAGAR"]), -18: ("Líquido Convocação Intermitente", ["SAL_PAGAR"]),
}
LIQUIDOS_PADRAO = [-1, -2, -3, -4, -5, -7]

ABA_INTEGRA, ABA_EVENTO = "integra", "evento"
COL_TIPO = ("Tipo da Integração (1 - Folha mensal; 2 - Empresa; 3 - Férias; 4 - Rescisao; "
            "5 - Prov. Férias; 6 - Prov. 13)")
COLS_INTEGRA = ["Código da Empresa", "Separador", "Código Sequencial da Integração", COL_TIPO,
                "Código da Rúbrica Selecionada"]
COLS_EVENTO = ["Código da Empresa", "Separador (0 quando é sem separador)",
               "Código Sequencial da Integração", COL_TIPO, "Descrição",
               "Código da Conta Débito", "Código da Conta Crédito", "Código do Histórico", "Complemento"]

# Importação via tabela (importação por leiaute): um TXT por tabela, campos na ordem do leiaute
TAB_EVENTO, TAB_INTEGRA = "FOINTEGCONT", "FOINTEGCONTEVE"
CAMPOS_EVENTO = ["CODI_EMP", "I_SEPARADOR", "I_CODIGO", "TIPO", "DESCRICAO", "CONTA_DEB", "CONTA_CRED",
                 "I_HISTORICO", "COMP_HIST"]
CAMPOS_INTEGRA = ["CODI_EMP", "I_SEPARADOR", "I_CODIGO", "TIPO", "I_EVENTOS"]
# parte -> (pasta, título, campo de FOINTEGCONT com Valor fixo = NULO no leiaute)
PARTES = {1: ("parte_1_duas_contas", "lançamentos com débito e crédito", None),
          2: ("parte_2_so_debito", "lançamentos só a débito (proventos)", "CONTA_CRED"),
          3: ("parte_3_so_credito", "lançamentos só a crédito (descontos e líquidos)", "CONTA_DEB")}

TIPOS_LAYOUT = {1: "Folha mensal", 2: "Empresa", 3: "Férias", 4: "Rescisão",
                5: "Prov. Férias", 6: "Prov. 13", 10: "Outras Informações"}
TIPO_INTEGRACAO = {  # seção do relatório -> código do layout (None = não importável)
    "Folha Normal": 1, "Adiantamento": 1, "13º Salário": 1,
    "Empresa": 2, "Férias": 3, "Rescisão": 4,
    "Provisão de Férias": 5, "Provisão de 13º": 6,
    "Outras Informações": 10,
}

# Modo "todos os eventos do cadastro": Tipo do layout -> seção interna usada pelas regras
SECAO_POR_TIPO = {1: "Folha Normal", 2: "Empresa", 3: "Férias", 4: "Rescisão", 5: "Provisão de Férias",
                  6: "Provisão de 13º", 10: "Outras Informações"}
TIPOS_CADASTRO = list(SECAO_POR_TIPO)

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

# Catálogo do sistema Domínio (FOINTEGCONTITENS): abas Provisão de Férias e Provisão de 13º (mesmos itens
# nos Tipos 5 e 6). Os itens 14 a 25 (transferência de provisão entre empresas) não são gerados.
ITENS_PROVISAO = {
    1: "Valor da Provisão", 3: "Encargos FGTS", 4: "Encargos PIS", 5: "Valor Estorno Provisão",
    6: "Valor Estorno INSS Empresa", 7: "Valor Estorno INSS RAT", 8: "Valor Estorno INSS Terceiros",
    9: "Valor Estorno FGTS", 10: "Valor Estorno PIS", 11: "Encargos INSS Empresa", 12: "Encargos INSS RAT",
    13: "Encargos INSS Terceiros",
}
ITENS_PROVISAO_PIS = {4, 10}   # só para entidade com PIS sobre a folha (filantrópica)

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


def e_ferias(d):
    """Férias de fato: 'FÉRIAS', 'FER' abreviado, 'ABONO PECUNIÁRIO' (também 'ABONO/ABON PECUN') e o abono com
    médias/vantagens/1/3 ou do intermitente (não 'Abono Salarial'/'Abono Natalino')."""
    if tem(d, "FERIAS", "FER", "ABONO PECUN*", "ABON PECUN*"):
        return True
    return (tem(d, "ABONO", "ABON")
            and tem(d, "MEDIA*", "VANTAGE*", "DIFERENC*", "DIF", "PARCELA*", "1 3", "INTERMITENT*")
            and not tem(d, "SALARIAL", "NATALINO", "PIS"))


def txt_cel(v) -> str:
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()


def para_int(v):
    s = txt_cel(v)
    return int(s) if s.isdigit() else s


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
                "descricao": "Descrição", "tipo": "Tipo A/S (opcional)", "ecf": "Classificação ECF (opcional)"}
CAMPOS_OPCIONAIS = ("tipo", "ecf")
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


# Abreviações usuais em nomes de contas (só no plano de contas; as rubricas continuam com norm()).
# Dependentes do contexto (a palavra antes/depois) ficam em _expande_token.
ABREV_CONTA = {
    "SALS": "SALARIOS", "SALAR": "SALARIOS", "ORD": "ORDENADOS", "ORDEN": "ORDENADOS",
    "ADTO": "ADIANTAMENTO", "ADTOS": "ADIANTAMENTOS", "ADT": "ADIANTAMENTO", "ADTS": "ADIANTAMENTOS",
    "ADIANT": "ADIANTAMENTO", "ADIANTS": "ADIANTAMENTOS", "ANTEC": "ANTECIPACAO",
    "PROV": "PROVISAO", "PROVS": "PROVISOES", "PROVIS": "PROVISAO", "FER": "FERIAS",
    "RECOL": "RECOLHER", "RECOLH": "RECOLHER", "RECUP": "RECUPERAR", "PG": "PAGAR", "PAG": "PAGAR",
    "CONTRIB": "CONTRIBUICAO", "CONTR": "CONTRIBUICAO", "CONTRIBS": "CONTRIBUICOES",
    "PREVID": "PREVIDENCIA", "PREVIDENC": "PREVIDENCIA", "ENC": "ENCARGOS", "ENCS": "ENCARGOS",
    "OBRIG": "OBRIGACOES", "OBRIGS": "OBRIGACOES", "FUNC": "FUNCIONARIOS", "FUNCS": "FUNCIONARIOS",
    "FUNCION": "FUNCIONARIOS", "EMPREG": "EMPREGADOS", "COLAB": "COLABORADORES",
    "GRAT": "GRATIFICACOES", "GRATIF": "GRATIFICACOES", "COMIS": "COMISSOES", "COMISS": "COMISSOES",
    "INDENIZ": "INDENIZACOES", "INDEN": "INDENIZACOES", "RESC": "RESCISOES", "RESCIS": "RESCISOES",
    "ASSIST": "ASSISTENCIA", "ODONT": "ODONTOLOGICA", "TRANSP": "TRANSPORTE", "REFEIC": "REFEICAO",
    "ALIM": "ALIMENTACAO", "ALIMENT": "ALIMENTACAO", "BENEF": "BENEFICIOS", "PROLABORE": "PRO LABORE",
    "PROLAB": "PRO LABORE", "IRF": "IRRF", "SOC": "SOCIAIS", "PESS": "PESSOAL", "DESP": "DESPESAS",
    "CUST": "CUSTOS", "ADM": "ADMINISTRATIVAS", "ADMIN": "ADMINISTRATIVAS", "ADMINIST": "ADMINISTRATIVAS",
    "VEND": "VENDAS", "OPER": "OPERACIONAIS", "OPERAC": "OPERACIONAIS", "HS": "HORAS",
    "SERV": "SERVICOS", "SERVS": "SERVICOS", "CONSIG": "CONSIGNADOS", "CONSIGN": "CONSIGNADOS",
    "EMPREST": "EMPRESTIMOS", "SIND": "SINDICAL", "MATERN": "MATERNIDADE", "PENS": "PENSAO",
    "VT": "VALE TRANSPORTE", "VR": "VALE REFEICAO", "VA": "VALE ALIMENTACAO", "TERC": "TERCEIROS",
    "CIRC": "CIRCULANTE", "CRED": "CREDITOS", "RECEB": "RECEBER",
}
_REC_RECEITA = ("BRUTA", "BRUTAS", "OPERACIONAL", "OPERACIONAIS", "FINANCEIRA", "FINANCEIRAS", "SERVICOS",
                "VENDAS", "LIQUIDA", "SERV")


def _expande_token(ant, w, prox):
    if w == "SAL":
        return "SALARIO" if prox in ("FAMILIA", "FAM", "MAT", "MATERN", "MATERNIDADE", "EDUCACAO") else "SALARIOS"
    if w == "REC":
        return "RECEITA" if prox in _REC_RECEITA else "RECOLHER" if ant == "A" else w
    if w == "COMP":
        return "COMPENSAR" if ant == "A" else w
    if w == "PREV":
        return "PREVIO" if ant in ("AVISO", "AV") else "PREVIDENCIA"
    if w == "AV" and prox in ("PREV", "PREVIO"):
        return "AVISO"
    if w == "TRAB":
        return ("TRABALHADOR" if ant.startswith("CRED") else "TRABALHO" if ant in ("ACID", "ACIDENTE")
                else "TRABALHISTAS")
    if w == "LAB" and ant == "PRO":
        return "LABORE"
    if w == "MAT" and ant in ("SAL", "SALARIO", "LIC", "LICENCA"):
        return "MATERNIDADE"
    if w == "FAM" and ant in ("SAL", "SALARIO"):
        return "FAMILIA"
    if w == "MED" and ant.startswith("ASSIST"):
        return "MEDICA"
    if w == "H" and prox.startswith("EXTRA"):
        return "HORAS"
    if w == "V" and prox.startswith(("TRANSP", "REFEIC", "ALIM")):
        return "VALE"
    if w == "DEC" and prox.startswith("TERC"):
        return "DECIMO"
    if w == "TERC" and ant in ("DEC", "DECIMO"):
        return "TERCEIRO"
    if w == "EMP" and prox.startswith("CONSIG"):
        return "EMPRESTIMOS"
    return ABREV_CONTA.get(w, w)


def norm_conta(txt) -> str:
    """norm() do nome da conta com as abreviações expandidas: 'ADTO 13 SAL' -> 'ADIANTAMENTO 13 SALARIOS',
    'INSS A REC' -> 'INSS A RECOLHER', 'FUNDO DE GARANTIA' -> 'FGTS'."""
    p = norm(txt).split()
    out = [_expande_token(p[i - 1] if i else "", w, p[i + 1] if i + 1 < len(p) else "") for i, w in enumerate(p)]
    s = " ".join(out)
    s = re.sub(r"\bIMP(?:OSTO)? (?:DE )?RENDA\b", "IMPOSTO DE RENDA", s)
    s = re.sub(r"\bFUNDO (?:DE )?GARANTIA(?: (?:DO|POR) TEMPO DE SERVICO)?\b", "FGTS", s)
    return s


def _ler_xls_tolerante(dados: bytes) -> pd.DataFrame:
    """.xls exportado pela Domínio que o xlrd recusa ('Expected BOF record'): o BOUNDSHEET aponta para registros
    soltos que ficam antes do BOF da planilha. Corrige o ponteiro no fluxo Workbook e lê de novo."""
    from xlrd import compdoc
    log = io.StringIO()
    cd = compdoc.CompDoc(dados, logfile=log)
    for nome_fluxo in ("Workbook", "Book"):
        mem, base, tam = cd.locate_named_stream(nome_fluxo)
        if mem:
            break
    else:
        raise ValueError("arquivo .xls sem o fluxo da planilha (Workbook)")
    fluxo = bytearray(mem[base:base + tam])
    pos, nivel, bofs, bounds = 0, 0, [], []
    while pos + 4 <= len(fluxo):
        tipo, tam_reg = struct.unpack_from("<HH", fluxo, pos)
        if tipo == 0x0809:                     # BOF: início dos globais e de cada planilha
            if nivel == 0:
                bofs.append(pos)
            nivel += 1
        elif tipo == 0x000A:                   # EOF
            nivel = max(0, nivel - 1)
        elif tipo == 0x0085:                   # BOUNDSHEET: posição da planilha no fluxo
            bounds.append(pos)
        pos += 4 + tam_reg
    for k, b in enumerate(bounds):
        if k + 1 < len(bofs):
            struct.pack_into("<I", fluxo, b + 4, bofs[k + 1])
    return pd.read_excel(io.BytesIO(bytes(fluxo)), engine="xlrd", dtype=str, header=None,
                         engine_kwargs={"logfile": log})


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
        try:
            df = pd.read_excel(io.BytesIO(dados), dtype=str, header=None)
        except Exception as erro:
            if not nome.lower().endswith(".xls"):
                raise
            try:
                df = _ler_xls_tolerante(dados)
            except Exception:
                raise ValueError(f"{erro} — abra o arquivo no Excel e salve como .xlsx") from erro
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
        if "ecf" not in m and ("ECF" in n.split() or "REFERENCIAL" in n):
            m["ecf"] = c
        elif "reduzido" not in m and "REDUZ" in n:
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
def preparar_plano(raw: pd.DataFrame, c_red, c_cla, c_des, c_tip, c_ecf=None):
    df = pd.DataFrame({
        "reduzido": raw[c_red], "classificacao": raw[c_cla], "descricao": raw[c_des],
        "tipo": raw[c_tip] if c_tip else "",
        "ecf": raw[c_ecf].fillna("").astype(str).str.replace(r"\.0$", "", regex=True).str.replace(r"\D", "", regex=True)
        if c_ecf else "",
    }).dropna(subset=["reduzido", "classificacao"])
    df["reduzido"] = df["reduzido"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    df["classificacao"] = df["classificacao"].map(norm_classif)
    df["descricao"] = df["descricao"].fillna("").astype(str).str.strip()
    df = df[df["reduzido"].str.fullmatch(r"\d+") & (df["classificacao"] != "")]
    # a exportação repete a conta quando ela tem mais de um código referencial: guarda todos
    ecf = df.groupby(["reduzido", "classificacao"])["ecf"].agg(lambda s: "|".join(sorted({x for x in s if x})))
    df = df.drop_duplicates(subset=["reduzido", "classificacao"]).reset_index(drop=True)
    df["ecf"] = [ecf.get((r, c), "") for r, c in zip(df["reduzido"], df["classificacao"])]
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
    df["desc_norm"] = df["descricao"].map(norm_conta)
    return df, pontuado


def _raiz_pelo_nome(d):
    if (tem(d, "ATIVO") and not tem(d, "PASSIVO")) or (tem(d, "BENS", "DIREITOS") and not tem(d, "OBRIGACOES")):
        return "ATIVO"
    if tem(d, "PASSIVO", "OBRIGACOES", "EXIGIVEL*"):
        return "PASSIVO"
    if tem(d, "RECEITA*", "RENDAS", "INGRESSOS") and not tem(d, "CUSTO*", "DESPESA*"):
        return "RECEITA"
    if tem(d, "CUSTO*", "DESPESA*", "RESULTADO*", "GASTOS"):
        return "RESULTADO"
    return None


def detectar_raizes(plano):
    """Raízes do plano pelo nome; a raiz sem nome reconhecível (ex.: '1 - Contas Patrimoniais') é classificada
    pelo referencial ECF das contas abaixo dela (1 = Ativo, 2 = Passivo) e, sem referencial, pelo 1º dígito
    (1 = Ativo, 2 = Passivo) quando o lado ainda não foi achado."""
    raizes = plano[plano["pai"] == ""]
    out = {k: [] for k in ROTULOS_RAIZ}
    sem_nome = []
    for _, r in raizes.iterrows():
        lado = _raiz_pelo_nome(r["desc_norm"])
        if lado:
            out[lado].append(r["classificacao"])
        elif not tem(r["desc_norm"], "PATRIMONIO", "COMPENSACAO", "COMPENSACOES"):
            sem_nome.append(r["classificacao"])
    tem_ecf = "ecf" in plano.columns
    pont = bool(plano["classificacao"].str.contains(".", regex=False).any())
    for c in sem_nome:
        filhos = plano[plano["classificacao"].map(lambda x: sob(x, c, pont)) & (plano["tipo"] == "A")]
        codigos = [x[:1] for e in (filhos["ecf"] if tem_ecf else []) for x in str(e).split("|") if x]
        lado = None
        if codigos:
            dig = max(set(codigos), key=codigos.count)
            lado = {"1": "ATIVO", "2": "PASSIVO"}.get(dig)
        elif c[:1] == "1" and not out["ATIVO"]:
            lado = "ATIVO"
        elif c[:1] == "2" and not out["PASSIVO"]:
            lado = "PASSIVO"
        if lado:
            out[lado].append(c)
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
    sal = an[an["desc_norm"].map(lambda d: tem(d, "SALARIO*", "ORDENADO*", "REMUNERAC*")
                                 and not tem(d, "PAGAR", "FAMILIA", "MATERN*", "PROVIS*", "EDUCACAO", "DIRETOR*")
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
        if tem(norm_conta(f"{g['caminho']} {g['descricao']}"), chave):
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


NAO_TERC = ["SERVICO*", "TERCEIROS", "ALUGUE*", "APLICAC*", "PJ", "PF", "AUTONOM*", "JUROS", "PROVIS*"]
# INSS de terceiros/PJ/PF/retido em nota fiscal não é a conta da folha
NAO_INSS_FOLHA = ["RECEITA BRUTA", "PROVIS*", "RETID*", "PJ", "PF", "SERVIC*", "AUTONOM*", "CPRB", "DESONERAC*"]
NAO_INSS_NF = ["PJ", "PF", "RETID*", "SERVIC*", "NF", "NOTA*", "CLIENTE*"]
NAO_SAL_PAGAR = ["FERIAS", "13O", "13A", "13", "DECIMO*", "PRO LABORE", "RESCIS*", "FAMILIA", "MATERN*",
                 "EDUCACAO", "PROVIS*", "INSS", "FGTS", "IRRF", "PIS", "ADIANT*"]

ALVOS = {
    # ---- DRE (restrito ao grupo do separador)
    "SALARIOS": (G, "Salários e Ordenados", [
        (["SALARIO*"], ["PAGAR", "FAMILIA", "MATERN*", "PROVIS*", "13O", "13A", "13", "DECIMO*",
                        "ADIANT*", "INSS", "FGTS", "EDUCACAO"]),
        (["ORDENADO*"], ["PAGAR"]),
        (["REMUNERAC*"], ["PAGAR", "DIRETOR*", "ADMINISTRADOR*", "SOCIO*", "PRO LABORE", "VARIAVE*"])]),
    "PRO_LABORE": (G, "Pró-labore", [(["PRO LABORE"], ["PAGAR", "INSS"]),
                                     (["HONORARIO*", "DIRETOR*"], []), (["RETIRADA*", "SOCIO*"], []),
                                     (["REMUNERAC*", "DIRETOR*"], ["PAGAR"]),
                                     (["REMUNERAC*", "ADMINISTRADOR*"], ["PAGAR"])]),
    "HE": (G, "Horas Extras", [(["HORAS EXTRA*"], []), (["HORA EXTRA*"], []), (["EXTRAORDINAR*"], [])]),
    "PREMIOS": (G, "Prêmios e Gratificações", [(["PREMIO*"], []), (["GRATIFICAC*"], []),
                                               (["BONIFICAC*"], [])]),
    "COMISSOES": (G, "Comissões", [(["COMISS*"], ["BANCARI*"])]),
    "DECIMO": (G, "13º Salário", _13((), ["PROVIS*", "INSS", "FGTS", "PIS", "PAGAR", "ADIANT*"])),
    "FERIAS": (G, "Férias", [(["FERIAS"], ["PROVIS*", "PAGAR", "INSS", "FGTS", "PIS", "ADIANT*"])]),
    "INSS": (G, "INSS (encargo)", [
        (["INSS"], ["PROVIS*", "RECOLHER", "PAGAR", "RECEITA BRUTA", "RETIDO", "COMPENSAR"]),
        (["PREVIDENCIA*"], ["RECOLHER", "PAGAR", "PRIVADA", "COMPLEMENTAR"]),
        (["CONTRIBUIC*", "PREVIDENCIARIA*"], ["RECOLHER", "PAGAR", "RECEITA BRUTA"]),
        (["ENCARGO*", "PREVIDENCIARI*"], ["RECOLHER", "PAGAR", "PROVIS*"])]),
    "FGTS": (G, "FGTS (encargo)", [(["FGTS"], ["PROVIS*", "RECOLHER", "PAGAR"])]),
    "PIS": (G, "PIS s/ Folha", [(["PIS"], ["PROVIS*", "RECOLHER", "PAGAR", "RETIDO", "FATURAMENTO",
                                           "RECEITA*", "COFINS"])]),
    "INDENIZ": (G, "Indenizações e Aviso Prévio", [(["INDENIZAC*"], []), (["AVISO PREVIO"], []),
                                                   (["RESCIS*"], ["PAGAR"])]),
    "ASSIST": (G, "Assistência Médica", [(["ASSISTENCIA MEDICA*"], []), (["PLANO*", "SAUDE"], []),
                                         (["CONVENIO*", "MEDIC*"], []), (["ODONTO*"], [])]),
    "VT": (G, "Vale-Transporte", [(["VALE TRANSP*"], []), (["TRANSPORTE*", "EMPREGADO*"], []),
                                  (["TRANSPORTE*", "FUNCIONARIO*"], []), (["BENEFICIO*", "TRANSPORTE*"], [])]),
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
        (["SALARIO*", "PAGAR"], NAO_SAL_PAGAR),
        (["ORDENADO*", "PAGAR"], []), (["FOLHA*", "PAGAR"], ["INSS", "FGTS", "IRRF", "PIS", "PROVIS*"]),
        (["REMUNERAC*", "PAGAR"], ["DIRETOR*", "PRO LABORE", "FERIAS", "13O", "13A", "13", "DECIMO*", "RESCIS*"]),
        (["OBRIGAC*", "PESSOAL"], ["INSS", "FGTS", "IRRF", "PROVIS*"])]),
    "FER_PAGAR": (PA, "Férias a Pagar", [(["FERIAS", "PAGAR"], ["PROVIS*"])]),
    "RESC_PAGAR": (PA, "Rescisões a Pagar", [(["RESCIS*", "PAGAR"], [])]),
    "PROLAB_PAGAR": (PA, "Pró-labore a Pagar", [(["PRO LABORE", "PAGAR"], []), (["PRO LABORE"], ["INSS"]),
                                                (["HONORARIO*", "DIRETOR*"], ["INSS"]), (["RETIRADA*", "SOCIO*"], []),
                                                (["REMUNERAC*", "DIRETOR*"], ["INSS"]),
                                                (["REMUNERAC*", "ADMINISTRADOR*"], ["INSS"])]),
    "INSS_REC": (PA, "INSS a Recolher", [
        (["INSS", "RECOLHER"], NAO_INSS_FOLHA), (["INSS", "PAGAR"], NAO_INSS_FOLHA),
        (["PREVIDENCIA SOCIAL"], NAO_INSS_FOLHA),
        (["CONTRIBUIC*", "PREVIDENCIARI*"], NAO_INSS_FOLHA + ["COMPENSAR", "RECUPERAR"]),
        (["OBRIGAC*", "PREVIDENCIARI*"], NAO_INSS_FOLHA), (["GPS"], NAO_INSS_FOLHA),
        (["INSS"], NAO_INSS_FOLHA + ["COMPENSAR", "RECUPERAR"]),
        (["ENCARGO*", "SOCIAIS"], NAO_INSS_FOLHA + ["FGTS", "IRRF", "COMPENSAR", "RECUPERAR"]),
        (["TRIBUTOS FEDERAIS"], []), (["IMPOSTOS FEDERAIS"], [])]),
    "FGTS_REC": (PA, "FGTS a Recolher", [(["FGTS", "RECOLHER"], ["PROVIS*"]),
                                         (["FGTS", "PAGAR"], ["PROVIS*"]), (["FGTS"], ["PROVIS*"])]),
    "IRRF_REC": (PA, "IRRF s/ Folha a Recolher", [
        (["IRRF", "FOLHA*"], ["PROVIS*"]), (["IRRF", "SALARIO*"], []), (["IRRF", "TRABALH*"], []),
        (["IRRF", "RECOLHER"], NAO_TERC), (["IRRF", "PAGAR"], NAO_TERC),
        (["IMPOSTO DE RENDA", "FONTE"], NAO_TERC), (["IR", "FONTE"], NAO_TERC), (["IRRF"], NAO_TERC),
        (["IMPOSTO*", "RENDA", "RETID*"], NAO_TERC),
        (["TRIBUTOS FEDERAIS"], NAO_TERC), (["IMPOSTOS FEDERAIS"], NAO_TERC)]),
    "PIS_REC": (PA, "PIS s/ Folha a Recolher", [(["PIS", "FOLHA*"], ["PROVIS*"]),
                                                (["PIS", "RECOLHER"], ["PROVIS*", "RETIDO", "FATURAMENTO"])]),
    "SIND_REC": (PA, "Contribuição Sindical a Recolher", [
        (["CONTRIBUIC*", "SINDICA*"], ["PATRONAL"]), (["SINDICA*"], ["PATRONAL"]),
        (["CONTRIBUIC*", "ASSISTENCIA*"], ["PATRONAL", "MEDICA*", "SAUDE", "SOCIAL", "ODONTO*"]),
        (["CONTRIBUIC*", "CONFEDERATIV*"], ["PATRONAL"])]),
    "SIND_PAT_REC": (PA, "Contribuição Sindical Patronal a Recolher", [(["SINDICA*", "PATRONAL"], [])]),
    "EMPREST": (PA, "Empréstimos Consignados a Repassar", [
        (["CONSIGNA*"], []), (["EMPRESTIMO*", "EMPREGADO*"], []),
        (["EMPRESTIMO*", "FUNCIONARIO*"], []), (["EMPRESTIMO*", "FOLHA"], []),
        (["CREDITO*", "TRABALHADOR*"], []), (["ECONSIGNADO*"], [])]),
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
        (["ADIANTAMENTO*", "SALARI*"], ["13O", "13A", "13", "DECIMO*", "FERIAS"]),
        (["ADIANTAMENTO*", "EMPREGADO*"], []), (["ADIANTAMENTO*", "FUNCIONARIO*"], []),
        (["ADIANTAMENTO*", "PESSOAL"], []), (["ADIANTAMENTO*", "COLABORADOR*"], []),
        (["ANTECIPAC*", "SALARI*"], ["13O", "13A", "13", "DECIMO*", "FERIAS"])]),
    "ADIANT_13": (AT, "Adiantamento de 13º", _13(["ADIANTAMENTO*"]) + _13(["ANTECIPAC*"])),
    "ADIANT_FER": (AT, "Adiantamento de Férias", [(["ADIANTAMENTO*", "FERIAS"], []), (["ANTECIPAC*", "FERIAS"], [])]),
    "INSS_COMP": (AT, "INSS a Compensar (retenções)", [(["INSS", "COMPENSAR"], ["MATERN*", "FAMILIA"]),
                                                       (["INSS", "RECUPERAR"], [])]),
    "BENEF_INSS": (AT, "Sal.-Família/Maternidade a Compensar", [
        (["MATERN*", "COMPENSAR"], []), (["FAMILIA", "COMPENSAR"], []),
        (["MATERN*", "RECUPERAR"], []), (["FAMILIA", "RECUPERAR"], []),
        (["INSS", "COMPENSAR"], NAO_INSS_NF), (["INSS", "RECUPERAR"], NAO_INSS_NF)]),
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

# Grupo sintético que indica obrigação/direito da folha: entre contas candidatas do Ativo/Passivo, as que estão
# dentro dele têm preferência ("INSS a pagar" em Salários e Encargos ganha de "INSS PJ a recolher" em Impostos).
TERMOS_GRUPO_FOLHA = ("SALARIO*", "ORDENADO*", "TRABALHIST*", "PREVIDENCIARI*", "SOCIAIS", "ENCARGO*", "PESSOAL",
                      "FOLHA*", "EMPREGADO*", "FUNCIONARIO*", "COLABORADOR*")
# Benefícios e comissões costumam ficar em outro grupo de despesa: sem conta no grupo do separador,
# procura em todo o resultado antes de cair em Salários.
FORA_DO_GRUPO = {"BOLSA", "ASSIST", "VT", "ALIM", "SEGURO", "TREIN", "COMISSOES"}

# Plano referencial da RFB (ECF) da coluna "Classificação ECF" do plano exportado da Domínio: é nacional e vale
# para qualquer plano de contas mapeado. Os códigos mudam conforme o referencial (Geral × Imunes/Isentas), por
# isso o referencial em uso é detectado pelas próprias contas. No Ativo/Passivo o código confirma ou descarta a
# conta; nas despesas só confirma (na prática as empresas agrupam encargos em "Outros gastos com pessoal").
ECF_REFERENCIAL = {
    "Geral (L100/L300)": {
        "SAL_PAGAR": ("201010101",), "FER_PAGAR": ("201010101",), "RESC_PAGAR": ("201010101", "201010160"),
        "PROLAB_PAGAR": ("201010101",), "INSS_REC": ("201010103",), "FGTS_REC": ("201010104",),
        "IRRF_REC": ("201010901", "201010109"), "PIS_REC": ("201010904", "201010109"), "SIND_REC": ("201010109",),
        "CPRB_REC": ("201010911",), "ISS_REC": ("201010908", "201010912"),
        "ADIANT_SAL": ("101020103",), "ADIANT_13": ("101020103",), "ADIANT_FER": ("101020103",),
        "BENEF_INSS": ("101020418",), "INSS_COMP": ("101020417", "101020418"), "IRRF_COMP": ("101020401",),
        "PIS_COMP": ("101020408",),
        "PRO_LABORE": ("30101070101",), "SALARIOS": ("30101070102",), "INSS": ("30101070105",),
        "FGTS": ("30101070106",), "ALIM": ("30101070113",), "ASSIST": ("30101070133",),
        "D_PROV_FER": ("30101070126",), "D_PROV_13": ("30101070127",), "SEGURO": ("30101070140",),
    },
    "Imunes e isentas (U100/U150)": {
        "SAL_PAGAR": ("201010101",), "FER_PAGAR": ("201010101",), "RESC_PAGAR": ("201010101",),
        "PROLAB_PAGAR": ("201010101",), "INSS_REC": ("201010102",), "FGTS_REC": ("201010103",),
        "ADIANT_SAL": ("101020102",), "ADIANT_13": ("101020102",), "ADIANT_FER": ("101020102",),
        "SALARIOS": ("402010201",), "INSS": ("402010203",), "FGTS": ("402010203",),
        "D_PROV_FER": ("402010205",), "D_PROV_13": ("402010205",), "PRO_LABORE": ("4020101",),
    },
}
# Código referencial específico: a conta mapeada nele vale mesmo com nome fora do padrão (desde que o nome não
# tenha termo que o alvo exclui: "INSS PJ a recolher" mapeada no INSS a recolher continua fora)
ECF_SUFICIENTE = {"INSS_REC", "FGTS_REC", "BENEF_INSS", "INSS_COMP", "IRRF_COMP", "PIS_COMP", "CPRB_REC"}

# Termos que, num grupo acima da conta (dentro do escopo da busca), indicam outro subrazão: em "Provisão de
# Férias > INSS" a conta não é o INSS a recolher da folha. Só vetam o alvo que já exclui o termo no nome da conta.
VETO_CONTEXTO = {"PROVIS*", "RECEITA BRUTA", "CPRB", "DESONERAC*", "RETID*", "RETIDO", "PJ", "PF", "TERCEIROS",
                 "SERVICO*", "SERVIC*", "AUTONOM*", "NF", "NOTA*", "CLIENTE*", "ALUGUE*", "APLICAC*", "JUROS",
                 "FATURAMENTO", "COFINS", "RECEITA*", "COMPENSAR", "RECUPERAR", "PRIVADA", "COMPLEMENTAR"}
# ...exceto num grupo que também é da folha ("Retenções da folha > INSS"); provisão e CPRB vetam sempre
VETO_SEMPRE = {"PROVIS*", "RECEITA BRUTA", "CPRB", "DESONERAC*", "PRIVADA", "COMPLEMENTAR"}
TERMOS_FOLHA_FORTE = ("FOLHA*", "EMPREGADO*", "SALARIO*", "FUNCIONARIO*", "COLABORADOR*", "TRABALHIST*", "PESSOAL")
# No Passivo o "a pagar"/"a recolher" fica implícito ("Obrigações Trabalhistas > Salários"); termo genérico
# demais não vira padrão sozinho
IMPLICITOS_PASSIVO = {"PAGAR", "RECOLHER"}
GENERICOS = {"FOLHA*", "OBRIGAC*", "PESSOAL", "OUTRAS"}
# termos genéricos que só valem no nome da própria conta, nunca vindos do grupo ("Outras Obrigações > Telefone a
# pagar" não é "Outras contas a pagar" da folha)
SO_NO_NOME = GENERICOS | {"ENCARGO*", "SOCIAIS", "TAXAS", "CONTRIBUIC*", "TRIBUTOS FEDERAIS", "IMPOSTOS FEDERAIS",
                          "BENEFICIO*"}
LONGO_PRAZO = ("NAO CIRCULANTE", "LONGO PRAZO", "EXIGIVEL A LONGO*", "REALIZAVEL A LONGO*")
# como a conta foi achada (0 a 4); as fracas e os empates vão para "Revisar"
COMO = {0: "pelo nome", 1: "pelo nome e grupo", 2: "só pelo grupo", 3: "pelo nome, no Passivo",
        4: "pelo referencial ECF"}
COMO_FRACO = {COMO[2], COMO[4]}


def referencial_do_plano(plano):
    """(nome, mapa alvo -> códigos) do referencial ECF que o plano segue; ('', {}) sem mapeamento confiável.
    Confere os códigos das contas de Salários, INSS e FGTS a pagar/recolher com os nomes delas."""
    if "ecf" not in plano.columns:
        return "", {}
    an = plano[(plano["tipo"] == "A") & (plano["ecf"] != "")]
    if an.empty:
        return "", {}
    pontos = {}
    for nome, mapa in ECF_REFERENCIAL.items():
        p = 0
        for alvo in ("SAL_PAGAR", "INSS_REC", "FGTS_REC"):
            must, nots = ALVOS[alvo][2][0] if alvo == "SAL_PAGAR" else ALVOS[alvo][2][1]
            for d, codigos in zip(an["desc_norm"], an["ecf"]):
                if all(tem(d, x) for x in must) and not any(tem(d, x) for x in nots):
                    p += 1 if any(c.startswith(mapa[alvo]) for c in codigos.split("|")) else -1
        pontos[nome] = p
    nome = max(pontos, key=pontos.get)
    return (nome, ECF_REFERENCIAL[nome]) if pontos[nome] > 0 else ("", {})


class Resolvedor:
    def __init__(self, plano, pontuado, raizes, overrides):
        self.an = plano[plano["tipo"] == "A"]
        self.p, self.raizes, self.ov = pontuado, raizes, overrides
        self.desc = dict(zip(plano["reduzido"], plano["descricao"]))
        self.cla = dict(zip(plano["reduzido"], plano["classificacao"]))
        self.cache, self.bases, self.fora_grupo, self.info = {}, {}, set(), {}
        desc_cla = dict(zip(plano["classificacao"], plano["desc_norm"]))
        self.existentes = set(plano["classificacao"])
        # grupos acima de cada conta analítica, do mais próximo para a raiz
        self.anc = {c: [(a, desc_cla[a]) for a in reversed(ancestrais(c, pontuado)) if a in desc_cla]
                    for c in self.an["classificacao"]}
        self.folha = {c: any(tem(d, *TERMOS_GRUPO_FOLHA) for _, d in g) for c, g in self.anc.items()}
        self.longo = {c: any(tem(d, *LONGO_PRAZO) for _, d in g) for c, g in self.anc.items()}
        self.referencial, self.ref = referencial_do_plano(plano)
        self.ecf = dict(zip(plano["reduzido"], plano["ecf"])) if "ecf" in plano.columns else {}

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

    def _ecf(self, reduzido, codigos):
        """1 = código referencial da conta confirma o alvo | -1 = contradiz | 0 = sem informação."""
        conta = self.ecf.get(reduzido, "")
        if not codigos or not conta:
            return 0
        return 1 if any(c.startswith(codigos) for c in conta.split("|")) else -1

    def _contexto(self, c, prefs):
        """Nomes dos grupos acima da conta dentro do escopo da busca (o grupo/raiz da busca não entra)."""
        p = next((x for x in prefs if sob(c, x, self.p)), "")
        return [d for a, d in self.anc.get(c, ()) if a != p and sob(a, p, self.p)]

    @staticmethod
    def _padroes(alvo):
        """[(termos, exclusões, implícito)]: no Passivo, também cada padrão sem o 'a pagar/a recolher'."""
        orig = [(must, nots, False) for must, nots in ALVOS[alvo][2]]
        if ALVOS[alvo][0] != PA:
            return orig
        extra = []
        for must, nots, _ in orig:
            resto = [x for x in must if x not in IMPLICITOS_PASSIVO]
            if (resto and len(resto) < len(must) and not set(resto) <= GENERICOS
                    and all((resto, nots) != (m, n) for m, n, _ in orig + extra)):
                extra.append((resto, nots, True))
        return orig + extra

    @staticmethod
    def _nivel(d, ctx, must, nots, nots_alvo, implicito):
        """0 = termos no nome da conta | 1 = parte no nome, o resto nos grupos acima | 2 = só pelos grupos |
        None = não casa."""
        if any(tem(d, x) for x in nots):
            return None
        falta = [x for x in must if not tem(d, x)]
        if (falta or implicito) and any(tem(d, x) for x in nots_alvo):
            return None          # sem o nome completo, a conta não pode ter termo que o alvo exclui
        if not falta:
            return 0
        if not ctx or any(tem(ctx[0], x) for x in nots) or set(falta) & SO_NO_NOME:
            return None
        if not all(any(tem(g, x) for g in ctx) for x in falta):
            return None
        return 1 if len(falta) < len(must) else 2

    @staticmethod
    def _vetado(ctx, veto):
        for g in ctx:
            achados = [x for x in veto if tem(g, x)]
            if achados and (set(achados) & VETO_SEMPRE or not tem(g, *TERMOS_FOLHA_FORTE)):
                return True
        return False

    def _buscar(self, alvo, prefs, patrimonial):
        """(reduzido, descrição, como, empatadas) da melhor conta do escopo ou None."""
        base = self._base(prefs) if prefs else self.an.iloc[0:0]
        if base.empty:
            return None
        codigos = self.ref.get(alvo, ())
        padroes = self._padroes(alvo)
        nots_alvo = sorted({x for _, ns, _ in padroes for x in ns})
        veto = [x for x in nots_alvo if x in VETO_CONTEXTO]
        cands = []
        for r in base.itertuples(index=False):
            ctx = self._contexto(r.classificacao, prefs)
            if veto and self._vetado(ctx, veto):
                continue
            for i, (must, nots, impl) in enumerate(padroes):
                nivel = self._nivel(r.desc_norm, ctx, must, nots, nots_alvo, impl)
                if nivel is not None:
                    cands.append((i, nivel, 3 if impl and nivel == 0 else nivel, r))
                    break
            else:
                if (codigos and alvo in ECF_SUFICIENTE and self._ecf(r.reduzido, codigos) == 1
                        and not any(tem(r.desc_norm, x) for x in nots_alvo)):
                    cands.append((len(padroes), 2, 4, r))
        if not cands:
            return None

        def ordem(item):
            i, nivel, como, r = item
            e = self._ecf(r.reduzido, codigos)
            # contradita pelo referencial vai por último; achada só pelo referencial fica depois das achadas pelo nome
            return (patrimonial and e == -1, como == 4, e != 1, i, nivel,
                    patrimonial and not self.folha.get(r.classificacao, False),
                    patrimonial and self.longo.get(r.classificacao, False), len(r.desc_norm), r.classificacao)
        cands.sort(key=ordem)
        topo = ordem(cands[0])[:-2]
        empatadas = [(c[3].reduzido, c[3].descricao) for c in cands[1:] if ordem(c)[:-2] == topo]
        r = cands[0][3]
        return r.reduzido, r.descricao, COMO[cands[0][2]], empatadas[:3]

    def sugestao(self, alvo, grupo):
        k = self.chave(alvo, grupo)
        if k not in self.cache:
            esc = ALVOS[alvo][0]
            res = self._buscar(alvo, self._prefixos(alvo, grupo), esc in (PA, AT))
            if res is None and esc == G and grupo:
                # encargos, provisões e benefícios costumam ficar num grupo irmão do mesmo departamento
                pai = next((a for a in reversed(ancestrais(grupo, self.p)) if a in self.existentes), "")
                if pai and pai not in self.raizes["RESULTADO"]:
                    res = self._buscar(alvo, (pai,), False)
            if res is None and esc == G and alvo in FORA_DO_GRUPO and self.raizes["RESULTADO"]:
                res = self._buscar(alvo, tuple(self.raizes["RESULTADO"]), False)
                if res:
                    self.fora_grupo.add(k)
            self.cache[k] = res[:2] if res else None
            self.info[k] = {"como": res[2], "empate": res[3]} if res else {}
        return self.cache[k]

    def alerta(self, alvo, grupo):
        """Aviso para a conta sugerida (não definida pelo usuário) achada de forma fraca ou empatada com outra."""
        k = self.chave(alvo, grupo)
        r = None if self.ov.get(k) else self.sugestao(alvo, grupo)
        if not r:
            return ""
        inf, msgs = self.info.get(k, {}), []
        como = inf.get("como")
        if como in COMO_FRACO or (como == COMO[3] and not self.folha.get(self.cla.get(r[0]), False)):
            msgs.append(f"{ALVOS[alvo][1]}: conta {r[0]} achada {como} — confirme")
        if inf.get("empate"):
            msgs.append(f"{ALVOS[alvo][1]}: {r[0]} empatou com " + ", ".join(x for x, _ in inf["empate"])
                        + " — confirme")
        return " | ".join(msgs)

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
                avisos = [] if a == alvos[0] else [f"fallback {ALVOS[alvos[0]][1]} → {ALVOS[a][1]}"]
                k = self.chave(a, grupo)
                if k in self.fora_grupo and not self.ov.get(k):
                    avisos.append(f"{ALVOS[a][1]}: conta fora do grupo do separador")
                aviso = self.alerta(a, grupo)
                if aviso:
                    avisos.append(aviso)
                return r[0], r[1], " | ".join(avisos)
        return "", "", f"❌ sem conta '{ALVOS[alvos[0]][1]}' (defina no Mapa de contas)"

# =====================================================================
# 4. LEITURA DOS RELATÓRIOS DA DOMÍNIO
# =====================================================================
def extrair_linhas(dados: bytes, nome: str):
    if nome.lower().endswith(".pdf"):
        # pypdfium2 lê as colunas coladas do relatório de rubricas sem embaralhar caracteres (pdfplumber embaralha).
        if pdfium is not None:
            pdf = pdfium.PdfDocument(dados)
            linhas = []
            try:
                for i in range(len(pdf)):
                    pagina = pdf[i]
                    tp = pagina.get_textpage()
                    txt = tp.get_text_range()
                    tp.close()
                    pagina.close()
                    linhas.extend(txt.replace("\r\n", "\n").replace("\r", "\n").split("\n"))
            finally:
                pdf.close()
            return linhas
        if pdfplumber is None:
            raise RuntimeError("Instale um leitor de PDF: pip install pypdfium2")
        linhas = []
        with pdfplumber.open(io.BytesIO(dados)) as pdf:
            for p in pdf.pages:
                linhas.extend((p.extract_text() or "").splitlines())
        return linhas
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return dados.decode(enc).splitlines()
        except UnicodeDecodeError:
            continue
    return []


# O tipo da rubrica é a ÚLTIMA ocorrência na linha (a descrição pode conter "Desconto", "Provento", "Informativa")
# e, nos PDFs da Domínio, a coluna Tipo vem COLADA à descrição ("...DOENCProvento Nenhuma"), por isso sem \s+.
RE_TIPO_RUB = r"(Provento|Desconto|Informativa|Informat|Inf\.\s*ded)"
RE_RUB = re.compile(r"^\s*(\d{1,5})\s+(.+)" + RE_TIPO_RUB)
RE_RUB_SEM_DESC = re.compile(r"^\s*(\d{1,5})\s*()" + RE_TIPO_RUB)


def _casa_rubrica(linha):
    return RE_RUB.match(linha) or RE_RUB_SEM_DESC.match(linha)


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
        m = _casa_rubrica(l)
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
        if _casa_rubrica(t):                       # já chegou nas rubricas
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


def itens_do_cadastro(cad: dict, tipos, separadores, filantropica=False):
    """Itens a contabilizar sem o relatório de pendências:
    Tipos 1/3/4 = todas as rubricas do cadastro; Tipo 2 = catálogo Empresa; Tipos 5/6 = catálogo Provisão
    (Encargos/Estorno PIS só para filantrópica); Tipo 10 = catálogo Outras Informações (itens 1 a 13),
    CADA ITEM EM LANÇAMENTO PRÓPRIO."""
    provisao = {c: v for c, v in ITENS_PROVISAO.items() if filantropica or c not in ITENS_PROVISAO_PIS}
    fontes = {
        "Empresa": {c: v[0] for c, v in ITENS_EMPRESA.items()},
        "Outras Informações": {c: v[0] for c, v in ITENS_OUTRAS.items()},
        "Provisão de Férias": provisao, "Provisão de 13º": provisao,
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


def unificar_com_folha(itens, mesma_ferias: bool, mesma_rescisao: bool, manter_exclusivas: bool = False):
    """'Usar a mesma configuração da folha normal para Férias/Rescisão' (Domínio):
    os itens dessas seções passam a usar o lançamento da Folha (Tipo 1).
    - rubrica que já existe na Folha (mesmo separador e código): o item de Férias/Rescisão é descartado;
    - rubrica exclusiva de Férias/Rescisão: com manter_exclusivas continua no próprio Tipo 3/4
      (como a Domínio grava, p.ex. Aviso Prévio e 51 Líquido Rescisão);
      sem manter_exclusivas é movida para a Folha (Tipo 1).
    Retorna (itens, quantidade unificada, quantidade de exclusivas mantidas)."""
    unir = {s for s, f in (("Férias", mesma_ferias), ("Rescisão", mesma_rescisao)) if f}
    if not unir:
        return itens, 0, 0
    folha = {(it["sep_cod"], it["codigo"]) for it in itens
             if it["secao"] not in unir and TIPO_INTEGRACAO.get(it["secao"]) == 1}
    out, movidos, mantidos = [], 0, 0
    for it in itens:
        if it["secao"] not in unir:
            out.append(it)
            continue
        k = (it["sep_cod"], it["codigo"])
        if k in folha:
            movidos += 1
            continue
        if manter_exclusivas:
            mantidos += 1
            out.append(it)
        else:
            movidos += 1
            folha.add(k)
            out.append({**it, "secao": "Folha Normal"})
    return out, movidos, mantidos

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


def passivo_secao(d):
    """Conta de passagem do cálculo. Férias e Rescisão também passam por Salários a Pagar, como a Domínio integra:
    o líquido das férias sai pela 937 (Adiantamento de Férias) e o da rescisão pela 51 (Rescisões a Pagar)."""
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
    """SALARIO MATERNIDADE / LIC.MATERN / LIC.MATER.INSS / LIC.MAT.INSS / SAL MAT"""
    return tem(d, "MATERN*", "MATER") or (tem(d, "LIC", "LICENC*", "SAL") and tem(d, "MAT"))


# Rubricas de natureza salarial sem regra própria: vão para Salários sem pedir revisão
TERMOS_SALARIO = ("SALARI*", "SAL", "ORDENADO*", "HORA*", "DIA", "DIAS", "DSR", "RSR", "REPOUSO", "ADICIONA*",
                  "ADIC", "INSALUB*", "PERICUL*", "NOTURN*", "ANUENIO", "BIENIO", "TRIENIO", "QUINQUENIO",
                  "SOBREAVISO", "PRONTIDAO", "GORJETA", "ACUMULO", "QUEBRA DE CAIXA", "REFLEXO*", "MEDIA*",
                  "VANTAGE*", "AFAST*", "ANTECIP*", "DIF*", "HORISTA", "MENSALISTA", "AULA*", "PROFESSOR*",
                  "JORNADA", "INTERVALO", "CONVOCAC*", "INTERMITENT*", "TEMPO DE SERVICO", "PISO", "SALDO")


def alvo_dre_provento(d):
    """(alvos, regra específica?) — sem regra específica a rubrica cai em Salários e vai para revisão."""
    if tem(d, "PRO LABORE"): return ["PRO_LABORE", "SALARIOS"], True
    if tem(d, "BOLSA", "ESTAGI*", "RECESSO"): return ["BOLSA", "SALARIOS"], True
    # Licença remunerada (férias coletivas s/ período aquisitivo) = natureza salarial;
    # a versão "GOZADA(S)" é férias de fato.
    if e_licenca_remunerada(d) and not tem(d, "GOZ*"): return ["SALARIOS"], True
    if tem(d, "HORAS EXTRA*", "EXTRAS", "BANCO DE HORAS"): return ["HE", "SALARIOS"], True
    if e13(d): return ["DECIMO", "SALARIOS"], True
    if e_ferias(d): return ["FERIAS", "SALARIOS"], True
    if tem(d, "AVISO PREVIO", "INDENIZ*", "MULTA", "ESTABILIDADE"): return ["INDENIZ", "SALARIOS"], True
    if tem(d, "PREMIO*", "GRATIFIC*", "BONUS", "PLR", "PARTICIPACAO LUCRO*", "PARTIC LUCRO*"):
        return ["PREMIOS", "SALARIOS"], True
    if tem(d, "COMISS*"): return ["COMISSOES", "PREMIOS", "SALARIOS"], True
    if tem(d, "VALE TRANSPORTE", "VT", "MOBILIDADE"): return ["VT"], True
    if tem(d, "PLANO DE SAUDE", "PLANO SAUDE", "ODONTO*", "ASSISTENCIA MEDICA"): return ["ASSIST"], True
    if tem(d, "REFEICAO", "ALIMENTACAO", "CESTA", "VR"): return ["ALIM"], True
    if tem(d, "SEGURO*"): return ["SEGURO", "ASSIST"], True
    if tem(d, "TREINAMENTO", "CURSO*"): return ["TREIN"], True
    return ["SALARIOS"], tem(d, *TERMOS_SALARIO)


def classificar_provento(d, baixa_prov):
    L = passivo_secao(d)
    if tem(d, "ESTOURO", "TROCO", "ARRED*"):
        return ["ADIANT_SAL"], L, "Estouro/troco/arredondamento: crédito a receber do empregado", False
    if tem(d, "ADIANT*", "ADTO"):
        return alvo_adiant(d), L, "Adiantamento pago: Ativo × obrigação", False
    if tem(d, "DEV", "DEVOLUCAO", "ESTORNO*") and tem(d, "EMPREST*", "CONSIG*"):
        return ["EMPREST"], L, "Devolução/estorno de consignado descontado", True
    # Maternidade reembolsada pelo INSS: SEMPRE pela conta-ponte (baixa no Tipo 10 "Compensação da Ded. Sal.
    # Maternidade"); lançar direto em INSS a Recolher e também compensar no Tipo 10 baixaria o INSS duas vezes.
    # Prorrogação (Empresa Cidadã) e bolsa de estagiária não são reembolsadas pelo INSS.
    if e_maternidade(d) and not tem(d, "DESC*", "DEDUC*", "PROR*", "EMPREGADOR", "BOLSA*", "ESTAGI*"):
        return ["BENEF_INSS"], L, "Salário-maternidade reembolsável → conta-ponte (baixa no Tipo 10)", False
    if tem(d, "INSS", "IRRF") and tem(d, "A MAIOR", "DIF*", "DEVOL*", "RESTITUI*"):
        return (["IRRF_REC"] if tem(d, "IRRF") else ["INSS_REC"]), L, "Restituição de retenção a maior", False
    if tem(d, "SALARIO FAMILIA", "SAL FAM*"):
        return ["BENEF_INSS"], L, "Salário-família → conta-ponte até a compensação na guia", False
    if tem(d, "1 12") and tem(d, "INDENIZ*"):
        # Projeção do aviso prévio indenizado: não foi provisionada, nunca baixa a provisão
        return (["FERIAS", "INDENIZ"] if e_ferias(d) else ["INDENIZ", "DECIMO"]), L, \
            "1/12 indenizado do aviso prévio (não provisionado)", False
    dre, especifica = alvo_dre_provento(d)
    if baixa_prov and dre[0] in ("FERIAS", "DECIMO"):
        prov = ["P_PROV_FER", "FERIAS"] if dre[0] == "FERIAS" else ["P_PROV_13", "DECIMO"]
        return prov, L, "Baixa contra a provisão", False
    if not especifica:
        return dre, L, "Sem regra específica — confirme a natureza da rubrica (padrão: Salários)", True
    return dre, L, "", False


def classificar_desconto(d, baixa_prov):
    L = passivo_secao(d)
    def r(c, obs="", rev=False): return L, c, obs, rev
    if tem(d, "ESTOURO", "TROCO", "ARRED*"):
        return r(["ADIANT_SAL"], "Baixa de estouro/troco/arredondamento anterior")
    if tem(d, "LIQUIDO") and tem(d, "RESCIS*"):
        # O cálculo de Rescisão passa por Salários a Pagar e esta rubrica (Tipo 4) leva o líquido para
        # Rescisões a Pagar, quitada no pagamento.
        return ((["SAL_PAGAR"], ["RESC_PAGAR"],
                 "Líquido da rescisão pago à parte — transfere p/ Rescisões a Pagar", True))
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
    if tem(d, "VALE TRANSPORTE", "VT", "MOBILIDADE"): return r(["VT"], "Recuperação do custo do vale-transporte/mobilidade")
    if tem(d, "REFEICAO", "ALIMENTACAO", "VR", "CESTA", "SUPERMERCADO"):
        return r(["ALIM"], "Recuperação do custo do benefício")
    if tem(d, "SEGURO*"): return r(["SEGURO", "ASSIST"], "Recuperação do custo do benefício")
    # Bolsa-auxílio (estágio): todo desconto da bolsa (afastamento, antecipação, 1/3 do recesso) volta para a
    # conta da bolsa, a mesma do provento
    if tem(d, "BOLSA*", "BOL"):
        if tem(d, "LIQUIDO", "LIQ") and tem(d, "RESCIS*"):
            return r(["BOLSA", "SALARIOS"], "Líquido da rescisão do estagiário já pago — estorna a bolsa repetida na "
                                         "folha", True)
        return r(["BOLSA", "SALARIOS"], "Redução do custo da bolsa-auxílio")
    if tem(d, "AVISO PREV*", "MULTA", "ESTABILIDADE", "INDENIZ*"):
        return r(["INDENIZ", "SALARIOS"], "Indenização devida pelo empregado", True)
    if e13(d):
        return r((["P_PROV_13"] if baixa_prov else []) + ["DECIMO", "SALARIOS"],
                 "Redução do 13º contra a provisão" if baixa_prov else "Redução do custo de 13º")
    if e_maternidade(d): return r(["BENEF_INSS"], "Estorno de salário-maternidade pago a maior", True)
    if e_ferias(d):
        return r((["P_PROV_FER"] if baixa_prov else []) + ["FERIAS", "SALARIOS"],
                 "Redução das férias contra a provisão" if baixa_prov else "Redução do custo de férias")
    if tem(d, "COMISS*"): return r(["COMISSOES", "PREMIOS", "SALARIOS"], "Estorno de comissões")
    if tem(d, "PREMIO*", "GRATIFIC*"): return r(["PREMIOS", "SALARIOS"], "Estorno de prêmio/gratificação")
    if tem(d, "GORJETA*"):
        return r(["SALARIOS"], "Gorjeta já recebida dos clientes: estorna o provento da estimativa, fica só o encargo "
                               "(Solução 12819)")
    if tem(d, "LIQUIDO", "LIQ") and tem(d, "CONVOCAC*", "INTERMITENT*"):
        return r(["ADIANT_SAL"], "Líquido já pago nas convocações anteriores do mês — baixa como adiantamento "
                                 "(confirme como as convocações são contabilizadas)", True)
    if tem(d, "AJUDA*", "APOIO*") and tem(d, "FINANC*"):
        return r(["SALARIOS"], "Apoio Financeiro do governo (MP 1.230/2024, Portaria MTE 991/2024) pago direto ao "
                               "empregado — reduz o salário pago pela empresa; confirme a conta", True)
    # Antecipação salarial (Solução 3472): no reajuste, a diferença retroativa entra cheia como provento e o que já
    # foi antecipado é descontado — estorna custo de salários, não é adiantamento a receber
    if tem(d, "ANTECIP*", "ANTEC") or (tem(d, "ANT") and tem(d, "DESC*", "PARC*")):
        return r(["SALARIOS"], "Desconto da antecipação salarial já paga — reduz o custo de salários")
    if (tem(d, "FALTA*", "ATRASO*", "DSR", "HORAS", "DIAS", "PAGO A MAIOR", "INATIV*", "SUSPENS*", "AFAST*")
            or (tem(d, "DIF*") and tem(d, "SALARI*", "SAL"))):
        return r(["SALARIOS"], "Redução do custo de salários")
    if tem(d, *TERMOS_SALARIO) and not tem(d, "SALDO"):
        return r(["SALARIOS"], "Desconto de natureza salarial — reduz o custo de salários (confirme)", True)
    return r([], "Desconto sem regra — definir conta manualmente", True)


def classificar_informativa(d, tipo, baixa_prov=False):
    """Regra: informativa só integra se a descrição tiver FGTS (com ou sem ponto)
    e não for base/dedução de base. Todo o resto: não integrar.
    Com baixa contra a provisão, o FGTS de férias/13º baixa o FGTS provisionado (exceto o do 1/12 indenizado)."""
    if e_base(d):
        return [], [], "Base/dedução de base — não gera lançamento", True
    if not e_fgts(d):
        return [], [], "Informativa sem FGTS na descrição — não integrar", True
    multa = tem(d, "40%", "20%")
    dre = ["INDENIZ", "FGTS"] if multa else ["FGTS"]
    if baixa_prov and not multa and not tem(d, "INDENIZ*"):
        if e13(d):
            dre = ["P_FGTS_13", "FGTS"]
        elif e_ferias(d):
            dre = ["P_FGTS_FER", "FGTS"]
    obs = "FGTS de férias/13º contra a provisão" if dre[0].startswith("P_") else "Encargo FGTS"
    if tem(d, "A MAIOR") or tipo == "Inf. dedutora":
        return ["FGTS_REC"], dre, f"Estorno de FGTS a maior ({obs})", True
    return dre, ["FGTS_REC"], obs, False


def classificar_provisao(d, secao):
    """Itens das seções Provisão de Férias / Provisão de 13º."""
    if tem(d, "TRANSFER*"):
        return [], [], "Transferência de provisão entre empresas — configure manualmente na Domínio", True
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


def classificar_item_empresa(codigo, d, baixa_prov, filantropica=False):
    cat = ITENS_EMPRESA.get(codigo)
    if cat and confere(d, norm(cat[0])):
        _, nat, comp = cat
        origem = "Catálogo Empresa"
    else:
        nat, comp = inferir_item_empresa(d)
        origem = (f"Inferido — no catálogo o item {codigo} é '{cat[0]}'" if cat
                  else f"Inferido — item {codigo} fora do catálogo")
    sfx = {"F": "FER", "13": "13"}.get(comp)
    # o 13º indenizado (projeção do aviso prévio) não foi provisionado: o INSS sobre ele vai para despesa
    prov = bool(baixa_prov and sfx and nat in ("INSS", "SENAI", "PIS", "ISENCAO") and not tem(d, "INDENIZ*"))
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
    elif nat == "ISENCAO" and filantropica:
        # Estorno do INSS patronal não devido: inverso do lançamento do item de INSS correspondente
        deb = ["INSS_REC"]
        cred = [f"P_INSS_{sfx}", f"P_PROV_{sfx}"] if prov else ["INSS"]
    else:
        deb, cred = [], []
    obs, rev = OBS_EMPRESA.get(nat, OBS_EMPRESA[None])
    if nat == "ISENCAO" and filantropica:
        obs = "Isenção filantropia: estorno do INSS patronal não devido (inverso do item de INSS correspondente)"
        rev = True
    if nat == "PIS" and filantropica:
        obs, rev = "PIS s/ folha — entidade filantrópica/sem fins lucrativos", False
    if prov:
        obs += " — contra a provisão"
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
# 5b. CONFERÊNCIA DO ARQUIVO CONTRA O LEIAUTE DA IMPORTAÇÃO DE TABELAS
#     aba "evento" = FOINTEGCONT  (codi_emp, i_separador, i_codigo, tipo, descricao char(40),
#                                  conta_deb, conta_cred, i_historico, comp_hist varchar(512))
#     aba "integra" = FOINTEGCONTEVE (codi_emp, i_separador, i_codigo, tipo, i_eventos)
#     Chave primária de FOINTEGCONT: (codi_emp, i_separador, i_codigo, tipo)
#     Chave primária de FOINTEGCONTEVE: (codi_emp, i_separador, i_codigo, i_eventos, tipo), com FK para
#     FOINTEGCONT; na prática um evento pertence a UM lançamento por (empresa, separador, tipo).
# =====================================================================
def _eh_int(v) -> bool:
    if isinstance(v, bool):
        return False
    if isinstance(v, numbers.Integral):
        return True
    return isinstance(v, str) and bool(re.fullmatch(r"-?\d+", v.strip()))


def e_nulo(v) -> bool:
    """Lado sem conta (partidas múltiplas), em qualquer caixa."""
    return txt_cel(v).upper() == NULO


def parte_do_lancamento(deb, cred) -> int:
    """1 = débito e crédito | 2 = só débito (CONTA_CRED nula) | 3 = só crédito (CONTA_DEB nula)."""
    return 3 if e_nulo(deb) else 2 if e_nulo(cred) else 1


def checar_layout(evento: pd.DataFrame, integra: pd.DataFrame, multi: bool = False):
    out = []

    def _conta_ok(v):
        return _eh_int(v) or e_nulo(v)
    def add(nome, nivel, problemas):
        ok = not problemas
        out.append({"Verificação": nome, "Nível": nivel,
                    "Resultado": "✅" if ok else ("❌" if nivel == "Erro" else "⚠️"),
                    "Detalhe": "ok" if ok else problemas})

    add("Aba evento: colunas e ordem do leiaute", "Erro",
        "" if list(evento.columns) == COLS_EVENTO else "colunas diferentes do leiaute")
    add("Aba integra: colunas e ordem do leiaute", "Erro",
        "" if list(integra.columns) == COLS_INTEGRA else "colunas diferentes do leiaute")

    c_emp, c_sep, c_cod, c_tip, c_des, c_deb, c_cred, c_his, c_com = COLS_EVENTO
    ruim = {c: int((~evento[c].map(_eh_int)).sum()) for c in (c_emp, c_sep, c_cod, c_tip)}
    ruim.update({c: int((~evento[c].map(_conta_ok)).sum()) for c in (c_deb, c_cred)})
    ruim = {c: n for c, n in ruim.items() if n}
    add(f"Aba evento: empresa, separador, código, tipo e contas numéricos (integer; '{NULO}' = lado sem conta)", "Erro",
        "; ".join(f"{c[:28]}: {n} valor(es)" for c, n in ruim.items()))
    n_vazio = int((evento[c_his].astype(str).str.strip() == "").sum())
    n_nao_num = int(((evento[c_his].astype(str).str.strip() != "") & (~evento[c_his].map(_eh_int))).sum())
    add("Aba evento: histórico numérico (i_historico integer)", "Erro",
        f"{n_nao_num} histórico(s) não numérico(s)" if n_nao_num else "")
    add("Aba evento: histórico preenchido (todos os lançamentos reais têm histórico)", "Aviso",
        f"{n_vazio} lançamento(s) sem histórico" if n_vazio else "")
    d = evento[c_des].astype(str)
    p_des = []
    if (d.str.strip() == "").any():
        p_des.append(f"{int((d.str.strip() == '').sum())} vazia(s)")
    if (d.str.len() > LIM_DESC_EVENTO).any():
        p_des.append(f"{int((d.str.len() > LIM_DESC_EVENTO).sum())} acima de {LIM_DESC_EVENTO} caracteres")
    add(f"Aba evento: descrição preenchida e com até {LIM_DESC_EVENTO} caracteres (char(40))", "Erro", "; ".join(p_des))
    add(f"Aba evento: complemento com até {LIM_COMP_HIST} caracteres (varchar(512))", "Erro",
        f"{int((evento[c_com].astype(str).str.len() > LIM_COMP_HIST).sum())} acima do limite"
        if (evento[c_com].astype(str).str.len() > LIM_COMP_HIST).any() else "")
    tipos_fora = sorted({t for t in evento[c_tip].tolist() if t not in TIPOS_LAYOUT})
    add("Aba evento: tipo da integração no conjunto gerado (1, 2, 3, 4, 5, 6, 10)", "Erro",
        f"tipos não previstos: {tipos_fora}" if tipos_fora else "")
    n_deb, n_cred = evento[c_deb].map(e_nulo), evento[c_cred].map(e_nulo)
    ambos_nulos = n_deb & n_cred
    add("Aba evento: todo lançamento tem ao menos uma conta (débito ou crédito)", "Erro",
        f"{int(ambos_nulos.sum())} lançamento(s) com as duas contas '{NULO}'" if ambos_nulos.any() else "")
    iguais = (evento[c_deb].map(txt_cel) == evento[c_cred].map(txt_cel)) & ~n_deb
    add("Aba evento: conta débito diferente da conta crédito", "Erro",
        f"{int(iguais.sum())} lançamento(s) com débito = crédito" if iguais.any() else "")
    if not multi and (n_deb | n_cred).any():
        add("Partidas dobradas: todo lançamento com débito e crédito", "Erro",
            f"{int((n_deb | n_cred).sum())} lançamento(s) com um lado '{NULO}'")
    pk_ev = [c_emp, c_sep, c_cod, c_tip]
    dup_ev = evento.duplicated(pk_ev, keep=False)
    add("Aba evento: chave única (empresa + separador + código + tipo)", "Erro",
        f"{int(dup_ev.sum())} linha(s) repetem a chave — a importação rejeitaria/sobrescreveria" if dup_ev.any() else "")

    i_emp, i_sep, i_cod, i_tip, i_rub = COLS_INTEGRA
    ruim_i = {c: int((~integra[c].map(_eh_int)).sum()) for c in COLS_INTEGRA}
    ruim_i = {c: n for c, n in ruim_i.items() if n}
    add("Aba integra: todas as colunas numéricas (integer)", "Erro",
        "; ".join(f"{c[:28]}: {n}" for c, n in ruim_i.items()))
    dup_in = integra.duplicated([i_emp, i_sep, i_cod, i_tip, i_rub], keep=False)
    add("Aba integra: chave única (empresa + separador + código + tipo + rubrica)", "Erro",
        f"{int(dup_in.sum())} linha(s) repetidas" if dup_in.any() else "")
    dup_rub = integra.duplicated([i_emp, i_sep, i_tip, i_rub], keep=False)
    add("Aba integra: cada rubrica/item em UM só lançamento por empresa + separador + tipo", "Erro",
        f"{int(dup_rub.sum())} vínculo(s) com a mesma rubrica em lançamentos diferentes" if dup_rub.any() else "")
    chaves_ev = set(map(tuple, evento[pk_ev].values.tolist()))
    chaves_in = set(map(tuple, integra[[i_emp, i_sep, i_cod, i_tip]].values.tolist()))
    orfaos = chaves_in - chaves_ev
    add("Aba integra → evento: todo vínculo aponta para um lançamento existente (chave estrangeira)", "Erro",
        f"{len(orfaos)} vínculo(s) sem lançamento correspondente" if orfaos else "")
    sem_vinculo = chaves_ev - chaves_in
    add("Aba evento → integra: todo lançamento tem ao menos uma rubrica vinculada", "Aviso",
        f"{len(sem_vinculo)} lançamento(s) sem rubrica vinculada" if sem_vinculo else "")
    if multi:
        t1 = evento[evento[c_tip] == 1]
        p_seq = []
        for (e_, s_), g_ in t1.groupby([c_emp, c_sep]):
            cods = sorted(int(x) for x in g_[c_cod])
            if cods != list(range(cods[0], cods[0] + len(cods))):
                p_seq.append(f"separador {s_}: códigos fora de sequência {cods[0]}..{cods[0] + len(cods) - 1}")
        add("Partidas múltiplas: lançamentos do Tipo 1 em sequência numérica (1, 2, 3…) por separador", "Erro",
            "; ".join(p_seq))
        liq_ok = integra[(integra[i_tip] == 1) & (integra[i_rub].map(lambda v: _eh_int(v) and int(v) < 0))]
        seps_t1 = set(t1[c_sep].tolist())
        sem_liq = sorted(seps_t1 - set(liq_ok[i_sep].tolist()))
        add("Partidas múltiplas: há lançamento de líquido (item negativo) em cada separador do Tipo 1", "Erro",
            f"separador(es) sem líquido: {sem_liq}" if sem_liq else "")
        mistura, ordem_ruim = [], []
        for (e_, s_), g_ in t1.groupby([c_emp, c_sep]):
            g_ = g_.sort_values(c_cod)
            partes = [parte_do_lancamento(d_, c_) for d_, c_ in zip(g_[c_deb], g_[c_cred])]
            if 1 in partes and not {2, 3} & set(partes):
                mistura.append(f"separador {s_}: nenhum lançamento de uma só conta")
            # o lançamento composto (proventos, descontos e líquidos) fica contíguo, antes dos de duas contas
            if 1 in partes and any(p != 1 for p in partes[partes.index(1):]):
                ordem_ruim.append(f"separador {s_}")
        add("Partidas múltiplas: Tipo 1 com lançamentos de uma só conta (proventos/descontos/líquido)", "Aviso",
            "; ".join(mistura))
        add("Partidas múltiplas: proventos, descontos e líquidos em sequência, antes dos lançamentos de duas contas",
            "Erro", ("lançamento de duas contas no meio do composto: " + ", ".join(ordem_ruim)) if ordem_ruim else "")
    return out

# =====================================================================
# 6. CONFIGURAÇÃO POR EMPRESA (.json)
# =====================================================================
CFG_WIDGETS = (("historico", "w_hist"), ("baixa_prov", "w_baixa"), ("socio_adm", "w_socio"),
               ("bloqueio_extra", "w_bloq"), ("empresa", "w_cod"), ("complemento", "w_compl"),
               ("tipos_cadastro", "w_tipos_cad"), ("separadores_cadastro", "w_seps_cad"),
               ("mesma_rescisao", "w_mesma_resc"), ("mesma_ferias", "w_mesma_fer"),
               ("filantropica", "w_filant"), ("manter_exclusivas", "w_exclusivas"),
               ("partidas_multiplas", "w_multi"), ("liquidos", "w_liquidos"), ("leiaute", "w_leiaute"),
               ("inicio_tipo1", "w_inicio_t1"))


def aplicar_config(dados: bytes):
    fid = hashlib.md5(dados).hexdigest()
    if st.session_state.get("cfg_id") == fid:
        return
    try:
        cfg = json.loads(dados.decode("utf-8-sig"))
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
            st.session_state[wk] = txt_cel(cfg[chave]) if wk in ("w_cod", "w_leiaute") else cfg[chave]
    for t, v in cfg.get("hist_tipo", {}).items():
        st.session_state[f"w_hist_{t}"] = str(v)


def md5(txt: str) -> str:
    return hashlib.md5(txt.encode("utf-8")).hexdigest()[:10]

# =====================================================================
# 6b. ARQUIVOS DA IMPORTAÇÃO VIA TABELA (importação por leiaute)
#     TXT: um registro por linha, campos separados por TAB, sem cabeçalho, Windows-1252, CR+LF, na ordem do
#     leiaute (evento = FOINTEGCONT, integra = FOINTEGCONTEVE). A importação grava 0 em campo numérico vazio lido
#     do arquivo; NULL só sai de campo com "Valor fixo = NULO" no leiaute, que não ocupa coluna no TXT. Em
#     partidas múltiplas a parte 2 vai sem a coluna CONTA_CRED e a parte 3 sem a coluna CONTA_DEB.
# =====================================================================
def _campo_txt(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    if isinstance(v, numbers.Integral) and not isinstance(v, bool):
        return str(int(v))
    return re.sub(r"[\t\r\n]+", " ", str(v)).strip()


def _txt(linhas) -> bytes:
    corpo = "".join("\t".join(_campo_txt(v) for v in lin) + "\r\n" for lin in linhas)
    return corpo.encode("cp1252", errors="replace")


def fora_do_cp1252(textos):
    """Textos com caractere que o Windows-1252 não tem (no TXT viraria '?')."""
    out = []
    for t in textos:
        try:
            str(t).encode("cp1252")
        except UnicodeEncodeError:
            out.append(str(t))
    return out


def _sql_leiaute(n, nulo):
    """Leiaute n da importação: o campo 'nulo' de FOINTEGCONT com Valor fixo = NULO; os demais lidos do arquivo."""
    out = []
    for campo in ("CONTA_DEB", "CONTA_CRED"):
        valor = "origem_valor = 'F', valor_fixo = 'NULO'" if campo == nulo else "origem_valor = 'I', valor_fixo = ''"
        out.append(f"update bethadba.gelayout_campos set {valor}\n where i_layout = {n} and upper(i_nome) = "
                   f"'{TAB_EVENTO}' and upper(trim(campo)) = '{campo}';")
    return out + ["commit;", "", "-- conferência",
                  f"select i_nome, i_sequencial, campo, origem_valor, valor_fixo\n  from bethadba.gelayout_campos\n"
                  f" where i_layout = {n} and upper(i_nome) = '{TAB_EVENTO}'\n"
                  f"   and upper(trim(campo)) in ('CONTA_DEB', 'CONTA_CRED')\n order by i_sequencial;"]


def _texto(linhas) -> bytes:
    return ("\r\n".join(linhas) + "\r\n").encode("cp1252", errors="replace")


def _sql(linhas) -> bytes:
    return ("\r\n".join(linhas) + "\r\n").encode("utf-8-sig")


def pacote_importacao(evento, integra, emp, leiaute="", gerado_em=None):
    """.zip da importação via tabela: TXT de cada parte, estrutura dos arquivos, LEIA-ME, scripts do leiaute (com o
    número do leiaute) e de conferência. Retorna (bytes, {parte: (lançamentos, vínculos)}, passos)."""
    c_emp, c_sep, c_cod, c_tip, *_ = COLS_EVENTO
    c_deb, c_cred = COLS_EVENTO[5], COLS_EVENTO[6]
    i_emp, i_sep, i_cod, i_tip, _ = COLS_INTEGRA
    ev = evento.assign(_parte=[parte_do_lancamento(d, c) for d, c in zip(evento[c_deb], evento[c_cred])])
    da_chave = {k: p for k, p in zip(map(tuple, ev[[c_emp, c_sep, c_cod, c_tip]].values.tolist()), ev["_parte"])}
    it = integra.assign(_parte=[da_chave.get(k, 1) for k in
                                map(tuple, integra[[i_emp, i_sep, i_cod, i_tip]].values.tolist())])
    usadas = sorted(set(ev["_parte"])) or [1]
    em_pastas = usadas != [1]
    leiaute = txt_cel(leiaute)
    arquivos, estrutura, resumo = {}, [], {}
    for p in usadas:
        pasta, _, nulo = PARTES[p]
        campos = [c for c in CAMPOS_EVENTO if c != nulo]
        cols = [COLS_EVENTO[CAMPOS_EVENTO.index(c)] for c in campos]
        e_p, i_p = ev[ev["_parte"] == p], it[it["_parte"] == p]
        pref = f"{pasta}/" if em_pastas else ""
        arquivos[f"{pref}evento.txt"] = _txt(e_p[cols].values.tolist())
        arquivos[f"{pref}integra.txt"] = _txt(i_p[COLS_INTEGRA].values.tolist())
        resumo[p] = (len(e_p), len(i_p))
        for pos, (c, d) in enumerate(zip(campos, cols), 1):
            estrutura.append([f"{pref}evento.txt", pos, TAB_EVENTO, c, d])
        if nulo:
            estrutura.append([f"{pref}evento.txt", "", TAB_EVENTO, nulo, "Valor fixo = NULO no leiaute (fora do TXT)"])
        for pos, (c, d) in enumerate(zip(CAMPOS_INTEGRA, COLS_INTEGRA), 1):
            estrutura.append([f"{pref}integra.txt", pos, TAB_INTEGRA, c, d])
    with io.StringIO() as s:
        w = csv.writer(s, delimiter=";")
        w.writerow(["arquivo", "posicao", "tabela", "campo", "descricao"])
        w.writerows(estrutura)
        arquivos["estrutura_arquivos.csv"] = s.getvalue().encode("utf-8-sig")

    n_lei = leiaute if leiaute.isdigit() else "N"
    passos = [f"Antes de importar, rode conferir_antes_de_importar_emp{emp}.sql (somente leitura): ele precisa "
              "voltar vazio. Se trouxer linhas, esses códigos de lançamento já existem na empresa"
              + (" — em partidas múltiplas, ajuste no app o 'Primeiro código do Tipo 1' para depois do último "
                 "código existente." if em_pastas else ".")]
    if em_pastas:
        estado = None
        for p in usadas:
            pasta, titulo, nulo = PARTES[p]
            ajuste = ("normal (CONTA_DEB e CONTA_CRED lidos do arquivo)" if not nulo else
                      f"{nulo} com Valor fixo = NULO" + ("" if estado is None else
                                                         f" e {estado} de volta para 'lido do arquivo'"))
            if p > 1 and leiaute.isdigit():
                ajuste += f" (script leiaute_{n_lei}_antes_da_parte_{p}.sql)"
            passos.append(f"Parte {p} — {titulo}: {resumo[p][0]} lançamento(s). Leiaute: {ajuste}. Importe "
                          f"{pasta}/evento.txt e depois {pasta}/integra.txt.")
            estado = nulo
        passos.append(f"No fim, volte {estado} para 'lido do arquivo' (leiaute normal)"
                      + (f" — script leiaute_{n_lei}_restaurar.sql." if leiaute.isdigit() else "."))
    else:
        passos.append(f"Importe evento.txt e depois integra.txt com o leiaute normal (todos os campos lidos do "
                      f"arquivo): {resumo[1][0]} lançamento(s) e {resumo[1][1]} vínculo(s).")

    agora = (gerado_em or dt.datetime.now()).strftime("%d/%m/%Y %H:%M")
    leia = ["INTEGRAÇÃO CONTÁBIL DA FOLHA — IMPORTAÇÃO VIA TABELA (importação por leiaute)",
            f"Empresa {emp} | gerado em {agora} | {len(ev)} lançamento(s) em {TAB_EVENTO} e {len(it)} vínculo(s) "
            f"em {TAB_INTEGRA}", "",
            "FORMATO DOS ARQUIVOS",
            f"   - evento.txt  -> {TAB_EVENTO}: " + ", ".join(CAMPOS_EVENTO),
            f"   - integra.txt -> {TAB_INTEGRA}: " + ", ".join(CAMPOS_INTEGRA),
            "   - Um registro por linha, campos separados por TAB, sem cabeçalho, Windows-1252, CR+LF, na ordem do "
            "leiaute (posições em estrutura_arquivos.csv).",
            "   - Importe sempre o evento.txt antes do integra.txt: o vínculo precisa do lançamento gravado."]
    if em_pastas:
        leia += ["", "POR QUE EM PARTES (lançamentos de uma só conta)",
                 "   A importação via tabela grava 0 em campo numérico vazio lido do arquivo, e nenhum texto no TXT "
                 "vira branco. O lado sem conta só fica em branco (NULL) com o campo no leiaute em Valor fixo = NULO — "
                 "e campo com valor fixo não ocupa coluna no TXT. Por isso a parte 2 vai sem a coluna CONTA_CRED e a "
                 "parte 3 sem a coluna CONTA_DEB.",
                 "", "AJUSTE DO LEIAUTE (escolha um jeito)",
                 "   a) Na tela da importação por leiaute: tabela FOINTEGCONT, campo indicado em cada parte -> "
                 "Valor fixo = NULO.",
                 "   b) Por SQL, no Interactive SQL (dbisql) com usuário com permissão de escrita; depois de cada "
                 "script, feche e abra a tela de importação."
                 + ("" if leiaute.isdigit() else " Para gerar os scripts, informe o número do leiaute no app "
                    "(Opções avançadas).")]
    leia += ["", "PASSO A PASSO"] + [f"   {i}. {t}" for i, t in enumerate(passos, 1)]
    leia += ["", "CONFERÊNCIA NA DOMÍNIO",
             "   - Arquivo > Contabilidade > Configurar Integração: confira lançamentos de cada guia"
             + ("; na guia Folha, proventos só com Conta Débito, descontos e líquidos só com Conta Crédito "
                "e os itens de líquido selecionados." if em_pastas else "."),
             f"   - conferir_contas_zero_emp{emp}.sql (somente leitura) não pode trazer linha: conta 0 não existe.",
             f"   - Lançamentos importados antes com 0 na contrapartida: corrigir_contas_zero_emp{emp}.sql troca o 0 por "
             "branco (NULL) só nesta empresa.",
             "   - Primeira importação neste banco: faça antes a cópia de segurança conforme o procedimento da sua "
             "área."]
    arquivos["LEIA-ME.txt"] = _texto(leia)

    if em_pastas and leiaute.isdigit():
        estado = None
        for p in usadas:
            nulo = PARTES[p][2]
            if p > 1:
                arquivos[f"leiaute_{n_lei}_antes_da_parte_{p}.sql"] = _sql(
                    [f"-- Leiaute {n_lei}: antes da parte {p} ({PARTES[p][1]})",
                     "-- Execute no Interactive SQL (dbisql); depois feche e abra a tela de importação.", ""]
                    + _sql_leiaute(n_lei, nulo))
            estado = nulo
        arquivos[f"leiaute_{n_lei}_restaurar.sql"] = _sql(
            [f"-- Leiaute {n_lei}: CONTA_DEB e CONTA_CRED de volta para 'lido do arquivo' (depois da última parte).", ""]
            + _sql_leiaute(n_lei, None))
    filtro = f"codi_emp = {int(emp)}" if _eh_int(emp) else "codi_emp = <EMPRESA>"
    grupos_chave = {}
    for s_, t_, c_ in zip(ev[c_sep], ev[c_tip], ev[c_cod]):
        grupos_chave.setdefault((s_, t_), []).append(int(c_))
    conds = []
    for (s_, t_), cods in sorted(grupos_chave.items(), key=lambda x: (str(x[0][0]), x[0][1])):
        cods = sorted(set(cods))
        for i in range(0, len(cods), 400):
            conds.append(f"(i_separador = {s_} and tipo = {t_} and i_codigo in "
                         f"({', '.join(str(x) for x in cods[i:i + 400])}))")
    arquivos[f"conferir_antes_de_importar_emp{emp}.sql"] = _sql(
        ["-- Códigos de lançamento deste arquivo que JÁ existem na empresa. Deve voltar vazio antes de importar.",
         "-- Somente leitura.",
         f"select i_separador, tipo, i_codigo, descricao, conta_deb, conta_cred\n  from bethadba.{TAB_EVENTO.lower()}\n"
         f" where {filtro}\n   and (" + "\n     or ".join(conds or ["1 = 0"]) + ")\n order by i_separador, tipo, i_codigo;"])
    arquivos[f"conferir_contas_zero_emp{emp}.sql"] = _sql(
        ["-- Lançamentos da integração com conta 0 (deve voltar vazio). Somente leitura.",
         f"select i_separador, tipo, i_codigo, descricao, conta_deb, conta_cred\n  from bethadba.{TAB_EVENTO.lower()}\n"
         f" where {filtro} and (conta_deb = 0 or conta_cred = 0)\n order by i_separador, tipo, i_codigo;"])
    arquivos[f"corrigir_contas_zero_emp{emp}.sql"] = _sql(
        ["-- Troca a conta 0 (inexistente no plano) por branco (NULL) nos lançamentos da integração desta empresa.",
         "-- Use só para lançamentos importados antes com 0 na contrapartida. Interactive SQL com permissão de escrita.",
         "", f"update bethadba.{TAB_EVENTO.lower()} set conta_deb = null where {filtro} and conta_deb = 0;",
         f"update bethadba.{TAB_EVENTO.lower()} set conta_cred = null where {filtro} and conta_cred = 0;",
         "commit;"])
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for nome, conteudo in arquivos.items():
            z.writestr(nome, conteudo)
    return buf.getvalue(), resumo, passos

# =====================================================================
# 7. INTERFACE
# =====================================================================
st.set_page_config(page_title="Integrador Contábil da Folha", page_icon="📒", layout="wide")
st.title("📒 Integrador Contábil da Folha — Domínio")
st.caption("Gera o arquivo de integração contábil da folha para importar na Domínio e confere se o líquido fecha. "
           "Regras fixas e auditáveis, sem IA; as contas são sempre lidas do plano de contas da empresa.")


def desenhar_passos(ph, atual: int):
    nomes = ["Arquivos", "Configuração", "Conferência", "Importação"]
    with ph.container():
        cols_p = st.columns(4)
        for i, nome_p in enumerate(nomes, 1):
            if i < atual:
                cols_p[i - 1].markdown(f":green[✅ **{i}. {nome_p}**]")
            elif i == atual:
                cols_p[i - 1].markdown(f":blue[▶ **{i}. {nome_p}**]")
            else:
                cols_p[i - 1].markdown(f":gray[○ {i}. {nome_p}]")
        st.progress((atual - 1) / 3)


ph_passos = st.empty()
ph_resumo = st.empty()
desenhar_passos(ph_passos, 1)

MODO_PEND = "Só o que falta configurar (relatório de pendências)"
MODO_TODOS = "Todos os eventos cadastrados"

for k, v in (("w_hist", ""), ("w_baixa", True), ("w_socio", True), ("w_bloq", ""),
             ("w_compl", COMPLEMENTO_PADRAO), ("w_mesma_resc", False), ("w_mesma_fer", False),
             ("w_filant", False), ("w_exclusivas", True), ("w_leiaute", ""), ("w_cod", ""),
             ("w_forcar", False), ("w_multi", False), ("w_liquidos", LIQUIDOS_PADRAO), ("w_modo", MODO_PEND)):
    st.session_state.setdefault(k, v)

with st.sidebar:
    box_arq = st.container()
    box_cfg = st.container()

    # A configuração salva é lida primeiro (ela preenche os campos abaixo), mas aparece depois dos arquivos.
    with box_cfg:
        with st.expander("💾 Já configurou esta empresa? Carregar configuração salva"):
            f_cfg = st.file_uploader("Arquivo .json salvo anteriormente", type=["json"])
            if f_cfg is not None:
                aplicar_config(f_cfg.getvalue())

    with box_arq:
        st.header("Passo 1 · Arquivos")
        st.text_input("Código da empresa na Domínio", key="w_cod", placeholder="ex.: 56",
                      help="Empresa que vai receber a integração: vai no CODI_EMP dos TXT, nos scripts SQL e no nome "
                           "dos arquivos. Em branco, o app usa o código lido do relatório de pendências (ou do "
                           "cadastro de rubricas, no modo 'Todos os eventos').")
        ph_empresa = st.empty()
        modo_txt = st.radio("O que contabilizar?", [MODO_PEND, MODO_TODOS], key="w_modo",
                            help="Pendências: gera só o que o relatório 'Rubricas/Itens não configurados' lista. "
                                 "Todos os eventos: gera o lançamento de todas as rubricas do cadastro geral, "
                                 "mais os itens das abas Empresa, Provisão de Férias/13º e Outras Informações.")
        modo_cad = modo_txt == MODO_TODOS
        f_pend = None
        if not modo_cad:
            f_pend = st.file_uploader("1. Rubricas/Itens não configurados (relatório da Domínio)", type=["pdf", "txt"])
        f_cad = st.file_uploader("2. Cadastro geral de rubricas", type=["pdf", "txt"])
        f_plano = st.file_uploader("3. Plano de contas da empresa", type=["xlsx", "xls", "csv", "txt"])

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

    st.divider()
    st.header("Regras desta empresa")

    st.markdown("**1 · Como lançar a folha**")
    modo_multi = st.checkbox("Lançar o líquido à parte (partidas múltiplas)", key="w_multi",
                             help="Modelo da Solução 6583: no Tipo 1 cada provento vira lançamento só de débito, cada "
                                  "desconto só de crédito e o líquido é lançado pelos itens de Líquido (códigos "
                                  "negativos). Os lançamentos do Tipo 1 seguem a numeração 1, 2, 3… (proventos, "
                                  "descontos e líquidos em sequência; por último os de duas contas, como o FGTS). "
                                  "Na importação via tabela o lado sem conta vai como Valor fixo = NULO no leiaute: "
                                  "o .zip separa as partes. A Rescisão continua com duas contas no Tipo 4, pois a "
                                  "Domínio não tem o item 'Líquido Rescisão' na guia Folha.")
    st.caption("Desligado: cada rubrica leva débito e crédito (Salários a Pagar). "
               "Ligado: proventos só débito, descontos só crédito e o líquido sai nos itens de líquido.")
    liquidos_sel, inicio_t1 = [], 1
    if modo_multi:
        st.session_state["w_liquidos"] = [int(x) for x in st.session_state["w_liquidos"] if int(x) in LIQUIDOS]
        liquidos_sel = st.multiselect("Itens de líquido a lançar", list(LIQUIDOS), key="w_liquidos",
                                      format_func=lambda c: f"{c} · {LIQUIDOS[c][0]}")
        st.session_state["w_inicio_t1"] = max(1, int(st.session_state.get("w_inicio_t1") or 1))
        inicio_t1 = st.number_input("Primeiro código dos lançamentos do Tipo 1", min_value=1, step=1,
                                    key="w_inicio_t1",
                                    help="1 para empresa sem integração configurada. Se a empresa já tem lançamentos "
                                         "no Tipo 1 (modo pendências), comece depois do último código existente — "
                                         "o script conferir_antes_de_importar do .zip mostra os que coincidem.")

    st.markdown("**2 · Férias e Rescisão**")
    st.caption("Usar a mesma configuração da folha normal para:")
    mesma_resc = st.checkbox("Rescisão", key="w_mesma_resc",
                             help="Igual à Domínio. Marcado: a Rescisão usa os lançamentos da Folha "
                                  "(Tipo 1) e o Tipo 4 não é gerado. Desmarcado: Tipo 4 próprio, com a mesma "
                                  "classificação por natureza; o cálculo passa por Salários a Pagar e a 51 "
                                  "Líquido Rescisão leva o líquido para Rescisões a Pagar.")
    mesma_fer = st.checkbox("Férias", key="w_mesma_fer",
                            help="Igual à Domínio. Marcado: as Férias usam os lançamentos da Folha "
                                 "(Tipo 1) e o Tipo 3 não é gerado. Desmarcado: Tipo 3 próprio, com a mesma "
                                 "classificação por natureza; o cálculo passa por Salários a Pagar e a 937 "
                                 "Adiantamento de Férias leva o líquido para Adiantamento de Férias.")
    manter_excl = st.checkbox("Manter no Tipo 3/4 as rubricas exclusivas de Férias/Rescisão",
                              key="w_exclusivas", disabled=not (mesma_resc or mesma_fer),
                              help="Só vale quando 'mesma configuração' está marcada. Rubricas que existem "
                                   "somente em Férias/Rescisão (ex.: Aviso Prévio, 51 Líquido Rescisão) não "
                                   "têm lançamento na Folha; a Domínio as grava no próprio Tipo 3/4. "
                                   "Desmarcado: vão para o Tipo 1.")

    st.markdown("**3 · Encargos, provisões e sócios**")
    filant = st.checkbox("Entidade filantrópica (isenta de INSS patronal)", key="w_filant",
                         help="Gera também os itens de isenção da aba Empresa (25 a 36): estorno do INSS "
                              "patronal não devido, inverso do lançamento do item de INSS correspondente "
                              "(débito INSS a Recolher × crédito INSS despesa, ou provisão se marcada a baixa). "
                              "Desmarcado: esses itens não são integrados.")
    baixa_prov = st.checkbox("Baixar férias/13º pagos contra a provisão", key="w_baixa",
                             help="Mantenha marcado quando a empresa integra a Provisão de Férias/13º (Tipos 5 e 6), "
                                  "como a Domínio contabiliza na prática: férias, 13º e o FGTS/INSS sobre eles, pagos "
                                  "no cálculo, debitam a provisão. O item 'Valor Estorno Provisão' é o ajuste do saldo "
                                  "da provisão, não a baixa das férias pagas. Desmarque só se a empresa não provisiona.")
    socio_adm = st.checkbox("Pró-labore e encargos do sócio sempre em Despesas Administrativas", key="w_socio")

    with st.expander("Opções avançadas"):
        bloq_extra = st.text_input("Contas extras de colaborador (reduzidos, separados por vírgula)", key="w_bloq",
                                   help="Além das detectadas automaticamente; itens patronais nunca poderão usá-las.")
        historico = HISTORICO_FIXO
        hist_tipo = {}
        st.text_input("Código do Histórico (fixo)", value=HISTORICO_FIXO, disabled=True)
        complemento = st.text_input("Complemento", key="w_compl")
        st.caption(f"1 lançamento por rubrica/item · Código Sequencial = código da rubrica/item · "
                   f"Descrição = 'código - descrição' em caixa mista, máx. {LIM_DESC_EVENTO} caracteres.")
        leiaute_num = st.text_input("Número do leiaute da importação via tabela (opcional)", key="w_leiaute",
                                    placeholder="ex.: 18",
                                    help="Com o número, o .zip traz os scripts SQL que deixam CONTA_DEB/CONTA_CRED "
                                         "com Valor fixo = NULO em cada parte e restauram o leiaute no fim.")

if not (f_cad and f_plano and (f_pend or modo_cad)):
    st.info("👈 Comece pelo Passo 1: informe o código da empresa na Domínio e envie na barra lateral o cadastro "
            "geral de rubricas e o plano de contas"
            + ("." if modo_cad else " e o relatório de rubricas/itens não configurados — ou escolha "
               "'Todos os eventos cadastrados' em 'O que contabilizar?'.")
            + " Se já configurou esta empresa antes, carregue também a configuração salva.")
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
        itens = itens_do_cadastro(cad, tipos_cad, parse_separadores(seps_cad), filant)
    else:
        cab, itens = parse_pendencias(f_pend.getvalue(), f_pend.name)
except Exception as e:
    st.error(f"Erro na leitura: {e}")
    st.stop()

mesma_resc_cfg = mesma_resc
if modo_multi and mesma_resc:
    st.info("Partidas múltiplas: a Rescisão será gerada no Tipo 4 com duas contas (crédito em Rescisões a Pagar), "
            "porque não existe o item 'Líquido Rescisão' na guia Folha. A opção 'mesma configuração da folha normal "
            "para Rescisão' foi ignorada.")
    mesma_resc = False
desenhar_passos(ph_passos, 2)
itens, n_unif, n_excl = unificar_com_folha(itens, mesma_fer, mesma_resc, manter_excl)
if n_unif or n_excl:
    nomes_unif = " e ".join(n for n, f in (("Férias", mesma_fer), ("Rescisão", mesma_resc)) if f)
    msg = (f"🔗 {nomes_unif}: mesma configuração da folha normal — {n_unif} item(ns) usam o "
           "lançamento da Folha (Tipo 1); o Tipo 3/4 correspondente não será gerado para eles.")
    if n_excl:
        msg += (f" {n_excl} rubrica(s) exclusivas de Férias/Rescisão (inexistentes na Folha) "
                "permanecem no próprio Tipo 3/4.")
    st.info(msg)

if modo_cad and not cad and any(t in (1, 3, 4) for t in tipos_cad):
    st.error("Nenhuma rubrica lida no cadastro geral — não há eventos para os Tipos 1, 3 e 4.")
    st.stop()
if not itens:
    st.error("Nenhuma rubrica/item encontrado" + (" no cadastro." if modo_cad else " no relatório de pendências."))
    st.stop()
if not cad and not modo_cad:
    st.warning("Nenhuma rubrica lida no cadastro geral — o Tipo de todas será inferido pela descrição.")
if cad and sum(1 for r in cad.values() if not r["descricao"]):
    st.warning(f"{sum(1 for r in cad.values() if not r['descricao'])} rubrica(s) do cadastro sem descrição lida — "
               "confira o layout do relatório.")

fonte_emp = "cadastro" if modo_cad else "relatório"
cod_lido = txt_cel(cab["codigo"])
cod_digitado = txt_cel(st.session_state.get("w_cod"))
cod_empresa = cod_digitado or cod_lido
if cod_lido:
    ph_empresa.caption(f"Lido do {fonte_emp}: {cod_lido} - {cab['nome'][:40]}" + ("" if cod_digitado else " · em uso"))
else:
    ph_empresa.caption(f"O {fonte_emp} não traz o código da empresa — informe acima.")
if cod_digitado and cod_lido and cod_digitado != cod_lido:
    origem_cod = " (da configuração carregada)" if txt_cel(cfg.get("empresa")) == cod_digitado else ""
    st.warning(f"⚠️ Os arquivos vão para a empresa {cod_digitado}{origem_cod}, mas o {fonte_emp} é da empresa "
               f"{cod_lido} — confira o código no Passo 1.")
empresa_ok = cod_empresa.isdigit()
if not empresa_ok:
    st.error("❌ Informe o código da empresa na Domínio (somente números) no Passo 1, na barra lateral.")

# ---------- 1. Estrutura do plano ----------
st.subheader("Passo 2 · Plano de contas")
cols = list(raw.columns)
auto = auto_colunas(cols)
cfg_cols = cfg.get("colunas", {})
exp_plano = st.expander("⚙️ Colunas e raízes do plano",
                        expanded=not all(auto.get(c) for c in ("reduzido", "classificacao", "descricao")))
with exp_plano:
    cc = st.columns(len(CAMPOS_PLANO))
    sel = {}
    for i, (campo, rot) in enumerate(CAMPOS_PLANO.items()):
        opts = (["(nenhuma)"] if campo in CAMPOS_OPCIONAIS else []) + cols
        wk = f"c_{campo}"
        if st.session_state.get(wk) not in opts:
            pref = cfg_cols.get(campo) if cfg_cols.get(campo) in opts else auto.get(campo)
            st.session_state[wk] = pref if pref in opts else opts[0]
        sel[campo] = cc[i].selectbox(rot, opts, key=wk, index=None)

try:
    plano, pontuado = preparar_plano(raw, sel["reduzido"], sel["classificacao"], sel["descricao"],
                                     None if sel["tipo"] in (None, "(nenhuma)") else sel["tipo"],
                                     None if sel["ecf"] in (None, "(nenhuma)") else sel["ecf"])
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
m4.metric("Empresa", f'{cod_empresa} - {cab["nome"][:25]}' if cod_empresa == cod_lido else (cod_empresa or "—"))
if modo_cad:
    cont = pd.Series([TIPO_INTEGRACAO[i["secao"]] for i in itens]).value_counts().sort_index()
    st.caption("Eventos por Tipo da Integração: "
               + " · ".join(f"{t} - {TIPOS_LAYOUT[t]}: {n}" for t, n in cont.items()))

# ---------- 2. Separadores ----------
st.subheader("Passo 2 · Onde lançar a despesa da folha")
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
    st.info("A folha não está dividida por centro de custo, filial ou serviço. "
            "Escolha abaixo o grupo de despesa em que a folha inteira será lançada.")

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
    st.info("Escolha o grupo de despesa acima para continuar.")
    st.stop()
desenhar_passos(ph_passos, 3)

grupo_socio = None
if socio_adm:
    adm = next((g["classif"] for g in cands
                if tem(norm_conta(f"{g['caminho']} {g['descricao']}"), "ADMINISTRATIV*")), None)
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
        deb, cred, obs, rev, origem, nat = classificar_item_empresa(it["codigo"], d, baixa_prov, filant)
    elif sec == "Outras Informações":
        tipo = "Item (Outras Informações)"
        deb, cred, obs, rev, origem, nat = classificar_item_outras(it["codigo"], d)
    elif sec in SECOES_DE_ITENS:
        tipo, origem = f"Item ({sec})", "Seção"
        deb, cred, obs, rev = classificar_provisao(d, sec)
    else:
        tipo, origem = tipo_rubrica(it["codigo"], d, cad)
        fgts_forcado = False
        if e_fgts(d) and tipo not in ("Informativa", "Inf. dedutora"):
            fgts_forcado, tipo_cad = True, tipo
            tipo = "Informativa"
        if tipo in ("Informativa", "Inf. dedutora"):
            deb, cred, obs, rev = classificar_informativa(d, tipo, baixa_prov)
            if fgts_forcado:
                obs = f"Rubrica de FGTS cadastrada como {tipo_cad} — tratada como informativa | {obs}"
                rev = True
        elif tipo == "Desconto":
            deb, cred, obs, rev = classificar_desconto(d, baixa_prov)
        else:
            deb, cred, obs, rev = classificar_provento(d, baixa_prov)
    if grupo_socio and (e_socio(d) or nat == "INSS_SOCIO" or (nat == "ISENCAO" and tem(d, "PRO LAB*"))):
        prefixo = grupo_socio
    regras.append(dict(it=it, d=d, sec=sec, k=k, prefixo=prefixo, tipo=tipo, origem=origem,
                       deb=deb, cred=cred, obs=obs, rev=rev, nat=nat))

n_isencao_ign = sum(1 for r in regras if r["nat"] == "ISENCAO") if not filant else 0
if any(TIPO_INTEGRACAO.get(i["secao"]) in (5, 6) for i in itens) and not baixa_prov:
    st.warning("⚠️ A empresa integra a Provisão de Férias/13º (Tipos 5 e 6), mas 'Baixar férias/13º pagos contra a "
               "provisão' está desmarcado: férias e 13º pagos irão para despesa e o custo será contado duas vezes "
               "(provisão mensal + pagamento).")
if n_isencao_ign:
    st.info(f"ℹ️ {n_isencao_ign} item(ns) de isenção de filantropia não serão integrados. "
            "Se a empresa é filantrópica, marque 'Entidade filantrópica' na barra lateral.")

# ---------- 3. Mapa de contas ----------
st.subheader("Passo 2 · Contas do plano usadas nos lançamentos")
base_res = Resolvedor(plano, pontuado, raizes, {})
pares = {Resolvedor.chave(a, ""): (a, "") for a in COLAB_ALVOS}
for r in regras:
    for a in r["deb"] + r["cred"]:
        pares.setdefault(Resolvedor.chave(a, r["prefixo"]), (a, r["prefixo"]))
cfg_contas = st.session_state.setdefault("cfg_contas", {})
linhas_map = []
for chave, (a, g) in sorted(pares.items()):
    sug = base_res.sugestao(a, g)
    inf = base_res.info.get(chave, {})
    linhas_map.append({
        "Chave": chave, "Alvo": ALVOS[a][1],
        "Escopo": f"grupo {g}" if ALVOS[a][0] == G else ESCOPO_ROTULO[ALVOS[a][0]],
        "Sugerida": sug[0] if sug else "", "Descrição sugerida": sug[1] if sug else "❌ não encontrada",
        "Como foi achada": ("⚠️ " if inf.get("como") in COMO_FRACO else "") + inf.get("como", ""),
        "Empatou com": ", ".join(f"{x} {d}" for x, d in inf.get("empate", [])),
        "Conta definida": cfg_contas.get(chave, ""),
    })
df_map = pd.DataFrame(linhas_map)
n_nf = int(((df_map["Sugerida"] == "") & (df_map["Conta definida"] == "")).sum())
n_conf = int((((df_map["Como foi achada"].str.startswith("⚠️")) | (df_map["Empatou com"] != ""))
              & (df_map["Conta definida"] == "")).sum())
with st.expander(f"Contas encontradas no plano — {n_nf} sem conta definida, {n_conf} para confirmar",
                 expanded=bool(n_nf or n_conf)):
    st.caption("O app procura cada conta no plano da empresa pelo nome (com as abreviações usuais), pelos grupos "
               "acima dela, pelo lado do balanço e pelo referencial ECF — sem depender do plano padrão. "
               "'Como foi achada' com ⚠️ ou 'Empatou com' preenchido: confira. Onde aparecer '❌ não encontrada', "
               "informe o código reduzido na coluna 'Conta definida' (também serve para trocar uma sugestão). "
               "Os valores ficam na configuração salva da empresa.")
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
    multi_row = modo_multi and sec in SECOES_TIPO1 and r["tipo"] in ("Provento", "Desconto")
    if multi_row:
        falta_conta = (not dc) if r["tipo"] == "Provento" else (not cc_)
    else:
        falta_conta = not dc or not cc_
    if not r["deb"] and not r["cred"]:
        status = ST_PEND if (sec in SECOES_CATALOGO and r["nat"] != "ISENCAO") else ST_NAO
    elif tipo_int is None:
        status = ST_MAN
        alertas.append("seção sem Tipo da Integração no layout — configurar na Domínio")
    elif falta_conta:
        status = ST_PEND
    elif r["rev"] or alertas or r["origem"].startswith("Inferido"):
        status = ST_REV
    else:
        status = ST_OK
    exportar_ok = bool(not falta_conta and tipo_int)
    conta_liq = ""
    if multi_row:
        if r["tipo"] == "Provento":
            cc_, cd = NULO, ""
        else:
            dc, dd = NULO, ""
    elif r["tipo"] in ("Provento", "Desconto"):
        conta_liq = res.resolver(passivo_secao(r["d"]), r["prefixo"])[0]
    linhas.append({
        "Status": status, "Seção": sec, "Tipo Integração": str(tipo_int) if tipo_int else "—",
        "Separador": r["k"], "Nome do separador": it["sep_nome"] or NOME_PADRAO_SEM_SEPARADOR,
        "Grupo": r["prefixo"], "Código": it["codigo"], "Descrição": it["descricao"],
        "Descrição Lançamento": descricao_lancamento(it["codigo"], it["descricao"]),
        "Tipo": r["tipo"], "Origem do tipo": r["origem"], "Débito": dc, "Desc. Débito": dd,
        "Crédito": cc_, "Desc. Crédito": cd, "Conta Líquido": conta_liq,
        "Observação": " | ".join([r["obs"]] + alertas).strip(" |"),
        "Exportar": exportar_ok,
    })
df = pd.DataFrame(linhas)

n_div = int(df["Origem do tipo"].str.startswith("Inferido").sum())
ne, nc = norm(cab["nome"]), norm(nome_cad)
if not modo_cad and ne and nc and ne not in nc and nc not in ne:
    if n_div:
        st.warning(f"⚠️ O cadastro enviado é de outro modelo (**{nome_cad}**): {n_div} rubrica(s) não conferem "
                   "em código/descrição e tiveram o tipo (provento/desconto) deduzido pelo nome — revise-as.")
    else:
        st.info(f"Cadastro do modelo **{nome_cad}**, mas todas as rubricas conferem por código + descrição.")

# ---------- 4. Conferência ----------
st.subheader("Passo 3 · Conferência das rubricas")
n_ok_ = int((df["Status"] == ST_OK).sum())
n_rev_ = int((df["Status"] == ST_REV).sum())
n_pend_ = int((df["Status"] == ST_PEND).sum())
mcols = st.columns(5)
for col, s in zip(mcols, (ST_OK, ST_REV, ST_PEND, ST_NAO, ST_MAN)):
    col.metric(s, int((df["Status"] == s).sum()))

EDITAVEIS = ("Débito", "Crédito", "Descrição Lançamento", "Exportar")
with st.expander("Ver e ajustar rubrica por rubrica", expanded=bool(n_pend_ or n_rev_)):
    st.caption("✅ OK = pronto · ⚠️ Revisar = confira a sugestão · ❌ Pendente = falta conta. "
               "Edições aqui valem só para este lote; para correções permanentes de contas use o Passo 2.")
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
        elif e_nulo(c):
            pass
        elif c not in idx_plano.index:
            msgs.append(f"{lado} {c} inexistente")
        elif idx_plano.loc[c, "tipo"] != "A":
            msgs.append(f"{lado} {c} é sintética")
        elif row["Seção"] in SECOES_DE_ITENS and c in colab:
            msgs.append(f"{lado} {c} é conta de colaborador — proibido em item patronal")
    if e_nulo(vals["Débito"]) and e_nulo(vals["Crédito"]):
        msgs.append(f"as duas contas estão '{NULO}'")
    elif (e_nulo(vals["Débito"]) or e_nulo(vals["Crédito"])) and not modo_multi:
        msgs.append(f"lado '{NULO}' só vale em partidas múltiplas")
    elif vals["Débito"] and vals["Débito"] == vals["Crédito"]:
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
    st.error(f"{len(bloqueadas)} linha(s) marcada(s) para exportar têm erro e ficarão fora do arquivo.")
    with st.expander("Ver linhas com erro", expanded=True):
        st.dataframe(bloqueadas[["Seção", "Separador", "Código", "Descrição", "Validação"]], hide_index=True)

manuais = ed[ed["Status"] == ST_MAN]
if not manuais.empty:
    st.warning(f"📝 {len(manuais)} item(ns) não têm tipo de integração no layout — configure-os manualmente na Domínio.")
    with st.expander("Ver itens para configurar manualmente", expanded=True):
        st.dataframe(manuais[["Seção", "Separador", "Código", "Descrição", "Débito", "Desc. Débito",
                              "Crédito", "Desc. Crédito"]], hide_index=True)

# ---------- 4b. Verificação do líquido ----------
# Regra: todo provento credita a conta de líquido da seção e todo desconto a debita; assim
# Σ proventos − Σ descontos = saldo da conta de líquido (o líquido da folha). Nenhuma rubrica com
# valor (provento/desconto) pode ficar fora do arquivo.
st.subheader("Passo 3 · O líquido da folha fecha?")
e_vd = ed["Tipo"].isin(["Provento", "Desconto"])
no_arquivo = ed["Exportar"] & (ed["Validação"] == "")
fora = ed[e_vd & ~no_arquivo]


def _fura_liquido(row):
    liq = txt_cel(row["Conta Líquido"])
    if not liq or row["Tipo"] not in ("Provento", "Desconto"):
        return False
    lado = row["Crédito"] if row["Tipo"] == "Provento" else row["Débito"]
    return txt_cel(lado) != liq


furam = ed[e_vd & no_arquivo & ed.apply(_fura_liquido, axis=1)]
l1, l2, l3 = st.columns(3)
l1.metric("Proventos/descontos no arquivo", int((e_vd & no_arquivo).sum()))
l2.metric("Sem conta (fora do arquivo)", len(fora))
l3.metric("Fora da conta de líquido", len(furam))
if not fora.empty:
    st.error(f"❌ {len(fora)} rubrica(s) ainda não têm conta definida e ficarão fora do arquivo — sem elas o líquido "
             "da folha não fecha. Informe as contas em 'Passo 2 · Contas do plano usadas nos lançamentos' "
             "(coluna 'Conta definida').")
    with st.expander("Ver rubricas sem conta", expanded=True):
        st.dataframe(fora[["Status", "Seção", "Separador", "Código", "Descrição", "Tipo", "Débito", "Crédito",
                           "Observação"]], hide_index=True)
if not furam.empty:
    st.warning(f"⚠️ {len(furam)} rubrica(s) não usam a conta de líquido da seção (coluna 'Conta Líquido'). "
               "Confirme se é intencional.")
    with st.expander("Ver rubricas fora da conta de líquido", expanded=True):
        st.dataframe(furam[["Seção", "Separador", "Código", "Descrição", "Tipo", "Débito", "Crédito",
                            "Conta Líquido"]], hide_index=True)
if fora.empty and furam.empty:
    if modo_multi:
        st.success("✅ Partidas múltiplas: proventos só com débito, descontos só com crédito e o líquido lançado "
                   "pelos itens de líquido; nenhuma rubrica com valor ficou de fora.")
    else:
        st.success("✅ O líquido fecha: todo provento é creditado e todo desconto é debitado na conta de líquido, "
                   "e nenhuma rubrica com valor ficou de fora.")
forcar = st.checkbox("Permitir exportar mesmo assim (o líquido não fechará)", key="w_forcar",
                     disabled=fora.empty)

# ---------- 5. Montagem do layout (1 lançamento por rubrica/item) ----------
st.subheader("Passo 4 · Arquivo para importar na Domínio")
emp = para_int(cod_empresa)

exp = exportar.copy()
exp["_sep"] = [para_int(v) for v in exp["Separador"]]
exp["_s"] = [str(v) for v in exp["_sep"]]
exp["_tipo"] = [int(TIPO_INTEGRACAO[s]) for s in exp["Seção"]]
exp["_cod"] = [int(c) for c in exp["Código"]]          # código da rubrica/item
exp["_lanc"] = exp["_cod"]                              # código do lançamento (sequencial no Tipo 1 em partidas múltiplas)
exp["_deb"] = [NULO if e_nulo(v) else txt_cel(v) for v in exp["Débito"]]
exp["_cred"] = [NULO if e_nulo(v) else txt_cel(v) for v in exp["Crédito"]]
exp["_hist"] = [txt_cel(hist_tipo.get(t)) or txt_cel(historico) for t in exp["_tipo"]]
exp["_desc"] = [txt_cel(v) for v in exp["Descrição Lançamento"]]

# A chave do lançamento (empresa + separador + código + tipo) é ÚNICA em FOINTEGCONT, e uma rubrica pertence
# a um só lançamento por empresa + separador + tipo. Linhas 100% idênticas (mesma rubrica em Folha Normal e
# Adiantamento com as mesmas contas) são unificadas; se as contas diferem, só a primeira linha segue.
chave_ev = ["_s", "_tipo", "_cod", "_deb", "_cred", "_hist", "_desc"]
exp = exp.drop_duplicates(chave_ev)
chave_pk = ["_s", "_tipo", "_cod"]
conflitos = exp[exp.duplicated(chave_pk, keep=False)]
exp = (exp.drop_duplicates(chave_pk, keep="first")
       .sort_values(chave_pk, kind="stable").reset_index(drop=True))
if not conflitos.empty:
    st.error(f"❌ {len(conflitos)} rubrica(s) disputam o mesmo código de lançamento (mesmo separador e tipo) com contas "
             "diferentes. A Domínio aceita um único lançamento por código: foi mantida a PRIMEIRA de cada grupo. "
             "Para escolher outra, desmarque 'Exportar' nas que não devem seguir.")
    with st.expander("Ver rubricas em conflito", expanded=True):
        st.dataframe(conflitos[["Seção", "Separador", "Tipo Integração", "Código", "Descrição Lançamento",
                                "Débito", "Crédito"]], hide_index=True)
n_lanc_liq, liq_sem_conta = 0, []
if modo_multi:
    liq_rows = []
    for k_sep in mapa_grupos:
        for cod_l in liquidos_sel:
            nome_l, alvos_l = LIQUIDOS[cod_l]
            conta_l = res.resolver(alvos_l, "")[0]
            if not conta_l:
                liq_sem_conta.append(f"{nome_l} (separador {k_sep})")
                continue
            sep_l = para_int(k_sep)
            liq_rows.append({"_sep": sep_l, "_s": str(sep_l), "_tipo": 1, "_cod": cod_l, "_lanc": cod_l,
                             "_deb": NULO, "_cred": conta_l, "_hist": txt_cel(historico),
                             "_desc": cortar(nome_l), "Descrição": nome_l})
    n_lanc_liq = len(liq_rows)
    if liq_rows:
        exp = pd.concat([exp, pd.DataFrame(liq_rows)], ignore_index=True)
    t1 = exp[exp["_tipo"] == 1].copy()
    if not t1.empty:
        # lançamento composto contíguo: só débito (proventos), só crédito (descontos), líquidos; depois duas contas
        t1["_grp"] = [2 if c < 0 else 0 if e_nulo(cr) else 1 if e_nulo(de) else 3
                      for c, de, cr in zip(t1["_cod"], t1["_deb"], t1["_cred"])]
        t1["_ord"] = [abs(c) for c in t1["_cod"]]
        t1 = t1.sort_values(["_s", "_grp", "_ord"], kind="stable")
        t1["_lanc"] = t1.groupby("_s").cumcount() + int(inicio_t1)
        exp.loc[t1.index, "_lanc"] = t1["_lanc"]
    exp = exp.sort_values(["_s", "_tipo", "_lanc"], kind="stable").reset_index(drop=True)
    if not liquidos_sel:
        st.error("❌ Partidas múltiplas: selecione ao menos um item de líquido na barra lateral.")
    if liq_sem_conta:
        st.error("❌ Sem conta para o(s) item(ns) de líquido: " + "; ".join(liq_sem_conta)
                 + ". Informe 'Salários a Pagar' (e Pró-labore/Férias a Pagar, se usar) no Passo 2.")
n = len(exp)

evento = pd.DataFrame({
    COLS_EVENTO[0]: [emp] * n,
    COLS_EVENTO[1]: exp["_sep"].tolist(),
    COLS_EVENTO[2]: exp["_lanc"].tolist(),          # Código Sequencial do lançamento (= rubrica; sequencial 1..N no Tipo 1 em partidas múltiplas)
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
    COLS_INTEGRA[2]: exp["_lanc"].tolist(),         # aponta para o lançamento da rubrica/item
    COLS_INTEGRA[3]: exp["_tipo"].tolist(),
    COLS_INTEGRA[4]: exp["_cod"].tolist(),          # código da rubrica/item (negativo = item de líquido)
}, columns=COLS_INTEGRA).drop_duplicates().reset_index(drop=True)

cortadas = sum(1 for c, d in zip(exp["_cod"], exp["Descrição"])
               if c > 0 and len(f"{c} - {caixa_mista(d)}") > LIM_DESC_EVENTO)
if cortadas:
    st.info(f"{cortadas} descrição(ões) cortada(s) em {LIM_DESC_EVENTO} caracteres — "
            "ajuste na coluna 'Descrição Lançamento' se quiser abreviar.")

# sem o código da empresa o bloqueio já aparece uma vez ("falta o código"): a conformidade confere o resto
emp_conf = {COLS_EVENTO[0]: emp if empresa_ok else 0}
checagens = (checar_layout(evento.assign(**emp_conf), integra.assign(**{COLS_INTEGRA[0]: emp_conf[COLS_EVENTO[0]]}),
                           multi=modo_multi) if n else [])
erros_layout = [c for c in checagens if c["Nível"] == "Erro" and c["Resultado"] == "❌"]
avisos_layout = [c for c in checagens if c["Nível"] == "Aviso" and c["Resultado"] == "⚠️"]
with st.expander(f"Conformidade com o layout da Domínio (FOINTEGCONT / FOINTEGCONTEVE) — "
                 f"{len(erros_layout)} erro(s), {len(avisos_layout)} aviso(s)",
                 expanded=bool(erros_layout or avisos_layout)):
    if checagens:
        st.dataframe(pd.DataFrame(checagens), hide_index=True, use_container_width=True)
    else:
        st.info("Nenhum lançamento para conferir.")
    st.caption("As verificações conferem ordem das colunas, tipos, limites de tamanho, chave única e vínculos do arquivo "
               "contra as tabelas de destino (FOINTEGCONT e FOINTEGCONTEVE).")

zip_imp, resumo_partes, passos_imp = (pacote_importacao(evento, integra, emp, leiaute_num) if n
                                      else (b"", {}, []))
sem_cp = fora_do_cp1252(evento[COLS_EVENTO[4]].tolist() + [complemento]) if n else []
if sem_cp:
    st.warning("⚠️ Texto com caractere fora do Windows-1252 (no TXT vira '?'): " + "; ".join(sem_cp[:5]))

e1, e2, e3 = st.columns(3)
e1.metric("Lançamentos (evento)", len(evento))
e2.metric("Vínculos (integra)", len(integra))
e3.metric("Configurar manualmente", len(manuais))
with st.expander("Como importar na Domínio (importação via tabela)", expanded=len(resumo_partes) > 1):
    if len(resumo_partes) > 1:
        st.info("A importação via tabela grava **0** em campo numérico vazio lido do arquivo — por isso a "
                "contrapartida vinha 0. O lado sem conta só fica **em branco** com o campo em **Valor fixo = NULO** "
                "no leiaute (campo com valor fixo não ocupa coluna no TXT). O .zip separa os lançamentos em partes: "
                "a parte 2 vai sem a coluna CONTA_CRED e a parte 3 sem a coluna CONTA_DEB.")
        st.dataframe(pd.DataFrame([{"Parte": p, "Pasta": PARTES[p][0], "O que vai": PARTES[p][1],
                                    "Campo com Valor fixo = NULO": PARTES[p][2] or "— (leiaute normal)",
                                    "Lançamentos": v[0], "Vínculos": v[1]} for p, v in resumo_partes.items()]),
                     hide_index=True, use_container_width=True)
    for i, t_ in enumerate(passos_imp, 1):
        st.markdown(f"{i}. {t_}")
    st.caption("O .zip traz o LEIA-ME com este passo a passo, a posição de cada coluna (estrutura_arquivos.csv), "
               "a consulta que mostra se algum código de lançamento já existe na empresa (antes de importar), a que "
               "confere se ficou conta 0 (depois) e o script que corrige lançamentos já importados com 0"
               + ("; com o número do leiaute (Opções avançadas), também os scripts do leiaute de cada parte."
                  if len(resumo_partes) > 1 else "."))
with st.expander("Pré-visualizar o arquivo (evento e integra)"):
    t1, t2 = st.tabs([ABA_EVENTO, ABA_INTEGRA])
    t1.dataframe(evento.astype(str), hide_index=True, use_container_width=True)
    t2.dataframe(integra.astype(str), hide_index=True, use_container_width=True)


def excel_bytes(abas: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for nome, d in abas.items():
            d.to_excel(w, sheet_name=nome[:31], index=False)
    return buf.getvalue()


cfg_out = {
    "versao": 10, "empresa": cod_empresa, "nome_empresa": cab["nome"],
    "grupos": mapa_grupos, "grupo_socio": grupo_socio, "contas": dict(cfg_contas),
    "historico": historico, "hist_tipo": {str(t): v for t, v in hist_tipo.items() if v},
    "complemento": complemento, "baixa_prov": baixa_prov, "socio_adm": socio_adm,
    "bloqueio_extra": bloq_extra, "raizes": raizes,
    "mesma_rescisao": mesma_resc_cfg, "mesma_ferias": mesma_fer, "manter_exclusivas": manter_excl,
    "partidas_multiplas": modo_multi, "liquidos": liquidos_sel, "filantropica": filant, "leiaute": leiaute_num,
    "inicio_tipo1": int(inicio_t1),
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
if checagens:
    conf_abas["conformidade_layout"] = pd.DataFrame(checagens)
if n:
    conf_abas["evento"] = evento.assign(**{"Parte da importação": [
        parte_do_lancamento(d_, c_) for d_, c_ in zip(evento[COLS_EVENTO[5]], evento[COLS_EVENTO[6]])]})
    conf_abas["integra"] = integra

bloqueio_import = (bool(erros_layout) or (not fora.empty and not forcar) or n == 0 or not empresa_ok
                   or bool(liq_sem_conta) or (modo_multi and not liquidos_sel))
motivos = []
if not empresa_ok:
    motivos.append("falta o código da empresa na Domínio (Passo 1)")
if n == 0:
    motivos.append("nenhum lançamento foi gerado")
if erros_layout:
    motivos.append(f"{len(erros_layout)} erro(s) de conformidade com o layout da Domínio")
if not fora.empty and not forcar:
    motivos.append(f"{len(fora)} rubrica(s) sem conta definida (o líquido não fecha)")
if liq_sem_conta:
    motivos.append(f"{len(liq_sem_conta)} item(ns) de líquido sem conta")
if modo_multi and not liquidos_sel:
    motivos.append("nenhum item de líquido selecionado")

desenhar_passos(ph_passos, 3 if bloqueio_import else 4)
with ph_resumo.container():
    if motivos:
        st.error("⛔ Ainda não dá para importar: " + "; ".join(motivos) + ". Veja os Passos 2 e 3 abaixo.")
    elif n_rev_ or not furam.empty or avisos_layout or not conflitos.empty or not fora.empty:
        st.warning(f"⚠️ Arquivo gerado, mas confira antes de importar: {n_rev_} rubrica(s) marcada(s) como 'Revisar'"
                   + ("; exportação liberada com o líquido não fechando" if not fora.empty else "") + ".")
    else:
        st.success("✅ Tudo certo: o arquivo está pronto para importar na Domínio.")
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Lançamentos prontos", len(evento))
    k2.metric("Rubricas OK", n_ok_)
    k3.metric("Para revisar", n_rev_)
    k4.metric("Sem conta definida", n_pend_)
    k5.metric("Líquido da folha",
              ("Pelos itens de líquido" if (modo_multi and fora.empty and not liq_sem_conta and liquidos_sel)
               else "Fecha") if (fora.empty and furam.empty and not liq_sem_conta) else "Não fecha" if not fora.empty or liq_sem_conta else "Confirmar")

if erros_layout:
    st.error("❌ O arquivo não está conforme o layout da Domínio (veja 'Conformidade com o layout' acima) — "
             "o download está bloqueado.")
elif not fora.empty and not forcar:
    st.error("❌ Download bloqueado: há rubricas sem conta definida, e o líquido não fecharia. Defina as contas no "
             "Passo 2 ou marque 'Permitir exportar mesmo assim' no Passo 3.")

sufixo = "_completo" if modo_cad else ""
st.download_button(f"⬇️ Baixar arquivos da importação via tabela (.zip com TXT — {len(evento)} lançamentos"
                   + (f", {len(resumo_partes)} partes)" if len(resumo_partes) > 1 else ")"),
                   zip_imp, file_name=f"integracao_folha_emp{cod_empresa}{sufixo}.zip", mime="application/zip",
                   disabled=bloqueio_import, type="primary")
c1_, c2, c3 = st.columns(3)
c1_.download_button("📄 Mesmo conteúdo em planilha (integra/evento)",
                    excel_bytes({ABA_INTEGRA: integra, ABA_EVENTO: evento}),
                    file_name=f"integracao_folha_emp{cod_empresa}{sufixo}.xlsx", disabled=bloqueio_import,
                    help="Só para consulta: em partidas múltiplas o lado sem conta aparece como NULO, e a importação "
                         "via tabela grava 0 nele. Para importar, use os TXT do .zip.")
c2.download_button("📋 Planilha de conferência completa", excel_bytes(conf_abas),
                   file_name=f"conferencia_folha_emp{cod_empresa}{sufixo}.xlsx")
c3.download_button("💾 Salvar configuração da empresa (.json)",
                   json.dumps(cfg_out, ensure_ascii=False, indent=2).encode("utf-8"),
                   file_name=f"config_folha_emp{cod_empresa}.json", mime="application/json")
