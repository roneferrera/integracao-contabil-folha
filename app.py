# ============================================================
# app.py  –  Integração Contábil Domínio V6.3
# Correção definitiva: tabelas de score recalibradas
# contra os dois planos de contas reais
# ============================================================

import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

VERSAO = "V6.3"

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
# NORMALIZAÇÃO
# ══════════════════════════════════════════════════════════════════════════
def _norm(texto: str) -> str:
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


# ══════════════════════════════════════════════════════════════════════════
# ANTI CURTO-CIRCUITO
# ══════════════════════════════════════════════════════════════════════════
def validar_par_contas(conta_debito: str, conta_credito: str) -> bool:
    if not conta_debito or not conta_credito:
        return True
    return str(conta_debito).strip() != str(conta_credito).strip()


# ══════════════════════════════════════════════════════════════════════════
# SCORE GERAL DE FOLHA
# ══════════════════════════════════════════════════════════════════════════
PALAVRAS_FOLHA_POSITIVO: list[tuple[str, int]] = [
    ("SALARIOS E ORDENADOS A PAGAR", 10), ("SALARIOS A PAGAR", 10),
    ("PRO-LABORE A PAGAR", 10), ("GRATIFICACOES A PAGAR", 10),
    ("FERIAS A PAGAR", 10), ("RESCISOES A PAGAR", 10),
    ("13 SALARIO A PAGAR", 10), ("PENSAO ALIMENTICIA A PAGAR", 10),
    ("COMISSOES A PAGAR", 10), ("AUTONOMOS A PAGAR", 10),
    ("INDENIZACOES A PAGAR", 10), ("PREMIOS E BONIFICACOES", 10),
    ("INSS A RECOLHER", 10), ("INSS SOBRE PROVISOES", 10),
    ("FGTS A RECOLHER", 10), ("FGTS SOBRE PROVISOES", 10),
    ("PIS S/ FOLHA A RECOLHER", 10), ("IRRF S/ FOLHA", 10),
    ("CONTRIBUICOES SINDICAIS", 10), ("PROVISOES PARA FERIAS", 10),
    ("PROVISOES PARA 13", 10), ("INSS SOBRE PROVISOES PARA FERIAS", 10),
    ("INSS SOBRE PROVISOES PARA 13", 10), ("FGTS SOBRE PROVISOES PARA FERIAS", 10),
    ("FGTS SOBRE PROVISOES PARA 13", 10), ("OBRIGACOES COM O PESSOAL", 10),
    ("OBRIGACOES SOCIAIS", 10), ("OBRIGACOES TRABALHISTA", 10),
    ("OBRIGACOES TRABALHISTAS E PREVIDENCIARIA", 10), ("PROVISOES", 8),
    ("DESPESAS COM PESSOAL", 10), ("SALARIOS E ORDENADOS", 8),
    ("PRO-LABORE", 8), ("PREMIOS E GRATIFICACOES", 8),
    ("INDENIZACOES E AVISO PREVIO", 8), ("ASSISTENCIA MEDICA E SOCIAL", 8),
    ("VALE TRANSPORTE", 7), ("VALE REFEICAO", 7),
    ("ALIMENTACAO/ CESTA BASICA", 7), ("DESPESAS COM ALIMENTACAO", 7),
    ("PENSAO ALIMENTICIA", 7), ("COMISSOES SOBRE VENDAS", 7),
    ("COMISSOES", 7), ("HORAS EXTRAS", 7), ("PIS S/ FOLHA", 7),
    ("MAO-DE-OBRA DIRETA", 10), ("MAO-DE-OBRA INDIRETA", 10),
    ("SALARIOS E ORDENADOS CUSTOS", 10), ("PRO-LABORE CUSTOS", 10),
    ("FERIAS CUSTOS", 10), ("INSS CUSTOS", 10), ("FGTS CUSTOS", 10),
    ("INSS EMPRESA", 10), ("INSS TERCEIROS", 10), ("INSS ACIDENTE", 10),
    ("INSS PATRONAL", 10), ("ENCARGOS SOCIAIS", 10),
    ("CONTRIBUICAO PREVIDENCIARIA", 10), ("CONTRIBUICAO PATRONAL", 10),
    ("IMPOSTO DE RENDA A RECOLHER", 6), ("IMPOSTO DE RENDA RETIDO", 6),
    ("IRRF", 6), ("INSS", 5), ("FGTS", 5),
    ("FOLHA DE PAGAMENTO", 8), ("REMUNERACAO", 6),
    ("PESSOAL", 5), ("TRABALHISTA", 5), ("PREVIDENCIARIA", 5),
    ("PATRONAL", 8), ("EMPREGADOS", 5), ("FUNCIONARIOS", 5),
    ("EMPRESTIMO / CONSIGNADO", 8), ("EMPRESTIMOS CONSIGNADOS", 8),
]

PALAVRAS_FOLHA_NEGATIVO: list[tuple[str, int]] = [
    ("FORNECEDORES NACIONAIS", -50), ("FORNECEDORES ESTRANGEIROS", -50),
    ("FORNECEDORES DO GRUPO", -50), ("FORNECEDORES", -40),
    ("CLIENTES NACIONAIS", -50), ("CLIENTES ESTRANGEIROS", -50),
    ("CLIENTES RELACIONADOS", -50), ("CLIENTES", -40),
    ("MERCADORIAS PARA REVENDA", -50), ("MATERIA-PRIMA", -30),
    ("ESTOQUE", -40), ("ALMOXARIFADO", -30),
    ("PRODUTOS ACABADOS", -40), ("PRODUTOS SEMI ACABADOS", -40),
    ("IMOVEIS", -40), ("MAQUINAS E EQUIPAMENTOS", -40),
    ("VEICULOS", -40), ("MOVEIS E UTENSILIOS", -40),
    ("COMPUTADORES E ACESSORIOS", -40), ("INSTALACOES", -40),
    ("FERRAMENTAS E ACESSORIOS", -40), ("SOFTWARES", -40),
    ("MARCAS E PATENTES", -40), ("DEPRECIACAO DE EDIFICIOS", -30),
    ("DEPRECIACAO DE MOVEIS", -30), ("DEPRECIACAO DE MAQUINAS", -30),
    ("DEPRECIACAO DE VEICULOS", -30), ("BANCO DO BRASIL", -30),
    ("BANCO ITAU", -30), ("BANCO BRADESCO", -30),
    ("BANCO SANTANDER", -30), ("BANCO INTER", -30),
    ("BANCO C6", -30), ("BANCO NU", -30), ("BANCO CORA", -30),
    ("BANCO DAYCOVAL", -30), ("CAIXA ECONOMICA", -30),
    ("CAIXA GERAL", -30), ("FUNDO FIXO DE CAIXA", -30),
    ("APLICACOES FINANCEIRAS", -30), ("APLICACOES BANCO", -30),
    ("CHEQUE ESPECIAL", -30), ("EMPRESTIMOS BANCOS", -30),
    ("FINANCIAMENTO BANCO", -30), ("IPI A RECOLHER", -30),
    ("ICMS A RECOLHER", -30), ("ISS A RECOLHER", -30),
    ("PIS A RECOLHER", -30), ("COFINS A RECOLHER", -30),
    ("SIMPLES NACIONAL A RECOLHER", -30), ("VENDA DE PRODUTOS", -50),
    ("VENDA DE MERCADORIAS", -50), ("SERVICOS PRESTADOS", -40),
    ("RECEITA", -40), ("CAPITAL SOCIAL", -50), ("RESERVAS", -40),
    ("LUCROS OU PREJUIZOS", -40), ("DIVIDENDOS", -40),
    ("RESULTADO DO EXERCICIO", -40), ("APURACAO DO RESULTADO", -40),
    ("ADIANTAMENTO A SOCIOS", -30), ("ADIANTAMENTO A FORNECEDORES", -20),
    ("TITULOS A RECEBER", -30), ("DEPOSITOS JUDICIAIS", -30),
    ("INVESTIMENTOS", -30), ("PARTICIPACOES SOCIETARIAS", -30),
    ("IPI A RECUPERAR", -30), ("ICMS A RECUPERAR", -30),
    ("PIS A RECUPERAR", -30), ("COFINS A RECUPERAR", -30),
    # CRÍTICO: contas de Ativo não devem ser débito de folha
    ("A COMPENSAR", -50),   # bloqueia INSS A COMPENSAR, IRRF RETIDO A COMPENSAR etc.
    ("A RECUPERAR", -50),   # bloqueia tributos a recuperar
    ("SALDO NEGATIVO", -50),
    ("ADIANTAMENTO DE SALARIO", -30),  # é Ativo, não DRE
    ("ADIANTAMENTO DE 13", -30),
    ("ADIANTAMENTO DE FERIAS", -30),
]

SCORE_MINIMO_FOLHA = 5


def calcular_score_folha(nome_conta_norm: str) -> int:
    score = 0
    for termo, peso in PALAVRAS_FOLHA_POSITIVO:
        if _norm(termo) in nome_conta_norm:
            score += peso
    for termo, peso in PALAVRAS_FOLHA_NEGATIVO:
        if _norm(termo) in nome_conta_norm:
            score += peso
    return score


def conta_e_de_folha(nome_conta_norm: str) -> bool:
    return calcular_score_folha(nome_conta_norm) >= SCORE_MINIMO_FOLHA


# ══════════════════════════════════════════════════════════════════════════
# CLASSIFICAÇÃO DE CONTAS POR POSIÇÃO NO PLANO
# Usa a Classificação para determinar se é DRE, Passivo, Ativo etc.
# ══════════════════════════════════════════════════════════════════════════
def _e_conta_dre(classificacao: str) -> bool:
    """Conta começa com 3 ou 4 = Resultado (DRE)."""
    c = str(classificacao).strip()
    return c.startswith("3") or c.startswith("4")

def _e_conta_passivo(classificacao: str) -> bool:
    """Conta começa com 2 = Passivo."""
    c = str(classificacao).strip()
    return c.startswith("2")

def _e_conta_ativo(classificacao: str) -> bool:
    """Conta começa com 1 = Ativo."""
    c = str(classificacao).strip()
    return c.startswith("1")


# ══════════════════════════════════════════════════════════════════════════
# MOTOR DE SCORING — TABELAS SEPARADAS D vs C
# Calibradas contra os planos reais 9999 e o plano alternativo
# ══════════════════════════════════════════════════════════════════════════

# ─── PROVENTOS ────────────────────────────────────────────────────────────
# Débito = conta de DESPESA/CUSTO na DRE (NUNCA provisão/passivo/ativo)
# Termos calibrados para encontrar contas como:
#   331 SALÁRIOS E ORDENADOS (Adm DRE)
#   298 SALÁRIOS E ORDENADOS CUSTOS (Custo DRE)
#   331 SALÁRIOS E ORDENADOS (Despesas Adm)
PROVENTO_DEBITO_POS = [
    # Contas de Despesas Administrativas com Pessoal — peso máximo
    ("DESPESAS COM PESSOAL",          300),
    ("SALARIOS E ORDENADOS CUSTOS",   200),   # 298 no plano alt.
    ("PRO-LABORE CUSTOS",             200),
    ("SALARIOS E ORDENADOS",          150),   # 331 Adm DRE
    ("PRO-LABORE",                    120),
    ("PREMIOS E GRATIFICACOES",       120),
    ("13 SALARIO CUSTOS",             180),
    ("13 SALARIO",                    100),
    ("FERIAS CUSTOS",                 180),
    ("FERIAS",                         80),
    ("INSS CUSTOS",                   180),   # 303 DRE
    ("FGTS CUSTOS",                   180),   # 304 DRE
    ("INSS",                           80),   # 336 DRE
    ("FGTS",                           80),   # 337 DRE
    ("INDENIZACOES E AVISO PREVIO",   120),
    ("ASSISTENCIA MEDICA E SOCIAL",   100),
    ("VALE TRANSPORTE",                80),
    ("HORAS EXTRAS",                   80),
    ("PIS S/ FOLHA",                   80),
    ("MAO-DE-OBRA DIRETA",            200),
    ("MAO-DE-OBRA INDIRETA",          200),
    ("DESPESA",                        30),
    ("CUSTO",                          30),
]
PROVENTO_DEBITO_NEG = [
    # Bloqueia passivos
    ("A PAGAR",                      -300),
    ("A RECOLHER",                   -300),
    # Bloqueia provisões no passivo
    ("PROVISOES PARA FERIAS",        -400),
    ("PROVISOES PARA 13",            -400),
    ("INSS SOBRE PROVISOES",         -400),
    ("FGTS SOBRE PROVISOES",         -400),
    ("PIS SOBRE PROVISOES",          -400),
    ("PROVISAO",                     -300),
    # Bloqueia ativos
    ("A COMPENSAR",                  -400),   # INSS A COMPENSAR é Ativo!
    ("A RECUPERAR",                  -300),
    ("ADIANTAMENTO DE SALARIO",      -300),
    ("ADIANTAMENTO DE 13",           -300),
    ("ADIANTAMENTO DE FERIAS",       -300),
    ("EMPRESTIMO / CONSIGNADO",      -400),   # É Ativo!
    ("EMPRESTIMOS",                  -200),
    # Bloqueia obrigações do passivo
    ("OBRIGACOES COM O PESSOAL",     -300),
    ("OBRIGACOES SOCIAIS",           -300),
    ("OBRIGACOES TRABALHISTA",       -300),
    # Bloqueia contas não-DRE
    ("FORNECEDORES",                 -300),
    ("RETENCOES",                    -300),
    ("RECEITA",                      -300),
    ("BANCO",                        -200),
    ("CLIENTES",                     -300),
    ("CAPITAL SOCIAL",               -300),
    ("DIVIDENDOS",                   -300),
    ("RESULTADO DO EXERCICIO",       -300),
    # Bloqueia contas de benefícios que têm nomes similares
    # mas não são contas de débito de folha
    ("DESPESAS COM ALIMENTACAO",     -100),   # pode ser confundida
    ("ALIMENTACAO/ CESTA BASICA",    -100),
    ("VALE REFEICAO",                -100),   # 672 não é conta de débito de salário
]

# Crédito = SALÁRIOS A PAGAR (passivo circulante)
PROVENTO_CREDITO_POS = [
    ("SALARIOS E ORDENADOS A PAGAR", 500),
    ("SALARIOS A PAGAR",             500),
    ("ORDENADOS A PAGAR",            500),
    ("FOLHA A PAGAR",                500),
    ("OBRIGACOES COM O PESSOAL",     250),
    ("PRO-LABORE A PAGAR",           200),
    ("GRATIFICACOES A PAGAR",        180),
    ("RESCISOES A PAGAR",            180),
    ("FERIAS A PAGAR",               180),
    ("13 SALARIO A PAGAR",           180),
    ("PENSAO ALIMENTICIA A PAGAR",   180),
    ("COMISSOES A PAGAR",            180),
    ("AUTONOMOS A PAGAR",            180),
    ("INDENIZACOES A PAGAR",         180),
    ("PREMIOS E BONIFICACOES",       160),
]
PROVENTO_CREDITO_NEG = [
    ("DESPESA",                      -300),
    ("CUSTO",                        -300),
    ("INSS A RECOLHER",              -200),
    ("FGTS A RECOLHER",              -200),
    ("IRRF S/ FOLHA",                -200),
    ("PIS S/ FOLHA A RECOLHER",      -200),
    ("PROVISAO",                     -200),
    ("RECEITA",                      -300),
    ("BANCO",                        -200),
    ("FORNECEDORES",                 -300),
    ("SALARIOS E ORDENADOS CUSTOS",  -300),
    ("A COMPENSAR",                  -400),
    ("A RECUPERAR",                  -300),
    ("EMPRESTIMO",                   -200),
]


# ─── DESCONTOS ────────────────────────────────────────────────────────────
DESCONTO_DEBITO_POS = [
    ("SALARIOS E ORDENADOS A PAGAR", 500),
    ("SALARIOS A PAGAR",             500),
    ("ORDENADOS A PAGAR",            500),
    ("FOLHA A PAGAR",                500),
    ("OBRIGACOES COM O PESSOAL",     250),
    ("PRO-LABORE A PAGAR",           200),
    ("GRATIFICACOES A PAGAR",        180),
    ("RESCISOES A PAGAR",            180),
]
DESCONTO_DEBITO_NEG = [
    ("DESPESA",                      -300),
    ("CUSTO",                        -300),
    ("INSS A RECOLHER",              -200),
    ("FGTS A RECOLHER",              -200),
    ("IRRF S/ FOLHA",                -200),
    ("PROVISAO",                     -200),
    ("RECEITA",                      -300),
    ("BANCO",                        -200),
    ("FORNECEDORES",                 -300),
    ("SALARIOS E ORDENADOS CUSTOS",  -300),
    ("SALARIOS E ORDENADOS",         -150),
    ("A COMPENSAR",                  -400),
    ("A RECUPERAR",                  -300),
    ("EMPRESTIMO",                   -200),
]

# Créditos específicos por tipo de desconto
DESCONTO_CRED_INSS_POS = [
    ("INSS A RECOLHER",              500),
    ("OBRIGACOES SOCIAIS",           200),
    ("OBRIGACOES TRABALHISTAS E PREVIDENCIARIA", 200),
    ("A RECOLHER",                   100),
]
DESCONTO_CRED_INSS_NEG = [
    ("FGTS",                        -100),
    ("IRRF",                        -150),
    ("SALARIOS A PAGAR",            -100),
    ("DESPESA",                     -300),
    ("CUSTO",                       -300),
    ("PROVISAO",                    -150),
    ("A COMPENSAR",                 -400),
    ("A RECUPERAR",                 -300),
]

DESCONTO_CRED_IRRF_POS = [
    ("IRRF S/ FOLHA",                500),
    ("IMPOSTO DE RENDA A RECOLHER",  400),
    ("IRRF",                         250),
    ("A RECOLHER",                   100),
]
DESCONTO_CRED_IRRF_NEG = [
    ("INSS",                        -150),
    ("FGTS",                        -150),
    ("SALARIOS A PAGAR",            -100),
    ("DESPESA",                     -300),
    ("CUSTO",                       -300),
    ("A COMPENSAR",                 -400),
    ("A RECUPERAR",                 -300),
    ("NF",                          -150),
    ("ALUGUEL",                     -150),
    ("APLICACAO",                   -150),
]

DESCONTO_CRED_VT_POS = [
    ("SALARIOS E ORDENADOS A PAGAR", 500),
    ("SALARIOS A PAGAR",             500),
    ("OBRIGACOES COM O PESSOAL",     250),
]
DESCONTO_CRED_VT_NEG = [
    ("INSS",                        -150),
    ("FGTS",                        -150),
    ("IRRF",                        -150),
    ("DESPESA",                     -300),
    ("CUSTO",                       -300),
    ("A COMPENSAR",                 -400),
    ("A RECUPERAR",                 -300),
]

DESCONTO_CRED_PLANO_POS = [
    ("SALARIOS E ORDENADOS A PAGAR", 500),
    ("SALARIOS A PAGAR",             500),
    ("OBRIGACOES COM O PESSOAL",     250),
]
DESCONTO_CRED_PLANO_NEG = [
    ("INSS",                        -150),
    ("FGTS",                        -150),
    ("IRRF",                        -150),
    ("DESPESA",                     -300),
    ("CUSTO",                       -300),
    ("A COMPENSAR",                 -400),
    ("A RECUPERAR",                 -300),
]

DESCONTO_CRED_ADIANT_POS = [
    ("SALARIOS E ORDENADOS A PAGAR", 500),
    ("SALARIOS A PAGAR",             500),
    ("OBRIGACOES COM O PESSOAL",     250),
]
DESCONTO_CRED_ADIANT_NEG = [
    ("INSS",                        -150),
    ("FGTS",                        -150),
    ("IRRF",                        -150),
    ("DESPESA",                     -300),
    ("CUSTO",                       -300),
    ("A COMPENSAR",                 -400),
    ("A RECUPERAR",                 -300),
]

DESCONTO_CRED_FECHAMENTO_POS = [
    ("SALARIOS E ORDENADOS A PAGAR", 500),
    ("SALARIOS A PAGAR",             500),
    ("OBRIGACOES COM O PESSOAL",     250),
]
DESCONTO_CRED_FECHAMENTO_NEG = [
    ("INSS",                        -150),
    ("FGTS",                        -150),
    ("IRRF",                        -150),
    ("DESPESA",                     -300),
    ("CUSTO",                       -300),
    ("A COMPENSAR",                 -400),
    ("A RECUPERAR",                 -300),
]

DESCONTO_CREDITO_GENERICO_POS = [
    ("INSS A RECOLHER",              400),
    ("FGTS A RECOLHER",              400),
    ("IRRF S/ FOLHA",                400),
    ("PIS S/ FOLHA A RECOLHER",      350),
    ("CONTRIBUICOES SINDICAIS",      300),
    ("OBRIGACOES SOCIAIS",           250),
    ("OBRIGACOES TRABALHISTAS E PREVIDENCIARIA", 250),
    ("PENSAO ALIMENTICIA A PAGAR",   250),
    ("SALARIOS E ORDENADOS A PAGAR", 200),
    ("SALARIOS A PAGAR",             200),
    ("A RECOLHER",                   100),
    ("A PAGAR",                       80),
    ("RETIDO",                       100),
    ("PASSIVO",                       50),
]
DESCONTO_CREDITO_GENERICO_NEG = [
    ("DESPESA",                     -300),
    ("CUSTO",                       -300),
    ("FORNECEDORES",                -300),
    ("RECEITA",                     -300),
    ("BANCO",                       -200),
    ("PROVISOES PARA FERIAS",       -200),
    ("PROVISOES PARA 13",           -200),
    ("INSS SOBRE PROVISOES",        -200),
    ("FGTS SOBRE PROVISOES",        -200),
    ("A COMPENSAR",                 -400),
    ("A RECUPERAR",                 -300),
    ("EMPRESTIMO",                  -100),
]


# ─── CONSIGNADO ───────────────────────────────────────────────────────────
CONSIGNADO_DEBITO_POS = [
    ("SALARIOS E ORDENADOS A PAGAR", 500),
    ("SALARIOS A PAGAR",             500),
    ("ORDENADOS A PAGAR",            500),
    ("FOLHA A PAGAR",                500),
    ("OBRIGACOES COM O PESSOAL",     250),
]
CONSIGNADO_DEBITO_NEG = [
    ("DESPESA",                     -300),
    ("CUSTO",                       -300),
    ("BANCOS",                      -300),
    ("FORNECEDORES",                -300),
    ("EMPRESTIMOS",                 -150),
    ("CONSIGNADO",                  -150),
    ("RECEITA",                     -300),
    ("INSS",                        -150),
    ("FGTS",                        -150),
    ("IRRF",                        -150),
    ("SALARIOS E ORDENADOS CUSTOS", -300),
    ("SALARIOS E ORDENADOS",        -150),
    ("A COMPENSAR",                 -400),
    ("A RECUPERAR",                 -300),
]

# Consignado crédito: busca conta de Ativo (empréstimo ao funcionário)
# OU conta de Passivo (repasse ao banco)
# Prioriza EMPRESTIMO / CONSIGNADO (Ativo 11307) ou conta passivo específica
CONSIGNADO_CREDITO_POS = [
    ("EMPRESTIMO / CONSIGNADO",      600),   # 54 no plano alt. — Ativo do empregado
    ("EMPRESTIMOS CONSIGNADOS",      500),
    ("CONSIGNADOS A PAGAR",          500),
    ("CREDITO TRABALHO",             480),
    ("EMPRESTIMOS DE FUNCIONARIOS",  450),
    ("EMPRESTIMO A EMPREGADOS",      400),   # grupo 11307 no plano 9999
    ("EMPRESTIMO",                   200),
    ("CONSIGNADO",                   200),
]
CONSIGNADO_CREDITO_NEG = [
    ("DESPESA",                     -300),
    ("CUSTO",                       -300),
    ("SALARIOS A PAGAR",            -150),
    ("ORDENADOS A PAGAR",           -150),
    ("FORNECEDORES",                -300),
    ("RECEITA",                     -300),
    ("INSS",                        -150),
    ("FGTS",                        -150),
    ("IRRF",                        -150),
    ("A COMPENSAR",                 -400),
    ("A RECUPERAR",                 -300),
    ("BANCO DO BRASIL",             -100),
    ("BANCO ITAU",                  -100),
    ("BANCO BRADESCO",              -100),
    ("FINANCIAMENTO",               -100),
]


# ─── INFORMATIVOS / ENCARGO PATRONAL ─────────────────────────────────────
# Débito = conta de DESPESA/CUSTO/ENCARGO na DRE
INFORMATIVO_DEBITO_POS = [
    ("DESPESAS COM PESSOAL",         300),
    ("INSS CUSTOS",                  300),   # 303 DRE
    ("FGTS CUSTOS",                  300),   # 304 DRE
    ("ENCARGOS SOCIAIS",             280),
    ("CONTRIBUICAO PATRONAL",        280),
    ("CONTRIBUICAO PREVIDENCIARIA",  280),
    ("INSS EMPRESA",                 260),
    ("INSS TERCEIROS",               260),
    ("INSS ACIDENTE",                260),
    ("PATRONAL",                     260),
    ("INSS",                         150),   # 336 DRE
    ("FGTS",                         150),   # 337 DRE
    ("DESPESA",                       50),
    ("CUSTO",                         50),
]
INFORMATIVO_DEBITO_NEG = [
    ("A RECOLHER",                  -300),
    ("A PAGAR",                     -300),
    ("PROVISOES PARA FERIAS",       -400),
    ("PROVISOES PARA 13",           -400),
    ("INSS SOBRE PROVISOES",        -400),
    ("FGTS SOBRE PROVISOES",        -400),
    ("PIS SOBRE PROVISOES",         -400),
    ("PROVISAO",                    -300),
    ("OBRIGACOES COM O PESSOAL",    -300),
    ("OBRIGACOES SOCIAIS",          -300),
    ("OBRIGACOES TRABALHISTA",      -300),
    ("FORNECEDORES",                -300),
    ("RECEITA",                     -300),
    ("BANCO",                       -200),
    # CRÍTICO: bloqueia contas de Ativo
    ("A COMPENSAR",                 -500),   # INSS A COMPENSAR é Ativo!
    ("A RECUPERAR",                 -400),
    ("ADIANTAMENTO",                -200),
    ("EMPRESTIMO",                  -300),
    ("SALDO NEGATIVO",              -300),
]

INFORMATIVO_CREDITO_POS = [
    ("INSS A RECOLHER",              500),
    ("FGTS A RECOLHER",              500),
    ("OBRIGACOES SOCIAIS",           250),
    ("OBRIGACOES TRABALHISTAS E PREVIDENCIARIA", 250),
    ("CONTRIBUICOES SINDICAIS",      200),
    ("PIS S/ FOLHA A RECOLHER",      200),
    ("A RECOLHER",                   100),
    ("PASSIVO",                       50),
]
INFORMATIVO_CREDITO_NEG = [
    ("DESPESA",                     -300),
    ("CUSTO",                       -300),
    ("SALARIOS A PAGAR",            -200),
    ("ORDENADOS A PAGAR",           -200),
    ("PROVISOES PARA FERIAS",       -200),
    ("PROVISOES PARA 13",           -200),
    ("FORNECEDORES",                -300),
    ("RECEITA",                     -300),
    ("BANCO",                       -200),
    ("A COMPENSAR",                 -400),
    ("A RECUPERAR",                 -300),
]


# ─── FGTS DINÂMICO ────────────────────────────────────────────────────────
FGTS_CREDITO_POS = [
    ("FGTS A RECOLHER",              500),
    ("FGTS A PAGAR",                 480),
    ("FUNDO DE GARANTIA",            400),
    ("OBRIGACOES SOCIAIS",           200),
]
FGTS_CREDITO_NEG = [
    ("DESPESA",                     -300),
    ("CUSTO",                       -300),
    ("PROVISAO",                    -200),
    ("SALARIOS A PAGAR",            -200),
    ("INSS",                        -100),
    ("FORNECEDOR",                  -300),
    ("A COMPENSAR",                 -400),
    ("A RECUPERAR",                 -300),
]

FGTS_DEBITO_NORMAL_POS = [
    ("FGTS CUSTOS",                  400),   # 304 DRE
    ("FGTS",                         250),   # 337 DRE
    ("ENCARGOS SOCIAIS",             230),
    ("DESPESAS COM PESSOAL",         200),
    ("DESPESA",                       80),
    ("CUSTO",                         80),
]
FGTS_DEBITO_NORMAL_NEG = [
    ("A RECOLHER",                  -300),
    ("A PAGAR",                     -300),
    ("PROVISOES PARA FERIAS",       -400),
    ("PROVISOES PARA 13",           -400),
    ("SOBRE PROVISOES",             -400),
    ("PROVISAO",                    -300),
    ("FORNECEDOR",                  -300),
    ("A COMPENSAR",                 -400),
    ("A RECUPERAR",                 -300),
]

FGTS_DEBITO_PROVISAO_POS = [
    ("FGTS SOBRE PROVISOES PARA FERIAS", 500),
    ("FGTS SOBRE PROVISOES PARA 13",     500),
    ("FGTS SOBRE PROVISOES",             450),
    ("PROVISOES PARA FERIAS",            300),
    ("PROVISOES PARA 13",                300),
    ("PROVISAO",                         200),
]
FGTS_DEBITO_PROVISAO_NEG = [
    ("A RECOLHER",                  -300),
    ("A PAGAR",                     -200),
    ("DESPESA",                     -150),
    ("CUSTO",                       -150),
    ("FORNECEDOR",                  -300),
    ("A COMPENSAR",                 -400),
    ("A RECUPERAR",                 -300),
]

FGTS_DEBITO_RESCISAO_POS = [
    ("FGTS RESCISORIO",              500),
    ("MULTA RESCISORIA",             500),
    ("GRRF",                         500),
    ("DESPESA COM RESCISAO",         400),
    ("FGTS CUSTOS",                  300),
    ("DESPESAS COM PESSOAL",         200),
    ("DESPESA",                       80),
    ("CUSTO",                         80),
]
FGTS_DEBITO_RESCISAO_NEG = [
    ("A RECOLHER",                  -300),
    ("A PAGAR",                     -300),
    ("PROVISOES PARA FERIAS",       -400),
    ("PROVISOES PARA 13",           -400),
    ("SOBRE PROVISOES",             -400),
    ("PROVISAO",                    -300),
    ("FORNECEDOR",                  -300),
    ("A COMPENSAR",                 -400),
    ("A RECUPERAR",                 -300),
]


# ══════════════════════════════════════════════════════════════════════════
# DETECÇÃO DE TIPO DE RUBRICA PELO NOME
# ══════════════════════════════════════════════════════════════════════════
def _e_consignado(nome_norm: str) -> bool:
    termos = ["EMPRESTIMO CONSIGNADO", "CONSIGNADO", "CREDITO TRABALHO",
              "EMP CRED TRAB", "EMPRESTIMO TRABALHADOR", "CONSIG",
              "EMP. CRED. TRAB", "EMPRESTIMO CONSIG"]
    return any(_norm(t) in nome_norm for t in termos)

def _e_fgts(nome_norm: str) -> bool:
    return "FGTS" in nome_norm or "F.G.T.S" in nome_norm

def _e_inss_desconto(nome_norm: str) -> bool:
    termos = ["I.N.S.S", "INSS FERIAS", "INSS SOBRE RESCISAO",
              "INSS 13", "INSS DIFERENCA", "INSS EMPREGADO",
              "INSS SOBRE RESCISAO"]
    return any(_norm(t) in nome_norm for t in termos)

def _e_irrf(nome_norm: str) -> bool:
    termos = ["IMPOSTO DE RENDA", "IRRF", "I.R.R.F", "IR FONTE",
              "IRRF FERIAS", "IRRF 13"]
    return any(_norm(t) in nome_norm for t in termos)

def _e_vale_transporte(nome_norm: str) -> bool:
    termos = ["VALE TRANSPORTE", "VT 6", "DESC VT", "DESCONTO VT",
              "VALE TRANSP", "DESCONTO VALE TRANSP"]
    return any(_norm(t) in nome_norm for t in termos)

def _e_plano_saude_odonto(nome_norm: str) -> bool:
    termos = ["PLANO SAUDE", "PLANO DE SAUDE", "ODONTOLOGICO", "ODONTO",
              "COPARTICIPACAO", "COPART", "PLANO ODONT", "DESCONTO PLANO"]
    return any(_norm(t) in nome_norm for t in termos)

def _e_adiantamento_desc(nome_norm: str) -> bool:
    termos = ["DESC.ADIANT", "DESC ADIANT", "DESCONTO ADIANT",
              "DESCONTO ADIANTAMENTO", "ADIANTAMENTO SALARIAL DESC",
              "DESC.ADIANT.SALARIAL", "ADIANTAMENTO DE FERIAS"]
    return any(_norm(t) in nome_norm for t in termos)

def _e_fechamento(nome_norm: str) -> bool:
    termos = ["ESTOURO", "TROCO", "FECHAMENTO", "SALDO DE SALARIO",
              "SALDO SALARIO", "LIQUIDO RESCISAO", "LIQUIDO FOLHA",
              "ESTOURO MES", "ESTOURO RESCISAO", "ESTOURO SEMANA",
              "ESTOURO CONVOCACAO"]
    return any(_norm(t) in nome_norm for t in termos)

def _e_pensao_alimenticia(nome_norm: str) -> bool:
    return "PENSAO ALIMENTICIA" in nome_norm or "PENSAO ALIMENT" in nome_norm

def _e_sindicato(nome_norm: str) -> bool:
    termos = ["SINDICATO", "SINDICAL", "MENSALIDADE SINDICAL",
              "CONTRIBUICAO SINDICAL", "CONTRIB SINDICAL",
              "CONFEDERATIVA", "ASSISTENCIAL"]
    return any(_norm(t) in nome_norm for t in termos)

def _e_vale_refeicao_desc(nome_norm: str) -> bool:
    termos = ["DESC VALE REFEICAO", "DESCONTO VALE REFEICAO",
              "DESC. VALE REFEICAO", "DESCONTO VR"]
    return any(_norm(t) in nome_norm for t in termos)


# ══════════════════════════════════════════════════════════════════════════
# MOTOR DE SCORING — com filtro por posição no plano
# ══════════════════════════════════════════════════════════════════════════
def _score_conta(nome_norm: str,
                 pos: list[tuple[str, int]],
                 neg: list[tuple[str, int]]) -> int:
    score = 0
    for t, p in pos:
        if _norm(t) in nome_norm:
            score += p
    for t, p in neg:
        if _norm(t) in nome_norm:
            score += p
    return score


def _melhor_conta(
    df_contas: pd.DataFrame,
    pos: list[tuple[str, int]],
    neg: list[tuple[str, int]],
    score_min: int = 1,
    filtro_posicao: str = "",   # "DRE", "PASSIVO", "ATIVO", "" = sem filtro
) -> tuple[str, str]:
    """
    Retorna a melhor conta analítica.
    filtro_posicao: se informado, filtra contas pela posição no plano
    (DRE=começa com 3 ou 4, PASSIVO=começa com 2, ATIVO=começa com 1)
    """
    df_a = df_contas[df_contas["tipo"] == "A"].copy()
    if df_a.empty:
        return "", ""

    # Aplica filtro de posição se solicitado
    if filtro_posicao == "DRE":
        mask = df_a["classificacao"].apply(_e_conta_dre)
        df_filtrado = df_a[mask]
        if df_filtrado.empty:
            df_filtrado = df_a  # fallback sem filtro
    elif filtro_posicao == "PASSIVO":
        mask = df_a["classificacao"].apply(_e_conta_passivo)
        df_filtrado = df_a[mask]
        if df_filtrado.empty:
            df_filtrado = df_a
    elif filtro_posicao == "ATIVO":
        mask = df_a["classificacao"].apply(_e_conta_ativo)
        df_filtrado = df_a[mask]
        if df_filtrado.empty:
            df_filtrado = df_a
    else:
        df_filtrado = df_a

    best_score, best_red, best_nome = -9999, "", ""
    for _, row in df_filtrado.iterrows():
        s = _score_conta(row["nome_conta"], pos, neg)
        if s > best_score:
            best_score, best_red, best_nome = s, row["reduzido"], row["nome_original"]

    return (best_red, best_nome) if best_score >= score_min else ("", "")


# ══════════════════════════════════════════════════════════════════════════
# MOTOR PRINCIPAL: gerar De/Para por evento — V6.3
# Usa filtro de posição no plano para garantir DRE vs Passivo vs Ativo
# ══════════════════════════════════════════════════════════════════════════
def gerar_depara_evento_conta(
    evento_cod:  str,
    evento_nome: str,
    evento_tipo: str,
    grupo:       str,
    df_contas:   pd.DataFrame,
    tipo_folha:  str = "1",
) -> dict:
    vazio = {"conta_debito": "", "conta_credito": "",
             "desc_debito": "", "desc_credito": ""}
    if df_contas is None or df_contas.empty:
        return vazio

    nome = _norm(evento_nome)
    e_encargo = (tipo_folha == "2") or (grupo == "Encargo Patronal")

    # ── CASO 1: Consignado ────────────────────────────────────────────────
    if _e_consignado(nome) and evento_tipo in ("Desconto", "Inf. Dedutora"):
        # Débito = Salários a Pagar (PASSIVO)
        cd, dd = _melhor_conta(df_contas, CONSIGNADO_DEBITO_POS, CONSIGNADO_DEBITO_NEG,
                               score_min=100, filtro_posicao="PASSIVO")
        # Crédito = Empréstimo/Consignado (pode ser Ativo ou Passivo dependendo do plano)
        cc, dc = _melhor_conta(df_contas, CONSIGNADO_CREDITO_POS, CONSIGNADO_CREDITO_NEG,
                               score_min=100, filtro_posicao="")
        if cd and cc and validar_par_contas(cd, cc):
            return {"conta_debito": cd, "conta_credito": cc,
                    "desc_debito": dd, "desc_credito": dc}

    # ── CASO 2: FGTS dinâmico ─────────────────────────────────────────────
    if _e_fgts(nome) and evento_tipo in ("Informativa", "Inf. Dedutora"):
        # Crédito sempre = FGTS a Recolher (PASSIVO)
        cc, dc = _melhor_conta(df_contas, FGTS_CREDITO_POS, FGTS_CREDITO_NEG,
                               score_min=100, filtro_posicao="PASSIVO")
        # Débito varia por tipo de folha
        if tipo_folha in ("3", "6"):
            cd, dd = _melhor_conta(df_contas, FGTS_DEBITO_PROVISAO_POS, FGTS_DEBITO_PROVISAO_NEG,
                                   score_min=50, filtro_posicao="PASSIVO")
        elif tipo_folha == "4":
            cd, dd = _melhor_conta(df_contas, FGTS_DEBITO_RESCISAO_POS, FGTS_DEBITO_RESCISAO_NEG,
                                   score_min=50, filtro_posicao="DRE")
        else:
            cd, dd = _melhor_conta(df_contas, FGTS_DEBITO_NORMAL_POS, FGTS_DEBITO_NORMAL_NEG,
                                   score_min=50, filtro_posicao="DRE")
        if cd and cc and validar_par_contas(cd, cc):
            return {"conta_debito": cd, "conta_credito": cc,
                    "desc_debito": dd, "desc_credito": dc}

    # ── CASO 3: Encargo Patronal ──────────────────────────────────────────
    if e_encargo:
        # Débito = conta de Encargo na DRE (NUNCA Ativo como INSS A COMPENSAR)
        cd, dd = _melhor_conta(df_contas, INFORMATIVO_DEBITO_POS, INFORMATIVO_DEBITO_NEG,
                               score_min=50, filtro_posicao="DRE")
        # Crédito = INSS a Recolher (PASSIVO)
        cc, dc = _melhor_conta(df_contas, INFORMATIVO_CREDITO_POS, INFORMATIVO_CREDITO_NEG,
                               score_min=50, filtro_posicao="PASSIVO")
        if cd and cc and validar_par_contas(cd, cc):
            return {"conta_debito": cd, "conta_credito": cc,
                    "desc_debito": dd, "desc_credito": dc}

    # ── CASO 4: Proventos ─────────────────────────────────────────────────
    if evento_tipo == "Provento":
        # Débito = conta de Despesa/Custo na DRE
        cd, dd = _melhor_conta(df_contas, PROVENTO_DEBITO_POS, PROVENTO_DEBITO_NEG,
                               score_min=50, filtro_posicao="DRE")
        # Crédito = Salários a Pagar (PASSIVO)
        cc, dc = _melhor_conta(df_contas, PROVENTO_CREDITO_POS, PROVENTO_CREDITO_NEG,
                               score_min=100, filtro_posicao="PASSIVO")
        if cd and cc and validar_par_contas(cd, cc):
            return {"conta_debito": cd, "conta_credito": cc,
                    "desc_debito": dd, "desc_credito": dc}
        # curto-circuito: retentar débito excluindo a conta crédito
        if cc:
            cd, dd = _melhor_conta(df_contas, PROVENTO_DEBITO_POS,
                                   PROVENTO_DEBITO_NEG + [(_norm(cc), -1000)],
                                   score_min=20, filtro_posicao="DRE")
        return {"conta_debito": cd, "conta_credito": cc,
                "desc_debito": dd, "desc_credito": dc}

    # ── CASO 5: Descontos — crédito específico por tipo ───────────────────
    if evento_tipo in ("Desconto", "Inf. Dedutora"):
        # Débito sempre = Salários a Pagar (PASSIVO)
        cd, dd = _melhor_conta(df_contas, DESCONTO_DEBITO_POS, DESCONTO_DEBITO_NEG,
                               score_min=100, filtro_posicao="PASSIVO")

        # Crédito: roteamento por tipo de desconto
        if _e_inss_desconto(nome):
            cc, dc = _melhor_conta(df_contas, DESCONTO_CRED_INSS_POS, DESCONTO_CRED_INSS_NEG,
                                   score_min=100, filtro_posicao="PASSIVO")
        elif _e_irrf(nome):
            cc, dc = _melhor_conta(df_contas, DESCONTO_CRED_IRRF_POS, DESCONTO_CRED_IRRF_NEG,
                                   score_min=100, filtro_posicao="PASSIVO")
        elif _e_vale_transporte(nome):
            cc, dc = _melhor_conta(df_contas, DESCONTO_CRED_VT_POS, DESCONTO_CRED_VT_NEG,
                                   score_min=100, filtro_posicao="PASSIVO")
        elif _e_plano_saude_odonto(nome):
            cc, dc = _melhor_conta(df_contas, DESCONTO_CRED_PLANO_POS, DESCONTO_CRED_PLANO_NEG,
                                   score_min=100, filtro_posicao="PASSIVO")
        elif _e_adiantamento_desc(nome):
            cc, dc = _melhor_conta(df_contas, DESCONTO_CRED_ADIANT_POS, DESCONTO_CRED_ADIANT_NEG,
                                   score_min=100, filtro_posicao="PASSIVO")
        elif _e_fechamento(nome):
            cc, dc = _melhor_conta(df_contas, DESCONTO_CRED_FECHAMENTO_POS, DESCONTO_CRED_FECHAMENTO_NEG,
                                   score_min=100, filtro_posicao="PASSIVO")
        elif _e_vale_refeicao_desc(nome):
            cc, dc = _melhor_conta(df_contas, DESCONTO_CRED_VT_POS, DESCONTO_CRED_VT_NEG,
                                   score_min=100, filtro_posicao="PASSIVO")
        elif _e_pensao_alimenticia(nome):
            cc, dc = _melhor_conta(df_contas, [
                ("PENSAO ALIMENTICIA A PAGAR", 500),
                ("SALARIOS E ORDENADOS A PAGAR", 250),
                ("A PAGAR", 100),
            ], [("DESPESA",-300),("CUSTO",-300),("INSS",-150),("FGTS",-150),
                ("A COMPENSAR",-400),("A RECUPERAR",-300)],
                score_min=100, filtro_posicao="PASSIVO")
        elif _e_sindicato(nome):
            cc, dc = _melhor_conta(df_contas, [
                ("CONTRIBUICOES SINDICAIS", 500),
                ("MENSALIDADE SINDICAL", 500),
                ("OBRIGACOES SOCIAIS", 250),
                ("A RECOLHER", 100),
            ], [("DESPESA",-300),("CUSTO",-300),("SALARIOS A PAGAR",-100),
                ("A COMPENSAR",-400),("A RECUPERAR",-300)],
                score_min=100, filtro_posicao="PASSIVO")
        else:
            cc, dc = _melhor_conta(df_contas, DESCONTO_CREDITO_GENERICO_POS,
                                   DESCONTO_CREDITO_GENERICO_NEG,
                                   score_min=50, filtro_posicao="PASSIVO")

        if cd and cc and validar_par_contas(cd, cc):
            return {"conta_debito": cd, "conta_credito": cc,
                    "desc_debito": dd, "desc_credito": dc}
        # curto-circuito: retentar crédito
        if cd:
            cc, dc = _melhor_conta(df_contas, DESCONTO_CREDITO_GENERICO_POS,
                                   DESCONTO_CREDITO_GENERICO_NEG + [(_norm(cd), -1000)],
                                   score_min=30, filtro_posicao="PASSIVO")
        return {"conta_debito": cd, "conta_credito": cc,
                "desc_debito": dd, "desc_credito": dc}

    # ── CASO 6: Informativas genéricas ────────────────────────────────────
    if evento_tipo == "Informativa":
        cd, dd = _melhor_conta(df_contas, INFORMATIVO_DEBITO_POS, INFORMATIVO_DEBITO_NEG,
                               score_min=30, filtro_posicao="DRE")
        cc, dc = _melhor_conta(df_contas, INFORMATIVO_CREDITO_POS, INFORMATIVO_CREDITO_NEG,
                               score_min=30, filtro_posicao="PASSIVO")
        if cd and cc and validar_par_contas(cd, cc):
            return {"conta_debito": cd, "conta_credito": cc,
                    "desc_debito": dd, "desc_credito": dc}

    return vazio


# ══════════════════════════════════════════════════════════════════════════
# CLASSIFICADOR LOCAL DE RUBRICAS
# ══════════════════════════════════════════════════════════════════════════
KWORDS_RUBRICA: dict[str, list[str]] = {
    "Custo Direto de Produção": [
        "MATERIA PRIMA","MATERIAL APLICADO","MAO DE OBRA DIRETA",
        "SALARIOS E ORDENADOS CUSTOS","PRO LABORE CUSTOS",
        "13 SALARIO CUSTOS","FERIAS CUSTOS","INSS CUSTOS",
        "FGTS CUSTOS","INDUSTRIALIZACAO","PRODUCAO",
    ],
    "Custo Direto de Serviços": [
        "SERVICOS PRESTADOS","MAO DE OBRA","PRESTACAO DE SERVICO",
        "TERCEIRIZACAO","SUBCONTRATACAO","CUSTOS SERVICOS",
    ],
    "Custo Indireto de Produção": [
        "OVERHEAD","MANUTENCAO","DEPRECIACAO","ENERGIA ELETRICA",
        "AGUA ESGOTO","COMBUSTIVEL","ALUGUEL EQUIPAMENTO",
        "MATERIAL CONSUMO INDIRETO","LOCACAO","CONDOMINIO",
    ],
    "Despesa com Vendas": [
        "COMISSAO","COMISSOES","PROPAGANDA","PUBLICIDADE",
        "BONIFICACAO","REPRESENTACAO","FRETE VENDA",
    ],
    "Despesa Financeira": [
        "JUROS","IOF","TARIFA BANCARIA","VARIACAO CAMBIAL",
        "VARIACAO MONETARIA","DESCONTO CONCEDIDO","MORA",
        "EMPRESTIMO","FINANCIAMENTO",
    ],
    "Despesa Não Operacional": [
        "IRPJ","CSLL","PROVISAO IR","PROVISAO CSLL",
        "PERDA","ALIENACAO","SINISTRO","BAIXA ATIVO",
    ],
    "Despesa Administrativa": [
        "SALARIO","ORDENADO","PRO LABORE","GRATIFICACAO",
        "BONUS","PREMIO","HORAS EXTRAS","ADICIONAL NOTURNO",
        "ADICIONAL","INSALUBRIDADE","PERICULOSIDADE",
        "FERIAS","ABONO","13 SALARIO","AVISO PREVIO",
        "RESCISAO","INDENIZACAO","LICENCA","AFASTAMENTO",
        "REPOUSO","DSR","REFLEXO","MEDIA","DIFERENCA",
        "ADIANTAMENTO","VALE TRANSPORTE","VALE REFEICAO",
        "ALIMENTACAO","CESTA BASICA","AUXILIO","REEMBOLSO",
        "PARTICIPACAO LUCROS","PLR","COMPL SALARIAL",
        "DESC ADIANT","DESCONTO ADIANT","INSS","FGTS",
        "IMPOSTO DE RENDA","IRRF","PENSAO ALIMENTICIA",
        "PLANO SAUDE","PLANO ODONTOLOGICO","COPARTICIPACAO",
        "EMPRESTIMO CONSIGNADO","SINDICATO","CONTRIBUICAO",
        "DIAS NORMAIS","HORAS NORMAIS","SALDO SALARIO",
        "ESTOURO","TROCO","FECHAMENTO","DIAS FERIAS",
    ],
    "Encargo Patronal": [
        "INSS EMPRESA","INSS TERCEIROS","INSS ACID","INSS ACIDENTE",
        "INSS PATRONAL","ENCARGO PATRONAL","CONTRIBUICAO PATRONAL",
        "CONTRIBUICAO PREVIDENCIARIA PATRONAL",
        "RAT","FAP","SISTEMA S","SESI","SENAI","SEBRAE","SESC","SENAC",
        "INSS 13","INSS FERIAS","INSS EMPRESA FERIAS","INSS EMPRESA 13",
    ],
}

TIPOS_NAO_CUSTO = {"Desconto", "Informativa", "Inf. Dedutora"}


def classificar_rubrica_local(nome_rubrica: str, tipo_rubrica: str, tipo_folha: str = "1") -> dict:
    if tipo_folha == "2":
        return {"grupo": "Encargo Patronal", "confianca": "alta"}
    nome_norm = _norm(nome_rubrica)
    grupos_candidatos = list(KWORDS_RUBRICA.keys())
    if tipo_rubrica in TIPOS_NAO_CUSTO:
        grupos_candidatos = [g for g in grupos_candidatos if "Custo" not in g]
    scores: dict[str, int] = {g: 0 for g in grupos_candidatos}
    for grupo in grupos_candidatos:
        for kw in KWORDS_RUBRICA[grupo]:
            if _norm(kw) in nome_norm:
                scores[grupo] += len(kw.split())
    melhor = max(scores, key=lambda g: scores[g])
    s = scores[melhor]
    if s >= 4:   c = "alta"
    elif s >= 2: c = "media"
    elif s >= 1: c = "baixa"
    else:
        melhor, c = "Despesa Administrativa", "baixa"
    return {"grupo": melhor, "confianca": c}


# ══════════════════════════════════════════════════════════════════════════
# PARSE DO PLANO DE CONTAS
# ══════════════════════════════════════════════════════════════════════════
def parse_plano_contas(file_bytes: bytes, filename: str, log: list) -> pd.DataFrame:
    ext = filename.lower().rsplit(".", 1)[-1] if "." in filename else "xlsx"
    df_raw = None
    if ext == "xlsx":
        try:
            df_raw = pd.read_excel(BytesIO(file_bytes), sheet_name=0, header=0, dtype=str, engine="openpyxl")
            log.append("Plano de Contas: lido como .xlsx.")
        except Exception as e:
            log.append(f"ERRO ao abrir .xlsx: {e}"); return pd.DataFrame()
    else:
        for engine in ["xlrd", "openpyxl"]:
            try:
                df_raw = pd.read_excel(BytesIO(file_bytes), sheet_name=0, header=0, dtype=str, engine=engine)
                log.append(f"Plano de Contas: lido como .xls (engine={engine})."); break
            except Exception as e:
                log.append(f"  engine={engine} falhou: {e}"); df_raw = None
    if df_raw is None:
        log.append("ERRO: Não foi possível abrir o Plano de Contas."); return pd.DataFrame()

    cols = [str(c).strip() for c in df_raw.columns]
    idx_empresa = idx_reduzido = idx_classificacao = idx_tipo = idx_descricao = None
    for i, c in enumerate(cols):
        cl = c.lower()
        if idx_empresa is None and (i == 0 or "plano de contas" in cl or (cl == "empresa" and i < 3)):
            idx_empresa = i
        if idx_reduzido is None and ("reduzido" in cl or "unnamed: 1" in cl):
            idx_reduzido = i
        if idx_classificacao is None and ("classifica" in cl or "unnamed: 2" in cl):
            idx_classificacao = i
        if idx_tipo is None and ("unnamed: 3" in cl or cl == "tipo" or ("tipo" in cl and "ecf" not in cl and i < 6)):
            idx_tipo = i
        if idx_descricao is None and ("descri" in cl or "unnamed: 4" in cl):
            idx_descricao = i

    if idx_empresa is None:       idx_empresa = 0
    if idx_reduzido is None:      idx_reduzido = 1
    if idx_classificacao is None: idx_classificacao = 2
    if idx_tipo is None:          idx_tipo = 3
    if idx_descricao is None:     idx_descricao = 4

    registros, ignorados = [], 0
    for _, row in df_raw.iterrows():
        empresa_val = str(row.iloc[idx_empresa]).strip()
        reduzido    = str(row.iloc[idx_reduzido]).strip()
        classif     = str(row.iloc[idx_classificacao]).strip()
        tipo_raw    = str(row.iloc[idx_tipo]).strip().upper()
        nome        = str(row.iloc[idx_descricao]).strip()
        if empresa_val.lower().startswith("total"): ignorados += 1; continue
        if empresa_val.lower() in ("nan","none",""): ignorados += 1; continue
        if classif.endswith(".0"): classif = classif[:-2]
        if not re.match(r'^\d+$', classif): ignorados += 1; continue
        if tipo_raw not in ("S","A"): ignorados += 1; continue
        if not nome or nome.lower() in ("nan","none",""): ignorados += 1; continue
        if reduzido.endswith(".0"): reduzido = reduzido[:-2]
        if reduzido.lower() in ("nan","none",""): reduzido = classif
        nome_norm   = _norm(nome)
        score_folha = calcular_score_folha(nome_norm)
        registros.append({
            "reduzido": reduzido, "classificacao": classif,
            "nome_conta": nome_norm, "nome_original": nome,
            "tipo": tipo_raw, "score_folha": score_folha,
        })

    df = pd.DataFrame(registros).drop_duplicates(subset=["classificacao"]).reset_index(drop=True)
    n_a = len(df[df["tipo"] == "A"]); n_s = len(df[df["tipo"] == "S"])
    n_folha = len(df[(df["tipo"] == "A") & (df["score_folha"] >= SCORE_MINIMO_FOLHA)])
    log.append(f"Plano de Contas OK: {len(df)} contas ({n_a} analíticas · {n_s} sintéticas · "
               f"{ignorados} ignoradas · {n_folha} analíticas de folha)")
    return df


# ══════════════════════════════════════════════════════════════════════════
# KEYWORDS POR GRUPO (mantidas para filtros de CC)
# ══════════════════════════════════════════════════════════════════════════
KWORDS_DEBITO: dict[str, list[str]] = {
    "Custo Direto de Produção": [
        "MATERIA-PRIMA","MATERIAL APLICADO","MAO-DE-OBRA DIRETA",
        "SALARIOS E ORDENADOS CUSTOS","PRO-LABORE CUSTOS",
        "13 SALARIO CUSTOS","FERIAS CUSTOS","INSS CUSTOS","FGTS CUSTOS",
        "INDUSTRIALIZACAO","CUSTOS DIRETOS DE PRODUCAO",
    ],
    "Custo Direto de Serviços": [
        "CUSTOS DIRETOS DA PRODUCAO DE SERVICOS","MAO-DE-OBRA DIRETA",
        "SALARIOS E ORDENADOS","INSS","FGTS","FERIAS","13 SALARIO",
        "CUSTOS SERVICOS","VALE TRANSPORTE","ALIMENTACAO",
    ],
    "Custo Indireto de Produção": [
        "MAO-DE-OBRA INDIRETA","MATERIAIS DE CONSUMO INDIRETO",
        "ALUGUEIS","DEPRECIACOES","COMBUSTIVEIS","ENERGIA ELETRICA",
        "AGUA E ESGOTO","CUSTOS INDIRETOS","MANUTENCAO","LOCACAO",
    ],
    "Despesa Administrativa": [
        "DESPESAS COM PESSOAL","SALARIOS E ORDENADOS","PRO-LABORE",
        "PREMIOS E GRATIFICACOES","13 SALARIO","FERIAS","INSS","FGTS",
        "INDENIZACOES","ASSISTENCIA MEDICA","VALE TRANSPORTE",
        "PIS S/ FOLHA","ALIMENTACAO","VALE REFEICAO","HORAS EXTRAS",
        "PENSAO ALIMENTICIA","CESTA BASICA","COMISSOES",
        "ALUGUEIS","ENERGIA ELETRICA","AGUA E ESGOTO","TELEFONE",
        "SEGUROS","MATERIAL DE ESCRITORIO","DEPRECIACAO",
        "COMBUSTIVEIS","MATERIAIS DE CONSUMO","CONDOMINIOS",
        "DESPESAS GERAIS","FRETES E CARRETOS","MANUTENCAO",
        "VIAGENS","REFEICOES","SERVICOS TOMADOS","BENS DE PEQUENO VALOR",
    ],
    "Despesa com Vendas": [
        "DESPESAS COM VENDAS","COMISSOES SOBRE VENDAS","COMISSOES",
        "PROPAGANDA E PUBLICIDADE","BONIFICACAO","FRETES E CARRETOS",
        "MANUTENCAO DE VEICULOS","VIAGENS","ALUGUEIS",
    ],
    "Despesa Financeira": [
        "DESPESAS FINANCEIRAS","JUROS PASSIVOS","VARIACOES MONETARIAS",
        "VARIACOES CAMBIAIS","DESCONTOS FINANCEIROS CONCEDIDOS",
        "JUROS DE MORA","JUROS E COMISSOES BANCARIAS",
        "JUROS SOBRE EMPRESTIMOS","MULTAS PASSIVAS","TARIFA BANCARIA",
        "EMPRESTIMO / FINANCIAMENTO","IOF",
    ],
    "Despesa Não Operacional": [
        "DESPESAS NAO OPERACIONAIS","PERDAS NA ALIENACAO",
        "RESULTADO NEGATIVO NA ALIENACAO","RESULTADO NEGATIVO DE SINISTRO",
        "OUTRAS BAIXAS DO ATIVO","BAIXAS DE INVESTIMENTOS",
        "BAIXAS DE IMOBILIZADO","PROVISAO IRPJ","PROVISAO CSLL",
        "PERDAS POR FALTA NO INVENTARIO",
    ],
    "Encargo Patronal": [
        "DESPESAS COM PESSOAL","INSS","ENCARGOS SOCIAIS",
        "CONTRIBUICAO PREVIDENCIARIA","CONTRIBUICAO PATRONAL",
        "SALARIOS E ORDENADOS","CUSTOS DIRETOS",
        "MAO-DE-OBRA DIRETA","MAO-DE-OBRA INDIRETA",
        "INSS CUSTOS","FGTS CUSTOS",
    ],
}

KWORDS_CREDITO: dict[str, list[str]] = {
    "Custo Direto de Produção": [
        "SALARIOS E ORDENADOS A PAGAR","PRO-LABORE A PAGAR",
        "FERIAS A PAGAR","13 SALARIO A PAGAR","INSS A RECOLHER",
        "FGTS A RECOLHER","PROVISOES PARA FERIAS","PROVISOES PARA 13",
        "OBRIGACOES COM O PESSOAL","OBRIGACOES SOCIAIS","PROVISOES",
        "OBRIGACOES TRABALHISTA",
    ],
    "Custo Direto de Serviços": [
        "SALARIOS E ORDENADOS A PAGAR","FERIAS A PAGAR","13 SALARIO A PAGAR",
        "INSS A RECOLHER","FGTS A RECOLHER","PROVISOES PARA FERIAS",
        "PROVISOES PARA 13","OBRIGACOES COM O PESSOAL","OBRIGACOES SOCIAIS",
        "PROVISOES","OBRIGACOES TRABALHISTA",
    ],
    "Custo Indireto de Produção": [
        "SALARIOS E ORDENADOS A PAGAR","FERIAS A PAGAR","INSS A RECOLHER",
        "FGTS A RECOLHER","PROVISOES PARA FERIAS","PROVISOES PARA 13",
        "OBRIGACOES COM O PESSOAL","OBRIGACOES SOCIAIS","PROVISOES",
        "CONTAS A PAGAR","ALUGUEIS A PAGAR",
    ],
    "Despesa Administrativa": [
        "SALARIOS E ORDENADOS A PAGAR","PRO-LABORE A PAGAR",
        "GRATIFICACOES A PAGAR","FERIAS A PAGAR","RESCISOES A PAGAR",
        "13 SALARIO A PAGAR","PENSAO ALIMENTICIA A PAGAR",
        "COMISSOES A PAGAR","AUTONOMOS A PAGAR","INDENIZACOES A PAGAR",
        "INSS A RECOLHER","FGTS A RECOLHER","PIS S/ FOLHA A RECOLHER",
        "IRRF S/ FOLHA","CONTRIBUICOES SINDICAIS",
        "PROVISOES PARA FERIAS","PROVISOES PARA 13",
        "INSS SOBRE PROVISOES","FGTS SOBRE PROVISOES",
        "OBRIGACOES COM O PESSOAL","OBRIGACOES SOCIAIS","PROVISOES",
        "OBRIGACOES TRABALHISTA","CONTAS A PAGAR",
        "ENERGIA ELETRICA A PAGAR","TELEFONE A PAGAR",
        "ALUGUEIS A PAGAR","OUTRAS OBRIGACOES",
    ],
    "Despesa com Vendas": [
        "SALARIOS E ORDENADOS A PAGAR","FERIAS A PAGAR","13 SALARIO A PAGAR",
        "INSS A RECOLHER","FGTS A RECOLHER","PROVISOES PARA FERIAS",
        "PROVISOES PARA 13","OBRIGACOES COM O PESSOAL","OBRIGACOES SOCIAIS",
        "PROVISOES","OBRIGACOES TRABALHISTA","CONTAS A PAGAR","OUTRAS OBRIGACOES",
    ],
    "Despesa Financeira": [
        "CONTAS A PAGAR","OUTRAS OBRIGACOES",
        "IMPOSTOS E CONTRIBUICOES A RECOLHER",
    ],
    "Despesa Não Operacional": [
        "CONTAS A PAGAR","OUTRAS OBRIGACOES",
        "IMPOSTOS E CONTRIBUICOES A RECOLHER",
        "PROVISAO PARA IMPOSTO DE RENDA S/ LUCRO",
        "PROVISAO P/ CONTRIBUICAO SOCIAL S/ LUCRO",
        "IMPOSTO DE RENDA A RECOLHER","CONTRIBUICAO SOCIAL A RECOLHER",
    ],
    "Encargo Patronal": [
        "INSS A RECOLHER","OBRIGACOES SOCIAIS","OBRIGACOES TRABALHISTA",
        "OBRIGACOES TRABALHISTAS E PREVIDENCIARIA",
        "IMPOSTOS E CONTRIBUICOES A RECOLHER",
        "CONTRIBUICOES SINDICAIS","PIS S/ FOLHA A RECOLHER",
    ],
}

GRUPOS_LISTA = [
    "Despesa Administrativa","Despesa com Vendas","Despesa Financeira",
    "Despesa Não Operacional","Custo Direto de Produção",
    "Custo Direto de Serviços","Custo Indireto de Produção",
    "Encargo Patronal","Outro",
]


# ══════════════════════════════════════════════════════════════════════════
# FUNÇÕES DE FILTRO / FORMATAÇÃO
# ══════════════════════════════════════════════════════════════════════════
def _analiticas(df: pd.DataFrame) -> pd.DataFrame:
    return df[df["tipo"] == "A"].copy() if not df.empty else df

def _analiticas_folha(df: pd.DataFrame) -> pd.DataFrame:
    df_a = _analiticas(df)
    if df_a.empty: return df_a
    if "score_folha" in df_a.columns:
        f = df_a[df_a["score_folha"] >= SCORE_MINIMO_FOLHA]
        return f if not f.empty else df_a
    return df_a

def _fmt_opcoes(df_f: pd.DataFrame) -> list[str]:
    return [""] + [f"{r['reduzido']} - {r['nome_original']}" for _, r in df_f.iterrows()]

def _conta_bate(nome: str, kws: list[str]) -> bool:
    return any(_norm(k) in nome for k in kws)

def filtrar_contas_por_grupo(df_contas: pd.DataFrame, grupo: str,
                              aplicar_filtro_folha: bool = True) -> tuple[pd.DataFrame, pd.DataFrame]:
    df_a = _analiticas_folha(df_contas) if (aplicar_filtro_folha and grupo != "Outro") else _analiticas(df_contas)
    if df_a.empty: return pd.DataFrame(), pd.DataFrame()
    kw_d = KWORDS_DEBITO.get(grupo, [])
    kw_c = KWORDS_CREDITO.get(grupo, [])
    df_d = df_a[df_a["nome_conta"].apply(lambda n: _conta_bate(n, kw_d))] if (kw_d and grupo != "Outro") else df_a
    df_c = df_a[df_a["nome_conta"].apply(lambda n: _conta_bate(n, kw_c))] if (kw_c and grupo != "Outro") else df_a
    return df_d, df_c

def classificar_contas(df_contas: pd.DataFrame, grupo: str) -> tuple[list[str], list[str]]:
    df_d, df_c = filtrar_contas_por_grupo(df_contas, grupo)
    return _fmt_opcoes(df_d), _fmt_opcoes(df_c)

def extrair_codigo(opcao: str) -> str:
    if not opcao or " - " not in opcao: return opcao or ""
    return opcao.split(" - ")[0].strip()

def extrair_descricao(opcao: str) -> str:
    if not opcao or " - " not in opcao: return ""
    p = opcao.split(" - ", 1)
    return p[1].strip() if len(p) > 1 else ""

def buscar_conta_por_reduzido(df_contas: pd.DataFrame, reduzido: str) -> str:
    if df_contas is None or df_contas.empty or not reduzido: return ""
    r = df_contas[df_contas["reduzido"] == str(reduzido).strip()]
    return r.iloc[0]["nome_original"] if not r.empty else ""

def _idx(opcoes: list[str], valor: str) -> int:
    if not valor: return 0
    for i, op in enumerate(opcoes):
        if op.startswith(valor): return i
    return 0


# ══════════════════════════════════════════════════════════════════════════
# CLASSIFICAÇÃO AUTOMÁTICA
# ══════════════════════════════════════════════════════════════════════════
def classificar_todos_eventos(eventos: list, catalog: dict,
                               df_contas: pd.DataFrame | None, log: list) -> dict:
    resultado: dict[str, dict] = {}
    chaves = {(ev["cod"], ev["tipo_folha"]) for ev in eventos}
    for cod, tipo_folha in chaves:
        info  = catalog.get(cod, {})
        nome  = info.get("descricao", cod)
        tipo  = info.get("tipo", "Provento")
        cl    = classificar_rubrica_local(nome, tipo, tipo_folha)
        grupo = cl["grupo"]; conf = cl["confianca"]
        cd = cc = dd = dc = ""
        curto = False
        if df_contas is not None and not df_contas.empty:
            dep = gerar_depara_evento_conta(cod, nome, tipo, grupo, df_contas, tipo_folha)
            cd, cc, dd, dc = dep["conta_debito"], dep["conta_credito"], dep["desc_debito"], dep["desc_credito"]
            if not validar_par_contas(cd, cc):
                curto = True; cd = dd = ""
                log.append(f"⚠️ CURTO-CIRCUITO: {cod} ({nome}) — Débito resetado.")
        resultado[f"{cod}_{tipo_folha}"] = {
            "grupo": grupo, "confianca": conf,
            "conta_debito": cd, "conta_credito": cc,
            "desc_debito": dd, "desc_credito": dc,
            "tipo_folha": tipo_folha, "curto_circuito": curto,
        }
    n_a = sum(1 for v in resultado.values() if v["confianca"] == "alta")
    n_m = sum(1 for v in resultado.values() if v["confianca"] == "media")
    n_b = sum(1 for v in resultado.values() if v["confianca"] == "baixa")
    n_p = sum(1 for v in resultado.values() if v["grupo"] == "Encargo Patronal")
    n_c = sum(1 for v in resultado.values() if v.get("curto_circuito"))
    log.append(f"Classificação: {len(resultado)} rubricas → 🟢{n_a} alta · 🟡{n_m} média · 🔴{n_b} baixa · "
               f"🏛️{n_p} Patronal · ⚡{n_c} curto-circuito")
    return resultado


def classificar_eventos_por_grupo_cc(eventos: list, catalog: dict, grupo_cc: str,
                                      df_contas: pd.DataFrame, log: list) -> dict:
    resultado: dict[str, dict] = {}
    for ev in eventos:
        cod = ev["cod"]; tf = ev["tipo_folha"]
        info = catalog.get(cod, {}); nome = info.get("descricao", cod); tipo = info.get("tipo", "Provento")
        grupo_ef = "Encargo Patronal" if tf == "2" else grupo_cc
        dep = gerar_depara_evento_conta(cod, nome, tipo, grupo_ef, df_contas, tf)
        curto = False
        if not validar_par_contas(dep["conta_debito"], dep["conta_credito"]):
            curto = True; dep["conta_debito"] = dep["desc_debito"] = ""
            log.append(f"⚠️ CURTO-CIRCUITO (CC): {cod} ({nome}) — Débito resetado.")
        resultado[f"{cod}_{tf}"] = {
            "grupo": grupo_ef, "confianca": "manual" if tf != "2" else "alta",
            "conta_debito": dep["conta_debito"], "conta_credito": dep["conta_credito"],
            "desc_debito": dep["desc_debito"], "desc_credito": dep["desc_credito"],
            "tipo_folha": tf, "curto_circuito": curto,
        }
    log.append(f"CC '{grupo_cc}': {len(resultado)} evento(s).")
    return resultado


# ══════════════════════════════════════════════════════════════════════════
# PARSE TXT RUBRICAS
# ══════════════════════════════════════════════════════════════════════════
def parse_rubricas_txt(file_bytes: bytes, log: list) -> dict:
    catalog = {}
    TIPO_MAP = {"P": "Provento", "D": "Desconto", "I": "Informativa", "ID": "Inf. Dedutora"}
    try:
        texto = file_bytes.decode("latin-1", errors="replace")
    except Exception as e:
        log.append(f"ERRO ao decodificar rubricas.txt: {e}"); return catalog
    for raw in texto.splitlines():
        raw = raw.strip()
        if not raw: continue
        partes = raw.split("\t")
        if len(partes) < 5: continue
        cod = partes[2].strip(); descricao = partes[3].strip(); tipo_raw = partes[4].strip().upper()
        if not cod: continue
        tipo_norm = TIPO_MAP.get(tipo_raw)
        if tipo_norm is None: continue
        if cod not in catalog:
            catalog[cod] = {"tipo": tipo_norm, "descricao": descricao}
    log.append(f"rubricas.txt: {len(catalog)} código(s).")
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
    r"Provisão de Férias|Provisão de 13º|Provisão de 13o)$",
    re.IGNORECASE,
)
RE_CC    = re.compile(r"^Centro de Custo:\s*(\d+)\s+(.+)$", re.IGNORECASE)
RE_EVENT = re.compile(r"^\s*(\d+)\s+(.+)$")


def should_ignore(line: str) -> bool:
    return any(re.search(p, line, re.IGNORECASE) for p in IGNORE_PATTERNS)


def parse_nao_configurados_pdf(file_bytes: bytes, log: list) -> list:
    eventos, vistos = [], set()
    tipo_folha_atual = "1"; cc_cod_atual = cc_nome_atual = ""
    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text: continue
            for raw in text.splitlines():
                line = raw.strip()
                if not line: continue
                m = RE_SECAO.match(line)
                if m:
                    sec = m.group(1).strip()
                    for k, v in SECAO_TIPO_FOLHA.items():
                        if k.lower() in sec.lower():
                            tipo_folha_atual = v; break
                    continue
                m = RE_CC.match(line)
                if m:
                    cc_cod_atual = m.group(1).strip(); cc_nome_atual = m.group(2).strip(); continue
                if should_ignore(line): continue
                m = RE_EVENT.match(line)
                if m:
                    cod = m.group(1).strip(); desc = m.group(2).strip()
                    if not cod.isdigit(): continue
                    chave = (cod, tipo_folha_atual, cc_cod_atual)
                    if chave not in vistos:
                        vistos.add(chave)
                        eventos.append({
                            "cod": cod, "descricao_pdf": desc,
                            "tipo_folha": tipo_folha_atual,
                            "tipo_folha_desc": SECAO_TIPO_FOLHA_DESC.get(tipo_folha_atual, tipo_folha_atual),
                            "centro_custo_cod": cc_cod_atual,
                            "centro_custo_nome": cc_nome_atual,
                        })
    por_folha: dict[str, int] = {}
    for ev in eventos:
        d = ev["tipo_folha_desc"]; por_folha[d] = por_folha.get(d, 0) + 1
    log.append(f"PDF: {len(eventos)} evento(s). [{' · '.join(f'{k}: {v}' for k,v in por_folha.items())}]")
    return eventos


def get_centros_custo_unicos(eventos: list) -> list[tuple[str, str]]:
    vistos: dict[str, str] = {}
    for ev in eventos:
        cod = ev["centro_custo_cod"]; nome = ev["centro_custo_nome"]
        if cod and cod not in vistos: vistos[cod] = nome
    return list(vistos.items())


def get_eventos_por_cc(eventos: list, cc_cod: str) -> list:
    return [ev for ev in eventos if ev["centro_custo_cod"] == cc_cod]


# ══════════════════════════════════════════════════════════════════════════
# GERA EXCEL — ETAPA 1
# ══════════════════════════════════════════════════════════════════════════
def gerar_excel_configuracao(eventos, catalog, cod_empresa, log,
                              usa_separador=False, config_cc=None,
                              df_contas=None, classif_auto=None) -> bytes:
    linhas = []
    for ev in eventos:
        cod = ev["cod"]; tf = ev["tipo_folha"]
        info = catalog.get(cod, {}); tipo = info.get("tipo", "⚠️ Não encontrado")
        desc_rubr = info.get("descricao", ev["descricao_pdf"]); cc_cod = ev["centro_custo_cod"]
        cd = cc = hist = grupo = dd = dc = ""
        curto = False
        chave_auto = f"{cod}_{tf}"
        if usa_separador and config_cc and cc_cod in config_cc:
            cfg = config_cc[cc_cod]; hist = cfg.get("historico", "")
            grupo = "Encargo Patronal" if tf == "2" else cfg.get("grupo", "")
            if df_contas is not None and not df_contas.empty and grupo:
                dep = gerar_depara_evento_conta(cod, desc_rubr, tipo, grupo, df_contas, tf)
                cd, cc, dd, dc = dep["conta_debito"], dep["conta_credito"], dep["desc_debito"], dep["desc_credito"]
                if not validar_par_contas(cd, cc):
                    curto = True; cd = dd = ""
        elif classif_auto and chave_auto in classif_auto:
            auto = classif_auto[chave_auto]
            grupo = auto.get("grupo",""); cd = auto.get("conta_debito",""); cc = auto.get("conta_credito","")
            dd = auto.get("desc_debito",""); dc = auto.get("desc_credito",""); curto = auto.get("curto_circuito", False)
        obs = "⚡ CURTO-CIRCUITO: Débito resetado" if curto else ""
        linhas.append({
            "Cód. Empresa": cod_empresa, "Cód. Evento": cod,
            "Descrição (PDF)": ev["descricao_pdf"], "Descrição (Rubricas)": desc_rubr,
            "Tipo Rubrica": tipo, "Tipo Folha (Nº)": tf,
            "Tipo Folha": ev["tipo_folha_desc"], "Cód. Centro de Custo": cc_cod,
            "Centro de Custo": ev["centro_custo_nome"], "Grupo de Despesa": grupo,
            "Usa Separador": "Sim" if usa_separador else "Não",
            "Conta Débito": cd, "Descrição Conta Débito": dd,
            "Conta Crédito": cc, "Descrição Conta Crédito": dc,
            "Cód. Histórico": "", "Histórico": hist, "Observação": obs,
        })
    df = pd.DataFrame(linhas)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Configuração", index=False)
        _formatar_planilha_config(writer.sheets["Configuração"], df)
        if df_contas is not None and not df_contas.empty:
            df_exp = df_contas[["reduzido","classificacao","nome_original","tipo","score_folha"]].copy()
            df_exp.columns = ["Código Reduzido","Classificação","Nome da Conta","Tipo (S/A)","Score Folha"]
            df_exp.to_excel(writer, sheet_name="Plano de Contas", index=False)
            _formatar_planilha_saida(writer.sheets["Plano de Contas"])
    output.seek(0); log.append(f"Excel gerado: {len(linhas)} linha(s).")
    return output.read()


def _formatar_planilha_config(ws, df: pd.DataFrame):
    borda = Border(left=Side(style="thin"), right=Side(style="thin"),
                   top=Side(style="thin"), bottom=Side(style="thin"))
    larguras = {"A":12,"B":12,"C":38,"D":38,"E":16,"F":14,"G":20,"H":18,
                "I":22,"J":22,"K":14,"L":16,"M":38,"N":16,"O":38,"P":14,"Q":42,"R":30}
    for col, w in larguras.items(): ws.column_dimensions[col].width = w
    COLS_EDIT = {12,14,16,17,18}; COLS_AUTO = {13,15}; COLS_INFO = {10,11}
    TIPO_COR = {"Provento":"D4EDDA","Desconto":"F8D7DA","Informativa":"CCE5FF","Inf. Dedutora":"FFF3CD"}
    COR_EMP = "E8D5FF"; COR_CC = "FFD700"
    for ci, cell in enumerate(ws[1], 1):
        if ci in COLS_EDIT:   cell.fill = PatternFill("solid", fgColor="FF8000"); cell.font = Font(bold=True, color="FFFFFF", size=10)
        elif ci in COLS_AUTO: cell.fill = PatternFill("solid", fgColor="28A745"); cell.font = Font(bold=True, color="FFFFFF", size=10)
        elif ci in COLS_INFO: cell.fill = PatternFill("solid", fgColor="6C757D"); cell.font = Font(bold=True, color="FFFFFF", size=10)
        else:                 cell.fill = PatternFill("solid", fgColor="444444"); cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = borda
    ws.row_dimensions[1].height = 32
    for ri, row in enumerate(ws.iter_rows(min_row=2), 2):
        tv = ws.cell(ri, 5).value or ""; tf = ws.cell(ri, 6).value or ""; obs = ws.cell(ri, 18).value or ""
        if "CURTO-CIRCUITO" in str(obs).upper(): cor = COR_CC
        elif str(tf).strip() == "2":             cor = COR_EMP
        else:                                    cor = TIPO_COR.get(tv, "E2E3E5")
        for ci, cell in enumerate(row, 1):
            cell.border = borda; cell.alignment = Alignment(vertical="center", wrap_text=True)
            if ci in COLS_EDIT:   cell.fill = PatternFill("solid", fgColor="FFF8F0"); cell.font = Font(size=10)
            elif ci in COLS_AUTO: cell.fill = PatternFill("solid", fgColor="F0FFF4"); cell.font = Font(size=10, italic=True)
            else:                 cell.fill = PatternFill("solid", fgColor=cor);      cell.font = Font(size=10)
        ws.row_dimensions[ri].height = 18
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions


# ══════════════════════════════════════════════════════════════════════════
# GERA ARQUIVOS FINAIS — ETAPA 2
# ══════════════════════════════════════════════════════════════════════════
def ler_excel_preenchido(file_bytes: bytes, log: list) -> pd.DataFrame | None:
    try: xls = pd.ExcelFile(BytesIO(file_bytes))
    except Exception as e: log.append(f"ERRO ao abrir Excel: {e}"); return None
    sheet = None
    for c in ["Configuração","configuracao","Plan1","Sheet1"]:
        if c in xls.sheet_names: sheet = c; break
    if not sheet: sheet = xls.sheet_names[0]
    try:
        df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet, dtype=str)
    except Exception as e: log.append(f"ERRO ao ler aba '{sheet}': {e}"); return None
    df.columns = [str(c).strip() for c in df.columns]; df = df.dropna(how="all")
    log.append(f"Excel preenchido: {len(df)} linha(s) na aba '{sheet}'.")
    return df


def _limpa(val) -> str:
    if val is None: return ""
    s = str(val).strip()
    return "" if s.lower() in ("nan","none","") else s


def gerar_arquivos_finais(df: pd.DataFrame, cod_empresa_padrao: str, log: list) -> tuple[bytes, bytes]:
    col_map: dict[str, str] = {}
    for col in df.columns:
        cl = col.lower()
        if   "cód. empresa"           in cl or "cod. empresa"    in cl: col_map["empresa"]       = col
        elif "cód. evento"            in cl or "cod. evento"     in cl: col_map["seq"]           = col
        elif "tipo folha (nº)"        in cl or "tipo folha (n"   in cl: col_map["tipo"]          = col
        elif "descrição (rubricas)"   in cl:                             col_map["desc"]          = col
        elif "descrição (pdf)"        in cl and "desc" not in col_map:  col_map["desc"]          = col
        elif "cód. centro de custo"   in cl:                             col_map["cc"]            = col
        elif "conta débito"           in cl or "conta debito"    in cl: col_map["debito"]        = col
        elif "descrição conta débito" in cl or "descricao conta debito" in cl: col_map["desc_deb"] = col
        elif "conta crédito"          in cl or "conta credito"   in cl: col_map["credito"]       = col
        elif "descrição conta crédito" in cl or "descricao conta credito" in cl: col_map["desc_cred"] = col
        elif "cód. histórico"         in cl or "cod. historico"  in cl: col_map["historico"]     = col
        elif "histórico"              in cl and "cód" not in cl and "cod" not in cl: col_map["historico_texto"] = col
        elif "observação"             in cl:                             col_map["observacao"]    = col
        elif "usa separador"          in cl:                             col_map["usa_separador"] = col

    TIPO_COL = ("Tipo da Integração (1 - Folha mensal; 2 - Empresa; "
                "3 - Férias; 4 - Rescisao; 5 - Prov. Férias; 6 - Prov. 13)")
    linhas_evento, linhas_integra, linhas_integra_xls = [], [], []
    sem_conta = com_conta = cc_count = 0

    for _, row in df.iterrows():
        empresa     = _limpa(row.get(col_map.get("empresa",""),"")) or cod_empresa_padrao
        seq         = _limpa(row.get(col_map.get("seq",""),""))
        tipo        = _limpa(row.get(col_map.get("tipo",""),""))
        desc        = _limpa(row.get(col_map.get("desc",""),""))
        cc          = _limpa(row.get(col_map.get("cc",""),""))
        debito      = _limpa(row.get(col_map.get("debito",""),""))
        credito     = _limpa(row.get(col_map.get("credito",""),""))
        historico   = _limpa(row.get(col_map.get("historico",""),""))
        complemento = _limpa(row.get(col_map.get("historico_texto",""),""))
        usa_sep     = _limpa(row.get(col_map.get("usa_separador",""),""))
        if not seq: continue
        if debito and credito and debito == credito:
            cc_count += 1; debito = ""
            log.append(f"⚠️ CURTO-CIRCUITO exportação: Seq {seq} — Débito removido.")
        sep_val = "1" if usa_sep.lower() == "sim" else "0"
        if debito or credito: com_conta += 1
        else: sem_conta += 1
        linhas_evento.append({
            "Código da Empresa": empresa, "Centro de custo": cc,
            "Código Sequencial da Integração": seq, TIPO_COL: tipo,
            "Descrição": desc, "Código da Conta Débito": debito,
            "Código da Conta Crédito": credito, "Código do Histórico": historico,
            "Complemento": complemento,
        })
        linhas_integra.append({
            "Código da Empresa": empresa, "Separador": sep_val,
            "Código Sequencial da Integração": seq, TIPO_COL: tipo,
            "Código da Rúbrica Selecionada": seq,
        })
        linhas_integra_xls.append({
            "Código da Empresa": empresa, "Centro de Custo": cc,
            "Código Sequencial da Integração": seq, TIPO_COL: tipo,
            "Descrição": desc, "Código da Conta Crédito": credito,
            "Código da Conta Débito": debito, "Código do Histórico": historico,
        })
    log.append(f"Arquivos → Com conta: {com_conta} | Sem conta: {sem_conta} | CC corrigido: {cc_count}")

    buf_ev = BytesIO()
    with pd.ExcelWriter(buf_ev, engine="openpyxl") as w:
        pd.DataFrame(linhas_integra).to_excel(w, sheet_name="integra", index=False)
        pd.DataFrame(linhas_evento).to_excel(w,  sheet_name="evento",  index=False)
        for sn in ["integra","evento"]: _formatar_planilha_saida(w.sheets[sn])
    buf_ev.seek(0)

    buf_int = BytesIO()
    with pd.ExcelWriter(buf_int, engine="openpyxl") as w:
        pd.DataFrame(linhas_integra_xls).to_excel(w, sheet_name="Plan1", index=False)
        _formatar_planilha_saida(w.sheets["Plan1"])
    buf_int.seek(0)
    return buf_ev.read(), buf_int.read()


def _formatar_planilha_saida(ws):
    borda = Border(left=Side(style="thin"), right=Side(style="thin"),
                   top=Side(style="thin"), bottom=Side(style="thin"))
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor="444444")
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = borda
    ws.row_dimensions[1].height = 32
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = borda; cell.alignment = Alignment(vertical="center"); cell.font = Font(size=10)
    for col in ws.columns:
        ml = 0; cl = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value: ml = max(ml, len(str(cell.value)))
            except: pass
        ws.column_dimensions[cl].width = min(max(ml + 2, 10), 50)
    ws.freeze_panes = "A2"


# ══════════════════════════════════════════════════════════════════════════
# INTERFACE STREAMLIT
# ══════════════════════════════════════════════════════════════════════════
def main():
    st.set_page_config(page_title="Domínio | Integração Contábil",
                       page_icon="🟠", layout="wide", initial_sidebar_state="expanded")
    apply_tr_theme()

    st.markdown(f"""
        <div style="background:#444444; padding:24px 28px 18px 28px;
                    border-radius:8px; border-top:6px solid #FF8000; margin-bottom:28px;">
            <h2 style="color:#FF8000; margin:0;">
                📊 Integração Contábil — Domínio Sistemas &nbsp;|&nbsp; {VERSAO}
            </h2>
            <p style="color:#DDDDDD; margin:6px 0 0 0;">
                <b>Etapa 1:</b> PDF + TXT + Plano de Contas → classifica automaticamente → gera Excel.<br>
                <b>Etapa 2:</b> Excel preenchido → gera <b>evento exemplo.xlsx</b> e <b>integra exemplo.xlsx</b>.<br>
                <span style="color:#CC99FF;">🏛️ Tipo Folha "Empresa" → <b>Encargo Patronal</b> automático (DRE, não Ativo).</span><br>
                <span style="color:#FFD700;">⚡ Filtro por posição no plano: Débito=DRE · Crédito=Passivo · Consignado=Ativo/Passivo.</span>
            </p>
        </div>""", unsafe_allow_html=True)

    with st.sidebar:
        st.markdown("### ⚙️ Configurações")
        cod_empresa = st.text_input("Código da empresa", value="45")
        st.markdown("---")
        st.markdown("### 🎨 Legenda")
        for cor, txt in [("🟢","Provento"),("🔴","Desconto"),("🔵","Informativa"),
                         ("🟡","Inf. Dedutora"),("🟣","Encargo Patronal"),
                         ("🟠","Campos editáveis"),("🌿","Preenchimento auto"),("⭐","Curto-circuito")]:
            st.markdown(f"{cor} {txt}")
        st.markdown("---")
        st.markdown("### 📋 Regras V6.3")
        st.markdown(
            "**Filtro por posição no plano:**\n"
            "- Débito Provento/Encargo → DRE (começa 3 ou 4)\n"
            "- Crédito → Passivo (começa 2)\n"
            "- Consignado Crédito → qualquer (Ativo 11307 ou Passivo)\n\n"
            "**Descontos por tipo:**\n"
            "- INSS → INSS a Recolher\n"
            "- IRRF → IRRF S/ Folha\n"
            "- Consignado → Emprestimo/Consignado\n"
            "- VT/Plano/Adiant/Estouro → Sal. a Pagar\n\n"
            "**FGTS:** C=FGTS a Recolher | D varia por tipo"
        )
        st.markdown(f"---\n**Versão:** {VERSAO}")

    _defaults = {
        "log": [f"Pronto. Versão {VERSAO}"], "excel_config": None,
        "evento_xlsx": None, "integra_xls": None, "df_preview": None,
        "n_eventos": 0, "df_contas": None, "eventos_parsed": None,
        "catalog_parsed": None, "config_cc": {}, "classif_auto": {},
        "_contas_fid": None, "_contas_name": None,
    }
    for k, v in _defaults.items():
        if k not in st.session_state: st.session_state[k] = v

    # ── ETAPA 1 ────────────────────────────────────────────────────────────
    st.markdown("## 📋 Etapa 1 — Gerar Excel para Preenchimento")
    col1, col2, col3 = st.columns(3)
    with col1: pdf_file = st.file_uploader("1️⃣ PDF — Rubricas Não Configuradas", type=["pdf"], key="pdf_e1")
    with col2: txt_file = st.file_uploader("2️⃣ TXT — Rubricas (catálogo)", type=["txt"], key="txt_e1")
    with col3: contas_file = st.file_uploader("3️⃣ XLS/XLSX — Plano de Contas", type=["xls","xlsx"], key="contas_e1")

    if contas_file is not None:
        fid = getattr(contas_file, "file_id", id(contas_file))
        if st.session_state._contas_fid != fid:
            lt: list[str] = []
            df_c = parse_plano_contas(contas_file.read(), contas_file.name, lt)
            st.session_state.df_contas    = df_c if not df_c.empty else None
            st.session_state._contas_fid  = fid
            st.session_state._contas_name = contas_file.name
            st.session_state.config_cc    = {}
            st.session_state.log.extend(lt)
    else:
        if st.session_state._contas_fid is not None:
            st.session_state.df_contas = None; st.session_state._contas_fid = None
            st.session_state._contas_name = None; st.session_state.config_cc = {}

    df_pc = st.session_state.df_contas
    if df_pc is not None and not df_pc.empty:
        n_a = len(df_pc[df_pc["tipo"]=="A"]); n_s = len(df_pc[df_pc["tipo"]=="S"])
        n_f = len(df_pc[(df_pc["tipo"]=="A") & (df_pc.get("score_folha", pd.Series(dtype=int)) >= SCORE_MINIMO_FOLHA)]) if "score_folha" in df_pc.columns else 0

        # Conta DRE e Passivo para diagnóstico
        n_dre = len(df_pc[(df_pc["tipo"]=="A") & df_pc["classificacao"].apply(_e_conta_dre)])
        n_pas = len(df_pc[(df_pc["tipo"]=="A") & df_pc["classificacao"].apply(_e_conta_passivo)])

        st.success(f"✅ **{st.session_state._contas_name}**: {len(df_pc)} contas "
                   f"({n_a} analíticas · {n_s} sintéticas · **{n_f}** de folha · "
                   f"**{n_dre}** DRE · **{n_pas}** Passivo)")

        with st.expander("🔍 Preview das contas-chave detectadas (V6.3)", expanded=False):
            # Testa o motor com filtro de posição
            previews = [
                ("Proventos — Débito (DRE)", PROVENTO_DEBITO_POS, PROVENTO_DEBITO_NEG, "DRE"),
                ("Proventos — Crédito (Passivo)", PROVENTO_CREDITO_POS, PROVENTO_CREDITO_NEG, "PASSIVO"),
                ("INSS desconto — Crédito (Passivo)", DESCONTO_CRED_INSS_POS, DESCONTO_CRED_INSS_NEG, "PASSIVO"),
                ("IRRF — Crédito (Passivo)", DESCONTO_CRED_IRRF_POS, DESCONTO_CRED_IRRF_NEG, "PASSIVO"),
                ("Consignado — Crédito", CONSIGNADO_CREDITO_POS, CONSIGNADO_CREDITO_NEG, ""),
                ("Encargo Patronal — Débito (DRE)", INFORMATIVO_DEBITO_POS, INFORMATIVO_DEBITO_NEG, "DRE"),
                ("FGTS — Crédito (Passivo)", FGTS_CREDITO_POS, FGTS_CREDITO_NEG, "PASSIVO"),
                ("FGTS Férias — Débito (Passivo Provisão)", FGTS_DEBITO_PROVISAO_POS, FGTS_DEBITO_PROVISAO_NEG, "PASSIVO"),
                ("Encargo Patronal — Crédito (Passivo)", INFORMATIVO_CREDITO_POS, INFORMATIVO_CREDITO_NEG, "PASSIVO"),
            ]
            cols_p = st.columns(3)
            for i, (lbl, pos, neg, filtro) in enumerate(previews):
                with cols_p[i % 3]:
                    st.markdown(f"**{lbl}**")
                    cod_t, desc_t = _melhor_conta(df_pc, pos, neg, 30, filtro_posicao=filtro)
                    st.info(f"`{cod_t}` — {desc_t}" if cod_t else "⚠️ Não encontrada")
    elif contas_file is not None:
        st.error("❌ Não foi possível carregar o Plano de Contas.")

    st.markdown("---")
    st.markdown("### ⚙️ Configuração de Separador")
    usa_separador = st.radio("Os lançamentos usam separador por Centro de Custo?",
                              ["Não","Sim"], index=0, horizontal=True)
    usa_sep_bool = (usa_separador == "Sim")

    if usa_sep_bool and st.session_state.eventos_parsed:
        ccs = get_centros_custo_unicos(st.session_state.eventos_parsed)
        if ccs:
            st.markdown("#### 🏢 Grupo de Despesa por Centro de Custo")
            nao_cl = [f"CC {cc} — {nm}" for cc, nm in ccs
                      if not st.session_state.config_cc.get(cc, {}).get("grupo")]
            if nao_cl:
                with st.expander(f"⚠️ {len(nao_cl)} CC(s) sem grupo", expanded=True):
                    for item in nao_cl: st.markdown(f"- {item}")
            else:
                st.success("✅ Todos os CCs têm grupo definido!")
            st.markdown("---")
            for cc_cod, cc_nome in ccs:
                cfg_at = st.session_state.config_cc.get(cc_cod, {})
                grupo_ok = bool(cfg_at.get("grupo"))
                evs_cc = get_eventos_por_cc(st.session_state.eventos_parsed, cc_cod)
                n_emp = sum(1 for ev in evs_cc if ev["tipo_folha"] == "2")
                titulo = f"{'✅' if grupo_ok else '⚠️'} CC {cc_cod} — {cc_nome} ({len(evs_cc)} evento(s)"
                if n_emp > 0: titulo += f" · 🏛️ {n_emp} Encargo Patronal"
                titulo += ")"
                with st.expander(titulo, expanded=not grupo_ok):
                    if n_emp > 0:
                        st.info(f"🏛️ **{n_emp} evento(s)** do Tipo Folha **Empresa** → **Encargo Patronal** automático.")
                    gi = GRUPOS_LISTA.index(cfg_at.get("grupo","Despesa Administrativa")) if cfg_at.get("grupo") in GRUPOS_LISTA else 0
                    gs = st.selectbox("📂 Grupo de Despesa do CC", GRUPOS_LISTA, index=gi, key=f"g_{cc_cod}")
                    hs = st.text_input("📋 Histórico padrão", value=cfg_at.get("historico",""), key=f"h_{cc_cod}", placeholder="Ex: 001")
                    st.session_state.config_cc[cc_cod] = {"grupo": gs, "historico": hs}
    elif usa_sep_bool:
        st.info("⬆️ Faça upload do PDF e clique em **▶ Gerar Excel** para configurar os CCs.")

    st.markdown("---")
    col_b1, col_b2 = st.columns(2)
    with col_b1:
        gerar_excel = st.button("▶ Gerar Excel de Configuração",
                                 disabled=(pdf_file is None or txt_file is None),
                                 use_container_width=True, type="primary")
    with col_b2:
        if st.button("🗑 Limpar tudo", use_container_width=True):
            for k in ["log","excel_config","evento_xlsx","integra_xls","df_preview","n_eventos",
                      "df_contas","eventos_parsed","catalog_parsed","config_cc","classif_auto",
                      "_contas_fid","_contas_name"]:
                st.session_state[k] = (["Campos limpos."] if k == "log" else
                                        0 if k == "n_eventos" else
                                        {} if k in ("config_cc","classif_auto") else None)
            st.rerun()

    if gerar_excel and pdf_file and txt_file:
        log: list[str] = ["[Etapa 1] Iniciando..."]
        with st.spinner("Lendo rubricas.txt..."): catalog = parse_rubricas_txt(txt_file.read(), log)
        with st.spinner("Lendo PDF..."): eventos = parse_nao_configurados_pdf(pdf_file.read(), log)
        st.session_state.eventos_parsed = eventos; st.session_state.catalog_parsed = catalog
        with st.spinner("🔍 Classificando com regras V6.3..."):
            ca = classificar_todos_eventos(eventos, catalog, df_pc, log)
            st.session_state.classif_auto = ca
        if usa_sep_bool:
            for cc_cod, _ in get_centros_custo_unicos(eventos):
                if cc_cod not in st.session_state.config_cc:
                    evs_cc = [ev for ev in eventos if ev["centro_custo_cod"] == cc_cod and ev["tipo_folha"] != "2"]
                    grupos_cc = [ca.get(f"{ev['cod']}_{ev['tipo_folha']}", {}).get("grupo","Despesa Administrativa") for ev in evs_cc]
                    gd = max(set(grupos_cc), key=grupos_cc.count) if grupos_cc else "Despesa Administrativa"
                    st.session_state.config_cc[cc_cod] = {"grupo": gd, "historico": ""}
                    log.append(f"CC {cc_cod}: grupo sugerido → {gd}")
        if not eventos: log.append("AVISO: Nenhum evento encontrado.")
        else:
            with st.spinner("Gerando Excel..."):
                eb = gerar_excel_configuracao(eventos, catalog, cod_empresa, log,
                                              usa_separador=usa_sep_bool,
                                              config_cc=st.session_state.config_cc if usa_sep_bool else None,
                                              df_contas=df_pc, classif_auto=ca)
            st.session_state.excel_config = eb; st.session_state.n_eventos = len(eventos)
            linhas_prev = []
            for ev in eventos:
                cod_ev = ev["cod"]; tf = ev["tipo_folha"]
                info = catalog.get(cod_ev, {}); cc_cod = ev["centro_custo_cod"]
                chave = f"{cod_ev}_{tf}"
                if usa_sep_bool and cc_cod in st.session_state.config_cc:
                    cfg_cc = st.session_state.config_cc[cc_cod]
                    grupo = "Encargo Patronal" if tf == "2" else cfg_cc.get("grupo","")
                    if df_pc is not None and not df_pc.empty and grupo:
                        dep = gerar_depara_evento_conta(cod_ev, info.get("descricao", ev["descricao_pdf"]),
                                                        info.get("tipo","Provento"), grupo, df_pc, tf)
                        cd, cc2 = dep["conta_debito"], dep["conta_credito"]
                    else: cd = cc2 = ""
                    conf = "manual" if tf != "2" else "alta"
                else:
                    auto = ca.get(chave, {}); grupo = auto.get("grupo","—")
                    cd, cc2, conf = auto.get("conta_debito",""), auto.get("conta_credito",""), auto.get("confianca","")
                cc_alert = ""
                if cd and cc2 and cd == cc2: cc_alert = "⚡"; cd = ""
                ok = "✅" if (cd and cc2) else ("⚡" if cc_alert else "⚠️")
                linhas_prev.append({"Código": cod_ev, "Descrição": ev["descricao_pdf"],
                                    "Tipo": info.get("tipo","⚠️"), "Tipo Folha": ev["tipo_folha_desc"],
                                    "Centro Custo": ev["centro_custo_nome"], "Grupo": grupo,
                                    "Confiança": conf, "Conta Débito": cd, "Conta Crédito": cc2, "Status": ok})
            st.session_state.df_preview = pd.DataFrame(linhas_prev)
        st.session_state.log = log; st.rerun()

    if st.session_state.excel_config is not None:
        st.success(f"✅ Excel gerado — {st.session_state.n_eventos} evento(s)")
        st.download_button("⬇ Baixar Excel de Configuração", data=st.session_state.excel_config,
                           file_name="configuracao_rubricas_dominio.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True, type="primary")
        if st.session_state.df_preview is not None:
            df = st.session_state.df_preview
            total = len(df); p = len(df[df["Tipo"]=="Provento"]); d = len(df[df["Tipo"]=="Desconto"])
            i = len(df[df["Tipo"]=="Informativa"]); id_ = len(df[df["Tipo"]=="Inf. Dedutora"])
            nf = len(df[df["Tipo"].str.startswith("⚠️",na=False)])
            ok = len(df[df["Status"]=="✅"]) if "Status" in df.columns else 0
            nok = len(df[df["Status"]=="⚠️"]) if "Status" in df.columns else 0
            ncc = len(df[df["Status"]=="⚡"]) if "Status" in df.columns else 0
            np_ = len(df[df["Grupo"]=="Encargo Patronal"]) if "Grupo" in df.columns else 0
            cols_m = st.columns(9)
            for cm, lbl, val in zip(cols_m,
                ["📋 Total","🟢 Proventos","🔴 Descontos","🔵 Informativas","🟡 Inf.Ded.",
                 "⚠️ Tipo n/id","✅ Completos","⚡ Curto-circ.","🏛️ Patronal"],
                [total, p, d, i, id_, nf, ok, ncc, np_]):
                cm.metric(lbl, val)
            if ncc > 0:
                st.warning(f"⚡ **{ncc} lançamento(s)** com curto-circuito — Débito resetado.")
            if nok > 0:
                st.info(f"⚠️ **{nok} lançamento(s)** sem conta definida — verifique o plano de contas.")
            def hl(row):
                t = str(row.get("Tipo","")); tf = str(row.get("Tipo Folha","")); g = str(row.get("Grupo",""))
                s = str(row.get("Status",""))
                if s == "⚡":                              return ["background-color:#FFD700"]*len(row)
                if tf == "Empresa" or g == "Encargo Patronal": return ["background-color:#E8D5FF"]*len(row)
                if t == "Provento":                        return ["background-color:#d4edda"]*len(row)
                if t == "Desconto":                        return ["background-color:#f8d7da"]*len(row)
                if t == "Informativa":                     return ["background-color:#cce5ff"]*len(row)
                if t == "Inf. Dedutora":                   return ["background-color:#fff3cd"]*len(row)
                return ["background-color:#e2e3e5"]*len(row)
            st.dataframe(df.head(150).style.apply(hl, axis=1), use_container_width=True)

    st.markdown("---")
    st.markdown("## 📥 Etapa 2 — Importar Excel Preenchido → Gerar Arquivos Finais")
    st.markdown("1. Baixe o Excel da Etapa 1 · 2. Revise contas · 3. Faça upload e clique em **▶ Gerar**")
    excel_preenchido = st.file_uploader("4️⃣ Excel Preenchido (.xlsx)", type=["xlsx","xls"], key="excel_e2")
    col_b3, _ = st.columns(2)
    with col_b3:
        gerar_finais = st.button("▶ Gerar Arquivos Finais", disabled=(excel_preenchido is None),
                                  use_container_width=True, type="primary")
    if gerar_finais and excel_preenchido:
        log = list(st.session_state.log); log.append("[Etapa 2] Iniciando...")
        with st.spinner("Lendo Excel..."):
            df_p = ler_excel_preenchido(excel_preenchido.read(), log)
        if df_p is not None:
            with st.spinner("Gerando arquivos finais..."):
                ev_b, int_b = gerar_arquivos_finais(df_p, cod_empresa, log)
            st.session_state.evento_xlsx = ev_b; st.session_state.integra_xls = int_b
        st.session_state.log = log; st.rerun()

    if st.session_state.evento_xlsx is not None:
        st.success("✅ Arquivos finais gerados!")
        cd1, cd2 = st.columns(2)
        with cd1:
            st.download_button("⬇ Baixar evento exemplo.xlsx", data=st.session_state.evento_xlsx,
                               file_name="evento exemplo.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True, type="primary")
        with cd2:
            st.download_button("⬇ Baixar integra exemplo.xlsx", data=st.session_state.integra_xls,
                               file_name="integra exemplo.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True, type="primary")

    st.markdown("---")
    st.markdown("**Log de processamento**")
    log_texto = "\n".join(st.session_state.log)
    tem_erro = any(str(l).upper().startswith("ERRO") for l in st.session_state.log)
    tem_cc   = any("CURTO-CIRCUITO" in str(l).upper() for l in st.session_state.log)
    cor = "#D32F2F" if tem_erro else ("#FF8C00" if tem_cc else "#388E3C")
    st.markdown(f"""<div style="background:#FCFCFC; border:1px solid {cor}; border-radius:6px;
        padding:14px; font-family:Consolas,monospace; font-size:13px; white-space:pre-wrap;
        max-height:300px; overflow-y:auto; color:#1F1F1F;">{log_texto}</div>""",
        unsafe_allow_html=True)


if __name__ == "__main__":
    main()
