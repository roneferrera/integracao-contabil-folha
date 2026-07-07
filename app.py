# ============================================================
# app_integracao_dominio.py  –  Integração Contábil Domínio V4.5
# ============================================================

import streamlit as st
import pdfplumber
import pandas as pd
import re
import json
import time
from io import BytesIO
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

VERSAO = "V4.5"

# ══════════════════════════════════════════════════════════════════════════
# GEMINI — importação opcional
# ══════════════════════════════════════════════════════════════════════════
try:
    import google.generativeai as genai
    GEMINI_DISPONIVEL = True
except ImportError:
    GEMINI_DISPONIVEL = False

GEMINI_MODEL = "gemini-1.5-flash"

GRUPOS_VALIDOS = [
    "Custo Direto de Produção",
    "Custo Direto de Serviços",
    "Custo Indireto de Produção",
    "Despesa Administrativa",
    "Despesa com Vendas",
    "Despesa Financeira",
    "Despesa Não Operacional",
    "Outro",
]

PROMPT_SISTEMA_GEMINI = """
Você é um especialista em contabilidade brasileira com foco em integração
de folha de pagamento com plano de contas (sistema Domínio Sistemas).

Sua tarefa é analisar nomes de rubricas de folha de pagamento e classificá-las
em grupos de despesa para lançamentos contábeis.

GRUPOS DISPONÍVEIS:
- Custo Direto de Produção: mão de obra direta, salários de produção industrial
- Custo Direto de Serviços: mão de obra direta de prestação de serviços
- Custo Indireto de Produção: overhead de fábrica, manutenção, utilidades
- Despesa Administrativa: salários administrativos, encargos, benefícios admin
- Despesa com Vendas: comissões, salários de vendas, representação
- Despesa Financeira: juros, tarifas bancárias, IOF, variações cambiais
- Despesa Não Operacional: perdas, provisões de IR/CSLL, resultados eventuais
- Outro: quando não se enquadra nos grupos acima

REGRAS:
1. Analise o nome da rubrica e o tipo (Provento, Desconto, Informativa, Inf. Dedutora)
2. Retorne APENAS JSON válido, sem markdown ou explicações extras

FORMATO DE RESPOSTA (JSON):
{
  "grupo": "nome_do_grupo",
  "confianca": "alta|media|baixa",
  "motivo": "explicação breve em até 15 palavras"
}
"""


# ══════════════════════════════════════════════════════════════════════════
# FUNÇÕES GEMINI
# ══════════════════════════════════════════════════════════════════════════

def gemini_testar_conexao(api_key: str) -> bool:
    if not GEMINI_DISPONIVEL:
        return False
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel(GEMINI_MODEL)
        model.generate_content("ok")
        return True
    except Exception:
        return False


def gemini_classificar_rubrica(
    nome_rubrica: str,
    tipo_rubrica: str,
    api_key: str,
    max_retries: int = 3,
) -> dict:
    """Classifica uma rubrica de folha usando o Gemini."""
    if not GEMINI_DISPONIVEL or not api_key:
        return {
            "grupo": "Despesa Administrativa",
            "confianca": "baixa",
            "motivo": "Gemini não disponível",
            "fonte": "fallback",
        }

    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(
        model_name=GEMINI_MODEL,
        system_instruction=PROMPT_SISTEMA_GEMINI,
    )

    prompt = f"""
Classifique esta rubrica de folha de pagamento:

Nome da rubrica: {nome_rubrica}
Tipo: {tipo_rubrica}

Grupos válidos: {', '.join(GRUPOS_VALIDOS)}

Responda APENAS com JSON válido no formato especificado.
"""

    for tentativa in range(max_retries):
        try:
            resposta = model.generate_content(prompt)
            texto = resposta.text.strip()
            if "```json" in texto:
                texto = texto.split("```json")[1].split("```")[0].strip()
            elif "```" in texto:
                texto = texto.split("```")[1].split("```")[0].strip()
            resultado = json.loads(texto)
            if resultado.get("grupo") not in GRUPOS_VALIDOS:
                resultado["grupo"] = "Despesa Administrativa"
            resultado["fonte"] = "gemini"
            return resultado
        except json.JSONDecodeError:
            if tentativa < max_retries - 1:
                time.sleep(1)
                continue
            return {
                "grupo": "Despesa Administrativa",
                "confianca": "baixa",
                "motivo": "Erro ao parsear resposta",
                "fonte": "erro",
            }
        except Exception as e:
            erro_str = str(e)
            if "429" in erro_str or "quota" in erro_str.lower():
                time.sleep((tentativa + 1) * 10)
                continue
            return {
                "grupo": "Despesa Administrativa",
                "confianca": "baixa",
                "motivo": "Erro na API Gemini",
                "fonte": "erro",
                "erro": erro_str[:100],
            }

    return {
        "grupo": "Despesa Administrativa",
        "confianca": "baixa",
        "motivo": "Máximo de tentativas atingido",
        "fonte": "erro",
    }


def gemini_sugerir_contas_para_rubrica(
    nome_rubrica: str,
    tipo_rubrica: str,
    grupo_sugerido: str,
    df_contas: pd.DataFrame,
    api_key: str,
    top_n: int = 5,
) -> dict:
    """Sugere contas de débito e crédito para uma rubrica via Gemini."""
    if not GEMINI_DISPONIVEL or not api_key or df_contas is None or df_contas.empty:
        return {"erro": "Gemini não disponível ou plano de contas não carregado"}

    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(
        model_name=GEMINI_MODEL,
        system_instruction=PROMPT_SISTEMA_GEMINI,
    )

    col_nome = "nome_original" if "nome_original" in df_contas.columns else "nome_conta"
    df_analiticas = df_contas[df_contas["tipo"] == "A"].copy()
    if len(df_analiticas) > 400:
        df_analiticas = df_analiticas.head(400)

    lista_contas = "\n".join([
        f"{row['classificacao']} | {row[col_nome]}"
        for _, row in df_analiticas.iterrows()
    ])

    prompt = f"""
Rubrica de folha de pagamento:
- Nome: {nome_rubrica}
- Tipo: {tipo_rubrica}
- Grupo de despesa classificado: {grupo_sugerido}

Plano de contas disponível (formato: classificação | nome):
{lista_contas}

Tarefa:
1. Selecione as {top_n} melhores contas de DÉBITO (despesa/custo) para esta rubrica
2. Selecione as {top_n} melhores contas de CRÉDITO (passivo/obrigação) para esta rubrica

Responda APENAS com JSON válido:
{{
  "contas_debito": [
    {{"classificacao": "xxx", "nome": "nome_da_conta", "score": "alta|media|baixa"}}
  ],
  "contas_credito": [
    {{"classificacao": "xxx", "nome": "nome_da_conta", "score": "alta|media|baixa"}}
  ],
  "explicacao": "explicação em até 20 palavras"
}}
"""

    try:
        resposta = model.generate_content(prompt)
        texto = resposta.text.strip()
        if "```json" in texto:
            texto = texto.split("```json")[1].split("```")[0].strip()
        elif "```" in texto:
            texto = texto.split("```")[1].split("```")[0].strip()
        return json.loads(texto)
    except Exception as e:
        return {"erro": str(e)[:200]}


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
#
# Formato confirmado do arquivo Kaph Numeric:
#   Header linha 0:
#     col[0] = "Plano de Contas - Completo" (Empresa)
#     col[1] = "Unnamed: 1:Reduzido"
#     col[2] = "Unnamed: 2:Classificação"
#     col[3] = "Unnamed: 3:Tipo"          → S ou A
#     col[4] = "Unnamed: 4:Descriçao"
#   Dados:
#     col[0] = 1000003 (código empresa)
#     col[2] = 1, 11, 111, 11101, 11101000001 (classificação numérica)
#     col[3] = S ou A
#     col[4] = Nome da conta
#   Última linha: "Total de : 746" → ignorada
# ══════════════════════════════════════════════════════════════════════════
def parse_plano_contas(file_bytes: bytes, filename: str, log: list) -> pd.DataFrame:
    ext = filename.lower().rsplit(".", 1)[-1] if "." in filename else "xlsx"
    df_raw = None

    if ext == "xlsx":
        try:
            df_raw = pd.read_excel(
                BytesIO(file_bytes), sheet_name=0, header=0,
                dtype=str, engine="openpyxl"
            )
            log.append("Plano de Contas: lido como .xlsx (openpyxl).")
        except Exception as e:
            log.append(f"ERRO ao abrir .xlsx: {e}")
            return pd.DataFrame()
    else:
        for engine in ["xlrd", "openpyxl"]:
            try:
                df_raw = pd.read_excel(
                    BytesIO(file_bytes), sheet_name=0, header=0,
                    dtype=str, engine=engine
                )
                log.append(f"Plano de Contas: lido como .xls (engine={engine}).")
                break
            except Exception as e:
                log.append(f"  engine={engine} falhou: {e}")
                df_raw = None

    if df_raw is None:
        log.append(
            "ERRO: Não foi possível abrir o Plano de Contas. "
            "Abra no Excel, salve como .xlsx e tente novamente."
        )
        return pd.DataFrame()

    log.append(
        f"Plano de Contas: {len(df_raw)} linhas brutas, "
        f"{len(df_raw.columns)} colunas."
    )

    # ── Detecta índices das colunas por nome (flexível) ────────────────────
    cols = [str(c).strip() for c in df_raw.columns]
    log.append(f"Colunas detectadas: {cols[:6]}")

    idx_empresa       = None
    idx_classificacao = None
    idx_tipo          = None
    idx_descricao     = None

    for i, c in enumerate(cols):
        cl = c.lower()
        # Empresa: primeira coluna ou contém "plano de contas" ou "empresa"
        if idx_empresa is None and (
            i == 0
            or "plano de contas" in cl
            or (cl == "empresa" and i < 3)
        ):
            idx_empresa = i

        # Classificação: "classifica" no nome ou "unnamed: 2"
        if idx_classificacao is None and (
            "classifica" in cl
            or "unnamed: 2" in cl
        ):
            idx_classificacao = i

        # Tipo: "tipo" isolado (não ecf) ou "unnamed: 3"
        if idx_tipo is None and (
            "unnamed: 3" in cl
            or (cl == "tipo")
            or ("tipo" in cl and "ecf" not in cl and "dlpa" not in cl and i < 6)
        ):
            idx_tipo = i

        # Descrição: "descri" ou "unnamed: 4"
        if idx_descricao is None and (
            "descri" in cl
            or "unnamed: 4" in cl
        ):
            idx_descricao = i

    # Fallback para posições fixas do formato Kaph Numeric
    if idx_empresa is None:       idx_empresa = 0
    if idx_classificacao is None: idx_classificacao = 2
    if idx_tipo is None:          idx_tipo = 3
    if idx_descricao is None:     idx_descricao = 4

    log.append(
        f"Índices → empresa:{idx_empresa} | "
        f"classificação:{idx_classificacao} | "
        f"tipo:{idx_tipo} | descrição:{idx_descricao}"
    )

    max_idx = max(idx_empresa, idx_classificacao, idx_tipo, idx_descricao)
    if len(df_raw.columns) <= max_idx:
        log.append(
            f"ERRO: Esperado mínimo {max_idx + 1} colunas, "
            f"encontrado {len(df_raw.columns)}."
        )
        return pd.DataFrame()

    # ── Processa as linhas ─────────────────────────────────────────────────
    registros = []
    ignorados = 0

    for _, row in df_raw.iterrows():
        empresa_val = str(row.iloc[idx_empresa]).strip()
        classif     = str(row.iloc[idx_classificacao]).strip()
        tipo_raw    = str(row.iloc[idx_tipo]).strip().upper()
        nome        = str(row.iloc[idx_descricao]).strip()

        # Ignora linha de totalizador e linhas vazias
        if empresa_val.lower().startswith("total"):
            ignorados += 1
            continue
        if empresa_val.lower() in ("nan", "none", ""):
            ignorados += 1
            continue

        # Remove sufixo ".0" gerado pelo pandas em números float
        if classif.endswith(".0"):
            classif = classif[:-2]

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
            "nome_conta":    _norm(nome),
            "nome_original": nome,
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
        f"Plano de Contas OK: {len(df)} contas válidas "
        f"({n_a} analíticas · {n_s} sintéticas · {ignorados} ignoradas)."
    )

    if n_a == 0:
        log.append(
            "AVISO: Nenhuma conta analítica (tipo=A) encontrada. "
            "Verifique se a coluna Tipo contém 'A' e 'S'."
        )

    return df


# ══════════════════════════════════════════════════════════════════════════
# CLASSIFICADOR POR PALAVRAS-CHAVE (fallback)
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


KWORDS_DEBITO: dict[str, list[str]] = {
    "Custo Direto de Produção": [
        "MATERIA-PRIMA","MATERIAL APLICADO","MAO-DE-OBRA DIRETA",
        "SALARIOS E ORDENADOS CUSTOS","PRO-LABORE CUSTOS",
        "PREMIOS DE GRATIFICACOES CUSTOS","13 SALARIO CUSTOS",
        "FERIAS CUSTOS","INSS CUSTOS","FGTS CUSTOS",
        "INDENIZACOES E AVISO PREVIO CUSTOS",
        "ASSISTENCIA MEDICA E SOCIAL CUSTOS","PIS S/ FOLHA CUSTOS",
        "INDUSTRIALIZACAO","CUSTOS DIRETOS DE PRODUCAO",
        "MAO-DE-OBRA DIRETA",
    ],
    "Custo Direto de Serviços": [
        "CUSTOS DIRETOS DA PRODUCAO DE SERVICOS",
        "MAO-DE-OBRA DIRETA","SALARIOS E ORDENADOS","PRO-LABORE",
        "13 SALARIO","FERIAS","INSS","FGTS","INDENIZACOES",
        "ASSISTENCIA MEDICA","VALE TRANSPORTE","PIS S/ FOLHA",
        "ALIMENTACAO","VALE REFEICAO","HORAS EXTRAS",
        "CUSTOS SERVICOS","CUSTOS SERVICOS PRESTADOS",
    ],
    "Custo Indireto de Produção": [
        "MAO-DE-OBRA INDIRETA","MATERIAIS DE CONSUMO INDIRETO",
        "CUSTOS ADMINISTRATIVOS","ALUGUEIS E ARRENDAMENTOS",
        "DEPRECIACOES","AMORTIZACOES","COMBUSTIVEIS","ENERGIA ELETRICA",
        "AGUA E ESGOTO","CUSTOS INDIRETOS DE PRODUCAO",
        "CUSTOS SERVICOS TOMADOS","LOCACAO","CONDOMINIO",
        "MANUTENCAO","PEDAGIOS","ESTACIONAMENTO",
    ],
    "Despesa Administrativa": [
        "DESPESAS ADMINISTRATIVAS","DESPESAS COM PESSOAL",
        "SALARIOS E ORDENADOS","PRO-LABORE","PREMIOS E GRATIFICACOES",
        "13 SALARIO","FERIAS","INSS","FGTS",
        "INDENIZACOES E AVISO PREVIO","ASSISTENCIA MEDICA E SOCIAL",
        "VALE TRANSPORTE","PIS S/ FOLHA","DESPESAS COM ALIMENTACAO",
        "VALE REFEICAO","HORAS EXTRAS","PENSAO ALIMENTICIA",
        "ALIMENTACAO/ CESTA BASICA","COMISSOES",
        "ALUGUEIS DE IMOVEIS","ALUGUEIS DE MAQUINAS",
        "ARRENDAMENTO MERCANTIL","LEASING",
        "PIS","COFINS","IPTU","IPVA","TAXAS DIVERSAS","MULTAS DE MORA",
        "ENERGIA ELETRICA","AGUA E ESGOTO","TELEFONE",
        "DESPESAS POSTAIS","SEGUROS","MATERIAL DE ESCRITORIO",
        "MATERIAL DE HIGIENE","DEPRECIACAO","DEPRECIACOES E AMORTIZACOES",
        "REPRODUCOES","DESPESAS LEGAIS","LIVROS, JORNAIS",
        "COMBUSTIVEIS E LUBRIFICANTES","MATERIAIS DE CONSUMO",
        "CONDOMINIOS","GAS","BENS DE PEQUENO VALOR",
        "DESPESA SERVICOS","SERVICOS TOMADOS DE PJ",
        "DESPESAS GERAIS","FRETES E CARRETOS","MANUTENCAO DE VEICULOS",
        "VIAGENS","REFEICOES",
    ],
    "Despesa com Vendas": [
        "DESPESAS COM VENDAS","COMISSOES SOBRE VENDAS","COMISSOES",
        "PROPAGANDA E PUBLICIDADE","BONIFICACAO E/OU AMOSTRAS GRATIS",
        "DESPESAS COM ENTREGA","FRETES E CARRETOS",
        "MANUTENCAO DE VEICULOS","DESPESAS COM VIAGENS",
        "VIAGENS TERRESTRES","VIAGENS AEREAS","HOSPEDAGEM",
        "ALUGUEIS DE IMOVEIS","ALUGUEIS DE MAQUINAS",
    ],
    "Despesa Financeira": [
        "DESPESAS FINANCEIRAS","JUROS PASSIVOS",
        "VARIACOES MONETARIAS PASSIVAS","VARIACOES CAMBIAIS PASSIVAS",
        "DESCONTOS FINANCEIROS CONCEDIDOS","JUROS DE MORA",
        "JUROS E COMISSOES BANCARIAS",
        "JUROS SOBRE EMPRESTIMOS E FINANCIAMENTOS",
        "MULTAS PASSIVAS","TARIFA BANCARIA",
        "EMPRESTIMO / FINANCIAMENTO","IOF",
    ],
    "Despesa Não Operacional": [
        "DESPESAS NAO OPERACIONAIS","RESULTADOS NAO OPERACIONAIS",
        "PERDAS NA ALIENACAO","RESULTADO NEGATIVO NA ALIENACAO",
        "RESULTADO NEGATIVO DE SINISTRO","OUTRAS BAIXAS DO ATIVO",
        "BAIXAS DE INVESTIMENTOS","BAIXAS DE IMOBILIZADO",
        "BAIXAS DE ATIVO DIFERIDO","PROVISOES PARA PERDAS PERMANENTE",
        "PROVISAO IRPJ","PROVISAO CSLL",
        "PERDAS POR FALTA NO INVENTARIO",
    ],
}

KWORDS_CREDITO: dict[str, list[str]] = {
    "Custo Direto de Produção": [
        "SALARIOS E ORDENADOS A PAGAR","PRO-LABORE A PAGAR",
        "GRATIFICACOES A PAGAR","FERIAS A PAGAR","RESCISOES A PAGAR",
        "13 SALARIO A PAGAR","PENSAO ALIMENTICIA A PAGAR",
        "INDENIZACOES A PAGAR","INSS A RECOLHER","FGTS A RECOLHER",
        "PIS S/ FOLHA A RECOLHER","IRRF S/ FOLHA",
        "PROVISOES PARA FERIAS","PROVISOES PARA 13",
        "INSS SOBRE PROVISOES","FGTS SOBRE PROVISOES",
        "OBRIGACOES COM O PESSOAL","OBRIGACOES SOCIAIS","PROVISOES",
        "OBRIGACOES TRABALHISTA",
    ],
    "Custo Direto de Serviços": [
        "SALARIOS E ORDENADOS A PAGAR","PRO-LABORE A PAGAR",
        "FERIAS A PAGAR","13 SALARIO A PAGAR","INDENIZACOES A PAGAR",
        "INSS A RECOLHER","FGTS A RECOLHER","PIS S/ FOLHA A RECOLHER",
        "PROVISOES PARA FERIAS","PROVISOES PARA 13",
        "OBRIGACOES COM O PESSOAL","OBRIGACOES SOCIAIS","PROVISOES",
        "OBRIGACOES TRABALHISTA","FORNECEDORES NACIONAIS","CONTAS A PAGAR",
    ],
    "Custo Indireto de Produção": [
        "SALARIOS E ORDENADOS A PAGAR","FERIAS A PAGAR",
        "13 SALARIO A PAGAR","INSS A RECOLHER","FGTS A RECOLHER",
        "PROVISOES PARA FERIAS","PROVISOES PARA 13",
        "OBRIGACOES COM O PESSOAL","OBRIGACOES SOCIAIS","PROVISOES",
        "OBRIGACOES TRABALHISTA","FORNECEDORES NACIONAIS","CONTAS A PAGAR",
        "ALUGUEIS A PAGAR",
    ],
    "Despesa Administrativa": [
        "SALARIOS E ORDENADOS A PAGAR","PRO-LABORE A PAGAR",
        "GRATIFICACOES A PAGAR","FERIAS A PAGAR","RESCISOES A PAGAR",
        "13 SALARIO A PAGAR","PENSAO ALIMENTICIA A PAGAR",
        "PREMIOS E BONIFICACOES","COMISSOES A PAGAR","AUTONOMOS A PAGAR",
        "INDENIZACOES A PAGAR","INSS A RECOLHER","FGTS A RECOLHER",
        "PIS S/ FOLHA A RECOLHER","IRRF S/ FOLHA","CONTRIBUICOES SINDICAIS",
        "PROVISOES PARA FERIAS","PROVISOES PARA 13",
        "INSS SOBRE PROVISOES PARA FERIAS","INSS SOBRE PROVISOES PARA 13",
        "FGTS SOBRE PROVISOES PARA FERIAS","FGTS SOBRE PROVISOES PARA 13",
        "PIS SOBRE PROVISOES","OBRIGACOES COM O PESSOAL",
        "OBRIGACOES SOCIAIS","PROVISOES","OBRIGACOES TRABALHISTA",
        "ISS A RECOLHER","IRRF S/ NF A RECOLHER","INSS RETIDO A RECOLHER",
        "FORNECEDORES NACIONAIS","CONTAS A PAGAR",
        "HONORARIOS CONTABEIS","HONORARIOS JURIDICOS",
        "ENERGIA ELETRICA A PAGAR","TELEFONE A PAGAR",
        "ALUGUEIS A PAGAR","OUTRAS OBRIGACOES",
    ],
    "Despesa com Vendas": [
        "SALARIOS E ORDENADOS A PAGAR","FERIAS A PAGAR",
        "13 SALARIO A PAGAR","INSS A RECOLHER","FGTS A RECOLHER",
        "PROVISOES PARA FERIAS","PROVISOES PARA 13",
        "OBRIGACOES COM O PESSOAL","OBRIGACOES SOCIAIS","PROVISOES",
        "OBRIGACOES TRABALHISTA","FORNECEDORES NACIONAIS","CONTAS A PAGAR",
        "OUTRAS OBRIGACOES",
    ],
    "Despesa Financeira": [
        "CONTAS A PAGAR","OUTRAS OBRIGACOES",
        "BANCO DO BRASIL","BANCO ITAU UNIBANCO","BANCO BRADESCO",
        "BANCO SANTANDER","BANCO INTER","BANCO C6 BANK",
        "BANCO NU PAGAMENTOS","BANCO CORA","BANCO DAYCOVAL",
        "FINANCIAMENTO BANCO NACIONAL",
        "IMPOSTOS E CONTRIBUICOES A RECOLHER",
    ],
    "Despesa Não Operacional": [
        "CONTAS A PAGAR","OUTRAS OBRIGACOES",
        "IMPOSTOS E CONTRIBUICOES A RECOLHER",
        "PROVISAO PARA IMPOSTO DE RENDA S/ LUCRO",
        "PROVISAO P/ CONTRIBUICAO SOCIAL S/ LUCRO",
        "IMPOSTO DE RENDA A RECOLHER",
        "CONTRIBUICAO SOCIAL A RECOLHER",
    ],
}

GRUPOS_LISTA = list(KWORDS_DEBITO.keys()) + ["Outro"]


def _conta_bate(nome_conta_norm: str, keywords: list[str]) -> bool:
    for kw in keywords:
        if _norm(kw) in nome_conta_norm:
            return True
    return False


def _analiticas(df: pd.DataFrame) -> pd.DataFrame:
    return df[df["tipo"] == "A"].copy() if not df.empty else df


def _fmt_opcoes(df_f: pd.DataFrame) -> list[str]:
    col_nome = "nome_original" if "nome_original" in df_f.columns else "nome_conta"
    return [""] + [
        f"{r['classificacao']} - {r[col_nome]}"
        for _, r in df_f.iterrows()
    ]


def classificar_contas(
    df_contas: pd.DataFrame, grupo: str
) -> tuple[list[str], list[str]]:
    df_a = _analiticas(df_contas)
    if df_a.empty:
        return [""], [""]

    kw_d = KWORDS_DEBITO.get(grupo, [])
    kw_c = KWORDS_CREDITO.get(grupo, [])

    if kw_d and grupo != "Outro":
        mask_d = df_a["nome_conta"].apply(lambda n: _conta_bate(n, kw_d))
        df_d   = df_a[mask_d] if mask_d.any() else df_a
    else:
        df_d = df_a

    if kw_c and grupo != "Outro":
        mask_c = df_a["nome_conta"].apply(lambda n: _conta_bate(n, kw_c))
        df_c   = df_a[mask_c] if mask_c.any() else df_a
    else:
        df_c = df_a

    return _fmt_opcoes(df_d), _fmt_opcoes(df_c)


def sugerir_contas(df_contas: pd.DataFrame, grupo: str) -> dict:
    ops_d, ops_c = classificar_contas(df_contas, grupo)
    return {
        "ops_deb":       ops_d,
        "ops_cred":      ops_c,
        "conta_debito":  extrair_codigo(ops_d[1]) if len(ops_d) > 1 else "",
        "conta_credito": extrair_codigo(ops_c[1]) if len(ops_c) > 1 else "",
        "n_deb":  len(ops_d) - 1,
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
# Formato confirmado:
#   col[0]=empresa  col[1]=?  col[2]=código  col[3]=descrição  col[4]=tipo(P/D/I/ID)
# ══════════════════════════════════════════════════════════════════════════
def parse_rubricas_txt(file_bytes: bytes, log: list) -> dict:
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
                            "cod":               cod,
                            "descricao_pdf":     desc,
                            "tipo_folha":        tipo_folha_atual,
                            "tipo_folha_desc":   SECAO_TIPO_FOLHA_DESC.get(
                                tipo_folha_atual, tipo_folha_atual
                            ),
                            "centro_custo_cod":  cc_cod_atual,
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
    log.append(
        f"Excel gerado: {len(linhas)} linha(s). "
        f"Separador: {'Sim' if usa_separador else 'Não'}."
    )
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
        "Provento":      "D4EDDA",
        "Desconto":      "F8D7DA",
        "Informativa":   "CCE5FF",
        "Inf. Dedutora": "FFF3CD",
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
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
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


def gerar_arquivos_finais(
    df: pd.DataFrame,
    cod_empresa_padrao: str,
    log: list,
) -> tuple[bytes, bytes]:
    col_map: dict[str, str] = {}
    for col in df.columns:
        cl = col.lower()
        if   "cód. empresa"         in cl or "cod. empresa"    in cl: col_map["empresa"]       = col
        elif "cód. evento"          in cl or "cod. evento"     in cl: col_map["seq"]           = col
        elif "tipo folha (nº)"      in cl or "tipo folha (n"   in cl: col_map["tipo"]          = col
        elif "descrição (rubricas)" in cl:                             col_map["desc"]          = col
        elif "descrição (pdf)"      in cl and "desc" not in col_map:  col_map["desc"]          = col
        elif "cód. centro de custo" in cl:                             col_map["cc"]            = col
        elif "conta débito"         in cl or "conta debito"    in cl: col_map["debito"]        = col
        elif "conta crédito"        in cl or "conta credito"   in cl: col_map["credito"]       = col
        elif "cód. histórico"       in cl or "cod. historico"  in cl: col_map["historico"]     = col
        elif "histórico"            in cl and "cód" not in cl and "cod" not in cl:
            col_map["historico_texto"] = col
        elif "observação"           in cl:                             col_map["observacao"]    = col
        elif "usa separador"        in cl:                             col_map["usa_separador"] = col

    TIPO_COL = (
        "Tipo da Integração (1 - Folha mensal; 2 - Empresa; "
        "3 - Férias; 4 - Rescisao; 5 - Prov. Férias; 6 - Prov. 13)"
    )

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
            "Código da Empresa":             empresa,
            "Centro de custo":               cc,
            "Código Sequencial da Integração": seq,
            TIPO_COL:                        tipo,
            "Descrição":                     desc,
            "Código da Conta Débito":        debito,
            "Código da Conta Crédito":       credito,
            "Código do Histórico":           historico,
            "Complemento":                   complemento,
        })
        linhas_integra.append({
            "Código da Empresa":             empresa,
            "Separador":                     sep_val,
            "Código Sequencial da Integração": seq,
            TIPO_COL:                        tipo,
            "Código da Rúbrica Selecionada": seq,
        })
        linhas_integra_xls.append({
            "Código da Empresa":             empresa,
            "Centro de Custo":               cc,
            "Código Sequencial da Integração": seq,
            TIPO_COL:                        tipo,
            "Descrição":                     desc,
            "Código da Conta Crédito":       credito,
            "Código da Conta Débito":        debito,
            "Código do Histórico":           historico,
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
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
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
# SEÇÃO GEMINI — Classificação de Rubricas
# ══════════════════════════════════════════════════════════════════════════
def render_secao_gemini(
    catalog: dict,
    df_contas: pd.DataFrame | None,
    api_key: str,
):
    st.markdown("---")
    st.markdown("### 🤖 Classificação Automática de Rubricas com Gemini AI")

    if not GEMINI_DISPONIVEL:
        st.warning("⚠️ Instale: `pip install google-generativeai`")
        return

    if not api_key:
        st.info(
            "💡 Configure a **API Key do Gemini** na sidebar "
            "para usar esta funcionalidade."
        )
        return

    if not catalog:
        st.info("⬆️ Faça upload do arquivo TXT de rubricas para habilitar.")
        return

    tab1, tab2 = st.tabs(["🔍 Rubrica Individual", "📦 Classificação em Lote"])

    # ── Tab 1: Rubrica individual ──────────────────────────────────────────
    with tab1:
        st.markdown("#### Buscar contas para uma rubrica específica")

        opcoes_rubricas = [""] + [
            f"{cod} — {info['descricao']} ({info['tipo']})"
            for cod, info in sorted(
                catalog.items(),
                key=lambda x: int(x[0]) if x[0].isdigit() else 0,
            )
        ]

        col_r1, col_r2 = st.columns([3, 1])
        with col_r1:
            rubrica_sel = st.selectbox(
                "Selecione a rubrica",
                options=opcoes_rubricas,
                key="gemini_rubrica_sel",
            )
        with col_r2:
            st.markdown("<br>", unsafe_allow_html=True)
            buscar = st.button(
                "🔍 Classificar",
                key="btn_gemini_individual",
                disabled=not rubrica_sel,
                use_container_width=True,
            )

        if buscar and rubrica_sel:
            cod_sel   = rubrica_sel.split(" — ")[0].strip()
            info_sel  = catalog.get(cod_sel, {})
            nome_rubr = info_sel.get("descricao", rubrica_sel)
            tipo_rubr = info_sel.get("tipo", "Provento")

            with st.spinner(f"Classificando '{nome_rubr}'..."):
                resultado_grupo = gemini_classificar_rubrica(
                    nome_rubrica=nome_rubr,
                    tipo_rubrica=tipo_rubr,
                    api_key=api_key,
                )

            grupo     = resultado_grupo.get("grupo", "Despesa Administrativa")
            confianca = resultado_grupo.get("confianca", "media")
            motivo    = resultado_grupo.get("motivo", "")

            cor_conf = {
                "alta":  "#d4edda",
                "media": "#fff3cd",
                "baixa": "#f8d7da",
            }.get(confianca, "#e2e3e5")
            emoji_conf = {
                "alta": "🟢", "media": "🟡", "baixa": "🔴"
            }.get(confianca, "⚪")

            st.markdown(
                f"""
                <div style="background:{cor_conf}; border-radius:6px;
                            padding:12px 16px; margin:8px 0;">
                    <b>Grupo sugerido:</b> {grupo}
                    &nbsp; {emoji_conf} {confianca.upper()}<br>
                    <small>📝 {motivo}</small>
                </div>
                """,
                unsafe_allow_html=True,
            )

            if df_contas is not None and not df_contas.empty:
                with st.spinner("Buscando contas no plano de contas..."):
                    resultado_contas = gemini_sugerir_contas_para_rubrica(
                        nome_rubrica=nome_rubr,
                        tipo_rubrica=tipo_rubr,
                        grupo_sugerido=grupo,
                        df_contas=df_contas,
                        api_key=api_key,
                        top_n=5,
                    )

                if "erro" in resultado_contas:
                    st.error(f"❌ Erro ao buscar contas: {resultado_contas['erro']}")
                else:
                    if resultado_contas.get("explicacao"):
                        st.caption(f"📝 {resultado_contas['explicacao']}")

                    col_d, col_c = st.columns(2)
                    with col_d:
                        st.markdown("**💸 Contas de Débito sugeridas:**")
                        for conta in resultado_contas.get("contas_debito", []):
                            score = conta.get("score", "")
                            emoji = (
                                "🟢" if score == "alta"
                                else "🟡" if score == "media"
                                else "🔴"
                            )
                            st.markdown(
                                f"{emoji} `{conta.get('classificacao','?')}` "
                                f"— {conta.get('nome','?')}"
                            )
                    with col_c:
                        st.markdown("**💰 Contas de Crédito sugeridas:**")
                        for conta in resultado_contas.get("contas_credito", []):
                            score = conta.get("score", "")
                            emoji = (
                                "🟢" if score == "alta"
                                else "🟡" if score == "media"
                                else "🔴"
                            )
                            st.markdown(
                                f"{emoji} `{conta.get('classificacao','?')}` "
                                f"— {conta.get('nome','?')}"
                            )
            else:
                st.info(
                    "💡 Carregue o Plano de Contas para ver sugestões "
                    "de contas de débito e crédito."
                )

    # ── Tab 2: Lote ────────────────────────────────────────────────────────
    with tab2:
        st.markdown("#### Classificar todas as rubricas do catálogo")

        n_rubricas = len(catalog)
        tempo_est  = round(n_rubricas * 1.5 / 60, 1)

        st.info(
            f"📋 **{n_rubricas} rubricas** no catálogo. "
            f"Tempo estimado: **~{tempo_est} min** "
            f"(limite gratuito: 15 req/min)."
        )

        cache_key = "gemini_cache_rubricas"
        cache     = st.session_state.get(cache_key, {})
        ja_class  = len(cache)

        col_i1, col_i2, col_i3 = st.columns(3)
        col_i1.metric("Total de rubricas",        n_rubricas)
        col_i2.metric("Já classificadas (cache)",  ja_class)
        col_i3.metric("A classificar",             n_rubricas - ja_class)

        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            iniciar_lote = st.button(
                f"▶ Classificar {n_rubricas - ja_class} rubricas",
                key="btn_gemini_lote",
                type="primary",
                disabled=(n_rubricas - ja_class == 0),
            )
        with col_btn2:
            if st.button("🗑 Limpar cache", key="btn_limpar_cache_gemini"):
                st.session_state[cache_key] = {}
                st.rerun()

        if iniciar_lote:
            rubricas_para = {
                cod: info
                for cod, info in catalog.items()
                if cod not in cache
            }
            progress_bar = st.progress(0)
            status_text  = st.empty()
            total  = len(rubricas_para)
            erros  = 0

            for i, (cod, info) in enumerate(rubricas_para.items()):
                status_text.text(
                    f"Classificando {i+1}/{total}: "
                    f"{info['descricao'][:45]}..."
                )
                resultado = gemini_classificar_rubrica(
                    nome_rubrica=info["descricao"],
                    tipo_rubrica=info["tipo"],
                    api_key=api_key,
                )
                cache[cod] = {
                    "cod":       cod,
                    "descricao": info["descricao"],
                    "tipo":      info["tipo"],
                    "grupo":     resultado.get("grupo", "Despesa Administrativa"),
                    "confianca": resultado.get("confianca", "baixa"),
                    "motivo":    resultado.get("motivo", ""),
                    "fonte":     resultado.get("fonte", "gemini"),
                }
                if resultado.get("fonte") == "erro":
                    erros += 1
                progress_bar.progress((i + 1) / total)
                if i < total - 1:
                    time.sleep(1.5)

            st.session_state[cache_key] = cache
            progress_bar.progress(1.0)
            status_text.empty()
            st.success(
                f"✅ {total} rubricas classificadas! "
                f"{'⚠️ ' + str(erros) + ' com erro' if erros else ''}"
            )

        # Exibe resultados do cache
        if cache:
            st.markdown("#### 📊 Resultados da Classificação")
            df_res = pd.DataFrame(list(cache.values()))

            col_m1, col_m2, col_m3, col_m4 = st.columns(4)
            col_m1.metric("Total",    len(df_res))
            col_m2.metric("🟢 Alta",  len(df_res[df_res["confianca"] == "alta"]))
            col_m3.metric("🟡 Média", len(df_res[df_res["confianca"] == "media"]))
            col_m4.metric("🔴 Baixa", len(df_res[df_res["confianca"] == "baixa"]))

            if not df_res.empty:
                st.bar_chart(df_res["grupo"].value_counts())

            grupos_presentes = ["Todos"] + sorted(df_res["grupo"].unique().tolist())
            grupo_filtro = st.selectbox(
                "Filtrar por grupo:",
                grupos_presentes,
                key="filtro_grupo_gemini",
            )
            df_exibir = (
                df_res if grupo_filtro == "Todos"
                else df_res[df_res["grupo"] == grupo_filtro]
            )

            cols_exibir = [
                c for c in ["cod","descricao","tipo","grupo","confianca","motivo"]
                if c in df_exibir.columns
            ]

            def highlight_conf(row):
                cores = {
                    "alta":  "background-color: #d4edda",
                    "media": "background-color: #fff3cd",
                    "baixa": "background-color: #f8d7da",
                }
                return [cores.get(row.get("confianca",""), "")] * len(row)

            st.dataframe(
                df_exibir[cols_exibir].rename(columns={
                    "cod":       "Código",
                    "descricao": "Descrição",
                    "tipo":      "Tipo",
                    "grupo":     "Grupo Sugerido",
                    "confianca": "Confiança",
                    "motivo":    "Motivo",
                }).style.apply(highlight_conf, axis=1),
                use_container_width=True,
                height=400,
            )

            csv = df_exibir.to_csv(index=False).encode("utf-8")
            st.download_button(
                "⬇ Baixar classificações (CSV)",
                data=csv,
                file_name="classificacoes_gemini_rubricas.csv",
                mime="text/csv",
            )


# ══════════════════════════════════════════════════════════════════════════
# INTERFACE STREAMLIT — MAIN
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
                    border-radius:8px; border-top:6px solid #FF8000;
                    margin-bottom:28px;">
            <h2 style="color:#FF8000; margin:0;">
                📊 Integração Contábil — Domínio Sistemas
                &nbsp;|&nbsp; {VERSAO}
            </h2>
            <p style="color:#DDDDDD; margin:6px 0 0 0;">
                <b>Etapa 1:</b> PDF + TXT + Plano de Contas
                → classifica automaticamente → gera Excel.<br>
                <b>Etapa 2:</b> Excel preenchido
                → gera <b>evento exemplo.xlsx</b>
                e <b>integra exemplo.xlsx</b>.<br>
                <span style="color:#FFB74D;">
                    🤖 Classificação automática de rubricas com Gemini AI
                </span>
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
        st.markdown("### 🤖 Gemini AI")
        gemini_api_key = ""

        if not GEMINI_DISPONIVEL:
            st.error("Instale: `pip install google-generativeai`")
        else:
            try:
                gemini_api_key = st.secrets["GEMINI_API_KEY"]
                st.success("✅ API Key carregada dos secrets")
            except Exception:
                gemini_api_key = st.text_input(
                    "API Key Gemini",
                    type="password",
                    placeholder="AIza...",
                    help="Obtenha gratuitamente em: aistudio.google.com",
                    key="gemini_key_input",
                )

            if gemini_api_key:
                if st.button("🔌 Testar conexão", key="test_gemini"):
                    with st.spinner("Testando..."):
                        ok = gemini_testar_conexao(gemini_api_key)
                    if ok:
                        st.success("✅ Conectado!")
                        st.session_state["gemini_api_key_validada"] = gemini_api_key
                    else:
                        st.error("❌ Falha na conexão")

                if st.session_state.get("gemini_api_key_validada") == gemini_api_key:
                    st.markdown("🟢 **Conectado** | Flash 1.5 Free")
                    st.caption("15 req/min · 1M tokens/dia")

        st.markdown("---")
        st.markdown("### 🎨 Legenda de Tipos")
        st.markdown("🟢 Verde → Provento")
        st.markdown("🔴 Vermelho → Desconto")
        st.markdown("🔵 Azul → Informativa")
        st.markdown("🟡 Amarelo → Inf. Dedutora")
        st.markdown("🟠 Laranja → Campos a preencher")
        st.markdown("---")
        st.markdown(f"**Versão:** {VERSAO}")

    api_key_ativa = st.session_state.get(
        "gemini_api_key_validada", gemini_api_key
    )

    # ── Session state ──────────────────────────────────────────────────────
    _defaults = {
        "log":                     [f"Pronto. Versão {VERSAO}"],
        "excel_config":            None,
        "evento_xlsx":             None,
        "integra_xls":             None,
        "df_preview":              None,
        "n_eventos":               0,
        "df_contas":               None,
        "eventos_parsed":          None,
        "catalog_parsed":          None,
        "config_cc":               {},
        "_contas_fid":             None,
        "_contas_name":            None,
        "gemini_api_key_validada": "",
        "gemini_cache_rubricas":   {},
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
        pdf_file = st.file_uploader(
            "1️⃣ PDF — Rubricas/Itens Não Configurados",
            type=["pdf"], key="pdf_etapa1",
        )
    with col2:
        txt_file = st.file_uploader(
            "2️⃣ TXT — Rubricas (catálogo de tipos)",
            type=["txt"], key="txt_etapa1",
        )
    with col3:
        contas_file = st.file_uploader(
            "3️⃣ XLS/XLSX — Plano de Contas (opcional)",
            type=["xls", "xlsx"], key="contas_etapa1",
            help="Arquivo: Exportação Plano de Contas - Completo (Kaph Numeric)",
        )

    # ── Carrega plano de contas ────────────────────────────────────────────
    if contas_file is not None:
        fid   = getattr(contas_file, "file_id", id(contas_file))
        fname = contas_file.name
        if st.session_state._contas_fid != fid:
            log_tmp: list[str] = []
            raw_bytes = contas_file.read()
            df_c = parse_plano_contas(raw_bytes, fname, log_tmp)
            st.session_state.df_contas    = df_c if not df_c.empty else None
            st.session_state._contas_fid  = fid
            st.session_state._contas_name = fname
            st.session_state.config_cc    = {}
            st.session_state.log.extend(log_tmp)
    else:
        if st.session_state._contas_fid is not None:
            st.session_state.df_contas    = None
            st.session_state._contas_fid  = None
            st.session_state._contas_name = None
            st.session_state.config_cc    = {}

    df_pc = st.session_state.df_contas

    if df_pc is not None and not df_pc.empty:
        n_a = len(df_pc[df_pc["tipo"] == "A"])
        n_s = len(df_pc[df_pc["tipo"] == "S"])
        st.success(
            f"✅ Plano de Contas **{st.session_state._contas_name}** carregado: "
            f"**{len(df_pc)}** contas ({n_a} analíticas · {n_s} sintéticas)"
        )
        with st.expander("🔍 Ver amostra das contas analíticas", expanded=False):
            col_nome = (
                "nome_original" if "nome_original" in df_pc.columns
                else "nome_conta"
            )
            df_am = (
                df_pc[df_pc["tipo"] == "A"][["classificacao", col_nome]]
                .head(30)
            )
            df_am.columns = ["Classificação", "Nome da Conta"]
            st.dataframe(df_am, use_container_width=True)
            st.caption(f"Exibindo 30 de {n_a} contas analíticas.")
    elif contas_file is not None:
        st.error(
            "❌ Não foi possível carregar o Plano de Contas. "
            "Verifique o log abaixo. Se o arquivo for .xls, "
            "abra no Excel → Salvar Como → .xlsx e tente novamente."
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
                            st.info(
                                "💡 Carregue o Plano de Contas "
                                "para classificação automática."
                            )
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
                            ops_deb = ops_cred = [""]
                            n_deb = n_cred = 0

                        col_m1, col_m2, col_m3 = st.columns(3)
                        col_m1.metric("Contas Débito encontradas",  n_deb)
                        col_m2.metric("Contas Crédito encontradas", n_cred)
                        col_m3.metric(
                            "Status",
                            "✅ OK" if (n_deb > 0 and n_cred > 0) else "⚠️ Verificar",
                        )

                        if df_pc is not None and n_deb == 0:
                            st.warning(
                                f"⚠️ Nenhuma conta de Débito para **{grupo_sel}**."
                            )
                        if df_pc is not None and n_cred == 0:
                            st.warning(
                                f"⚠️ Nenhuma conta de Crédito para **{grupo_sel}**."
                            )
                        if df_pc is None:
                            st.info(
                                "💡 Carregue o Plano de Contas "
                                "para sugestões automáticas."
                            )

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
                            col_n = (
                                "nome_original"
                                if "nome_original" in df_pc.columns
                                else "nome_conta"
                            )
                            r = df_pc[df_pc["classificacao"] == cod]
                            return r.iloc[0][col_n] if not r.empty else cod

                        if deb_cod or cred_cod:
                            cor = "#e8f5e9" if (deb_cod and cred_cod) else "#fff3e0"
                            brd = "#4caf50" if (deb_cod and cred_cod) else "#FF8000"
                            st.markdown(
                                f"""
                                <div style="background:{cor};
                                            border-left:4px solid {brd};
                                            padding:8px 12px; border-radius:4px;
                                            margin-top:6px; font-size:13px;">
                                    <b>D:</b>
                                    <code>{deb_cod or '—'}</code>
                                    {_nome_conta(deb_cod)}
                                    &nbsp;&nbsp;
                                    <b>C:</b>
                                    <code>{cred_cod or '—'}</code>
                                    {_nome_conta(cred_cod)}
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
                st.warning(
                    "Nenhum Centro de Custo encontrado. "
                    "Processe o PDF primeiro."
                )
        else:
            st.info(
                "⬆️ Faça upload do PDF e clique em "
                "**▶ Gerar Excel** para configurar os CCs."
            )

    st.markdown("---")

    # ── Botões ────────────────────────────────────────────────────────────
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        gerar_excel = st.button(
            "▶ Gerar Excel de Configuração",
            disabled=(pdf_file is None or txt_file is None),
            use_container_width=True,
            type="primary",
        )
    with col_btn2:
        if st.button("🗑 Limpar tudo", use_container_width=True):
            for k in [
                "log","excel_config","evento_xlsx","integra_xls",
                "df_preview","n_eventos","df_contas","eventos_parsed",
                "catalog_parsed","config_cc","_contas_fid","_contas_name",
                "gemini_cache_rubricas",
            ]:
                if k == "log":
                    st.session_state[k] = ["Campos limpos."]
                elif k == "n_eventos":
                    st.session_state[k] = 0
                elif k in ("config_cc", "gemini_cache_rubricas"):
                    st.session_state[k] = {}
                else:
                    st.session_state[k] = None
            st.rerun()

    if gerar_excel and pdf_file and txt_file:
        log: list[str] = ["[Etapa 1] Iniciando..."]
        with st.spinner("Lendo rubricas.txt..."):
            catalog = parse_rubricas_txt(txt_file.read(), log)
        with st.spinner("Lendo PDF..."):
            eventos = parse_nao_configurados_pdf(pdf_file.read(), log)

        st.session_state.eventos_parsed = eventos
        st.session_state.catalog_parsed = catalog

        if df_pc is not None and not df_pc.empty and usa_sep_bool:
            ccs_novos = get_centros_custo_unicos(eventos)
            for cc_cod, _ in ccs_novos:
                if (
                    cc_cod not in st.session_state.config_cc
                    or not st.session_state.config_cc[cc_cod].get("conta_debito")
                ):
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
                cfg_cc = (
                    st.session_state.config_cc.get(cc_cod, {})
                    if usa_sep_bool else {}
                )
                ok = bool(
                    cfg_cc.get("conta_debito") and cfg_cc.get("conta_credito")
                )
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
        st.success(
            f"✅ Excel gerado — {st.session_state.n_eventos} evento(s)"
        )
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
            p   = len(df[df["Tipo"] == "Provento"])
            d   = len(df[df["Tipo"] == "Desconto"])
            i   = len(df[df["Tipo"] == "Informativa"])
            id_ = len(df[df["Tipo"] == "Inf. Dedutora"])
            nf  = len(df[df["Tipo"].str.startswith("⚠️", na=False)])
            ok  = (
                len(df[df["Classif."] == "✅"])
                if "Classif." in df.columns else 0
            )
            nok = (
                len(df[df["Classif."] == "⚠️"])
                if "Classif." in df.columns else 0
            )

            cols_m = st.columns(8)
            for col_m, lbl, val in zip(cols_m, [
                "📋 Total","🟢 Proventos","🔴 Descontos","🔵 Informativas",
                "🟡 Inf.Ded.","⚠️ Tipo n/id","✅ Classif.","⚠️ Sem conta",
            ], [total, p, d, i, id_, nf, ok, nok]):
                col_m.metric(lbl, val)

            if nok > 0 and usa_sep_bool and "Classif." in df.columns:
                df_nok = df[df["Classif."] == "⚠️"][
                    ["Código","Descrição","Centro Custo",
                     "Conta Débito","Conta Crédito"]
                ]
                with st.expander(
                    f"⚠️ {nok} evento(s) sem classificação completa",
                    expanded=True,
                ):
                    st.warning(
                        "Ajuste os CCs acima ou preencha manualmente no Excel."
                    )
                    st.dataframe(df_nok, use_container_width=True)

            def hl(row):
                t = str(row.get("Tipo", ""))
                if t == "Provento":      return ["background-color:#d4edda"] * len(row)
                if t == "Desconto":      return ["background-color:#f8d7da"] * len(row)
                if t == "Informativa":   return ["background-color:#cce5ff"] * len(row)
                if t == "Inf. Dedutora": return ["background-color:#fff3cd"] * len(row)
                return ["background-color:#e2e3e5"] * len(row)

            st.dataframe(
                df.head(100).style.apply(hl, axis=1),
                use_container_width=True,
            )

    # ── Seção Gemini ───────────────────────────────────────────────────────
    catalog_atual = st.session_state.get("catalog_parsed") or {}
    if not catalog_atual and txt_file is not None:
        try:
            log_tmp = []
            catalog_atual = parse_rubricas_txt(txt_file.read(), log_tmp)
            st.session_state.catalog_parsed = catalog_atual
        except Exception:
            catalog_atual = {}

    if catalog_atual:
        render_secao_gemini(
            catalog=catalog_atual,
            df_contas=df_pc,
            api_key=api_key_ativa,
        )

    st.markdown("---")

    # ══════════════════════════════════════════════════════════════════════
    # ETAPA 2
    # ══════════════════════════════════════════════════════════════════════
    st.markdown(
        "## 📥 Etapa 2 — Importar Excel Preenchido → Gerar Arquivos Finais"
    )
    st.markdown(
        "1. Baixe o Excel da Etapa 1 · "
        "2. Preencha Conta Débito e Conta Crédito · "
        "3. Faça upload e clique em **▶ Gerar**"
    )

    excel_preenchido = st.file_uploader(
        "4️⃣ Excel Preenchido (.xlsx)",
        type=["xlsx", "xls"],
        key="excel_etapa2",
    )
    col_btn3, _ = st.columns(2)
    with col_btn3:
        gerar_finais = st.button(
            "▶ Gerar Arquivos Finais",
            disabled=(excel_preenchido is None),
            use_container_width=True,
            type="primary",
        )

    if gerar_finais and excel_preenchido:
        log = list(st.session_state.log)
        log.append("[Etapa 2] Iniciando...")
        with st.spinner("Lendo Excel preenchido..."):
            df_preen = ler_excel_preenchido(excel_preenchido.read(), log)
        if df_preen is not None:
            with st.spinner("Gerando arquivos finais..."):
                evento_bytes, integra_bytes = gerar_arquivos_finais(
                    df_preen, cod_empresa, log
                )
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

    # ── Log ────────────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("**Log de processamento**")
    log_texto = "\n".join(st.session_state.log)
    tem_erro  = any(
        str(l).upper().startswith("ERRO")
        for l in st.session_state.log
    )
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
