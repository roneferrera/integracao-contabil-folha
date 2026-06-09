# ============================================================
# app_integracao_dominio.py  –  Integração Contábil Domínio V2.0
# Entradas:
#   1. RubricasItens não Configurados.pdf  → eventos sem config contábil
#   2. Rubricas.pdf                        → catálogo de tipos de rubrica
#   3. evento exemplo.xlsx                 → contas contábeis já configuradas
# Saída:
#   arquivo TXT no formato Domínio Separador
# ============================================================

import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

VERSAO = "V2.0"

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
        hr { border-color: #FF8000; }
        [data-testid="metric-container"] {
            background-color: #E9E9E9; border-left: 4px solid #FF8000;
            border-radius: 4px; padding: 10px;
        }
        .instrucoes-box {
            background-color: #E9E9E9; border-left: 4px solid #FF8000;
            border-radius: 4px; padding: 16px 20px; margin: 12px 0;
            color: #444444;
        }
        .instrucoes-box h4 { color: #FF8000; margin-top: 14px; margin-bottom: 6px; }
        .instrucoes-box h4:first-child { margin-top: 0; }
        </style>
    """, unsafe_allow_html=True)


# ==============================
# LINHAS A IGNORAR NO PDF DE RUBRICAS
# ==============================
IGNORE_PATTERNS_RUBRICAS = [
    r"^EMPRESA PADR",
    r"^RUBRICAS",
    r"^Emiss",
    r"^Hora:",
    r"^Pág",
    r"^Cód\.",
    r"^\s*$",
    r"^[A-Z]\.\s",       # letras de acumuladores ex: "A. I.R.R.F"
    r"^Soma na base",
]

def should_ignore_rubricas(line: str) -> bool:
    for pat in IGNORE_PATTERNS_RUBRICAS:
        if re.search(pat, line, re.IGNORECASE):
            return True
    return False


# ==============================
# LINHAS A IGNORAR NO PDF DE ITENS NÃO CONFIGURADOS
# ==============================
IGNORE_PATTERNS_NAO_CONFIG = [
    r"^RELAÇÃO DE RUBRICAS",
    r"^Página",
    r"^Emissão",
    r"^Hora:",
    r"^Empresa:",
    r"^Folha",
    r"^Centro de Custo:",
    r"^Código\s+Descrição",
    r"^\s*$",
    r"^Rescisão",
    r"^Férias",
    r"^Provisão",
    r"^Empresa$",
]

def should_ignore_nao_config(line: str) -> bool:
    for pat in IGNORE_PATTERNS_NAO_CONFIG:
        if re.search(pat, line, re.IGNORECASE):
            return True
    return False


# ==============================
# PARSE DO PDF DE RUBRICAS (catálogo de tipos)
# ==============================
def parse_rubricas_pdf(file_bytes: bytes, log: list) -> dict:
    """
    Lê o Rubricas.pdf e retorna dict {cod (str): tipo_norm (str)}
    Captura tipos colados ou separados na linha.
    """
    catalog = {}

    RE_LINHA = re.compile(
        r"^\s*(\d+)\s+"
        r"(.+?)"
        r"\s*(Provento|Desconto|Inf\.\s*ded(?:utora)?|Informativa)"
        r"[\s\w]",
        re.IGNORECASE,
    )

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for raw in text.splitlines():
                line = raw.strip()
                if not line or should_ignore_rubricas(line):
                    continue
                m = RE_LINHA.match(line)
                if m:
                    cod      = m.group(1).strip()
                    tipo_raw = m.group(3).strip().lower()
                    if "provento"   in tipo_raw: tipo_norm = "Provento"
                    elif "desconto" in tipo_raw: tipo_norm = "Desconto"
                    elif "inf. ded" in tipo_raw or "inf.ded" in tipo_raw:
                        tipo_norm = "Inf. dedutora"
                    elif "informat" in tipo_raw: tipo_norm = "Informativa"
                    else: tipo_norm = m.group(3).strip()
                    if cod not in catalog:
                        catalog[cod] = tipo_norm

    log.append(f"Rubricas.pdf: {len(catalog)} tipo(s) identificado(s) no catálogo.")
    return catalog


# ==============================
# PARSE DO PDF DE ITENS NÃO CONFIGURADOS
# ==============================
def parse_nao_configurados_pdf(file_bytes: bytes, log: list) -> list:
    """
    Lê o PDF 'Rubricas/Itens não Configurados' e retorna lista de dicts:
    [{ 'cod': str, 'descricao': str, 'tipo_folha': str, 'centro_custo': str }, ...]

    Estrutura do PDF:
      Folha Normal / Férias / Rescisão / Provisão de Férias / Provisão de 13º / Empresa
      Centro de Custo: N NOME
        Código  Descrição
        208     REEMBOLSO DE QUILOMETRAGEM
        ...
    """
    eventos = []
    vistos  = set()   # (cod, tipo_folha, centro_custo) para evitar duplicatas

    # Mapeamento de seção para tipo_integ
    SECAO_TIPO = {
        "Folha Normal":       "1",
        "Empresa":            "2",
        "Férias":             "3",
        "Rescisão":           "4",
        "Provisão de Férias": "5",
        "Provisão de 13º":    "6",
        "Provisão de 13o":    "6",
    }

    # Regex para linha de evento: código numérico + descrição
    RE_EVENTO = re.compile(r"^\s*(\d+)\s+(.+)$")

    # Regex para linha de seção
    RE_SECAO = re.compile(
        r"^(Folha Normal|Empresa|Férias|Rescisão|"
        r"Provisão de Férias|Provisão de 13º|Provisão de 13o)$",
        re.IGNORECASE,
    )

    # Regex para Centro de Custo
    RE_CC = re.compile(r"^Centro de Custo:\s*(\d+)\s+(.+)$", re.IGNORECASE)

    tipo_folha_atual  = "1"
    centro_custo_atual = ""

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for raw in text.splitlines():
                line = raw.strip()
                if not line:
                    continue

                # Detecta seção (tipo de folha)
                m_sec = RE_SECAO.match(line)
                if m_sec:
                    sec = m_sec.group(1).strip()
                    # Normaliza
                    for k, v in SECAO_TIPO.items():
                        if k.lower() in sec.lower():
                            tipo_folha_atual = v
                            break
                    continue

                # Detecta Centro de Custo
                m_cc = RE_CC.match(line)
                if m_cc:
                    centro_custo_atual = m_cc.group(1).strip()
                    continue

                # Ignora cabeçalhos e rodapés
                if should_ignore_nao_config(line):
                    continue

                # Tenta capturar linha de evento
                m_ev = RE_EVENTO.match(line)
                if m_ev:
                    cod  = m_ev.group(1).strip()
                    desc = m_ev.group(2).strip()

                    # Filtra falsos positivos (ex: "1/2 FERIAS" começa com dígito)
                    # Código válido: puramente numérico
                    if not cod.isdigit():
                        continue

                    chave = (cod, tipo_folha_atual, centro_custo_atual)
                    if chave not in vistos:
                        vistos.add(chave)
                        eventos.append({
                            "cod":          cod,
                            "descricao":    desc,
                            "tipo_folha":   tipo_folha_atual,
                            "centro_custo": centro_custo_atual,
                        })

    log.append(
        f"PDF Itens Não Configurados: {len(eventos)} evento(s) encontrado(s) "
        f"(únicos por código + tipo + centro de custo)."
    )
    return eventos


# ==============================
# LEITURA DO EXCEL DE CONTAS CONTÁBEIS
# ==============================
def ler_excel_contas(file_bytes: bytes, log: list) -> dict:
    """
    Lê o Excel (aba 'evento') e retorna dict:
    { (cod_seq, tipo_integ): { conta_debito, conta_credito, historico, complemento } }
    """
    contas = {}
    try:
        xls = pd.ExcelFile(BytesIO(file_bytes))
    except Exception as e:
        log.append(f"AVISO: Não foi possível abrir o Excel: {e}")
        return contas

    sheet = None
    for candidate in ["evento", "Evento", "EVENTO", "Plan1"]:
        if candidate in xls.sheet_names:
            sheet = candidate
            break

    if not sheet:
        log.append(f"AVISO: Aba 'evento' não encontrada no Excel. Abas: {xls.sheet_names}")
        return contas

    try:
        df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet, dtype=str)
    except Exception as e:
        log.append(f"AVISO: Erro ao ler aba '{sheet}': {e}")
        return contas

    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")

    # Mapeia nomes de colunas
    col_map = {}
    for col in df.columns:
        cl = col.lower()
        if "sequencial" in cl or "seq" in cl:
            col_map["seq"] = col
        elif "tipo" in cl and "integr" in cl:
            col_map["tipo"] = col
        elif "débito" in cl or "debito" in cl:
            col_map["debito"] = col
        elif "crédito" in cl or "credito" in cl:
            col_map["credito"] = col
        elif "histórico" in cl or "historico" in cl:
            col_map["historico"] = col
        elif "complemento" in cl:
            col_map["complemento"] = col

    for _, row in df.iterrows():
        seq  = str(row.get(col_map.get("seq",  ""), "") or "").strip()
        tipo = str(row.get(col_map.get("tipo", ""), "") or "").strip()
        deb  = str(row.get(col_map.get("debito",  ""), "") or "").strip()
        cred = str(row.get(col_map.get("credito", ""), "") or "").strip()
        hist = str(row.get(col_map.get("historico", ""), "") or "").strip()
        comp = str(row.get(col_map.get("complemento", ""), "") or "").strip()

        # Limpa "nan"
        deb  = "" if deb.lower()  == "nan" else deb
        cred = "" if cred.lower() == "nan" else cred
        hist = "" if hist.lower() == "nan" else hist
        comp = "" if comp.lower() == "nan" else comp

        if seq and seq.lower() != "nan":
            chave = (seq, tipo)
            if chave not in contas:
                contas[chave] = {
                    "conta_debito":  deb,
                    "conta_credito": cred,
                    "historico":     hist,
                    "complemento":   comp,
                }

    log.append(f"Excel: {len(contas)} configuração(ões) de contas carregada(s).")
    return contas


# ==============================
# GERAÇÃO DO TXT DOMÍNIO
# ==============================
def gerar_txt_dominio(
    eventos_nao_config: list,
    catalog_tipos: dict,
    contas_excel: dict,
    cod_empresa_padrao: str,
    log: list,
) -> tuple[str, pd.DataFrame]:
    """
    Cruza eventos não configurados + tipos do catálogo + contas do Excel.
    Gera o TXT no formato Domínio Separador.
    """
    linhas_txt   = []
    dados_tabela = []

    sem_tipo  = 0
    sem_conta = 0
    completos = 0

    for ev in eventos_nao_config:
        cod          = ev["cod"]
        descricao    = ev["descricao"]
        tipo_integ   = ev["tipo_folha"]
        centro_custo = ev["centro_custo"]

        # Busca tipo no catálogo
        tipo_rubrica = catalog_tipos.get(cod, "")

        # Busca contas no Excel
        conta_info = contas_excel.get((cod, tipo_integ), {})
        # Tenta também sem o tipo (chave só pelo código)
        if not conta_info:
            for k, v in contas_excel.items():
                if k[0] == cod:
                    conta_info = v
                    break

        conta_deb  = conta_info.get("conta_debito",  "")
        conta_cred = conta_info.get("conta_credito", "")
        historico  = conta_info.get("historico",     "")
        complemento = conta_info.get("complemento",  "")

        # Status
        if not tipo_rubrica:
            status = "⚠️ Tipo não encontrado no catálogo"
            sem_tipo += 1
        elif not conta_deb and not conta_cred:
            status = "ℹ️ Sem conta configurada no Excel"
            sem_conta += 1
        else:
            status = f"✅ {tipo_rubrica}"
            completos += 1

        # Linha do TXT
        linha = (
            f"{cod_empresa_padrao}|{centro_custo}|{cod}|{tipo_integ}|"
            f"{descricao}|{conta_deb}|{conta_cred}|{historico}|{complemento}|\n"
        )
        linhas_txt.append(linha)

        dados_tabela.append({
            "Código":          cod,
            "Descrição":       descricao,
            "Tipo Integração": tipo_integ,
            "Centro Custo":    centro_custo,
            "Tipo Rubrica":    tipo_rubrica or "—",
            "Conta Débito":    conta_deb,
            "Conta Crédito":   conta_cred,
            "Histórico":       historico,
            "Status":          status,
        })

    log.append(
        f"Geração concluída → "
        f"Completos: {completos} | "
        f"Sem tipo no catálogo: {sem_tipo} | "
        f"Sem conta no Excel: {sem_conta}"
    )

    return "".join(linhas_txt), pd.DataFrame(dados_tabela)


# ==============================
# INTERFACE STREAMLIT
# ==============================
def main():
    st.set_page_config(
        page_title="Domínio Sistemas | Integração Contábil",
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
                📊 Integração Contábil — Itens Não Configurados → Domínio
                &nbsp;|&nbsp; {VERSAO}
            </h2>
            <p style="color:#DDDDDD; margin:6px 0 0 0;
                      font-family:'Segoe UI',Arial,sans-serif;">
                Faça upload dos 3 arquivos e clique em
                <b>▶ Gerar arquivo Domínio</b>.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── Sidebar ───────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("### ⚙️ Configurações")
        cod_empresa_padrao = st.text_input(
            "Código da empresa (padrão)",
            value="45",
            help="Código da empresa no Domínio.",
        )
        st.markdown("---")
        st.markdown("### ℹ Sobre")
        st.markdown(f"**Versão:** {VERSAO}")
        st.markdown("**Thomson Reuters | Domínio Sistemas**")

    # ── Instruções ────────────────────────────────────────────────────
    with st.expander("📖 **Instruções de Uso** — clique para expandir", expanded=False):
        st.markdown(
            """
            <div class="instrucoes-box">

            <h4>🔹 Arquivos necessários</h4>
            <ol>
                <li><b>Rubricas/Itens não Configurados.pdf</b>: relatório gerado no Domínio
                    em <i>Plano e Acumuladores → Rubricas/Itens não Configurados</i>.</li>
                <li><b>Rubricas.pdf</b>: relatório gerado no Domínio em
                    <i>Plano e Acumuladores → Rubricas</i> (catálogo completo com tipos).</li>
                <li><b>Excel de Contas (.xlsx)</b>: planilha com as contas contábeis
                    já configuradas (aba <code>evento</code>).</li>
            </ol>

            <h4>🔹 O que o sistema faz</h4>
            <ul>
                <li>Lê todos os eventos sem configuração contábil do PDF 1.</li>
                <li>Busca o <b>tipo de rubrica</b> (Provento/Desconto/Informativa/Inf. dedutora)
                    de cada evento no PDF 2.</li>
                <li>Busca as <b>contas contábeis</b> (débito/crédito/histórico) no Excel.</li>
                <li>Gera o <b>arquivo TXT</b> no formato Domínio Separador.</li>
            </ul>

            <h4>🔹 Importar no Domínio</h4>
            <p><b>Utilitários → Importação → Importação Padrão →
            Leiaute Domínio Sistemas com Separador</b>.</p>

            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("---")

    # ── Session state ─────────────────────────────────────────────────
    defaults = {
        "log":          [f"Aplicação pronta. Versão: {VERSAO}"],
        "txt_gerado":   None,
        "nome_arquivo": "dominio_integracao.txt",
        "df_resultado": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

    # ── Uploads ───────────────────────────────────────────────────────
    col1, col2, col3 = st.columns(3)
    with col1:
        pdf_nao_config = st.file_uploader(
            "1️⃣ PDF — Itens Não Configurados",
            type=["pdf"],
            help="Relatório 'Rubricas/Itens não Configurados' do Domínio.",
        )
    with col2:
        pdf_rubricas = st.file_uploader(
            "2️⃣ PDF — Rubricas (catálogo)",
            type=["pdf"],
            help="Relatório 'Rubricas' completo do Domínio (com tipos).",
        )
    with col3:
        excel_contas = st.file_uploader(
            "3️⃣ Excel — Contas Contábeis (opcional)",
            type=["xlsx", "xls"],
            help="Planilha com contas débito/crédito já configuradas (aba 'evento').",
        )

    arquivos_ok = pdf_nao_config is not None and pdf_rubricas is not None

    col_btn1, col_btn2 = st.columns([1, 1])
    with col_btn1:
        gerar = st.button(
            "▶ Gerar arquivo Domínio",
            disabled=not arquivos_ok,
            use_container_width=True,
            type="primary",
        )
    with col_btn2:
        limpar = st.button("🗑 Limpar", use_container_width=True)

    if limpar:
        for k, v in defaults.items():
            st.session_state[k] = v
        st.session_state.log = ["Campos limpos."]
        st.rerun()

    # ── Processamento ─────────────────────────────────────────────────
    if gerar and arquivos_ok:
        log = ["Iniciando processamento..."]

        # 1. Catálogo de tipos (Rubricas.pdf)
        with st.spinner("Lendo Rubricas.pdf (catálogo de tipos)..."):
            catalog_tipos = parse_rubricas_pdf(pdf_rubricas.read(), log)

        # 2. Eventos não configurados
        with st.spinner("Lendo PDF de Itens Não Configurados..."):
            eventos = parse_nao_configurados_pdf(pdf_nao_config.read(), log)

        if not eventos:
            log.append("AVISO: Nenhum evento encontrado no PDF de Itens Não Configurados.")

        # 3. Contas do Excel (opcional)
        contas_excel = {}
        if excel_contas is not None:
            with st.spinner("Lendo Excel de contas..."):
                contas_excel = ler_excel_contas(excel_contas.read(), log)

        # 4. Geração do TXT
        txt, df_resultado = gerar_txt_dominio(
            eventos, catalog_tipos, contas_excel, cod_empresa_padrao, log
        )

        st.session_state.txt_gerado   = txt.encode("latin-1", errors="replace")
        st.session_state.nome_arquivo = "dominio_integracao.txt"
        st.session_state.df_resultado = df_resultado
        st.session_state.log          = log
        st.rerun()

    # ── Resultados ────────────────────────────────────────────────────
    if st.session_state.txt_gerado is not None:
        st.success("✅ Arquivo gerado com sucesso!")
        st.download_button(
            label="⬇ Baixar TXT (Domínio Separador)",
            data=st.session_state.txt_gerado,
            file_name=st.session_state.nome_arquivo,
            mime="text/plain",
            use_container_width=True,
            type="primary",
        )

        df = st.session_state.df_resultado
        if df is not None and not df.empty:
            total     = len(df)
            completos = len(df[df["Status"].str.startswith("✅")])
            sem_tipo  = len(df[df["Status"].str.startswith("⚠️")])
            sem_conta = len(df[df["Status"].str.startswith("ℹ️")])

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("📋 Total eventos", total)
            m2.metric("✅ Completos",      completos)
            m3.metric("⚠️ Sem tipo",       sem_tipo)
            m4.metric("ℹ️ Sem conta",      sem_conta)

            # Tabela com highlight
            def highlight_row(row):
                s = str(row.get("Status", ""))
                if s.startswith("✅"):  return ["background-color:#d4edda"] * len(row)
                if s.startswith("⚠️"): return ["background-color:#fff3cd"] * len(row)
                if s.startswith("ℹ️"): return ["background-color:#cce5ff"] * len(row)
                return [""] * len(row)

            st.dataframe(
                df.style.apply(highlight_row, axis=1),
                use_container_width=True,
            )

            # Expander: sem tipo no catálogo
            sem_tipo_df = df[df["Status"].str.startswith("⚠️")]
            if not sem_tipo_df.empty:
                with st.expander(
                    f"⚠️ Eventos sem tipo no catálogo ({len(sem_tipo_df)}) "
                    f"— verificar Rubricas.pdf"
                ):
                    st.dataframe(
                        sem_tipo_df[["Código", "Descrição", "Tipo Integração", "Centro Custo"]],
                        use_container_width=True,
                    )

            # Expander: sem conta
            sem_conta_df = df[df["Status"].str.startswith("ℹ️")]
            if not sem_conta_df.empty:
                with st.expander(
                    f"ℹ️ Eventos sem conta configurada ({len(sem_conta_df)}) "
                    f"— configurar no Excel"
                ):
                    st.dataframe(
                        sem_conta_df[["Código", "Descrição", "Tipo Integração",
                                      "Tipo Rubrica", "Centro Custo"]],
                        use_container_width=True,
                    )

            # Prévia do arquivo
            with st.expander("👁️ Prévia do arquivo gerado (primeiras 30 linhas)"):
                preview = "".join(
                    st.session_state.txt_gerado
                    .decode("latin-1", errors="replace")
                    .splitlines(True)[:30]
                )
                st.code(preview, language="text")

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
                    white-space:pre-wrap; max-height:340px;
                    overflow-y:auto; color:#1F1F1F;">
{log_texto}
        </div>
        """,
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
