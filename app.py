import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO
import re

st.set_page_config(page_title="PDF Rubricas → Excel", page_icon="📊", layout="wide")

st.title("📊 Conversor: Rubricas/Itens Não Configurados → Excel")
st.markdown("Faça o upload do PDF para converter as rubricas em planilha Excel estruturada.")

# ── Mapeamento de tipos de integração ──────────────────────────────────────────
TIPO_MAP = {
    "Folha Normal": 1,
    "Empresa":      2,
    "Férias":       3,
    "Rescisão":     4,
    "Provisão de Férias": 5,
    "Provisão de 13º":    6,
}

TIPO_DESC = {
    1: "Folha mensal",
    2: "Empresa",
    3: "Férias",
    4: "Rescisão",
    5: "Prov. Férias",
    6: "Prov. 13",
}

# ── Mapeamento de centros de custo ─────────────────────────────────────────────
CC_MAP = {
    "1":  "ADMINISTRAÇÃO",
    "2":  "ESTAGIÁRIOS",
    "6":  "ROCHE",
    "10": "BAYER",
    "13": "GERAL",
}


def parse_pdf(file_bytes: bytes) -> pd.DataFrame:
    """Extrai todas as linhas do PDF e monta o DataFrame."""
    rows = []

    # Padrões de detecção
    re_tipo   = re.compile(
        r"^(Folha Normal|Empresa|Férias|Rescisão|Provisão de Férias|Provisão de 13[oº°])\s*$",
        re.IGNORECASE,
    )
    re_cc     = re.compile(r"^Centro de Custo:\s*(\d+)\s+(.+)$", re.IGNORECASE)
    re_event  = re.compile(r"^\s*(\d+)\s+(.+)$")
    re_header = re.compile(r"^(Código|RELAÇÃO DE RUBRICAS|Página|Emissão|Hora|Empresa:)", re.IGNORECASE)

    current_tipo = None
    current_cc_cod  = None
    current_cc_desc = None

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            for raw_line in text.splitlines():
                line = raw_line.strip()

                if not line:
                    continue

                # Ignora cabeçalhos de página
                if re_header.match(line):
                    continue

                # ── Tipo de integração ─────────────────────────────────────
                m_tipo = re_tipo.match(line)
                if m_tipo:
                    # Normaliza "Provisão de 13o" → "Provisão de 13º"
                    raw = m_tipo.group(1)
                    if re.search(r"13[oO°]", raw):
                        raw = "Provisão de 13º"
                    current_tipo = TIPO_MAP.get(raw)
                    continue

                # ── Centro de Custo ────────────────────────────────────────
                m_cc = re_cc.match(line)
                if m_cc:
                    current_cc_cod  = m_cc.group(1).strip()
                    current_cc_desc = m_cc.group(2).strip()
                    continue

                # ── Evento (código + descrição) ────────────────────────────
                m_ev = re_event.match(line)
                if m_ev and current_tipo and current_cc_cod:
                    cod_ev  = m_ev.group(1).strip()
                    desc_ev = m_ev.group(2).strip()
                    tipo_num  = current_tipo
                    tipo_desc = TIPO_DESC.get(tipo_num, "")

                    rows.append({
                        "Cód Centro de Custo":       current_cc_cod,
                        "Desc. Centro de Custo":     current_cc_desc,
                        "Tipo da Integração":        tipo_num,
                        "Desc. Tipo Integração":     tipo_desc,
                        "Cod Evento":                cod_ev,
                        "Descrição Evento":          desc_ev,
                        "Cod + Descrição Evento":    f"{cod_ev} - {desc_ev}",
                    })

    return pd.DataFrame(rows)


def to_excel(df: pd.DataFrame) -> bytes:
    """Converte o DataFrame para bytes de um arquivo .xlsx estilizado."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Rubricas")

        workbook  = writer.book
        worksheet = writer.sheets["Rubricas"]

        # ── Formatos ───────────────────────────────────────────────────────
        header_fmt = workbook.add_format({
            "bold":       True,
            "bg_color":   "#1F4E79",
            "font_color": "#FFFFFF",
            "border":     1,
            "align":      "center",
            "valign":     "vcenter",
            "text_wrap":  True,
        })
        cell_fmt = workbook.add_format({
            "border":  1,
            "valign":  "vcenter",
            "align":   "left",
        })
        num_fmt = workbook.add_format({
            "border":  1,
            "valign":  "vcenter",
            "align":   "center",
        })
        alt_fmt = workbook.add_format({        # linha alternada
            "border":    1,
            "valign":    "vcenter",
            "align":     "left",
            "bg_color":  "#D6E4F0",
        })
        alt_num_fmt = workbook.add_format({
            "border":    1,
            "valign":    "vcenter",
            "align":     "center",
            "bg_color":  "#D6E4F0",
        })

        # ── Cabeçalho ──────────────────────────────────────────────────────
        for col_num, col_name in enumerate(df.columns):
            worksheet.write(0, col_num, col_name, header_fmt)

        # ── Larguras das colunas ───────────────────────────────────────────
        col_widths = {
            "Cód Centro de Custo":    8,
            "Desc. Centro de Custo": 22,
            "Tipo da Integração":    10,
            "Desc. Tipo Integração": 18,
            "Cod Evento":             8,
            "Descrição Evento":      40,
            "Cod + Descrição Evento":45,
        }
        for col_num, col_name in enumerate(df.columns):
            worksheet.set_column(col_num, col_num, col_widths.get(col_name, 18))

        # ── Dados com zebra striping ───────────────────────────────────────
        num_cols = {"Cód Centro de Custo", "Tipo da Integração", "Cod Evento"}
        for row_num, row_data in enumerate(df.itertuples(index=False), start=1):
            is_alt = row_num % 2 == 0
            for col_num, col_name in enumerate(df.columns):
                value = row_data[col_num]
                if col_name in num_cols:
                    fmt = alt_num_fmt if is_alt else num_fmt
                else:
                    fmt = alt_fmt if is_alt else cell_fmt
                worksheet.write(row_num, col_num, value, fmt)

        # ── Altura do cabeçalho ────────────────────────────────────────────
        worksheet.set_row(0, 30)

        # ── Freeze pane no cabeçalho ───────────────────────────────────────
        worksheet.freeze_panes(1, 0)

        # ── AutoFilter ────────────────────────────────────────────────────
        worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)

    return output.getvalue()


# ── Interface Streamlit ────────────────────────────────────────────────────────
uploaded_file = st.file_uploader("📁 Selecione o arquivo PDF", type=["pdf"])

if uploaded_file:
    with st.spinner("🔄 Processando PDF..."):
        file_bytes = uploaded_file.read()
        df = parse_pdf(file_bytes)

    if df.empty:
        st.error("⚠️ Nenhum dado foi extraído. Verifique se o PDF está correto.")
    else:
        st.success(f"✅ {len(df)} registros extraídos com sucesso!")

        # ── Métricas resumo ────────────────────────────────────────────────
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total de Registros",      len(df))
        col2.metric("Centros de Custo",         df["Cód Centro de Custo"].nunique())
        col3.metric("Tipos de Integração",      df["Tipo da Integração"].nunique())
        col4.metric("Eventos Únicos",           df["Cod Evento"].nunique())

        st.divider()

        # ── Filtros interativos ────────────────────────────────────────────
        st.subheader("🔍 Filtros")
        fcol1, fcol2, fcol3 = st.columns(3)

        with fcol1:
            cc_options = ["Todos"] + sorted(
                df["Cód Centro de Custo"].unique().tolist(), key=lambda x: int(x)
            )
            sel_cc = st.selectbox("Centro de Custo", cc_options)

        with fcol2:
            tipo_options = ["Todos"] + sorted(df["Tipo da Integração"].unique().tolist())
            sel_tipo = st.selectbox(
                "Tipo de Integração",
                tipo_options,
                format_func=lambda x: f"{x} - {TIPO_DESC[x]}" if x != "Todos" else x,
            )

        with fcol3:
            search_ev = st.text_input("🔎 Buscar Evento (código ou descrição)")

        # Aplica filtros
        df_filtered = df.copy()
        if sel_cc != "Todos":
            df_filtered = df_filtered[df_filtered["Cód Centro de Custo"] == sel_cc]
        if sel_tipo != "Todos":
            df_filtered = df_filtered[df_filtered["Tipo da Integração"] == sel_tipo]
        if search_ev:
            mask = (
                df_filtered["Cod Evento"].str.contains(search_ev, case=False, na=False)
                | df_filtered["Descrição Evento"].str.contains(search_ev, case=False, na=False)
            )
            df_filtered = df_filtered[mask]

        st.dataframe(df_filtered, use_container_width=True, height=450)
        st.caption(f"Exibindo {len(df_filtered)} de {len(df)} registros")

        st.divider()

        # ── Download ───────────────────────────────────────────────────────
        st.subheader("⬇️ Exportar")
        excel_bytes = to_excel(df)  # exporta sempre o df completo

        dcol1, dcol2 = st.columns([1, 3])
        with dcol1:
            st.download_button(
                label="📥 Baixar Excel Completo",
                data=excel_bytes,
                file_name="rubricas_nao_configuradas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        # Download filtrado
        if len(df_filtered) < len(df):
            excel_filtered = to_excel(df_filtered)
            with dcol2:
                st.download_button(
                    label=f"📥 Baixar Excel Filtrado ({len(df_filtered)} registros)",
                    data=excel_filtered,
                    file_name="rubricas_nao_configuradas_filtrado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
else:
    st.info("👆 Faça o upload do PDF 'RubricasItens não Configurados.pdf' para começar.")

    # ── Legenda de tipos ───────────────────────────────────────────────────
    st.subheader("📋 Legenda — Tipos de Integração")
    leg_data = [{"Código": k, "Descrição": v} for k, v in TIPO_DESC.items()]
    st.table(pd.DataFrame(leg_data))
