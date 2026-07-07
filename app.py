if gerar_excel and pdf_file and txt_file:
    log = ["[Etapa 1] Iniciando..."]

    with st.spinner("Lendo rubricas.txt..."):
        catalog = parse_rubricas_txt(txt_file.read(), log)

    with st.spinner("Lendo PDF..."):
        eventos = parse_nao_configurados_pdf(pdf_file.read(), log)

    st.session_state.eventos_parsed = eventos
    st.session_state.catalog_parsed = catalog

    # ✅ NOVO: Classifica automaticamente ANTES de gerar o Excel
    with st.spinner("🔍 Classificando rubricas automaticamente..."):
        classif_auto = classificar_todos_eventos(
            eventos, catalog, df_pc, log
        )
        st.session_state["classif_auto"] = classif_auto

    # Aplica classificação automática ao config_cc
    if df_pc is not None and not df_pc.empty and usa_sep_bool:
        ccs_novos = get_centros_custo_unicos(eventos)
        for cc_cod, _ in ccs_novos:
            if cc_cod not in st.session_state.config_cc:
                # Pega o grupo mais frequente entre os eventos deste CC
                grupos_cc = [
                    classif_auto.get(ev["cod"], {}).get("grupo", "Despesa Administrativa")
                    for ev in eventos
                    if ev["centro_custo_cod"] == cc_cod
                ]
                grupo_dominante = max(set(grupos_cc), key=grupos_cc.count)
                auto = sugerir_contas(df_pc, grupo_dominante)
                st.session_state.config_cc[cc_cod] = {
                    "grupo":         grupo_dominante,
                    "conta_debito":  auto["conta_debito"],
                    "conta_credito": auto["conta_credito"],
                    "historico":     "",
                }
