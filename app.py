"""
Integrador Contábil da Folha — Domínio Sistemas
Motor 100% determinístico (regex + regras de substring). Sem IA, sem APIs, sem chaves.

- Nenhuma conta, código reduzido ou classificação de plano é fixada no código.
- Exportação no layout de importação da Domínio (1 lançamento por rubrica/item):
    aba "evento"  -> lançamento (tabela FOINTEGCONT): Código Sequencial = código da rubrica/item,
                     Descrição = "código - descrição" (caixa mista, máx. 40 caracteres)
    aba "integra" -> vínculo da rubrica/item ao seu lançamento (tabela FOINTEGCONTEVE)
- Tipos: 1 Folha mensal | 2 Empresa | 3 Férias | 4 Rescisão | 5 Prov. Férias | 6 Prov. 13 | 10 Outras Informações
- Dois modos:
    * Pendências: lê o relatório "Rubricas/Itens não configurados"
    * Todos os eventos cadastrados: usa só o Cadastro geral de rubricas + Plano de contas
- "Usar a mesma configuração da folha normal para Férias/Rescisão" (igual à Domínio):
    marcado    = o Tipo 3/4 não é gerado; os itens usam o lançamento da Folha (Tipo 1).
                 Rubricas exclusivas de Férias/Rescisão (que não existem na Folha) podem permanecer
                 no próprio Tipo 3/4 (opção "Manter exclusivas"), como a Domínio grava.
    desmarcado = Tipo 3/4 próprio: passivo próprio (Férias/Rescisões a Pagar) E despesa própria
                 (Férias / Rescisão), com exceções automáticas e por código
- Entidade filantrópica: gera também os itens de isenção (25 a 36) como estorno do INSS patronal.
- Conferências: chave única do lançamento (empresa+separador+código+tipo), uma rubrica por lançamento
  em cada empresa+separador+tipo, verificação do líquido e conformidade com FOINTEGCONT/FOINTEGCONTEVE.

Executar:
    pip install -r requirements.txt
    streamlit run app_corrigido.py
"""
import csv
import hashlib
import io
import json
import numbers
import re
import unicodedata
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
