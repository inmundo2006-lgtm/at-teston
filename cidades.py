"""
cidades.py — Lista de municípios brasileiros para os campos de deslocamento

Lê o cidades_brasil.csv gerado uma vez pelo gerar_cidades_ibge.py. Sem
dependência da API do IBGE em produção.

Se o CSV não estiver na pasta, o app não quebra: os campos de cidade
viram texto livre e aparece um aviso.
"""

import csv
import os

import streamlit as st

ARQUIVO = "cidades_brasil.csv"

# Cidades da operação que não são município no IBGE (distritos, fazendas,
# unidades). Entram no topo da lista, junto com as prioritárias.
LOCALIDADES_PROPRIAS = [
    "Nova Produtiva - PR",      # distrito de Astorga
]

PRIORITARIAS = [
    "Cianorte - PR",
    "Maringá - PR",
    "Astorga - PR",
    "Naviraí - MS",
]

SEM_CIDADE = "— Selecione —"


@st.cache_data(ttl=3600)
def carregar_cidades() -> list[str]:
    """
    Retorna a lista para o selectbox: um marcador vazio, depois as
    localidades mais usadas, depois todos os municípios em ordem.
    """
    caminho = os.path.join(os.path.dirname(__file__), ARQUIVO)
    if not os.path.exists(caminho):
        return []

    todas = []
    try:
        with open(caminho, encoding="utf-8", newline="") as f:
            for linha in csv.DictReader(f):
                nome = (linha.get("nome_completo") or "").strip()
                if nome:
                    todas.append(nome)
    except Exception:
        return []

    if not todas:
        return []

    # dict.fromkeys preserva a ordem e remove repetição
    topo  = list(dict.fromkeys(PRIORITARIAS + LOCALIDADES_PROPRIAS))
    resto = [c for c in todas if c not in topo]

    return [SEM_CIDADE] + topo + resto


def campo_cidade(label: str, key: str, valor_atual: str = "") -> str:
    """
    Renderiza o campo de cidade. Usa selectbox com busca quando o CSV
    existe; cai para texto livre quando não existe, para o lançamento
    não ficar bloqueado por falta do arquivo.
    """
    cidades = carregar_cidades()

    if not cidades:
        st.caption(f"⚠️ {ARQUIVO} não encontrado — digite a cidade manualmente.")
        return st.text_input(label, value=valor_atual, key=key,
                             placeholder="Ex: Cianorte - PR")

    idx = cidades.index(valor_atual) if valor_atual in cidades else 0
    escolha = st.selectbox(label, cidades, index=idx, key=key)
    return "" if escolha == SEM_CIDADE else escolha
