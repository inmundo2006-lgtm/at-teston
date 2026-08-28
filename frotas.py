"""
frotas.py — Cadastro de frotas lido da lista KanbanFrotas (SharePoint)

Substitui a leitura do CADASTRO.xlsx. Motivo: a planilha só era atualizada
quando alguém lembrava de reenviá-la ao GitHub, enquanto a KanbanFrotas é
mantida pela operação e já alimenta o app de Checklist. Uma fonte só para
os dois apps.

O que a lista traz por frota:
    Title      → "1173 - STRADA ENDURANCE 1.4"  (código + descrição)
    Tipo       → Carro, Caminhão, Trator...
    Chassi, Ano
    CCNome     → "038 - AGRO ASTORGA"
    FrenteNome → "AGROASTORGA"
    Status     → Ativo, DESTINADO A VENDA, ENTREGA FUTURA...

Secrets no at-teston (os três primeiros já existem). Configure de UMA das
duas formas — a primeira é a mais legível:

    FROTAS_SITE_HOST = "metalcana.sharepoint.com"
    FROTAS_SITE_NAME = "AppKanbanFrotas"

ou, se preferir o id pronto:

    FROTAS_SITE_ID = "metalcana.sharepoint.com,<guid>,<guid>"

Com host+nome o app resolve o id sozinho, numa chamada extra que fica em
cache. Vale o custo: daqui a seis meses "AppKanbanFrotas" diz o que é,
enquanto o GUID não diz nada.

FALLBACK: se a lista não responder, cai para o CADASTRO.xlsx e usa apenas a
DESCRIÇÃO de lá. O centro de custo NUNCA vem do fallback — uma planilha
defasada preencheria o CC errado em silêncio, que é pior que não preencher.
"""

import os

import requests
import streamlit as st

LISTA_FROTAS = "KanbanFrotas"

# Status que indicam veículo fora de operação. Não bloqueiam a abertura de
# OS — um carro destinado a venda ainda pode precisar de reparo — mas geram
# aviso, porque costuma ser engano.
STATUS_ALERTA = ("DESTINADO A VENDA", "ENTREGA FUTURA", "VENDIDO", "BAIXADO")


def _cfg(chave: str, padrao=None):
    return st.secrets.get(chave, padrao)


@st.cache_data(ttl=3000)
def _token() -> str:
    r = requests.post(
        f"https://login.microsoftonline.com/{_cfg('TENANT_ID')}/oauth2/v2.0/token",
        data={
            "grant_type":    "client_credentials",
            "client_id":     _cfg("CLIENT_ID"),
            "client_secret": _cfg("CLIENT_SECRET"),
            "scope":         "https://graph.microsoft.com/.default",
        },
        timeout=15,
    )
    r.raise_for_status()
    return r.json()["access_token"]


@st.cache_data(ttl=3000)
def _site_id() -> str:
    """
    Id do site que hospeda a KanbanFrotas.

    Aceita o FROTAS_SITE_ID pronto; na falta dele, resolve a partir de
    host + nome, do mesmo jeito que o app de checklist faz. Em cache,
    então a chamada extra acontece uma vez a cada 50 minutos.
    """
    pronto = _cfg("FROTAS_SITE_ID")
    if pronto:
        return pronto

    host = _cfg("FROTAS_SITE_HOST", "metalcana.sharepoint.com")
    nome = _cfg("FROTAS_SITE_NAME", "AppKanbanFrotas")
    if not host or not nome:
        raise RuntimeError(
            "Configure FROTAS_SITE_HOST e FROTAS_SITE_NAME (ou FROTAS_SITE_ID) "
            "nos secrets."
        )

    r = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{host}:/sites/{nome}",
        headers={"Authorization": f"Bearer {_token()}"},
        timeout=15,
    )
    r.raise_for_status()
    return r.json()["id"]


def _codigo_e_descricao(titulo: str) -> tuple[str, str]:
    """'1173 - STRADA ENDURANCE 1.4' → ('1173', '1173 - STRADA ENDURANCE 1.4')."""
    titulo = str(titulo or "").strip()
    if not titulo:
        return "", ""
    codigo = titulo.split("-", 1)[0].strip()
    return codigo, titulo


def _do_sharepoint() -> dict:
    site = _site_id()

    url = (f"https://graph.microsoft.com/v1.0/sites/{site}"
           f"/lists/{LISTA_FROTAS}/items?$expand=fields&$top=999")
    headers = {"Authorization": f"Bearer {_token()}",
               "Content-Type": "application/json"}

    itens = []
    while url:
        r = requests.get(url, headers=headers, timeout=25)
        r.raise_for_status()
        data = r.json()
        itens.extend(data.get("value", []))
        url = data.get("@odata.nextLink")

    frotas = {}
    for i in itens:
        f = i.get("fields", {})
        codigo, descricao = _codigo_e_descricao(f.get("Title", ""))
        if not codigo:
            continue
        frotas[codigo] = {
            "descricao":   descricao,
            "tipo":        f.get("Tipo", ""),
            "chassi":      f.get("Chassi", ""),
            "ano":         f.get("Ano", ""),
            "cc_nome":     f.get("CCNome", ""),
            "frente_nome": f.get("FrenteNome", ""),
            "status":      f.get("Status", "") or "Ativo",
            "origem":      "sharepoint",
        }
    return frotas


def _do_excel() -> dict:
    """Fallback: só a descrição. Nunca o centro de custo (ver cabeçalho)."""
    try:
        import pandas as pd
        caminho = os.path.join(os.path.dirname(__file__), "CADASTRO.xlsx")
        df = pd.read_excel(caminho)
        df.columns = df.columns.str.strip()
        df["FROTA"]     = df["FROTA"].astype(str).str.strip()
        df["DESCRICAO"] = df["DESCRICAO"].astype(str).str.strip()
        return {
            row["FROTA"]: {
                "descricao": f'{row["FROTA"]} - {row["DESCRICAO"]}',
                "tipo": "", "chassi": "", "ano": "",
                "cc_nome": "",          # deliberadamente vazio
                "frente_nome": "",
                "status": "",
                "origem": "excel",
            }
            for _, row in df.iterrows()
        }
    except Exception:
        return {}


@st.cache_data(ttl=1800)
def carregar_frotas() -> dict:
    """
    {codigo: {descricao, tipo, chassi, ano, cc_nome, frente_nome, status, origem}}

    Tenta o SharePoint; se falhar, cai para o Excel. Nunca levanta exceção —
    a abertura de OS continua possível mesmo sem cadastro, só sem
    preenchimento automático.
    """
    try:
        frotas = _do_sharepoint()
        if frotas:
            return frotas
    except Exception:
        pass
    return _do_excel()


def info_frota(frota: str) -> dict:
    return carregar_frotas().get(str(frota).strip(), {})


def alerta_status(status: str) -> str:
    """Mensagem de aviso se o veículo estiver fora de operação, senão ''."""
    s = str(status or "").strip().upper()
    for marcado in STATUS_ALERTA:
        if marcado in s:
            return (f"⚠️ Esta frota está marcada como **{status}** no cadastro. "
                    "Confirme se a OS é mesmo para ela.")
    return ""