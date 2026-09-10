"""
dados.py — Configurações, tabelas de valores e persistência das OS
Persistência: SharePoint via MS Graph API (substitui os_data.json)
"""

import json
import hashlib
import uuid
from datetime import datetime

import requests
import streamlit as st

# ─────────────────────────────────────────────
#  TABELAS DE VALORES
# ─────────────────────────────────────────────
# Fonte única: calculos.py, que espelha a aba "Valores" da planilha.
# Antes existiam duas cópias (aqui e lá) com valores DIFERENTES — foi
# assim que a tabela errada (a da aba "Inf.") entrou no sistema e as
# comissões saíram ~8,6% acima do devido. Não redeclare nada aqui.

from calculos import (
    TABELA_KM,
    TABELA_HORA,
    HORA_DESLOCAMENTO,
    HORA_MUNCK,
    PERCENTUAL_DESLOCAMENTO,
    PERCENTUAIS_LOCAL as _PCT_FRACAO,
    ALMOCO_PADRAO_H,
    CAFE_PADRAO_H,
)

# A UI exibe o percentual como inteiro (4, 8, 10); o cálculo usa fração.
PERCENTUAIS_LOCAL = {k: round(v * 100) for k, v in _PCT_FRACAO.items()}

HORA_DESLOCAMENTO = 60
HORA_MUNCK        = 120


COMISSAO_DESLOCAMENTO_CAMPO   = 8.0
COMISSAO_DESLOCAMENTO_INTERNO = 4.0
COMISSAO_DESLOCAMENTO_MS      = 10.0

# ─────────────────────────────────────────────
#  TÉCNICOS  (sem alteração)
# ─────────────────────────────────────────────

TECNICOS = {
    1:  {"nome": "Zaqueu",           "funcao": "Mecânico",         "tipo_km": "Km Um",   "nivel": "Técnico Cinco"},
    2:  {"nome": "Cristiano B.",     "funcao": "Borracheiro",       "tipo_km": "Km Três", "nivel": "Técnico Três"},
    3:  {"nome": "Rodrigo J.",       "funcao": "Troca de Óleo",     "tipo_km": "Km Dois", "nivel": "Técnico Três"},
    4:  {"nome": "Marcos C.",        "funcao": "Eletricista",       "tipo_km": "Km Um",   "nivel": "Técnico Cinco"},
    5:  {"nome": "Cristiano dos S.", "funcao": "Mecânico",          "tipo_km": "Km Um",   "nivel": "Técnico Quatro"},
    6:  {"nome": "João P.",          "funcao": "Eletricista",       "tipo_km": "Km Um",   "nivel": "Técnico Cinco"},
    7:  {"nome": "Loran H.",         "funcao": "Eletricista",       "tipo_km": "Km Um",   "nivel": "Técnico Três"},
    8:  {"nome": "Bruno H.",         "funcao": "Mecânico",          "tipo_km": "Km Um",   "nivel": "Técnico Três"},
    9:  {"nome": "Gabriel",          "funcao": "Auxiliar Mecânico", "tipo_km": "Km Um",   "nivel": "Técnico Um"},
    10: {"nome": "Tiago Bardu",      "funcao": "Mecânico",          "tipo_km": "Km Um",   "nivel": "Técnico Quatro"},
    11: {"nome": "Felipe",           "funcao": "Lavador",           "tipo_km": "Km Um",   "nivel": "Técnico Dois"},
    12: {"nome": "Gustavo",          "funcao": "Auxiliar Mecânico", "tipo_km": "Km Um",   "nivel": "Técnico Um"},
    13: {"nome": "Vanderlei",        "funcao": "Auxiliar Mecânico", "tipo_km": "Km Um",   "nivel": "Técnico Um"},
    14: {"nome": "Claudecir",        "funcao": "Mecânico",          "tipo_km": "Km Um",   "nivel": "Técnico Três"},
    15: {"nome": "Rodrigo M.",       "funcao": "Mecânico",          "tipo_km": "Km Um",   "nivel": "Técnico Três"},
    16: {"nome": "José",             "funcao": "Auxiliar",          "tipo_km": "Km Um",   "nivel": "Técnico Um"},
    17: {"nome": "Luiz",             "funcao": "Lavador",           "tipo_km": "Km Um",   "nivel": "Técnico Um"},
}

# ─────────────────────────────────────────────
#  TIPOS DE SERVIÇO  (sem alteração)
# ─────────────────────────────────────────────

TIPOS_SERVICO = [
    "Mecânica",
    "Elétrica",
    "Hidráulica",
    "Borracharia",
    "Troca de Óleo",
    "Lavagem/Lubrificação/Abastecimento",
    "Serviços Gerais",
    "Carga/descarga",
    "Munck",
    "Prensagem",
    "Assistência",
]

# ─────────────────────────────────────────────
#  CENTROS DE CUSTO  (sem alteração)
# ─────────────────────────────────────────────

CENTROS_CUSTO = {
    1:  "SERVIÇOS PARTICULARES",
    3:  "COLHEITA COOPERVAL",
    4:  "VALE DO IVAI (RENUKA)",
    5:  "NOVA PRODUTIVA COLHEITA",
    7:  "TRANSPORTE",
    11: "TESTON",
    14: "AGRO CIANORTE",
    16: "ESTOQUE",
    21: "ASSISTÊNCIA ELÉTRICA",
    22: "ASSISTÊNCIA TÉCNICA TESTON",
    23: "PLANTIO MECÂNICO",
    28: "AGRO VALE DO IVAI",
    29: "RIO AMAMBAI COLHEITA",
    37: "AGRO NAVIRAÍ",
    38: "AGRO ASTORGA PLANTIO/CUL",
    41: "RAIZEN",
    42: "CONCESSIONÁRIA SERTÃOZINHO",
    44: "SOL NASCENTE",
    46: "SERVIÇOS DE PREPARO DE SOLO",
    49: "AGRO SANTA CÂNDIDA",
    50: "LOBO GUARÁ",
    51: "COLHEITA COGO",
    54: "AGRO RORAIMA",
    100: "METALCANA",
}

NIVEIS_TECNICOS = list(TABELA_HORA.keys())

# ─────────────────────────────────────────────
#  TIPO E ORIGEM DA OS  (novo)
# ─────────────────────────────────────────────
#  tipo_os define QUEM pode abrir; origem registra DE ONDE veio.
#  A matriz abaixo é a única fonte da regra — a UI apenas a consulta.

TIPOS_OS   = ("interna", "externa")
ORIGENS_OS = ("manual", "checklist")

# Locais de serviço considerados internos (para consistência de lançamento)
LOCAIS_INTERNOS = ("interno barracão",)

# Locais de serviço. Deslocamento SAIU daqui: deixou de ser um local e
# virou uma natureza de lançamento própria (ver NATUREZAS abaixo), com
# campos e cálculo próprios. O tipo da OS (interna/externa) é DERIVADO
# do local, não escolhido à parte.
LOCAIS_OS = ("interno barracão", "campo", "m.s")

LABELS_LOCAIS_OS = {
    "interno barracão": "🏠 Interna (barracão)",
    "campo":            "🚜 Campo",
    "m.s":              "🌎 M.S",
}

# ─────────────────────────────────────────────
#  NATUREZA DO LANÇAMENTO
# ─────────────────────────────────────────────
# Um lançamento é um SERVIÇO (tem frota, tipo de serviço e local) ou um
# DESLOCAMENTO (tem cidades de origem/destino, trajeto de ida e de volta,
# e não tem tipo de serviço). Os dois convivem dentro da mesma OS.

NATUREZAS = ("servico", "deslocamento", "lavagem")

LABELS_NATUREZA = {
    "servico":      "🔧 Serviço",
    "deslocamento": "🚚 Deslocamento",
    "lavagem":      "🧼 Lavagem",
}

# Uma OS pode ser aberta como deslocamento puro — sem frota, só com
# centro de custo. Nesse caso só aceita lançamentos de deslocamento.
NATUREZA_OS_PADRAO = "servico"


def tipo_os_do_local(local: str) -> str:
    """'interno barracão' → 'interna'; qualquer outro local → 'externa'.

    Concentrar a derivação aqui evita que a tela e a persistência discordem
    sobre o que conta como OS interna.
    """
    return "interna" if (local or "").strip().lower() in LOCAIS_INTERNOS else "externa"


def _normalizar_txt(txt: str) -> str:
    import unicodedata
    txt = unicodedata.normalize("NFKD", str(txt or ""))
    txt = "".join(c for c in txt if not unicodedata.combining(c))
    return " ".join(txt.upper().split())


def resolver_cod_cc(cc_nome: str) -> tuple[int | None, str]:
    """
    Converte o nome do centro de custo (como vem da lista KanbanFrotas,
    ex: '038 - AGRO ASTORGA') no cod_cc inteiro da oficina.

    A comparação por NOME vem primeiro de propósito: o prefixo numérico
    das listas do Kanban não é necessariamente o mesmo código do centro de
    custo daqui. Só depois, e só se o nome conferir, o prefixo é aceito.

    Devolve (None, texto_original) quando não reconhece — quem chama decide
    se bloqueia ou se marca como pendente.
    """
    bruto = str(cc_nome or "").strip()
    if not bruto:
        return None, ""

    alvo = _normalizar_txt(bruto)
    prefixo, _, resto = alvo.partition("-")
    sem_prefixo = resto.strip() if prefixo.strip().isdigit() else alvo

    for cod, nome in CENTROS_CUSTO.items():
        if _normalizar_txt(nome) in (alvo, sem_prefixo):
            return cod, nome

    if prefixo.strip().isdigit():
        cod = int(prefixo.strip())
        nome_oficial = CENTROS_CUSTO.get(cod)
        if nome_oficial and _normalizar_txt(nome_oficial)[:8] == sem_prefixo[:8]:
            return cod, nome_oficial

    return None, bruto


# Técnicos autorizados NOMINALMENTE a abrir OS interna (barracão).
# É uma exceção à regra por perfil: técnico fora desta lista não abre interna.
# Para autorizar outro, basta acrescentar o login aqui.
TECNICOS_OS_INTERNA = ("belote",)


def pode_abrir_os(perfil: str, tipo_os: str, origem: str = "manual",
                  usuario: str = "") -> bool:
    """
    Regra de permissão de abertura de OS.

      - OS vinda do checklist: sempre permitida (é interna por definição).
      - OS externa: técnico, supervisor e admin podem abrir.
      - OS interna: supervisor, admin e os técnicos de TECNICOS_OS_INTERNA.

    'usuario' é o login (chave de USUARIOS), não o nome de exibição — a
    exceção é por pessoa, então o perfil sozinho não basta para decidir.
    """
    if origem == "checklist":
        return True
    if tipo_os == "externa":
        return perfil in ("tecnico", "supervisor", "admin")
    if tipo_os == "interna":
        if perfil in ("supervisor", "admin"):
            return True
        return (perfil == "tecnico"
                and (usuario or "").strip().lower() in TECNICOS_OS_INTERNA)
    return False


def motivo_bloqueio_abertura(perfil: str, tipo_os: str, origem: str = "manual",
                             usuario: str = "") -> str:
    """Mensagem explicando por que a abertura foi negada (para exibir na UI)."""
    if tipo_os not in TIPOS_OS:
        return f"Tipo de OS inválido: '{tipo_os}'. Use 'interna' ou 'externa'."
    if origem not in ORIGENS_OS:
        return f"Origem de OS inválida: '{origem}'."
    if tipo_os == "interna" and perfil == "tecnico":
        return ("OS interna (barracão) só pode ser aberta pelo técnico autorizado, "
                "pelo supervisor ou pelo administrador. Se o veículo passou por "
                "checklist, abra a OS pelo app de Checklist.")
    return "Seu perfil não tem permissão para abrir esta OS."

# ─────────────────────────────────────────────
#  EQUIPAMENTOS  (sem alteração)
# ─────────────────────────────────────────────

EQUIPAMENTOS = [
    "COLHEDORA JOHN DEERE CH570",
    "COLHEDORA JOHN DEERE CH 570",
    "COLHEDORA JOHN DEERE 3520",
    "COLHEDORA AGNES 367K",
    "NOVA AGNES",
    "TRATOR NEW HOLLAND T7.245",
    "TRATOR NEW HOLLAND TL 75E",
    "TRATOR NEW HOLLAND TL5.100",
    "TRATOR NEW HOLLAND TL 95E",
    "TRATOR MASSEY FERGUSON 275 4X2",
    "GIGANTE TRACTOR",
    "VOLVO VM 330 6X4",
    "VOLVO VM 330 6X4R",
    "VOLVO VM 330 8X2R",
    "TRANSBORDO GIGANTE 22000",
    "PA CARREGADEIRA 938K",
    "PA CARREGADEIRA 920K",
    "PULVERIZADOR AUTOPROPELIDO PLA 125J",
    "RAPTOR ARADO 16000S",
    "CANTERIZADOR PROTOTIPO RAPTOR",
    "KOMBI",
    "KOMBI 2013",
    "KOMBI LOTAÇÃO",
    "STRADA",
    "OROCH PRO 1.6",
    "GERADOR CATERPILLAR",
    "BAÚ OFICINA",
    "VM CAMINHÃO (COMBOIO)",
    "CARRETA PRANCHA",
]

# ─────────────────────────────────────────────
#  USUÁRIOS  (sem alteração)
# ─────────────────────────────────────────────

def _h(s):
    return hashlib.sha256(s.encode()).hexdigest()

USUARIOS = {
    "admin": {
        "nome": "Administrador",
        "perfil": "admin",
        "senha": _h("admin123"),
        "cod_tecnico": None,
    },
    "supervisor": {
        "nome": "Supervisor",
        "perfil": "supervisor",
        "senha": _h("super123"),
        "cod_tecnico": None,
    },
    "zaqueu":       {"nome": "Zaqueu",           "perfil": "tecnico", "senha": _h("zaqueu123"), "cod_tecnico": 1},
    "belote":       {"nome": "Cristiano B.",     "perfil": "tecnico", "senha": _h("teston123"), "cod_tecnico": 2},
    "rodrigo.j":    {"nome": "Rodrigo J.",       "perfil": "tecnico", "senha": _h("teston123"), "cod_tecnico": 3},
    "marcos.c":     {"nome": "Marcos C.",        "perfil": "tecnico", "senha": _h("teston123"), "cod_tecnico": 4},
    "cristiano.s":  {"nome": "Cristiano dos S.", "perfil": "tecnico", "senha": _h("teston123"), "cod_tecnico": 5},
    "joao.p":       {"nome": "João P.",          "perfil": "tecnico", "senha": _h("teston123"), "cod_tecnico": 6},
    "loran":        {"nome": "Loran H.",         "perfil": "tecnico", "senha": _h("teston123"), "cod_tecnico": 7},
    "bruno.h":      {"nome": "Bruno H.",         "perfil": "tecnico", "senha": _h("teston123"), "cod_tecnico": 8},
    "gabriel":      {"nome": "Gabriel",          "perfil": "tecnico", "senha": _h("teston123"), "cod_tecnico": 9},
    "tiago":        {"nome": "Tiago Bardu",      "perfil": "tecnico", "senha": _h("teston123"), "cod_tecnico": 10},
    "felipe":       {"nome": "Felipe",           "perfil": "tecnico", "senha": _h("teston123"), "cod_tecnico": 11},
    "gustavo":      {"nome": "Gustavo",          "perfil": "tecnico", "senha": _h("teston123"), "cod_tecnico": 12},
    "vanderlei":    {"nome": "Vanderlei",        "perfil": "tecnico", "senha": _h("teston123"), "cod_tecnico": 13},
    "claudecir":    {"nome": "Claudecir",        "perfil": "tecnico", "senha": _h("teston123"), "cod_tecnico": 14},
    "rodrigo.m":    {"nome": "Rodrigo M.",       "perfil": "tecnico", "senha": _h("teston123"), "cod_tecnico": 15},
    "jose":         {"nome": "José",             "perfil": "tecnico", "senha": _h("teston123"), "cod_tecnico": 16},
    "luiz":         {"nome": "Luiz",             "perfil": "tecnico", "senha": _h("teston123"), "cod_tecnico": 17},
}

# ═══════════════════════════════════════════════════════════════
#  SHAREPOINT — MS Graph API
#  Credenciais lidas de st.secrets (local: .streamlit/secrets.toml)
# ═══════════════════════════════════════════════════════════════

def _cfg(key: str) -> str:
    return st.secrets[key]

def _base_url() -> str:
    return (
        "https://graph.microsoft.com/v1.0"
        f"/sites/{_cfg('SITE_ID')}"
        f"/lists/{_cfg('LIST_ID')}"
        "/items"
    )

# ── Token OAuth2 (client credentials) ──────────────────────────

@st.cache_data(ttl=3000)   # renova ~50 min (token válido 60 min)
def _get_token() -> str:
    url = (
        f"https://login.microsoftonline.com"
        f"/{_cfg('TENANT_ID')}/oauth2/v2.0/token"
    )
    resp = requests.post(url, data={
        "grant_type":    "client_credentials",
        "client_id":     _cfg("CLIENT_ID"),
        "client_secret": _cfg("CLIENT_SECRET"),
        "scope":         "https://graph.microsoft.com/.default",
    }, timeout=15)
    resp.raise_for_status()
    return resp.json()["access_token"]

def _headers() -> dict:
    return {
        "Authorization": f"Bearer {_get_token()}",
        "Content-Type":  "application/json",
        # Necessário para $filter em colunas não indexadas (ex: fields/Title).
        # Sem isso o Graph pode retornar 400 Bad Request de forma intermitente.
        "Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly",
    }

# ── Helpers internos ────────────────────────────────────────────

def _parse_item(item: dict) -> dict | None:
    """Converte item SharePoint em dict de OS."""
    try:
        dados_json = item["fields"].get("DadosOS", "")
        if not dados_json:
            return None
        os_dict = json.loads(dados_json)
        os_dict["_sp_item_id"] = item["id"]   # ID interno do SP (para updates)
        # cod_cc é int quando existe, mas pode ser None em OS vinda do
        # checklist cujo centro de custo não foi reconhecido (cc_pendente).
        # int(None) levantaria TypeError e faria a OS sumir da lista.
        if os_dict.get("cod_cc") is not None:
            try:
                os_dict["cod_cc"] = int(os_dict["cod_cc"])
            except (TypeError, ValueError):
                os_dict["cod_cc"]     = None
                os_dict["cc_pendente"] = True
        return os_dict
    except Exception:
        return None

def _raise_com_detalhe(r: requests.Response) -> None:
    """Levanta HTTPError mostrando o motivo real retornado pelo Graph API,
    em vez de deixar a mensagem genérica do Streamlit Cloud esconder tudo."""
    if not r.ok:
        try:
            detalhe = r.json().get("error", {}).get("message", r.text[:300])
        except Exception:
            detalhe = r.text[:300]
        st.error(f"Graph API — erro {r.status_code}: {detalhe}")
    r.raise_for_status()

def _fetch_all() -> list[dict]:
    """Busca todos os itens da lista com paginação automática."""
    url = _base_url() + "?$expand=fields&$top=999"
    items = []
    while url:
        r = requests.get(url, headers=_headers(), timeout=20)
        _raise_com_detalhe(r)
        data = r.json()
        for item in data.get("value", []):
            parsed = _parse_item(item)
            if parsed:
                items.append(parsed)
        url = data.get("@odata.nextLink")
    return items

def _escapar_odata(valor: str) -> str:
    """Escapa aspas simples para uso dentro de um $filter OData."""
    return str(valor).replace("'", "''")


def _itens_brutos_por_titulo(numero_os: str) -> list[dict]:
    """Retorna TODOS os itens do SharePoint com este Title (normalmente 0 ou 1).
    Mais de um significa colisão de numeração — tratada em _post_com_numero_unico."""
    url = (
        _base_url()
        + f"?$expand=fields&$filter=fields/Title eq '{_escapar_odata(numero_os)}'"
    )
    r = requests.get(url, headers=_headers(), timeout=15)
    _raise_com_detalhe(r)
    return r.json().get("value", [])


def _itens_brutos_por_prefixo(prefixo: str) -> list[dict]:
    """Itens cujo Title começa com o prefixo (ex: 'OS-2026-').
    Evita varrer a lista inteira só para descobrir o próximo número.
    Se o SharePoint recusar o startswith, cai para a varredura completa."""
    url = (
        _base_url()
        + f"?$expand=fields&$top=999"
        + f"&$filter=startswith(fields/Title,'{_escapar_odata(prefixo)}')"
    )
    itens = []
    try:
        while url:
            r = requests.get(url, headers=_headers(), timeout=20)
            r.raise_for_status()
            data = r.json()
            itens.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
        return itens
    except requests.HTTPError:
        # Fallback: lista completa, filtrando na memória.
        url = _base_url() + "?$expand=fields&$top=999"
        itens = []
        while url:
            r = requests.get(url, headers=_headers(), timeout=20)
            _raise_com_detalhe(r)
            data = r.json()
            itens.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
        return [i for i in itens
                if str(i.get("fields", {}).get("Title", "")).startswith(prefixo)]


def _fetch_by_numero(numero_os: str) -> tuple[str | None, dict | None]:
    """Retorna (sp_item_id, os_dict) para um numero_os."""
    items = _itens_brutos_por_titulo(numero_os)
    if not items:
        return None, None
    item = items[0]
    return item["id"], _parse_item(item)

def _patch(sp_item_id: str, os_dict: dict) -> None:
    """Salva alterações em um item existente."""
    # Remove chave interna antes de serializar
    os_clean = {k: v for k, v in os_dict.items() if k != "_sp_item_id"}
    url = _base_url() + f"/{sp_item_id}/fields"
    payload = {
        "DadosOS": json.dumps(os_clean, ensure_ascii=False),
        "Status":  os_dict.get("status", ""),
    }
    # O Title precisa acompanhar o numero_os — sem isso, uma renumeração
    # por colisão deixaria o Title antigo e a busca por número quebraria.
    if os_dict.get("numero_os"):
        payload["Title"] = os_dict["numero_os"]
    r = requests.patch(url, headers=_headers(), json=payload, timeout=15)
    _raise_com_detalhe(r)

def _post(os_dict: dict) -> str:
    """Cria novo item e retorna o sp_item_id."""
    os_clean = {k: v for k, v in os_dict.items() if k != "_sp_item_id"}
    payload = {
        "fields": {
            "Title":   os_dict["numero_os"],
            "DadosOS": json.dumps(os_clean, ensure_ascii=False),
            "Status":  os_dict.get("status", ""),
        }
    }
    r = requests.post(_base_url(), headers=_headers(), json=payload, timeout=15)
    _raise_com_detalhe(r)
    return r.json()["id"]

def _proximo_numero_livre(ano: int) -> str:
    """
    Próximo número livre do ano, lido DIRETO do SharePoint (sem cache).

    A versão anterior lia de carregar_os(), que tem cache de 30 s — com dois
    apps escrevendo (oficina + checklist), isso gerava números repetidos.
    """
    prefixo = f"OS-{ano}-"
    max_n = 0
    for item in _itens_brutos_por_prefixo(prefixo):
        titulo = str(item.get("fields", {}).get("Title", ""))
        try:
            max_n = max(max_n, int(titulo.rsplit("-", 1)[-1]))
        except ValueError:
            pass
    return f"{prefixo}{(max_n + 1):04d}"


def _post_com_numero_unico(os_dict: dict, tentativas: int = 8) -> tuple[str, str]:
    """
    Cria a OS garantindo número único, mesmo com dois apps gravando ao
    mesmo tempo. Retorna (sp_item_id, numero_os_final).

    Como funciona o desempate, sem precisar de lock:
      1. Grava com o próximo número livre.
      2. Reconsulta o Title. Se só existe um item, terminou.
      3. Se existem dois (colisão), quem tem o MENOR id do SharePoint fica
         com o número — o id é atômico e único, então os dois lados chegam
         à mesma conclusão sem se falarem.
      4. Quem perdeu o desempate assume o próximo número livre e regrava.
    """
    ano = datetime.now().year
    os_dict["numero_os"] = _proximo_numero_livre(ano)
    sp_id = _post(os_dict)

    for _ in range(tentativas):
        conflitantes = _itens_brutos_por_titulo(os_dict["numero_os"])
        if len(conflitantes) <= 1:
            return sp_id, os_dict["numero_os"]

        vencedor = min(conflitantes, key=lambda i: int(i["id"]))
        if str(vencedor["id"]) == str(sp_id):
            return sp_id, os_dict["numero_os"]

        # Perdemos o desempate: pega o próximo livre e regrava este item.
        os_dict["numero_os"] = _proximo_numero_livre(ano)
        _patch(sp_id, os_dict)

    raise RuntimeError(
        "Não foi possível obter um número de OS único após várias tentativas. "
        "Verifique se há gravações concorrentes na lista AT_Teston_OS."
    )

# ═══════════════════════════════════════════════════════════════
#  API PÚBLICA  —  mesma interface do dados.py original
# ═══════════════════════════════════════════════════════════════

@st.cache_data(ttl=30)
def carregar_os() -> list:
    """
    Retorna todas as OS como lista de dicts.
    Cache 30 s — invalidado por .clear() após cada mutação.
    """
    items = _fetch_all()
    return sorted(items, key=lambda x: x.get("data_abertura", ""), reverse=True)

def buscar_os_por_numero(numero_os: str) -> dict | None:
    """Retorna a OS pelo número (ex: OS-2026-0001) ou None."""
    _, os_dict = _fetch_by_numero(numero_os)
    return os_dict

class PermissaoNegada(Exception):
    """Levantada quando o perfil não pode abrir o tipo de OS solicitado."""


def criar_os(frota: str, equipamento: str, cod_cc: int | None,
             aberto_por: str, data_abertura: str | None = None,
             tecnico_designado: int | None = None,
             *,
             perfil_usuario: str = "admin",
             tipo_os: str = "interna",
             local_previsto: str = "",
             natureza_os: str = "servico",
             avaliacao: str = "",
             origem: str = "manual",
             checklist_id: str | None = None,
             cliente_texto: str | None = None) -> dict:
    """Cria uma nova OS (sem serviços ainda) e salva no SharePoint.

    tecnico_designado: cod_tecnico (chave de TECNICOS) responsável por esta OS.
    Enquanto a OS estiver aberta/em_andamento, apenas este técnico (além de
    supervisor/admin) pode visualizá-la para lançar serviços ou encerrá-la.

    perfil_usuario / tipo_os / origem: governam a permissão de abertura.
    A validação vive AQUI, e não só na tela, porque o app de checklist chama
    esta função direto — esconder o formulário no app.py não protegeria nada.

    avaliacao: descrição livre do problema apresentado pelo veículo.
    checklist_id: id do item em ChecklistVeiculos que originou a OS.
    cliente_texto: nome do CC quando cod_cc não pôde ser resolvido (deixa a
    OS rastreável em vez de gravar cliente vazio).
    """
    # OS de deslocamento puro: não tem frota nem local de serviço, e é
    # sempre externa — ninguém se desloca para o próprio barracão.
    if natureza_os == "deslocamento":
        tipo_os        = "externa"
        local_previsto = ""
        frota          = (frota or "").strip()
        if not equipamento:
            equipamento = "Deslocamento"
    # Caso normal: o local manda, e o tipo é derivado dele.
    elif local_previsto:
        tipo_os = tipo_os_do_local(local_previsto)

    # aberto_por já é o login de quem está abrindo — a exceção nominal da OS
    # interna é decidida aqui, e não só na tela (o checklist chama direto).
    if not pode_abrir_os(perfil_usuario, tipo_os, origem, aberto_por):
        raise PermissaoNegada(
            motivo_bloqueio_abertura(perfil_usuario, tipo_os, origem, aberto_por)
        )

    if origem == "checklist":
        tipo_os        = "interna"          # decisão de negócio: checklist é sempre interna
        local_previsto = "interno barracão"

    cliente = CENTROS_CUSTO.get(cod_cc, "") if cod_cc is not None else ""
    if not cliente and cliente_texto:
        cliente = cliente_texto

    nova = {
        "numero_os":                   None,   # definido por _post_com_numero_unico
        "data_abertura":               data_abertura or datetime.now().strftime("%Y-%m-%d"),
        "frota":                       frota,
        "equipamento":                 equipamento,
        "cod_cc":                      cod_cc,
        "cliente":                     cliente,
        "cc_pendente":                 cod_cc is None,
        "tipo_os":                     tipo_os,
        "natureza_os":                 natureza_os,
        "local_previsto":              local_previsto or "",
        "origem":                      origem,
        "checklist_id":                checklist_id,
        "avaliacao":                   (avaliacao or "").strip(),
        "status":                      "aberta",
        "aberto_por":                  aberto_por,
        "aberto_em":                   datetime.now().isoformat(),
        "tecnico_designado":           tecnico_designado,
        "fechado_para_aprovacao_por":  None,
        "fechado_para_aprovacao_em":   None,
        "validado_por":                None,
        "validado_em":                 None,
        "observacao_validacao":        "",
        "procedimentos":               [],
    }
    _, numero = _post_com_numero_unico(nova)
    nova["numero_os"] = numero
    carregar_os.clear()
    return nova


def registrar_avaliacao(numero_os: str, avaliacao: str, editado_por: str) -> bool:
    """Atualiza o campo Avaliação de uma OS ainda aberta/em andamento."""
    sp_id, os_dict = _fetch_by_numero(numero_os)
    if not sp_id:
        return False
    if os_dict["status"] not in ("aberta", "em_andamento"):
        return False
    os_dict["avaliacao"] = (avaliacao or "").strip()
    os_dict.setdefault("historico_avaliacao", []).append({
        "editado_por": editado_por,
        "editado_em":  datetime.now().isoformat(),
    })
    _patch(sp_id, os_dict)
    carregar_os.clear()
    return True

def definir_tecnico_designado(numero_os: str, cod_tecnico: int | None,
                               definido_por: str) -> bool:
    """
    Define ou altera o técnico designado de uma OS já existente
    (usado pelo supervisor/admin — inclusive para corrigir OS legadas
    sem técnico designado). Só permitido enquanto a OS está aberta/em_andamento.
    """
    sp_id, os_dict = _fetch_by_numero(numero_os)
    if not sp_id:
        return False
    if os_dict["status"] not in ("aberta", "em_andamento"):
        return False
    os_dict["tecnico_designado"] = cod_tecnico
    os_dict.setdefault("historico_designacao", []).append({
        "cod_tecnico":  cod_tecnico,
        "definido_por": definido_por,
        "definido_em":  datetime.now().isoformat(),
    })
    _patch(sp_id, os_dict)
    carregar_os.clear()
    return True

def adicionar_procedimento(numero_os: str, proc: dict) -> bool:
    """
    Adiciona um serviço a uma OS existente (aberta ou em_andamento).
    Importado pelo app.py como:  adicionar_procedimento as adicionar_servico
    """
    sp_id, os_dict = _fetch_by_numero(numero_os)
    if not sp_id:
        return False
    if os_dict["status"] not in ("aberta", "em_andamento"):
        return False
    proc["seq"]     = len(os_dict["procedimentos"]) + 1
    proc["proc_id"] = datetime.now().isoformat()   # mantido igual ao original
    os_dict["procedimentos"].append(proc)
    os_dict["status"] = "em_andamento"
    _patch(sp_id, os_dict)
    carregar_os.clear()
    return True

def editar_procedimento(numero_os: str, proc_id: str, proc_atualizado: dict) -> bool:
    """
    Substitui um serviço (por proc_id) numa OS aberta/em_andamento ou
    aguardando_aprovacao. O status aguardando_aprovacao é liberado para o
    ajuste de tempo que o adm/supervisor faz na própria tela de validação
    (quando o alerta de tempo acima da média não se justifica). O controle
    de acesso é feito na UI (página de validação é só de gestão). A OS
    permanece em aguardando_aprovacao após o ajuste — não sai da fila.
    Importado pelo app.py como:  editar_procedimento as editar_servico
    """
    sp_id, os_dict = _fetch_by_numero(numero_os)
    if not sp_id:
        return False
    if os_dict["status"] not in ("aberta", "em_andamento", "aguardando_aprovacao"):
        return False
    for idx, p in enumerate(os_dict["procedimentos"]):
        if p.get("proc_id") == proc_id:
            proc_atualizado["seq"]     = p["seq"]
            proc_atualizado["proc_id"] = proc_id
            os_dict["procedimentos"][idx] = proc_atualizado
            _patch(sp_id, os_dict)
            carregar_os.clear()
            return True
    return False

def excluir_procedimento(numero_os: str, proc_id: str, excluido_por: str) -> bool:
    """
    Remove um serviço (por proc_id) de uma OS em aberta/em_andamento.
    Controle de acesso (apenas admin) é feito na camada de UI (app.py).
    Renumera os 'seq' restantes e, se a OS ficar sem serviços,
    volta o status para 'aberta'.
    """
    sp_id, os_dict = _fetch_by_numero(numero_os)
    if not sp_id:
        return False
    if os_dict["status"] not in ("aberta", "em_andamento"):
        return False
    procs = os_dict.get("procedimentos", [])
    restantes = [p for p in procs if p.get("proc_id") != proc_id]
    if len(restantes) == len(procs):
        return False   # proc_id não encontrado

    for i, p in enumerate(restantes, start=1):
        p["seq"] = i
    os_dict["procedimentos"] = restantes
    if not restantes:
        os_dict["status"] = "aberta"

    os_dict.setdefault("historico_exclusao", []).append({
        "proc_id":      proc_id,
        "excluido_por": excluido_por,
        "excluido_em":  datetime.now().isoformat(),
    })
    _patch(sp_id, os_dict)
    carregar_os.clear()
    return True

def fechar_os_para_aprovacao(numero_os: str, fechado_por: str) -> bool:
    """Muda status: em_andamento → aguardando_aprovacao."""
    sp_id, os_dict = _fetch_by_numero(numero_os)
    if not sp_id:
        return False
    if os_dict["status"] != "em_andamento":
        return False
    os_dict["status"]                     = "aguardando_aprovacao"
    os_dict["fechado_para_aprovacao_por"] = fechado_por
    os_dict["fechado_para_aprovacao_em"]  = datetime.now().isoformat()
    _patch(sp_id, os_dict)
    carregar_os.clear()
    return True

def reabrir_os(numero_os: str, reaberto_por: str) -> bool:
    """Supervisor/Admin reabre OS de aguardando_aprovacao → em_andamento."""
    sp_id, os_dict = _fetch_by_numero(numero_os)
    if not sp_id:
        return False
    if os_dict["status"] != "aguardando_aprovacao":
        return False
    os_dict["status"]                     = "em_andamento"
    os_dict["fechado_para_aprovacao_por"] = None
    os_dict["fechado_para_aprovacao_em"]  = None
    os_dict.setdefault("historico_reabertura", []).append({
        "reaberto_por": reaberto_por,
        "reaberto_em":  datetime.now().isoformat(),
    })
    _patch(sp_id, os_dict)
    carregar_os.clear()
    return True

def validar_os(numero_os: str, novo_status: str,
               validado_por: str, observacao: str = "") -> bool:
    """Aprova ou rejeita OS em aguardando_aprovacao."""
    sp_id, os_dict = _fetch_by_numero(numero_os)
    if not sp_id:
        return False
    if os_dict["status"] != "aguardando_aprovacao":
        return False
    os_dict["status"]               = novo_status   # "aprovada" | "rejeitada"
    os_dict["validado_por"]         = validado_por
    os_dict["validado_em"]          = datetime.now().isoformat()
    os_dict["observacao_validacao"] = observacao
    _patch(sp_id, os_dict)
    carregar_os.clear()
    return True

# ── Funções compat (legacy — mantidas para compatibilidade) ────

def salvar_os(nova_os: dict) -> None:
    """Compat: insere OS diretamente (sem validação de status)."""
    sp_id, _ = _fetch_by_numero(nova_os.get("numero_os", ""))
    if sp_id:
        _patch(sp_id, nova_os)
    else:
        _post(nova_os)
    carregar_os.clear()

def atualizar_status_os(criado_em: str, novo_status: str,
                         validado_por: str, observacao: str = "") -> None:
    """Compat: atualiza campo status pela chave criado_em."""
    for os_dict in _fetch_all():
        if os_dict.get("criado_em") == criado_em:
            sp_id = os_dict.pop("_sp_item_id", None)
            os_dict["status"]               = novo_status
            os_dict["validado_por"]         = validado_por
            os_dict["validado_em"]          = datetime.now().isoformat()
            os_dict["observacao_validacao"] = observacao
            if sp_id:
                _patch(sp_id, os_dict)
            carregar_os.clear()
            return

def editar_os(criado_em: str, os_atualizada: dict) -> None:
    """Compat: substitui OS pela chave criado_em."""
    for os_dict in _fetch_all():
        if os_dict.get("criado_em") == criado_em:
            if os_dict.get("status") != "pendente":
                raise ValueError("Apenas OS pendentes podem ser editadas.")
            sp_id = os_dict.pop("_sp_item_id", None)
            if sp_id:
                _patch(sp_id, os_atualizada)
            carregar_os.clear()
            return