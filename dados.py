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
#  TABELAS DE VALORES  (sem alteração)
# ─────────────────────────────────────────────

TABELA_KM = {
    "Km Um":   1.70,
    "Km Dois": 4.10,
    "Km Três": 5.00,
}

TABELA_HORA = {
    "Técnico Um":     50,
    "Técnico Dois":   60,
    "Técnico Três":   80,
    "Técnico Quatro": 90,
    "Técnico Cinco": 100,
}

HORA_DESLOCAMENTO = 60
HORA_MUNCK        = 120

PERCENTUAIS_LOCAL = {
    "interno barracão": 4,
    "campo":            8,
    "deslocamento":     6,
    "m.s":             10,
}

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
    "Deslocamento",
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
    "cristiano.b":  {"nome": "Cristiano B.",     "perfil": "tecnico", "senha": _h("teston123"), "cod_tecnico": 2},
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
        # garante que cod_cc é int (JSON deserializa como int normalmente)
        if "cod_cc" in os_dict:
            os_dict["cod_cc"] = int(os_dict["cod_cc"])
        return os_dict
    except Exception:
        return None

def _fetch_all() -> list[dict]:
    """Busca todos os itens da lista com paginação automática."""
    url = _base_url() + "?$expand=fields&$top=999"
    items = []
    while url:
        r = requests.get(url, headers=_headers(), timeout=20)
        r.raise_for_status()
        data = r.json()
        for item in data.get("value", []):
            parsed = _parse_item(item)
            if parsed:
                items.append(parsed)
        url = data.get("@odata.nextLink")
    return items

def _fetch_by_numero(numero_os: str) -> tuple[str | None, dict | None]:
    """Retorna (sp_item_id, os_dict) para um numero_os."""
    url = (
        _base_url()
        + f"?$expand=fields&$filter=fields/Title eq '{numero_os}'"
    )
    r = requests.get(url, headers=_headers(), timeout=15)
    r.raise_for_status()
    items = r.json().get("value", [])
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
    r = requests.patch(url, headers=_headers(), json=payload, timeout=15)
    r.raise_for_status()

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
    r.raise_for_status()
    return r.json()["id"]

def _gerar_numero_os() -> str:
    """Gera próximo número de OS do ano corrente (ex: OS-2026-0006)."""
    todas = carregar_os()
    ano   = datetime.now().year
    max_n = 0
    for o in todas:
        n = o.get("numero_os", "")
        if n.startswith(f"OS-{ano}-"):
            try:
                max_n = max(max_n, int(n.rsplit("-", 1)[-1]))
            except ValueError:
                pass
    return f"OS-{ano}-{(max_n + 1):04d}"

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

def criar_os(frota: str, equipamento: str, cod_cc: int,
             aberto_por: str, data_abertura: str | None = None) -> dict:
    """Cria uma nova OS (sem serviços ainda) e salva no SharePoint."""
    numero = _gerar_numero_os()
    nova = {
        "numero_os":                   numero,
        "data_abertura":               data_abertura or datetime.now().strftime("%Y-%m-%d"),
        "frota":                       frota,
        "equipamento":                 equipamento,
        "cod_cc":                      cod_cc,
        "cliente":                     CENTROS_CUSTO.get(cod_cc, ""),
        "status":                      "aberta",
        "aberto_por":                  aberto_por,
        "aberto_em":                   datetime.now().isoformat(),
        "fechado_para_aprovacao_por":  None,
        "fechado_para_aprovacao_em":   None,
        "validado_por":                None,
        "validado_em":                 None,
        "observacao_validacao":        "",
        "procedimentos":               [],
    }
    _post(nova)
    carregar_os.clear()
    return nova

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
    Substitui um serviço (por proc_id) numa OS em aberta/em_andamento.
    Importado pelo app.py como:  editar_procedimento as editar_servico
    """
    sp_id, os_dict = _fetch_by_numero(numero_os)
    if not sp_id:
        return False
    if os_dict["status"] not in ("aberta", "em_andamento"):
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
