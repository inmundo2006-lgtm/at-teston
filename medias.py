"""
medias.py — Alerta de tempo acima da média histórica + base viva.

Duas fontes se somam para formar a média de cada procedimento:
  1. SEMENTE  → medias_historicas.csv (dados passados, da planilha)
  2. APP      → serviços APROVADOS lançados no próprio aplicativo

A semente guarda estatísticas suficientes (n, soma, soma_quad); os aprovados do
app são agregados do mesmo jeito e somados. Assim a média evolui sozinha à
medida que a oficina passa a lançar só pelo app — sem recalcular a planilha.

Régua do alerta: dispara quando  horas > média + K_DESVIOS × desvio.
Com K=1, ~15% dos lançamentos caem acima. Para mudar, mexa só em K_DESVIOS.

"Mesmo procedimento":
  1º) Tipo de Serviço + Equipamento (família canônica), se o grupo tiver
      pelo menos MIN_AMOSTRAS lançamentos (somando semente + app);
  2º) Tipo de Serviço sozinho — fallback.
Sem histórico para o Tipo → sem alerta (nada a comparar).
"""

import csv
import math
import os
import re
import unicodedata

import streamlit as st

ARQUIVO = "medias_historicas.csv"
K_DESVIOS = 1.0          # média + 1 desvio-padrão (ajuste a régua aqui)
MIN_AMOSTRAS = 5         # nº mínimo para um grupo específico (Tipo+Equip) valer


# ── Canonizador (idêntico ao do gerador — manter iguais) ──────────

def _norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s or ""))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s.upper()).strip()

REGRAS_EQUIP = [
    ("TRANSBORDO GIGANTE",          ["TRANSBORDO"]),
    ("GIGANTE TRACTOR",             ["GIGANTE"]),
    ("COLHEDORA JOHN DEERE CH570",  ["CH570", "CH 570"]),
    ("COLHEDORA JOHN DEERE CH670",  ["CH670", "CH 670"]),
    ("COLHEDORA JOHN DEERE 3520",   ["3520"]),
    ("COLHEDORA AGNES 367K",        ["AGNES 367", "367K"]),
    ("NOVA AGNES",                  ["NOVA AGNES"]),
    ("COLHEDORA CANA INTEIRA",      ["CANA INTEIRA"]),
    ("TRATOR NEW HOLLAND T7.245",   ["T7.245", "T7245", "T7 245"]),
    ("TRATOR NEW HOLLAND TL95E",    ["TL95", "TL 95"]),
    ("TRATOR NEW HOLLAND TL75E",    ["TL75", "TL 75"]),
    ("TRATOR NEW HOLLAND TL5.100",  ["TL5.100", "TL5100", "TL 5.100"]),
    ("TRATOR MASSEY 275",           ["MASSEY", "275"]),
    ("CARREGADEIRA MASSEY 292",     ["292"]),
    ("PA CARREGADEIRA 938K",        ["938"]),
    ("PA CARREGADEIRA 920K",        ["920"]),
    ("PULVERIZADOR ORION 250",      ["ORION"]),
    ("PULVERIZADOR PLA 125J",       ["PLA 125", "PLA125"]),
    ("VOLVO VM 330",                ["VM 330", "VM330"]),
    ("RAPTOR",                      ["RAPTOR"]),
    ("KOMBI",                       ["KOMBI"]),
    ("STRADA",                      ["STRADA"]),
    ("OROCH",                       ["OROCH"]),
    ("S10",                         ["S10", "S-10"]),
    ("UNO",                         ["UNO"]),
    ("GERADOR CATERPILLAR",         ["GERADOR", "CATERPILLAR"]),
    ("BAU OFICINA",                 ["BAU"]),
    ("VM CAMINHAO COMBOIO",         ["COMBOIO"]),
    ("CARRETA PRANCHA",             ["PRANCHA"]),
]

def canonizar_equipamento(bruto: str) -> str:
    n = _norm(bruto)
    if not n:
        return ""
    for chave, tokens in REGRAS_EQUIP:
        if any(t in n for t in tokens):
            return chave
    return n


# ── Semente (CSV) ─────────────────────────────────────────────────

@st.cache_data
def _carregar_seed() -> dict:
    """{(tipo, equip): [n, soma, soma_quad]} a partir do CSV. {} se ausente."""
    caminho = os.path.join(os.path.dirname(__file__), ARQUIVO)
    if not os.path.exists(caminho):
        return {}
    seed = {}
    try:
        with open(caminho, encoding="utf-8", newline="") as f:
            for lin in csv.DictReader(f):
                chave = (lin["tipo_servico"].strip().lower(), lin["equipamento"].strip())
                seed[chave] = [int(lin["n"]), float(lin["soma"]), float(lin["soma_quad"])]
    except Exception:
        return {}
    return seed


# ── Agregação dos aprovados do app ────────────────────────────────

def _acumular(agg: dict, tipo: str, equip: str, h: float) -> None:
    for chave in ({(tipo, equip), (tipo, "")} if equip else {(tipo, "")}):
        d = agg.setdefault(chave, [0, 0.0, 0.0])
        d[0] += 1; d[1] += h; d[2] += h * h


def agregar_servicos_app(os_list: list) -> dict:
    """Soma os serviços APROVADOS do app (mesmo escopo do alerta) em stats suf."""
    agg = {}
    for o in (os_list or []):
        if o.get("status") != "aprovada":
            continue
        equip_canon = canonizar_equipamento(o.get("equipamento", ""))
        for p in o.get("procedimentos", []):
            if p.get("natureza", "servico") != "servico" or p.get("tempo_fixo_aplicado"):
                continue
            tipo = (p.get("tipo_servico") or "").strip().lower()
            if not tipo:
                continue
            h = float(p.get("horas_trabalhadas", 0) or 0)
            if h <= 0 or h > 24:
                continue
            _acumular(agg, tipo, equip_canon, h)
    return agg


def _combinar(seed: dict, app_agg: dict) -> dict:
    """Soma as duas fontes e deriva {(tipo,equip): {n, media, desvio}}."""
    tabela = {}
    for k in (set(seed) | set(app_agg)):
        n, soma, sq = 0, 0.0, 0.0
        for src in (seed, app_agg):
            if k in src:
                n += src[k][0]; soma += src[k][1]; sq += src[k][2]
        tipo, equip = k
        if equip and n < MIN_AMOSTRAS:   # grupo específico raro fica fora
            continue
        if n < 2:
            continue
        media = soma / n
        var = max(0.0, (sq - n * media * media) / (n - 1))
        tabela[k] = {"n": n, "media": media, "desvio": math.sqrt(var)}
    return tabela


@st.cache_data(ttl=30)
def tabela_ativa() -> dict:
    """
    Tabela combinada (semente + aprovados do app), pronta para consulta.
    Cache de 30 s, alinhado ao carregar_os(). Importa carregar_os na hora
    para não criar dependência circular no import.
    """
    try:
        from dados import carregar_os
        os_list = carregar_os()
    except Exception:
        os_list = []
    return _combinar(_carregar_seed(), agregar_servicos_app(os_list))


# ── Avaliação de um lançamento ────────────────────────────────────

def avaliar_tempo(tipo_servico: str, equipamento: str, horas: float,
                  tabela: dict | None = None) -> dict | None:
    """
    Avalia UM serviço. Devolve dict (ver campos abaixo) ou None quando não há
    histórico comparável. Passe 'tabela' para reaproveitar uma já construída
    (dashboard); sem ela, usa tabela_ativa() (cacheada).
    """
    if tabela is None:
        tabela = tabela_ativa()
    if not tabela or not horas or horas <= 0:
        return None

    tipo = (tipo_servico or "").strip().lower()
    equip = canonizar_equipamento(equipamento)

    grupo = tabela.get((tipo, equip)) if equip else None
    escopo = "equipamento"
    if grupo is None:
        grupo = tabela.get((tipo, ""))
        escopo = "tipo"
    if grupo is None:
        return None

    media, desvio, n = grupo["media"], grupo["desvio"], grupo["n"]
    limite = media + K_DESVIOS * desvio
    acima = horas > limite
    rotulo = f"{tipo_servico} · {equip}" if escopo == "equipamento" else f"{tipo_servico} (geral)"

    return {
        "acima":         acima,
        "horas":         round(horas, 2),
        "media":         round(media, 2),
        "desvio":        round(desvio, 2),
        "limite":        round(limite, 2),
        "n":             n,
        "excedente_h":   round(horas - limite, 2) if acima else 0.0,
        "excedente_pct": round((horas / media - 1) * 100, 0) if media else 0.0,
        "escopo":        escopo,
        "rotulo":        rotulo,
    }


# ── Ranking para o dashboard ──────────────────────────────────────

def ranking_acima_media(os_list: list, tabela: dict | None = None) -> list:
    """
    Histórico de alertas por técnico, sobre OS APROVADAS.

    Usa o carimbo gravado no lançamento (alerta_tempo): "disparou_original"
    conta o alerta MESMO que a gestão tenha corrigido o tempo depois — assim o
    histórico do mecânico não é apagado por um ajuste. "corrigidos" conta
    quantos desses foram ajustados pela gestão. Lançamentos antigos sem carimbo
    caem numa avaliação ao vivo contra a base atual.

    Devolve lista de dicts, ordenada por "acima" desc:
      {tecnico, acima, corrigidos, avaliados, pct}
    """
    contagem = {}   # tecnico -> [acima, corrigidos, avaliados]
    for o in (os_list or []):
        if o.get("status") != "aprovada":
            continue
        equip = o.get("equipamento", "")
        for p in o.get("procedimentos", []):
            if p.get("natureza", "servico") != "servico" or p.get("tempo_fixo_aplicado"):
                continue
            at = p.get("alerta_tempo") if isinstance(p.get("alerta_tempo"), dict) else None
            if at is not None and at.get("avaliavel"):
                disparou   = bool(at.get("disparou_original", at.get("disparou", False)))
                foi_corrig = bool(at.get("corrigido", False))
            else:
                # Legado sem carimbo: avalia ao vivo.
                h = float(p.get("horas_trabalhadas", 0) or 0)
                av = avaliar_tempo(p.get("tipo_servico", ""), equip, h, tabela=tabela)
                if av is None:
                    continue
                disparou, foi_corrig = av["acima"], False
            nome = p.get("nome_tecnico", "?")
            c = contagem.setdefault(nome, [0, 0, 0])
            c[2] += 1
            if disparou:
                c[0] += 1
            if foi_corrig:
                c[1] += 1

    ranking = [
        {"tecnico": t, "acima": v[0], "corrigidos": v[1], "avaliados": v[2],
         "pct": round(100 * v[0] / v[2], 1) if v[2] else 0.0}
        for t, v in contagem.items()
    ]
    ranking.sort(key=lambda d: (-d["acima"], -d["pct"]))
    return ranking
