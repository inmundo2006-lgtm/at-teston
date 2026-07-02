"""
tempos_fixos.py — Tabela de tempos fixos por componente (planilha TEMPO_SERVIÇO.xlsx)

Regra de negócio:
- Se o serviço executado for a troca/reparo de um COMPONENTE que está nesta tabela
  (combinação Equipamento + Componente), o tempo do serviço já é o TEMPO TOTAL
  definido na planilha — não é calculado por hora início/fim.
- Se não estiver na tabela, o tempo continua sendo hora_chegada − hora_saida
  (comportamento original do app).
- Deslocamento continua sendo lançado como um serviço à parte (com seu próprio
  horário), e soma normalmente ao total da OS — nada muda nisso.

Conversão de "dias" em horas:
  Jornada semanal de 44h distribuída em 5 dias (seg-sex, 7h-17h, 1h12 de almoço)
  = 8,8h por dia. Usado para textos como "1 DIA E MEIO", "DOIS DIAS", etc.
"""

import os
import re
import datetime as dt

import pandas as pd
import streamlit as st

HORAS_POR_DIA_UTIL = 44 / 5  # 8.8h — jornada semanal (CLT) distribuída em 5 dias

NENHUM_COMPONENTE = "— Nenhum (calcular por horário) —"

_NUM_PT = {
    "UM": 1, "UMA": 1,
    "DOIS": 2, "DUAS": 2,
    "TRÊS": 3, "TRES": 3,
    "QUATRO": 4, "CINCO": 5, "SEIS": 6,
    "SETE": 7, "OITO": 8, "NOVE": 9, "DEZ": 10,
}


def _caminho_planilha() -> str:
    base = os.path.dirname(__file__)
    for nome in ("TEMPO_SERVIÇO.xlsx", "TEMPO_SERVICO.xlsx"):
        caminho = os.path.join(base, nome)
        if os.path.exists(caminho):
            return caminho
    return os.path.join(base, "TEMPO_SERVIÇO.xlsx")


def _parse_tempo(valor) -> float:
    """Converte um valor da coluna TEMPO TOTAL em horas decimais.
    Aceita: datetime.time (HH:MM:SS do Excel), string 'HH:MM:SS',
    ou texto em dias ('1 DIA E MEIO', 'DOIS DIAS', 'TRÊS DIAS E MEIO')."""
    if valor is None:
        return 0.0
    if isinstance(valor, float) and pd.isna(valor):
        return 0.0
    if isinstance(valor, (dt.datetime, pd.Timestamp)):
        return valor.hour + valor.minute / 60 + valor.second / 3600
    if isinstance(valor, dt.time):
        return valor.hour + valor.minute / 60 + valor.second / 3600

    texto = str(valor).strip().upper()
    if not texto:
        return 0.0

    # "HH:MM:SS" ou "HH:MM" como string
    m = re.match(r'^(\d{1,3}):(\d{2})(?::(\d{2}))?$', texto)
    if m:
        h, mi, s = int(m.group(1)), int(m.group(2)), int(m.group(3) or 0)
        return h + mi / 60 + s / 3600

    # Texto em dias: "1 DIA E MEIO", "DOIS DIAS", "TRÊS DIAS E MEIO", "MEIO DIA"
    if texto.startswith("MEIO DIA"):
        return 0.5 * HORAS_POR_DIA_UTIL

    dias = 0.0
    alternativas = r'\d+|' + '|'.join(_NUM_PT.keys())
    mnum = re.search(r'(' + alternativas + r')\s*DIAS?', texto)
    if mnum:
        token = mnum.group(1)
        dias = float(token) if token.isdigit() else float(_NUM_PT.get(token, 0))
    if "MEIO" in texto or "MEIA" in texto:
        dias += 0.5

    return dias * HORAS_POR_DIA_UTIL


@st.cache_data
def carregar_tabela_tempos():
    """
    Retorna:
      tabela          -> dict {(equipamento, componente): horas_decimais}
      por_equipamento -> dict {equipamento: [componentes...]}
    Ambos vazios se a planilha não for encontrada.
    """
    caminho = _caminho_planilha()
    if not os.path.exists(caminho):
        return {}, {}

    try:
        df = pd.read_excel(caminho, header=1)  # linha 1 (0-indexed) é o cabeçalho real
        df = df.iloc[:, :5]
        df.columns = ["equipamento", "componente", "desmontagem", "montagem", "total"]
        df = df.dropna(subset=["equipamento", "componente"])
        df["equipamento"] = df["equipamento"].astype(str).str.strip().str.upper()
        df["componente"]  = df["componente"].astype(str).str.strip().str.upper()
    except Exception:
        return {}, {}

    tabela = {}
    por_equipamento = {}
    for _, r in df.iterrows():
        eq, comp = r["equipamento"], r["componente"]
        horas = _parse_tempo(r["total"])
        tabela[(eq, comp)] = horas
        por_equipamento.setdefault(eq, []).append(comp)

    for eq in por_equipamento:
        por_equipamento[eq] = sorted(set(por_equipamento[eq]))

    return tabela, por_equipamento


def formatar_horas(horas: float) -> str:
    """Ex: 22.0 -> '22h'; 13.2 -> '13h12min'."""
    total_min = round(horas * 60)
    hh, mm = divmod(total_min, 60)
    return f"{hh}h{mm:02d}min" if mm else f"{hh}h"
