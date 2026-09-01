"""
calculos.py — Motor de cálculo de comissões

Reescrito a partir das FÓRMULAS da planilha Relatório Assist. Téc.
(abas OS, Valores e Comissão), não do README — que descrevia regras
divergentes das que a planilha realmente aplica.

═══════════════════════════════════════════════════════════════════════
REGRAS (conferidas contra 570 lançamentos reais)
═══════════════════════════════════════════════════════════════════════

1. VALOR HORA POR NÍVEL — aba "Valores", células D3:D7
   A aba "Inf." tem uma segunda tabela (50/60/80/90/100) que NÃO é usada
   por fórmula nenhuma. Foi dela que a versão anterior do sistema copiou
   os valores, e por isso calculava a menos. A tabela válida é esta:

       Técnico Um 55 · Dois 80 · Três 90 · Quatro 100 · Cinco 105

2. HORAS TRABALHADAS (coluna AF)
       = Hora Final − Hora Inicial − Almoço − Café
   O sistema desconta automaticamente 1h12 de almoço + 15min de café
   quando o serviço atravessa o meio-dia. Quem parar mais ou menos que
   isso ajusta o horário de saída para compensar (convenção do time).

3. VALOR DO SERVIÇO (coluna AR)
       = horas líquidas × valor hora do nível

4. VALOR KM (coluna AO)
       = km rodado × preço do tipo de KM do técnico

5. VALOR MUNCK (coluna AP)
       = horas munck × R$120

6. DESLOCAMENTO (coluna AQ)
       = (horas de ida + horas de retorno) × R$60
   R$60 é fixo, independe do nível do técnico. Não desconta intervalo:
   trajeto é trajeto.

7. COMISSÃO (colunas AT, AU, AV, AW)
       Campo            → valor do serviço × 8%
       Interno Barracão → valor do serviço × 4%
       M.S              → valor do serviço × 10%
       Deslocamento     → valor do deslocamento × 6%

   ATENÇÃO: a comissão incide SÓ sobre o valor das horas. KM e Munck
   entram no valor cobrado do cliente (coluna AS), mas ficam FORA da
   base da comissão. A versão anterior incluía os dois e inflava o
   pagamento em ~8,6%.

8. VALOR TOTAL COBRADO (coluna AS)
       = valor serviço + valor KM + valor munck + valor deslocamento
═══════════════════════════════════════════════════════════════════════
"""

# ─────────────────────────────────────────────
#  CONSTANTES  (espelham a aba "Valores")
# ─────────────────────────────────────────────

TABELA_KM = {
    "Km Um":   1.70,   # Valores!B3
    "Km Dois": 4.10,   # Valores!B4
    "Km Três": 5.00,   # Valores!B5
}

# Valores!D3:D7 — esta é a tabela que as fórmulas usam.
TABELA_HORA = {
    "Técnico Um":     55.0,
    "Técnico Dois":   80.0,
    "Técnico Três":   90.0,
    "Técnico Quatro": 100.0,
    "Técnico Cinco":  105.0,
}

HORA_DESLOCAMENTO = 60.0    # Valores!D8 — fixo, independe do nível
HORA_MUNCK        = 120.0   # Valores!B6

# Lavagem — valor fixo por tipo de veículo. "Outros" NÃO está aqui: cai na
# regra de tempo (horas × valor-hora do nível), porque varia demais
# (colhedora, implemento agrícola, etc.).
VALOR_LAVAGEM = {
    "carro":  80.0,
    "onibus": 250.0,
}

PERCENTUAIS_LOCAL = {
    "interno barracão": 0.04,
    "campo":            0.08,
    "m.s":              0.10,
}
PERCENTUAL_DESLOCAMENTO = 0.06

# ─────────────────────────────────────────────
#  INTERVALOS
# ─────────────────────────────────────────────
# Descontados juntos quando o serviço atravessa o meio-dia.
# Se o café precisar de gatilho próprio, separe aqui.

ALMOCO_PADRAO_H = 1.20      # 1h12
CAFE_PADRAO_H   = 0.25      # 15min
HORA_LIMITE_ALMOCO = 12.0   # meio-dia


def _para_horas(valor) -> float:
    """Aceita float (horas), datetime.time ou 'HH:MM' e devolve horas decimais."""
    if valor is None:
        return 0.0
    if isinstance(valor, bool):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    if hasattr(valor, "hour"):
        return valor.hour + valor.minute / 60 + getattr(valor, "second", 0) / 3600
    texto = str(valor).strip()
    if not texto:
        return 0.0
    try:
        partes = texto.split(":")
        return (int(partes[0]) + int(partes[1]) / 60
                + (int(partes[2]) if len(partes) > 2 else 0) / 3600)
    except (ValueError, IndexError):
        return 0.0


def calcular_intervalo(hora_saida, hora_chegada) -> float:
    """
    Quanto descontar de intervalo, em horas.

    Regra: o serviço atravessou o meio-dia → desconta almoço + café.
    Testada contra os dados reais: acerta 156 dos 159 lançamentos que
    tinham almoço na planilha, e não desconta em 387 dos 411 que não
    tinham. Os casos restantes são resolvidos pela convenção de ajustar
    o horário de saída.
    """
    ini, fim = _para_horas(hora_saida), _para_horas(hora_chegada)
    if fim <= ini:
        return 0.0
    if ini < HORA_LIMITE_ALMOCO < fim:
        return ALMOCO_PADRAO_H + CAFE_PADRAO_H
    return 0.0


def calcular_horas_liquidas(hora_saida, hora_chegada) -> dict:
    """
    Horas efetivamente trabalhadas, já descontado o intervalo.

    Nunca devolve horas negativas: se o intervalo for maior que a
    jornada (serviço curto em cima do meio-dia), zera e sinaliza, para
    a tela poder avisar em vez de gravar um valor sem sentido.
    """
    ini, fim = _para_horas(hora_saida), _para_horas(hora_chegada)
    brutas = fim - ini
    if brutas <= 0:
        return {"horas_brutas": 0.0, "intervalo": 0.0,
                "horas_liquidas": 0.0, "intervalo_maior_que_jornada": False}

    intervalo = calcular_intervalo(hora_saida, hora_chegada)
    estourou = intervalo >= brutas
    if estourou:
        intervalo = brutas          # zera em vez de virar negativo

    return {
        "horas_brutas":   round(brutas, 4),
        "intervalo":      round(intervalo, 4),
        "horas_liquidas": round(brutas - intervalo, 4),
        "intervalo_maior_que_jornada": estourou,
    }


# ─────────────────────────────────────────────
#  SERVIÇO
# ─────────────────────────────────────────────

def calcular_servico(
    local_servico: str,
    nivel_tecnico: str,
    hora_saida=None,
    hora_chegada=None,
    horas_trabalhadas: float | None = None,
    km_rodado: float = 0.0,
    tipo_km: str = "Km Um",
    horas_munck: float = 0.0,
) -> dict:
    """
    Calcula um lançamento de SERVIÇO (não deslocamento).

    horas_trabalhadas: passe um valor quando o tempo vier da tabela de
    tempos fixos por componente — nesse caso não há intervalo a descontar,
    porque o tempo da planilha já é o tempo de execução. Deixe None para
    calcular a partir dos horários.
    """
    local = (local_servico or "").strip().lower()

    if horas_trabalhadas is not None:
        det = {"horas_brutas": float(horas_trabalhadas), "intervalo": 0.0,
               "horas_liquidas": float(horas_trabalhadas),
               "intervalo_maior_que_jornada": False}
    else:
        det = calcular_horas_liquidas(hora_saida, hora_chegada)

    valor_hora  = TABELA_HORA.get(nivel_tecnico, TABELA_HORA["Técnico Um"])
    val_servico = det["horas_liquidas"] * valor_hora
    val_km      = km_rodado * TABELA_KM.get(tipo_km, TABELA_KM["Km Um"])
    val_munck   = horas_munck * HORA_MUNCK

    perc     = PERCENTUAIS_LOCAL.get(local, PERCENTUAIS_LOCAL["interno barracão"])
    comissao = val_servico * perc          # só o valor das horas

    return {
        "natureza":          "servico",
        "horas_brutas":      det["horas_brutas"],
        "intervalo":         det["intervalo"],
        "horas_trabalhadas": det["horas_liquidas"],
        "intervalo_maior_que_jornada": det["intervalo_maior_que_jornada"],
        "valor_hora":        valor_hora,
        "valor_servico":     round(val_servico, 4),
        "valor_km":          round(val_km, 4),
        "valor_munck":       round(val_munck, 4),
        "valor_deslocamento": 0.0,
        "valor_total":       round(val_servico + val_km + val_munck, 4),
        "base_comissao":     round(val_servico, 4),
        "percentual":        round(perc * 100, 2),
        "comissao":          round(comissao, 4),
    }


# ─────────────────────────────────────────────
#  DESLOCAMENTO
# ─────────────────────────────────────────────

def calcular_deslocamento(
    ida_saida=None, ida_chegada=None,
    retorno_saida=None, retorno_chegada=None,
    km_rodado: float = 0.0,
    tipo_km: str = "Km Um",
) -> dict:
    """
    Calcula um lançamento de DESLOCAMENTO.

    São dois trajetos com horários próprios (ida e retorno), como na
    planilha. O valor é (ida + retorno) × R$60, sem desconto de
    intervalo — trajeto não tem almoço a descontar.

    A velocidade média serve de conferência: se sair um número absurdo,
    algum horário ou KM foi digitado errado.
    """
    h_ida = max(0.0, _para_horas(ida_chegada) - _para_horas(ida_saida))
    h_ret = max(0.0, _para_horas(retorno_chegada) - _para_horas(retorno_saida))
    horas = h_ida + h_ret

    val_desloc = horas * HORA_DESLOCAMENTO
    val_km     = km_rodado * TABELA_KM.get(tipo_km, TABELA_KM["Km Um"])
    comissao   = val_desloc * PERCENTUAL_DESLOCAMENTO

    return {
        "natureza":           "deslocamento",
        "horas_ida":          round(h_ida, 4),
        "horas_retorno":      round(h_ret, 4),
        "horas_trabalhadas":  round(horas, 4),
        "valor_hora":         HORA_DESLOCAMENTO,
        "valor_servico":      0.0,
        "valor_km":           round(val_km, 4),
        "valor_munck":        0.0,
        "valor_deslocamento": round(val_desloc, 4),
        "valor_total":        round(val_desloc + val_km, 4),
        "base_comissao":      round(val_desloc, 4),
        "velocidade_media":   round(km_rodado / horas, 2) if horas > 0 else 0.0,
        "percentual":         round(PERCENTUAL_DESLOCAMENTO * 100, 2),
        "comissao":           round(comissao, 4),
    }


# ─────────────────────────────────────────────
#  LAVAGEM
# ─────────────────────────────────────────────

def calcular_lavagem(
    tipo_veiculo: str,
    local_servico: str,
    nivel_tecnico: str = "Técnico Um",
    hora_saida=None,
    hora_chegada=None,
    horas_trabalhadas: float | None = None,
) -> dict:
    """
    Calcula um lançamento de LAVAGEM.

    Regra de valor:
      carro  → R$80 fixo
      onibus → R$250 fixo
      outros → horas líquidas × valor-hora do nível (mesma conta do serviço,
               com desconto de almoço/café quando atravessa o meio-dia)

    Comissão: valor × percentual DO LOCAL (mesma tabela por local dos demais
    lançamentos — Interno 4%, Campo 8%, M.S 10%). KM e Munck não se aplicam.
    """
    tv    = (tipo_veiculo or "").strip().lower()
    local = (local_servico or "").strip().lower()

    if tv in VALOR_LAVAGEM:
        det = {"horas_brutas": 0.0, "intervalo": 0.0, "horas_liquidas": 0.0,
               "intervalo_maior_que_jornada": False}
        valor_hora  = 0.0
        val_servico = VALOR_LAVAGEM[tv]
    else:  # "outros" — por tempo
        if horas_trabalhadas is not None:
            det = {"horas_brutas": float(horas_trabalhadas), "intervalo": 0.0,
                   "horas_liquidas": float(horas_trabalhadas),
                   "intervalo_maior_que_jornada": False}
        else:
            det = calcular_horas_liquidas(hora_saida, hora_chegada)
        valor_hora  = TABELA_HORA.get(nivel_tecnico, TABELA_HORA["Técnico Um"])
        val_servico = det["horas_liquidas"] * valor_hora

    perc     = PERCENTUAIS_LOCAL.get(local, PERCENTUAIS_LOCAL["interno barracão"])
    comissao = val_servico * perc

    return {
        "natureza":          "lavagem",
        "tipo_veiculo":      tv,
        "horas_brutas":      det["horas_brutas"],
        "intervalo":         det["intervalo"],
        "horas_trabalhadas": det["horas_liquidas"],
        "intervalo_maior_que_jornada": det["intervalo_maior_que_jornada"],
        "valor_hora":        valor_hora,
        "valor_servico":     round(val_servico, 4),
        "valor_km":          0.0,
        "valor_munck":       0.0,
        "valor_deslocamento": 0.0,
        "valor_total":       round(val_servico, 4),
        "base_comissao":     round(val_servico, 4),
        "percentual":        round(perc * 100, 2),
        "comissao":          round(comissao, 4),
    }


# ─────────────────────────────────────────────
#  RESUMO POR TÉCNICO
# ─────────────────────────────────────────────

def calcular_resumo_tecnico(lancamentos: list) -> dict:
    """
    Consolida os lançamentos de um técnico, no mesmo recorte da aba
    "Comissão" da planilha: valores e comissão separados por local,
    com o deslocamento à parte.
    """
    res = {
        "horas_totais": 0.0, "km_total": 0.0,
        "val_interno": 0.0, "val_campo": 0.0, "val_ms": 0.0,
        "val_deslocamento": 0.0, "val_km": 0.0, "val_munck": 0.0,
        "comissao_interno": 0.0, "comissao_campo": 0.0, "comissao_ms": 0.0,
        "comissao_deslocamento": 0.0, "total_comissao": 0.0,
    }

    for it in lancamentos:
        res["horas_totais"] += float(it.get("horas_trabalhadas", 0) or 0)
        res["km_total"]     += float(it.get("km_rodado", 0) or 0)
        res["val_km"]       += float(it.get("valor_km", 0) or 0)
        res["val_munck"]    += float(it.get("valor_munck", 0) or 0)
        comissao = float(it.get("comissao", 0) or 0)

        if it.get("natureza") == "deslocamento":
            res["val_deslocamento"]      += float(it.get("valor_deslocamento", 0) or 0)
            res["comissao_deslocamento"] += comissao
        else:
            local = (it.get("local_servico") or "").strip().lower()
            valor = float(it.get("valor_servico", 0) or 0)
            if local == "campo":
                res["val_campo"] += valor
                res["comissao_campo"] += comissao
            elif local == "m.s":
                res["val_ms"] += valor
                res["comissao_ms"] += comissao
            else:
                res["val_interno"] += valor
                res["comissao_interno"] += comissao

    res["total_comissao"] = (res["comissao_interno"] + res["comissao_campo"]
                             + res["comissao_ms"] + res["comissao_deslocamento"])
    return {k: round(v, 2) for k, v in res.items()}
