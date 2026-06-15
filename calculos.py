"""
calculos.py — Motor de cálculo de comissões
Baseado na planilha Relatório Assist. Téc. (aba Valores + Comissão)

REGRAS (extraídas da planilha):
──────────────────────────────────────────────────────────────────────────
1. VALOR DO SERVIÇO = Horas Trabalhadas × Valor Hora do Técnico
   (exceto deslocamento, que tem valor fixo R$60/h)

2. VALOR KM RODADO = KM Rodado × Valor por KM (depende do tipo: Km Um/Dois/Três)

3. VALOR MUNCK = Horas Munck × R$120/h

4. VALOR TOTAL DA OS = Valor Serviço + Valor KM + Valor Munck

5. COMISSÃO = Valor Total × Percentual do Local:
   - interno barracão → 4%
   - campo            → 8%
   - deslocamento     → 6%
   - m.s              → 10%

OBSERVAÇÕES ESPECIAIS (observadas na planilha):
- No deslocamento, o valor hora usado é R$60 (fixo, não o do técnico)
- Para "m.s" (Mato do Sul), comissão é sobre serviço + km + munck
- O Cristiano B. usa Km Três (R$5,00/km), outros usam Km Um ou Km Dois
- Rodrigo J. usa Km Dois (R$4,10/km)
──────────────────────────────────────────────────────────────────────────
"""

# Constantes (espelham aba "Valores" da planilha)
TABELA_KM = {
    "Km Um":   1.70,
    "Km Dois": 4.10,
    "Km Três": 5.00,
}

HORA_DESLOCAMENTO = 60.0   # R$/hora (linha "Deslocamento" na aba Valores)
HORA_MUNCK = 120.0          # R$/hora (linha "Hora Munck" na aba Valores)

PERCENTUAIS_LOCAL = {
    "interno barracão": 0.04,   # 4%
    "campo":            0.08,   # 8%
    "deslocamento":     0.06,   # 6%
    "m.s":              0.10,   # 10%
}


def calcular_comissao_os(
    local_servico: str,
    horas_trabalhadas: float,
    valor_hora: float,          # R$/hora do técnico (conforme nível)
    km_rodado: float = 0.0,
    tipo_km: str = "Km Um",     # "Km Um", "Km Dois" ou "Km Três"
    horas_munck: float = 0.0,
) -> dict:
    """
    Calcula o valor do serviço e a comissão de uma OS.

    Retorna dict com:
      valor_hora_usado, valor_servico, valor_km, valor_munck,
      valor_total, percentual, comissao
    """
    local_norm = local_servico.strip().lower()

    # 1. Hora usada: deslocamento tem taxa fixa R$60/h
    hora_usada = HORA_DESLOCAMENTO if local_norm == "deslocamento" else valor_hora

    # 2. Valor do serviço (horas × valor hora)
    val_servico = horas_trabalhadas * hora_usada

    # 3. Valor KM
    preco_km = TABELA_KM.get(tipo_km, TABELA_KM["Km Um"])
    val_km = km_rodado * preco_km

    # 4. Valor Munck
    val_munck = horas_munck * HORA_MUNCK

    # 5. Total
    val_total = val_servico + val_km + val_munck

    # 6. Percentual e comissão
    perc = PERCENTUAIS_LOCAL.get(local_norm, 0.04)
    comissao = val_total * perc

    return {
        "valor_hora_usado": hora_usada,
        "valor_servico":    round(val_servico, 4),
        "valor_km":         round(val_km, 4),
        "valor_munck":      round(val_munck, 4),
        "valor_total":      round(val_total, 4),
        "percentual":       int(perc * 100),
        "comissao":         round(comissao, 4),
    }


def calcular_resumo_tecnico(lista_os: list) -> dict:
    """
    Dado uma lista de OS aprovadas de um técnico,
    retorna o resumo consolidado de comissões.
    """
    total_interno = 0.0
    total_campo   = 0.0
    total_desloc  = 0.0
    total_ms      = 0.0
    total_km_val  = 0.0
    total_munck_val = 0.0
    total_horas   = 0.0
    total_km      = 0.0

    for os_item in lista_os:
        local = (os_item.get("local_servico") or "").strip().lower()
        val_serv = float(os_item.get("valor_servico", 0))
        val_km   = float(os_item.get("valor_km", 0))
        val_mun  = float(os_item.get("valor_munck", 0))
        total_km_val   += val_km
        total_munck_val += val_mun
        total_horas    += float(os_item.get("horas_trabalhadas", 0))
        total_km       += float(os_item.get("km_rodado", 0))

        if local == "interno barracão":
            total_interno += val_serv
        elif local == "campo":
            total_campo += val_serv
        elif local == "deslocamento":
            total_desloc += val_serv
        elif local == "m.s":
            total_ms += val_serv

    comissao_interno = total_interno * PERCENTUAIS_LOCAL["interno barracão"]
    comissao_campo   = total_campo   * PERCENTUAIS_LOCAL["campo"]
    comissao_desloc  = total_desloc  * PERCENTUAIS_LOCAL["deslocamento"]
    comissao_ms      = total_ms      * PERCENTUAIS_LOCAL["m.s"]
    comissao_km      = total_km_val  * PERCENTUAIS_LOCAL["campo"]   # KM em campo = 8%
    comissao_munck   = total_munck_val * PERCENTUAIS_LOCAL["campo"] # Munck = 8%

    total_comissao = (comissao_interno + comissao_campo +
                      comissao_desloc + comissao_ms +
                      comissao_km + comissao_munck)

    return {
        "horas_totais":      round(total_horas, 2),
        "km_total":          round(total_km, 2),
        "val_interno":       round(total_interno, 2),
        "val_campo":         round(total_campo, 2),
        "val_deslocamento":  round(total_desloc, 2),
        "val_ms":            round(total_ms, 2),
        "val_km":            round(total_km_val, 2),
        "val_munck":         round(total_munck_val, 2),
        "comissao_interno":  round(comissao_interno, 2),
        "comissao_campo":    round(comissao_campo, 2),
        "comissao_deslocamento": round(comissao_desloc, 2),
        "comissao_ms":       round(comissao_ms, 2),
        "comissao_km":       round(comissao_km, 2),
        "comissao_munck":    round(comissao_munck, 2),
        "total_comissao":    round(total_comissao, 2),
    }
