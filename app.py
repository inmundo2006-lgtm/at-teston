import io
import streamlit as st
import pandas as pd
import hashlib
import os
from datetime import date, datetime, time

from calculos import calcular_comissao_os
from tempos_fixos import carregar_tabela_tempos, formatar_horas, NENHUM_COMPONENTE

# ─────────────────────────────────────────────
#  TABELA DE FROTAS
# ─────────────────────────────────────────────
@st.cache_data
def carregar_frotas():
    caminho = os.path.join(os.path.dirname(__file__), "CADASTRO.xlsx")
    try:
        df = pd.read_excel(caminho)
        df.columns = df.columns.str.strip()
        df["FROTA"]     = df["FROTA"].astype(str).str.strip()
        df["DESCRICAO"] = df["DESCRICAO"].astype(str).str.strip()
        return dict(zip(df["FROTA"], df["DESCRICAO"]))
    except Exception:
        return {}

# ─────────────────────────────────────────────
#  CONFIGURAÇÃO
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Assistência Técnica Teston",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
#  ESTILOS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

[data-testid="stAppViewContainer"] > .main { background: transparent; }
[data-testid="block-container"] { padding-top: 1rem; }
.stApp { background: #0f1420; }
section[data-testid="stSidebar"] + div { background: #0f1420; }

[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #1a1f2e 0%, #0f1420 100%);
    border-right: 1px solid #2d3748;
}
[data-testid="stSidebar"] * { color: #e2e8f0 !important; }
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stTextInput label { color: #94a3b8 !important; font-size: 0.75rem !important; }

label, .stTextInput label, .stSelectbox label,
.stDateInput label, .stTimeInput label, .stNumberInput label,
.stTextArea label, p, span, div { color: #e2e8f0 !important; }
h1, h2, h3, h4 { color: #f8fafc !important; }

input, textarea, select,
[data-testid="stTextInput"] input,
[data-testid="stNumberInput"] input {
    background: #1e2738 !important;
    color: #f1f5f9 !important;
    border-color: #3d4a5c !important;
}

.stDataFrame { background: #1a1f2e !important; }
.stTabs [data-baseweb="tab-list"] { background: #1a1f2e; border-radius: 8px; }
.stTabs [data-baseweb="tab"] { color: #94a3b8 !important; }
.stTabs [aria-selected="true"] { color: #f8fafc !important; }

/* ── Componentes ── */
.main-header {
    background: linear-gradient(135deg, #1a1f2e 0%, #2d3748 100%);
    border-radius: 12px;
    padding: 1.5rem 2rem;
    margin-bottom: 1.5rem;
    border: 1px solid #3d4a5c;
    display: flex;
    align-items: center;
    gap: 1rem;
}
.main-header h1 { color: #f8fafc !important; margin: 0; font-size: 1.5rem; font-weight: 700; }
.main-header p  { color: #94a3b8 !important; margin: 0; font-size: 0.875rem; }
.main-header .icon { font-size: 2.5rem; }

.metric-card {
    background: #1e2738;
    border-radius: 10px;
    padding: 1.25rem;
    border: 1px solid #3d4a5c;
    border-left: 4px solid #3b82f6;
    box-shadow: 0 1px 3px rgba(0,0,0,0.3);
}
.metric-card.green  { border-left-color: #10b981; }
.metric-card.yellow { border-left-color: #f59e0b; }
.metric-card.red    { border-left-color: #ef4444; }
.metric-card.purple { border-left-color: #8b5cf6; }
.metric-label { color: #94a3b8 !important; font-size: 0.75rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em; }
.metric-value { color: #f8fafc !important; font-size: 1.75rem; font-weight: 700; line-height: 1.2; }
.metric-sub   { color: #64748b !important; font-size: 0.8rem; }

.login-container {
    max-width: 400px;
    margin: 5rem auto;
    background: #1e2738;
    border-radius: 16px;
    padding: 2.5rem;
    box-shadow: 0 10px 40px rgba(0,0,0,0.4);
    border: 1px solid #3d4a5c;
}
.login-title { text-align: center; color: #f8fafc !important; font-size: 1.5rem; font-weight: 700; margin-bottom: 0.25rem; }
.login-sub   { text-align: center; color: #94a3b8 !important; font-size: 0.875rem; margin-bottom: 2rem; }
.logo-area   { text-align: center; font-size: 3rem; margin-bottom: 1rem; }

.section-header {
    display: flex;
    align-items: center;
    gap: 0.5rem;
    margin-bottom: 1rem;
    padding-bottom: 0.75rem;
    border-bottom: 2px solid #2d3748;
}
.section-header h3 { margin: 0; color: #f1f5f9 !important; font-size: 1rem; font-weight: 700; }

/* ── OS Card ── */
.os-card {
    background: #1a1f2e;
    border: 1px solid #2d3748;
    border-radius: 12px;
    padding: 1rem 1.25rem;
    margin-bottom: 0.75rem;
    position: relative;
}
.os-card:hover { border-color: #4a5568; }
.os-numero { font-size: 1.1rem; font-weight: 800; color: #60a5fa !important; letter-spacing: 0.02em; }
.os-frota  { font-size: 0.8rem; color: #94a3b8 !important; }

/* ── Status badges ── */
.badge {
    display: inline-block;
    padding: 0.18rem 0.7rem;
    border-radius: 9999px;
    font-size: 0.72rem;
    font-weight: 700;
    letter-spacing: 0.03em;
}
.badge-aberta           { background:#0c2a3d; color:#38bdf8 !important; border:1px solid #0369a1; }
.badge-em_andamento     { background:#422006; color:#fbbf24 !important; border:1px solid #d97706; }
.badge-aguardando       { background:#1e1b4b; color:#a5b4fc !important; border:1px solid #6366f1; }
.badge-aprovada         { background:#052e16; color:#4ade80 !important; border:1px solid #16a34a; }
.badge-rejeitada        { background:#3f0000; color:#fca5a5 !important; border:1px solid #dc2626; }

/* ── Serviço card ── */
.serv-card {
    background: #0f1420;
    border: 1px solid #2d3748;
    border-left: 3px solid #3b82f6;
    border-radius: 8px;
    padding: 0.875rem 1rem;
    margin-bottom: 0.5rem;
}
.serv-seq  { font-size: 0.7rem; font-weight: 700; color: #60a5fa !important; text-transform: uppercase; letter-spacing: 0.06em; }
.serv-tec  { font-size: 0.9rem; font-weight: 600; color: #f1f5f9 !important; }
.serv-tipo { font-size: 0.78rem; color: #94a3b8 !important; }

/* ── Preview ── */
.preview-box {
    background: linear-gradient(135deg, #0c2a3d 0%, #0a1f30 100%);
    border: 2px solid #1e6fa8;
    border-radius: 12px;
    padding: 1.25rem 1.5rem;
    margin: 1rem 0;
}
.preview-box .title { font-size: 0.8rem; font-weight: 700; color: #38bdf8 !important;
    text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 0.75rem; }
.preview-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 1rem; }
.preview-grid-2 { display: grid; grid-template-columns: repeat(2, 1fr); gap: 1rem; }
.preview-item { text-align: center; }
.preview-item .lbl { font-size: 0.7rem; color: #7dd3fc !important; font-weight: 600; }
.preview-item .val { font-size: 1.3rem; font-weight: 800; color: #e0f2fe !important; }
.preview-item .sub { font-size: 0.7rem; color: #7dd3fc !important; }

.comissao-box {
    background: linear-gradient(135deg, #052e16 0%, #064e3b 100%);
    border: 1px solid #166534;
    border-radius: 10px;
    padding: 1rem 1.5rem;
    text-align: center;
}
.comissao-box .valor { font-size: 2rem; font-weight: 800; color: #4ade80 !important; }
.comissao-box .label { font-size: 0.8rem; color: #86efac !important; font-weight: 500; }

.info-pill {
    display: inline-block;
    background: #1e3a5f;
    color: #93c5fd !important;
    padding: 0.2rem 0.8rem;
    border-radius: 9999px;
    font-size: 0.75rem;
    font-weight: 600;
    margin: 0.15rem;
}

.edit-banner {
    background: #422006;
    border: 1px solid #f59e0b;
    border-radius: 8px;
    padding: 0.75rem 1rem;
    margin-bottom: 1rem;
    font-size: 0.875rem;
    color: #fcd34d !important;
    font-weight: 500;
}

.nova-os-banner {
    background: #0c2a3d;
    border: 1px solid #0369a1;
    border-radius: 8px;
    padding: 0.75rem 1rem;
    margin-bottom: 1rem;
    font-size: 0.875rem;
    color: #7dd3fc !important;
    font-weight: 500;
}

.concluir-box {
    background: #1e1b4b;
    border: 1px solid #6366f1;
    border-radius: 10px;
    padding: 1rem 1.25rem;
    margin-top: 1rem;
}

.timeline-item {
    display: flex;
    gap: 0.75rem;
    margin-bottom: 0.75rem;
    align-items: flex-start;
}
.timeline-dot {
    width: 28px;
    height: 28px;
    border-radius: 50%;
    background: #1e3a5f;
    border: 2px solid #3b82f6;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 0.7rem;
    font-weight: 800;
    color: #60a5fa !important;
    flex-shrink: 0;
    margin-top: 0.1rem;
}
.timeline-content { flex: 1; }

.tempo-fixo-box {
    background: linear-gradient(135deg, #052e2e 0%, #0a3d3d 100%);
    border: 1px solid #0e7490;
    border-radius: 10px;
    padding: 0.9rem 1.1rem;
    margin-top: 0.75rem;
    font-size: 0.85rem;
    color: #67e8f9 !important;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  DADOS
# ─────────────────────────────────────────────
from dados import (
    USUARIOS, TECNICOS, CENTROS_CUSTO, TIPOS_SERVICO,
    EQUIPAMENTOS, NIVEIS_TECNICOS, TABELA_KM, TABELA_HORA,
    PERCENTUAIS_LOCAL,
    # novas funções
    carregar_os, criar_os,
    adicionar_procedimento as adicionar_servico,
    editar_procedimento    as editar_servico,
    excluir_procedimento   as excluir_servico,
    fechar_os_para_aprovacao,
    reabrir_os, validar_os, buscar_os_por_numero,
    definir_tecnico_designado,
    # novas — tipo/permissão de OS
    pode_abrir_os, motivo_bloqueio_abertura, PermissaoNegada,
    registrar_avaliacao, TIPOS_OS,
    # compat
    salvar_os, atualizar_status_os, editar_os,
)

# ─────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────

def df_para_excel(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="OS")
    return buf.getvalue()

def _badge(status: str) -> str:
    labels = {
        "aberta":                "🔵 ABERTA",
        "em_andamento":          "🟡 EM ANDAMENTO",
        "aguardando_aprovacao":  "🟣 AGUARDANDO APROVAÇÃO",
        "aprovada":              "🟢 APROVADA",
        "rejeitada":             "🔴 REJEITADA",
    }
    cls = {
        "aberta":               "badge-aberta",
        "em_andamento":         "badge-em_andamento",
        "aguardando_aprovacao": "badge-aguardando",
        "aprovada":             "badge-aprovada",
        "rejeitada":            "badge-rejeitada",
    }
    return (f'<span class="badge {cls.get(status,"")}">'
            f'{labels.get(status, status.upper())}</span>')

def _total_comissao_os(os_item: dict) -> float:
    return sum(float(p.get("comissao", 0)) for p in os_item.get("procedimentos", []))

def _total_servico_os(os_item: dict) -> float:
    return sum(float(p.get("valor_servico", 0)) for p in os_item.get("procedimentos", []))

def _resumo_tecnicos(os_item: dict) -> str:
    nomes = list(dict.fromkeys(
        p.get("nome_tecnico", "?") for p in os_item.get("procedimentos", [])
    ))
    return ", ".join(nomes) if nomes else "—"

# ─────────────────────────────────────────────
#  AUTENTICAÇÃO
# ─────────────────────────────────────────────

def _hash(s):
    return hashlib.sha256(s.encode()).hexdigest()

def verificar_login(usuario, senha):
    if usuario in USUARIOS and USUARIOS[usuario]["senha"] == _hash(senha):
        return USUARIOS[usuario]
    return None

def tela_login():
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("""
        <div class="login-container">
            <div class="logo-area">🔧</div>
            <div class="login-title">Assistência Técnica</div>
            <div class="login-sub">Teston — Sistema de OS e Comissões</div>
        </div>
        """, unsafe_allow_html=True)
        with st.form("login_form"):
            usuario = st.text_input("Usuário", placeholder="seu.usuario")
            senha   = st.text_input("Senha", type="password", placeholder="••••••••")
            if st.form_submit_button("Entrar", use_container_width=True, type="primary"):
                ud = verificar_login(usuario, senha)
                if ud:
                    st.session_state.update({"logado": True, "usuario": usuario, "user_data": ud})
                    st.rerun()
                else:
                    st.error("Usuário ou senha incorretos.")

# ─────────────────────────────────────────────
#  SIDEBAR
# ─────────────────────────────────────────────

def render_sidebar():
    ud = st.session_state["user_data"]
    with st.sidebar:
        st.markdown(f"""
        <div style="padding:1rem 0;border-bottom:1px solid #3d4a5c;margin-bottom:1rem;">
            <div style="font-size:1.8rem;text-align:center;">👤</div>
            <div style="text-align:center;font-weight:700;font-size:0.95rem;">{ud['nome']}</div>
            <div style="text-align:center;font-size:0.75rem;color:#94a3b8;">{ud['perfil'].upper()}</div>
        </div>
        """, unsafe_allow_html=True)

        if ud["perfil"] == "tecnico":
            pages = ["📋 Minhas OS",  "➕ Abrir Nova OS","🔧 Adicionar Serviço"]
        else:
            pages = ["📊 Dashboard", "✅ Validações", "📋 Todas as OS",
                     "➕ Abrir Nova OS", "🔧 Adicionar Serviço",
                     "💰 Comissões", "📤 Exportar Relatório", "⚙️ Configurações"]

        page = st.radio("Navegação", pages, label_visibility="collapsed")
        st.session_state["page"] = page
        st.markdown("---")
        if st.button("🚪 Sair", use_container_width=True):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.rerun()
    return page

# ─────────────────────────────────────────────
#  RECALC PREVIEW (serviço)
# ─────────────────────────────────────────────

def _recalc_preview():
    cod_tec = st.session_state.get("_pr_cod_tecnico", 1)
    local   = st.session_state.get("_pr_local", "interno barracão")
    h_s     = st.session_state.get("_pr_hora_saida",  time(7,  30))
    h_c     = st.session_state.get("_pr_hora_cheg",   time(16, 30))
    km_i    = st.session_state.get("_pr_km_ini",  0.0)
    km_f    = st.session_state.get("_pr_km_fim",  0.0)
    m_i     = st.session_state.get("_pr_mun_ini", 0.0)
    m_f     = st.session_state.get("_pr_mun_fim", 0.0)

    tec_info   = TECNICOS.get(cod_tec, {})
    valor_hora = TABELA_HORA.get(tec_info.get("nivel", "Técnico Um"), 50)
    km_rodado  = max(0.0, km_f - km_i)
    horas_munck  = max(0.0, m_f - m_i)

    # ── Tempo fixo por componente (se aplicável) ──
    tabela_tempos, _ = carregar_tabela_tempos()
    equip_fixo = st.session_state.get("_pr_equip_fixo", NENHUM_COMPONENTE)
    comp_fixo  = st.session_state.get("_pr_componente_fixo", NENHUM_COMPONENTE)
    tempo_fixo_aplicado = (equip_fixo != NENHUM_COMPONENTE and comp_fixo != NENHUM_COMPONENTE)

    if tempo_fixo_aplicado:
        horas_trab = tabela_tempos.get((equip_fixo, comp_fixo), 0.0)
    else:
        horas_trab = (datetime.combine(date.today(), h_c) -
                      datetime.combine(date.today(), h_s)).seconds / 3600

    est = calcular_comissao_os(
        local_servico=local,
        horas_trabalhadas=horas_trab,
        valor_hora=valor_hora,
        km_rodado=km_rodado,
        tipo_km=tec_info.get("tipo_km", "Km Um"),
        horas_munck=horas_munck,
    )
    return est, horas_trab, km_rodado, horas_munck, tempo_fixo_aplicado

# ─────────────────────────────────────────────
#  FORMULÁRIO DE SERVIÇO
#  (um técnico adicionando seu trabalho a uma OS)
# ─────────────────────────────────────────────

def _form_servico(ud, os_item: dict, serv_existente: dict | None = None):
    """
    Renderiza o formulário de serviço.
    - os_item:        OS pai (dict completo)
    - serv_existente: se não None, estamos editando um serviço já salvo
    Retorna o dict do serviço salvo, ou None se ainda não submetido.
    """
    editando   = serv_existente is not None
    e          = serv_existente or {}
    eh_tecnico = ud["perfil"] == "tecnico"

    # Inicializa session_state
    if "_pr_km_ini" not in st.session_state or st.session_state.get("_pr_reset"):
        def _pt(v, default):
            if isinstance(v, time): return v
            if isinstance(v, str):
                try:
                    p = v.split(":")
                    return time(int(p[0]), int(p[1]))
                except Exception:
                    return default
            return default

        st.session_state["_pr_km_ini"]     = float(e.get("km_inicial", 0))
        st.session_state["_pr_km_fim"]     = float(e.get("km_final",   0))
        st.session_state["_pr_mun_ini"]    = float(e.get("hora_munck_inicial", 0))
        st.session_state["_pr_mun_fim"]    = float(e.get("hora_munck_final",   0))
        st.session_state["_pr_hora_saida"] = _pt(e.get("hora_saida"),   time(7,  30))
        st.session_state["_pr_hora_cheg"]  = _pt(e.get("hora_chegada"), time(16, 30))
        st.session_state["_pr_equip_fixo"] = e.get("equipamento_tempo_fixo") or NENHUM_COMPONENTE
        st.session_state["_pr_reset"] = False

    # Cabeçalho da OS
    n_servs   = len(os_item.get("procedimentos", []))
    seq_atual = e.get("seq", n_servs + 1)
    st.markdown(f"""
    <div class="nova-os-banner">
        📋 <strong>{os_item['numero_os']}</strong> &nbsp;·&nbsp;
        Frota <strong>{os_item['frota']}</strong> &nbsp;·&nbsp;
        {os_item['equipamento']} &nbsp;·&nbsp;
        {os_item['cliente']} &nbsp;·&nbsp;
        {"Editando serviço #" + str(seq_atual) if editando else
         f"Novo serviço (#{seq_atual})"}
    </div>
    """, unsafe_allow_html=True)

    # ── Linha 1 ──
    c1, c2, c3 = st.columns(3)
    with c1:
        data_serv = st.date_input("Data do Serviço",
            value=date.fromisoformat(e["data"]) if e.get("data") else date.today())
        if eh_tecnico and ud.get("cod_tecnico"):
            cod_tecnico = ud["cod_tecnico"]
            st.text_input("Técnico", value=TECNICOS[cod_tecnico]["nome"], disabled=True)
        else:
            lista_tec = {v["nome"]: k for k, v in TECNICOS.items()}
            nomes = list(lista_tec.keys())
            idx = nomes.index(e["nome_tecnico"]) if e.get("nome_tecnico") in nomes else 0
            cod_tecnico = lista_tec[st.selectbox("Técnico", nomes, index=idx)]
        st.session_state["_pr_cod_tecnico"] = cod_tecnico

    with c2:
        idx_tp = TIPOS_SERVICO.index(e["tipo_servico"]) if e.get("tipo_servico") in TIPOS_SERVICO else 0
        tipo_servico = st.selectbox("Tipo de Serviço", TIPOS_SERVICO, index=idx_tp)
        locais = ["interno barracão", "campo", "deslocamento", "m.s"]
        idx_lc = locais.index(e["local_servico"]) if e.get("local_servico") in locais else 0
        local_servico = st.selectbox("Local do Serviço", locais, index=idx_lc, key="_pr_local")

    with c3:
        # CC herdado da OS — exibe apenas como info, não editável no serviço
        st.text_input("Centro de Custo", value=f"{os_item['cod_cc']} - {os_item['cliente']}", disabled=True)
        st.caption("Herdado da OS — altere na OS se necessário")

    st.markdown("---")

    # ── Tempo Fixo por Componente (opcional) ──
    st.markdown("##### ⏱️ Tempo Fixo por Componente (opcional)")
    st.caption(
        "Se este serviço é a troca/reparo de um componente da tabela, o tempo já vem "
        "pronto da planilha e o campo Hora Saída/Chegada abaixo não é usado no cálculo. "
        "Se não for um item da tabela, o tempo continua sendo calculado por hora início/fim."
    )
    tabela_tempos, por_equipamento = carregar_tabela_tempos()

    if not por_equipamento:
        st.warning("⚠️ Planilha TEMPO_SERVIÇO.xlsx não encontrada na pasta do app — "
                    "tempo fixo indisponível, todos os serviços usam hora início/fim.")
        equip_sel, comp_sel = NENHUM_COMPONENTE, NENHUM_COMPONENTE
    else:
        equipamentos_tabela = [NENHUM_COMPONENTE] + sorted(por_equipamento.keys())
        if st.session_state.get("_pr_equip_fixo") not in equipamentos_tabela:
            st.session_state["_pr_equip_fixo"] = NENHUM_COMPONENTE

        c_eq, c_comp = st.columns(2)
        equip_sel = c_eq.selectbox("Equipamento (tabela de tempos)", equipamentos_tabela,
                                    key="_pr_equip_fixo")

        if equip_sel != NENHUM_COMPONENTE:
            comps = [NENHUM_COMPONENTE] + por_equipamento.get(equip_sel, [])
        else:
            comps = [NENHUM_COMPONENTE]
        if st.session_state.get("_pr_componente_fixo") not in comps:
            st.session_state["_pr_componente_fixo"] = (
                e.get("componente_tempo_fixo") if e.get("componente_tempo_fixo") in comps
                else NENHUM_COMPONENTE
            )
        comp_sel = c_comp.selectbox("Componente", comps, key="_pr_componente_fixo",
                                     disabled=(equip_sel == NENHUM_COMPONENTE))

        if equip_sel != NENHUM_COMPONENTE and comp_sel != NENHUM_COMPONENTE:
            horas_fixas = tabela_tempos.get((equip_sel, comp_sel), 0.0)
            st.markdown(f"""
            <div class="tempo-fixo-box">
                ⏱️ <strong>Tempo fixo aplicado: {horas_fixas:.2f}h ({formatar_horas(horas_fixas)})</strong><br>
                Hora Saída/Chegada abaixo servem apenas de registro — não entram no cálculo.
            </div>
            """, unsafe_allow_html=True)

    st.markdown("---")

    # ── Linha 2 ──
    c4, c5 = st.columns(2)
    with c4:
        hora_saida   = st.time_input("Hora Saída",   key="_pr_hora_saida")
        hora_chegada = st.time_input("Hora Chegada", key="_pr_hora_cheg")
    with c5:
        km_ini  = st.number_input("KM Inicial",        min_value=0.0, step=1.0,  format="%.1f", key="_pr_km_ini")
        km_fim  = st.number_input("KM Final",           min_value=0.0, step=1.0,  format="%.1f", key="_pr_km_fim")
        mun_ini = st.number_input("Hora Munck Inicial", min_value=0.0, step=0.1,  format="%.2f", key="_pr_mun_ini")
        mun_fim = st.number_input("Hora Munck Final",   min_value=0.0, step=0.1,  format="%.2f", key="_pr_mun_fim")

    descricao = st.text_area("Descrição do Serviço Executado",
        value=e.get("descricao", ""), height=100,
        placeholder="Descreva detalhadamente o serviço realizado...")

    # ── Preview ──
    st.markdown("---")
    comissao_est, horas_trab, km_rodado, horas_munck, tempo_fixo_aplicado = _recalc_preview()

    if eh_tecnico:
        st.markdown(f"""
        <div class="preview-box">
            <div class="title">📊 Resumo</div>
            <div class="preview-grid-2">
                <div class="preview-item">
                    <div class="lbl">{"Horas (tempo fixo)" if tempo_fixo_aplicado else "Horas Trabalhadas"}</div>
                    <div class="val">{horas_trab:.2f}h</div>
                </div>
                <div class="preview-item">
                    <div class="lbl">KM Rodado</div>
                    <div class="val">{km_rodado:.1f} km</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown(f"""
        <div class="preview-box">
            <div class="title">📊 Preview da Comissão</div>
            <div class="preview-grid">
                <div class="preview-item">
                    <div class="lbl">{"Horas (fixo)" if tempo_fixo_aplicado else "Horas"}</div>
                    <div class="val">{horas_trab:.2f}h</div>
                </div>
                <div class="preview-item">
                    <div class="lbl">KM Rodado</div>
                    <div class="val">{km_rodado:.1f} km</div>
                    <div class="sub">{km_ini:.1f} → {km_fim:.1f}</div>
                </div>
                <div class="preview-item">
                    <div class="lbl">Valor Serviço</div>
                    <div class="val">R$ {comissao_est['valor_servico']:,.2f}</div>
                    <div class="sub">+ KM R$ {comissao_est['valor_km']:,.2f}</div>
                </div>
                <div class="preview-item">
                    <div class="lbl">Comissão ({comissao_est['percentual']}%)</div>
                    <div class="val" style="color:#4ade80;">R$ {comissao_est['comissao']:,.2f}</div>
                    <div class="sub">Local: {local_servico}</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    # ── Botão salvar ──
    label_btn = "💾 Salvar Alterações" if editando else "💾 Salvar Serviço"
    if st.button(label_btn, type="primary", use_container_width=True):
        if not descricao.strip():
            st.error("A descrição do serviço é obrigatória.")
            return None
        if horas_trab <= 0:
            if tempo_fixo_aplicado:
                st.error("Não foi possível obter o tempo fixo para este componente.")
            else:
                st.error("Hora chegada deve ser posterior à hora saída.")
            return None

        serv = {
            "data":                str(data_serv),
            "cod_tecnico":         cod_tecnico,
            "nome_tecnico":        TECNICOS[cod_tecnico]["nome"],
            "nivel_tecnico":       TECNICOS[cod_tecnico]["nivel"],
            "tipo_servico":        tipo_servico,
            "local_servico":       local_servico,
            "hora_saida":          str(hora_saida),
            "hora_chegada":        str(hora_chegada),
            "horas_trabalhadas":   round(horas_trab, 4),
            "equipamento_tempo_fixo": equip_sel if equip_sel != NENHUM_COMPONENTE else None,
            "componente_tempo_fixo":  comp_sel  if comp_sel  != NENHUM_COMPONENTE else None,
            "tempo_fixo_aplicado":    tempo_fixo_aplicado,
            "km_inicial":          km_ini,
            "km_final":            km_fim,
            "km_rodado":           round(km_rodado, 2),
            "hora_munck_inicial":  mun_ini,
            "hora_munck_final":    mun_fim,
            "horas_munck":         round(horas_munck, 4),
            "descricao":           descricao,
            "valor_servico":       round(comissao_est["valor_servico"], 2),
            "comissao":            round(comissao_est["comissao"], 2),
            "percentual_comissao": comissao_est["percentual"],
            "valor_km":            round(comissao_est.get("valor_km", 0), 2),
            "valor_munck":         round(comissao_est.get("valor_munck", 0), 2),
            "registrado_por":      st.session_state["usuario"],
            "registrado_em":       datetime.now().isoformat(),
            "editado_por":         None,
            "editado_em":          None,
        }

        if editando:
            serv["editado_por"]    = st.session_state["usuario"]
            serv["editado_em"]     = datetime.now().isoformat()
            serv["registrado_por"] = e.get("registrado_por", st.session_state["usuario"])
            serv["registrado_em"]  = e.get("registrado_em",  datetime.now().isoformat())
            ok = editar_servico(os_item["numero_os"], e["proc_id"], serv)
            if not ok:
                st.error("Não foi possível editar. OS pode ter sido enviada para aprovação.")
                return None
        else:
            ok = adicionar_servico(os_item["numero_os"], serv)
            if not ok:
                st.error("Não foi possível adicionar serviço. Verifique o status da OS.")
                return None

        # Limpa session_state do form
        for k in ["_pr_km_ini","_pr_km_fim","_pr_mun_ini","_pr_mun_fim",
                  "_pr_hora_saida","_pr_hora_cheg","_pr_cod_tecnico","_pr_local",
                  "_pr_equip_fixo","_pr_componente_fixo"]:
            st.session_state.pop(k, None)
        st.session_state["_pr_reset"] = True
        return serv
    return None

# ─────────────────────────────────────────────
#  COMPONENTE: timeline de serviços de uma OS
# ─────────────────────────────────────────────

def _render_servicos(os_item: dict, ud: dict, pode_editar: bool = False):
    servs = os_item.get("procedimentos", [])
    if not servs:
        st.info("Nenhum serviço registrado ainda.")
        return

    eh_tecnico = ud["perfil"] == "tecnico"

    for p in servs:
        tempo_fixo_txt = ""
        if p.get("tempo_fixo_aplicado"):
            tempo_fixo_txt = (
                f" &nbsp;·&nbsp; ⏱️ Tempo fixo: <strong>{p.get('componente_tempo_fixo','?')}</strong>"
            )
        c_left, c_right = st.columns([3, 1])
        with c_left:
            st.markdown(f"""
            <div class="serv-card">
                <div class="serv-seq">Serviço #{p.get('seq','?')}</div>
                <div class="serv-tec">{p.get('nome_tecnico','?')} <span style="color:#64748b;font-size:0.78rem;">— {p.get('nivel_tecnico','?')}</span></div>
                <div class="serv-tipo">{p.get('tipo_servico','?')} &nbsp;·&nbsp; {p.get('local_servico','?')} &nbsp;·&nbsp; {p.get('data','?')}{tempo_fixo_txt}</div>
                <div style="margin-top:0.5rem;font-size:0.82rem;color:#cbd5e1;">
                    <strong>Horário:</strong> {p.get('hora_saida','?')} → {p.get('hora_chegada','?')}
                    &nbsp;|&nbsp; <strong>Horas:</strong> {float(p.get('horas_trabalhadas',0)):.2f}h
                    &nbsp;|&nbsp; <strong>KM:</strong> {float(p.get('km_rodado',0)):.1f} km
                </div>
                <div style="margin-top:0.4rem;font-size:0.82rem;color:#94a3b8;">{p.get('descricao','-')}</div>
                {"<div style='margin-top:0.3rem;font-size:0.72rem;color:#4a5568;'>✏️ Editado por " + str(p.get('editado_por','')) + " em " + str(p.get('editado_em',''))[:16].replace('T',' ') + "</div>" if p.get('editado_em') else ""}
            </div>
            """, unsafe_allow_html=True)
        with c_right:
            if eh_tecnico:
                st.markdown(f"""
                <div class="comissao-box" style="background:linear-gradient(135deg,#0c1a2e,#1e2738);border-color:#3d4a5c;">
                    <div class="label" style="color:#94a3b8!important;">Horas</div>
                    <div class="valor" style="color:#60a5fa!important;font-size:1.3rem;">{float(p.get('horas_trabalhadas',0)):.2f}h</div>
                    <div style="margin-top:0.5rem;"><span class="info-pill">KM {float(p.get('km_rodado',0)):.1f}</span></div>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown(f"""
                <div class="comissao-box">
                    <div class="label">Serviço</div>
                    <div class="valor" style="font-size:1.2rem;">R$ {float(p.get('valor_servico',0)):,.2f}</div>
                    <div style="margin-top:0.4rem;"><span class="info-pill">KM R$ {float(p.get('valor_km',0)):,.2f}</span></div>
                    <div style="margin-top:0.5rem;padding-top:0.5rem;border-top:1px solid #166534;">
                        <div class="label">Comissão ({p.get('percentual_comissao','?')}%)</div>
                        <div class="valor">R$ {float(p.get('comissao',0)):,.2f}</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

            # Botão editar serviço
            if pode_editar:
                eh_meu = p.get("registrado_por") == st.session_state["usuario"]
                pode_editar_serv = (
                    (not eh_tecnico) or    # supervisor/admin pode editar qualquer um
                    eh_meu                 # técnico só edita o próprio
                )
                if pode_editar_serv:
                    bc1, bc2 = st.columns(2) if ud["perfil"] == "admin" else (st.container(), None)
                    with bc1:
                        if st.button(f"✏️ Editar", key=f"edit_serv_{p.get('proc_id','')}"):
                            st.session_state["_editando_serv"] = {
                                "numero_os": os_item["numero_os"],
                                "serv":      p,
                            }
                            st.session_state["_pr_reset"] = True
                            st.rerun()

                    # Botão excluir — somente admin
                    if ud["perfil"] == "admin":
                        confirm_key = f"_confirm_excl_{p.get('proc_id','')}"
                        with bc2:
                            if st.session_state.get(confirm_key):
                                if st.button("✅ Confirmar", key=f"conf_excl_{p.get('proc_id','')}"):
                                    ok = excluir_servico(
                                        os_item["numero_os"],
                                        p.get("proc_id", ""),
                                        st.session_state["usuario"],
                                    )
                                    st.session_state.pop(confirm_key, None)
                                    if ok:
                                        st.success("Serviço excluído!")
                                    else:
                                        st.error("Não foi possível excluir o serviço.")
                                    st.rerun()
                            else:
                                if st.button("🗑️ Excluir", key=f"del_serv_{p.get('proc_id','')}"):
                                    st.session_state[confirm_key] = True
                                    st.rerun()
                        if st.session_state.get(confirm_key):
                            st.caption("⚠️ Confirma a exclusão deste serviço? Essa ação não pode ser desfeita.")
                            if st.button("✖️ Cancelar exclusão", key=f"canc_excl_{p.get('proc_id','')}"):
                                st.session_state.pop(confirm_key, None)
                                st.rerun()

# ─────────────────────────────────────────────
#  PÁGINA: ABRIR NOVA OS
# ─────────────────────────────────────────────

def pagina_abrir_os():
    ud     = st.session_state["user_data"]
    perfil = ud["perfil"]
    st.markdown("""<div class="section-header">
        <span style="font-size:1.3rem">➕</span><h3>Abrir Nova Ordem de Serviço</h3>
    </div>""", unsafe_allow_html=True)

    # Quais tipos este perfil pode abrir (a regra vive no dados.py)
    tipos_permitidos = [t for t in TIPOS_OS if pode_abrir_os(perfil, t, "manual")]
    if not tipos_permitidos:
        st.error("Seu perfil não pode abrir OS.")
        return

    if perfil == "tecnico":
        st.info(
            "Você pode abrir **OS externa** (campo, deslocamento, M.S). "
            "OS interna é aberta pelo supervisor, ou automaticamente pelo "
            "app de Checklist de Veículos."
        )
    else:
        st.info("Informe os dados do equipamento. Os técnicos adicionarão os serviços depois.")

    FROTAS = carregar_frotas()

    with st.form("form_nova_os"):
        tipo_os = st.radio(
            "Tipo de OS", tipos_permitidos, horizontal=True,
            format_func=lambda t: {
                "interna": "🏠 Interna (barracão)",
                "externa": "🚚 Externa (campo / deslocamento / M.S)",
            }.get(t, t),
        )

        c1, c2 = st.columns(2)
        with c1:
            data_ab   = st.date_input("Data de Abertura", value=date.today())
            frota_val = st.text_input("Frota", placeholder="Ex: 1201")
        with c2:
            lista_cc = {f"{k} - {v}": k for k, v in CENTROS_CUSTO.items()}
            ccs = list(lista_cc.keys())
            cod_cc = lista_cc[st.selectbox("Centro de Custo / Cliente", ccs)]

        frota = frota_val.strip()
        desc_frota = FROTAS.get(frota, "")

        equipamento = st.text_input(
            "Equipamento / Descrição",
            value=desc_frota,
            placeholder="Será preenchido automaticamente ao digitar a frota",
        )

        avaliacao = st.text_area(
            "Avaliação",
            placeholder=("Descreva o problema apresentado pelo veículo — ex: para-brisa "
                         "trincado, motor com fumaça excessiva, vazamento hidráulico."),
            height=110,
            help="Obrigatório. É o que orienta quem vai executar o serviço.",
        )

        # Técnico: mecânico abre para si mesmo; supervisor escolhe.
        if perfil == "tecnico" and ud.get("cod_tecnico"):
            cod_tecnico_sel = ud["cod_tecnico"]
            nome_tec_sel    = TECNICOS[cod_tecnico_sel]["nome"]
            st.text_input("Técnico Responsável", value=nome_tec_sel, disabled=True)
        else:
            lista_tec_nomes = {v["nome"]: k for k, v in TECNICOS.items()}
            nome_tec_sel = st.selectbox(
                "Técnico Responsável", list(lista_tec_nomes.keys()),
                help="Apenas este técnico verá esta OS para lançar serviços e encerrá-la.",
            )
            cod_tecnico_sel = lista_tec_nomes[nome_tec_sel]

        if st.form_submit_button("🚀 Abrir OS", type="primary", use_container_width=True):
            if not frota:
                st.error("Informe a frota.")
            elif not avaliacao.strip():
                st.error("Preencha a Avaliação — descreva o problema do veículo.")
            else:
                try:
                    nova = criar_os(
                        frota=frota,
                        equipamento=equipamento or desc_frota or frota,
                        cod_cc=cod_cc,
                        aberto_por=st.session_state["usuario"],
                        data_abertura=str(data_ab),
                        tecnico_designado=cod_tecnico_sel,
                        perfil_usuario=perfil,
                        tipo_os=tipo_os,
                        avaliacao=avaliacao,
                        origem="manual",
                    )
                except PermissaoNegada as e:
                    st.error(f"🚫 {e}")
                    return

                st.success(f"✅ OS **{nova['numero_os']}** criada com sucesso!")
                st.balloons()
                st.markdown(f"""
                <div style="background:#052e16;border:2px solid #16a34a;border-radius:12px;
                     padding:1.5rem;text-align:center;margin-top:1rem;">
                    <div style="font-size:0.8rem;color:#86efac;font-weight:600;text-transform:uppercase;
                         letter-spacing:0.08em;">Número da OS</div>
                    <div style="font-size:2.5rem;font-weight:900;color:#4ade80;letter-spacing:0.1em;">
                        {nova['numero_os']}
                    </div>
                    <div style="font-size:0.85rem;color:#6ee7b7;margin-top:0.25rem;">
                        Frota {frota} · {CENTROS_CUSTO[cod_cc]} · OS {tipo_os}
                    </div>
                    <div style="font-size:0.85rem;color:#6ee7b7;margin-top:0.15rem;">
                        Técnico responsável: {nome_tec_sel}
                    </div>
                </div>
                """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  PÁGINA: ADICIONAR SERVIÇO (técnico)
# ─────────────────────────────────────────────

def pagina_adicionar_servico():
    ud = st.session_state["user_data"]

    # Modo edição de serviço
    if st.session_state.get("_editando_serv"):
        info  = st.session_state["_editando_serv"]
        os_it = buscar_os_por_numero(info["numero_os"])
        if not os_it:
            st.error("OS não encontrada.")
            st.session_state.pop("_editando_serv", None)
            st.rerun()
            return
        st.markdown("""<div class="section-header">
            <span style="font-size:1.3rem">✏️</span><h3>Editar Serviço</h3>
        </div>""", unsafe_allow_html=True)
        if st.button("← Voltar sem salvar"):
            st.session_state.pop("_editando_serv", None)
            st.session_state["_pr_reset"] = True
            st.rerun()
        res = _form_servico(ud, os_it, serv_existente=info["serv"])
        if res:
            st.success("✅ Serviço atualizado!")
            st.session_state.pop("_editando_serv", None)
            st.rerun()
        return

    st.markdown("""<div class="section-header">
        <span style="font-size:1.3rem">🔧</span><h3>Adicionar Serviço a uma OS</h3>
    </div>""", unsafe_allow_html=True)

    # ── Se já tem uma OS selecionada, mostra o formulário ──
    numero_os_sel = st.session_state.get("_serv_numero_os")
    if numero_os_sel and st.session_state.get("_serv_os_carregada"):
        os_item = buscar_os_por_numero(numero_os_sel)

        if not os_item or os_item["status"] not in ("aberta", "em_andamento"):
            st.session_state.pop("_serv_os_carregada", None)
            st.session_state.pop("_serv_numero_os", None)
            st.rerun()
            return

        # Trava de segurança: técnico só acessa OS designada a ele
        # (ou legada, sem designação).
        if ud["perfil"] == "tecnico":
            tec_desig = os_item.get("tecnico_designado")
            if tec_desig not in (None, ud.get("cod_tecnico")):
                st.error("🚫 Esta OS está designada a outro técnico. Você não pode lançar serviços nela.")
                st.session_state.pop("_serv_os_carregada", None)
                st.session_state.pop("_serv_numero_os", None)
                if st.button("← Voltar para lista de OS"):
                    st.rerun()
                return

        if st.button("← Voltar para lista de OS"):
            st.session_state.pop("_serv_os_carregada", None)
            st.session_state.pop("_serv_numero_os", None)
            st.session_state["_pr_reset"] = True
            st.rerun()

        # Card da OS selecionada
        st.markdown(f"""
        <div class="os-card" style="margin-bottom:1rem;">
            <div class="os-numero">{os_item.get('numero_os','—')}</div>
            <div class="os-frota">Frota {os_item.get('frota','-')} · {os_item.get('equipamento','-')} · {os_item.get('cliente','-')}</div>
            <div style="margin-top:0.4rem;">
                {_badge(os_item.get('status',''))}
                &nbsp; <span style="font-size:0.78rem;color:#64748b;">
                    {len(os_item.get('procedimentos',[]))} serviço(s) já registrado(s)
                </span>
            </div>
        </div>
        """, unsafe_allow_html=True)

        if os_item.get("procedimentos"):
            with st.expander(f"📋 Ver {len(os_item['procedimentos'])} serviço(s) anteriores"):
                _render_servicos(os_item, ud, pode_editar=False)

        st.markdown("---")
        st.markdown("**Novo Serviço**")
        res = _form_servico(ud, os_item)
        if res:
            st.success(f"✅ Serviço #{res.get('seq','?')} adicionado à OS {numero_os_sel}!")

            os_atualizada = buscar_os_por_numero(numero_os_sel)
            if os_atualizada and os_atualizada["status"] == "em_andamento":
                st.markdown("""
                <div class="concluir-box">
                    <div style="font-size:0.9rem;font-weight:700;color:#a5b4fc;">🏁 Encerrar esta OS?</div>
                    <div style="font-size:0.82rem;color:#818cf8;margin-top:0.25rem;">
                        Se você é o último técnico, clique abaixo para enviar a OS para aprovação do supervisor.
                        Outros técnicos ainda podem adicionar serviços antes disso.
                    </div>
                </div>
                """, unsafe_allow_html=True)
                if st.button("🏁 Concluir OS — Enviar para Aprovação", use_container_width=True):
                    fechar_os_para_aprovacao(numero_os_sel, st.session_state["usuario"])
                    st.success("✅ OS enviada para aprovação!")
                    st.session_state.pop("_serv_os_carregada", None)
                    st.session_state.pop("_serv_numero_os", None)
                    st.rerun()
        return

    # ── Lista de OS abertas/em andamento ──
    todas = carregar_os()
    disponiveis = [
        o for o in todas
        if o.get("status") in ("aberta", "em_andamento")
    ]

    # Técnico só vê OS designadas a ele (ou OS legadas sem técnico
    # designado, que continuam abertas a todos até serem fechadas).
    if ud["perfil"] == "tecnico":
        meu_cod = ud.get("cod_tecnico")
        disponiveis = [
            o for o in disponiveis
            if o.get("tecnico_designado") in (None, meu_cod)
        ]

    if not disponiveis:
        st.info("Nenhuma OS aberta no momento. Aguarde o supervisor abrir uma OS.")
        return

    # Filtro rápido por frota
    busca = st.text_input("🔍 Filtrar por frota ou número", placeholder="Ex: 1201 ou OS-2026-0003")
    if busca:
        b = busca.lower()
        disponiveis = [
            o for o in disponiveis
            if b in o.get("frota","").lower() or b in o.get("numero_os","").lower()
        ]

    st.markdown(f"**{len(disponiveis)} OS disponível(is)**")
    st.markdown("")

    for os_item in sorted(disponiveis, key=lambda x: x.get("data_abertura",""), reverse=True):
        servs      = os_item.get("procedimentos", [])
        n_servs    = len(servs)
        tecs_str   = _resumo_tecnicos(os_item) if n_servs else ""
        tecs_linha = ("&nbsp;&middot;&nbsp; Técnicos: " + tecs_str) if n_servs else ""
        badge_html = _badge(os_item.get("status", ""))
        numero     = os_item.get("numero_os", "—")
        frota      = os_item.get("frota", "-")
        equip      = os_item.get("equipamento", "-")
        cliente    = os_item.get("cliente", "-")
        dt_ab      = os_item.get("data_abertura", "-")
        tec_desig  = os_item.get("tecnico_designado")
        tec_desig_nome = TECNICOS.get(tec_desig, {}).get("nome") if tec_desig else None
        tec_desig_linha = (
            f'<div style="margin-top:0.2rem;font-size:0.78rem;color:#a5b4fc;">'
            f'👤 Designada a: <strong>{tec_desig_nome}</strong></div>'
            if tec_desig_nome else
            '<div style="margin-top:0.2rem;font-size:0.78rem;color:#f59e0b;">'
            '⚠️ Sem técnico designado (visível a todos)</div>'
        )

        html_card = (
            '<div class="os-card" style="margin-bottom:0;">'
            '<div style="display:flex;align-items:center;gap:0.75rem;flex-wrap:wrap;">'
            f'<div class="os-numero">{numero}</div>'
            f'{badge_html}'
            '</div>'
            '<div style="margin-top:0.35rem;font-size:0.85rem;color:#cbd5e1;">'
            f'<strong>Frota {frota}</strong>'
            f'&nbsp;&middot;&nbsp; {equip}'
            f'&nbsp;&middot;&nbsp; {cliente}'
            '</div>'
            f'{tec_desig_linha}'
            '<div style="margin-top:0.25rem;font-size:0.78rem;color:#64748b;">'
            f'Aberta em {dt_ab}'
            f'&nbsp;&middot;&nbsp; {n_servs} serviço(s)'
            f'{tecs_linha}'
            '</div>'
            '</div>'
        )

        col_card, col_btn = st.columns([5, 1])
        with col_card:
            st.markdown(html_card, unsafe_allow_html=True)
        with col_btn:
            st.markdown("<div style='height:1.6rem'></div>", unsafe_allow_html=True)
            if st.button("Trabalhar →", key=f"sel_{os_item.get('numero_os','')}",
                         use_container_width=True, type="primary"):
                st.session_state["_serv_numero_os"]    = os_item.get("numero_os")
                st.session_state["_serv_os_carregada"] = True
                st.session_state["_pr_reset"]          = True
                st.rerun()
        st.markdown("<div style='margin-bottom:0.5rem'></div>", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  PÁGINA: MINHAS OS / TODAS AS OS
# ─────────────────────────────────────────────

def pagina_os(somente_minhas: bool = False):
    ud         = st.session_state["user_data"]
    eh_tecnico = ud["perfil"] == "tecnico"

    # Se tem edição de serviço pendente, delega para pagina_adicionar_servico
    if st.session_state.get("_editando_serv"):
        pagina_adicionar_servico()
        return

    titulo = "📋 Minhas OS" if somente_minhas else "📋 Todas as OS"

    st.markdown(f"""<div class="section-header">
        <span style="font-size:1.3rem">📋</span><h3>{titulo}</h3>
    </div>""", unsafe_allow_html=True)

    todas = carregar_os()

    if somente_minhas and eh_tecnico:
        meu_cod = ud.get("cod_tecnico")
        todas = [
            o for o in todas
            if o.get("tecnico_designado") == meu_cod
            or any(p.get("cod_tecnico") == meu_cod for p in o.get("procedimentos", []))
        ]

    if not todas:
        st.info("Nenhuma OS encontrada.")
        return

    # Filtros
    cf1, cf2, cf3 = st.columns(3)
    with cf1:
        status_opts = ["Todos", "aberta", "em_andamento", "aguardando_aprovacao", "aprovada", "rejeitada"]
        filtro_status = st.selectbox("Status", status_opts)
    with cf2:
        busca_frota = st.text_input("🔍 Frota / Número OS")
    with cf3:
        todas_datas = sorted({o.get("data_abertura","") for o in todas if o.get("data_abertura")}, reverse=True)
        if todas_datas:
            meses = list(dict.fromkeys(d[:7] for d in todas_datas))
            filtro_mes = st.selectbox("Mês", ["Todos"] + meses)
        else:
            filtro_mes = "Todos"

    lista = todas
    if filtro_status != "Todos":
        lista = [o for o in lista if o.get("status") == filtro_status]
    if busca_frota:
        b = busca_frota.lower()
        lista = [o for o in lista
                 if b in o.get("frota","").lower() or b in o.get("numero_os","").lower()]
    if filtro_mes != "Todos":
        lista = [o for o in lista if o.get("data_abertura","").startswith(filtro_mes)]

    st.markdown(f"**{len(lista)} OS encontradas**")

    for os_item in sorted(lista, key=lambda x: x.get("data_abertura",""), reverse=True):
        servs   = os_item.get("procedimentos", [])
        n_servs = len(servs)
        total_com  = _total_comissao_os(os_item)
        total_serv = _total_servico_os(os_item)

        with st.expander(
            f"{os_item.get('numero_os','—')}  ·  Frota {os_item.get('frota','-')}  ·  "
            f"{os_item.get('equipamento','-')}  ·  {os_item.get('cliente','-')}  ·  "
            f"{os_item.get('data_abertura','-')}  ·  {n_servs} serviço(s)"
        ):
            # Header da OS
            col_info, col_totais = st.columns([3, 1])
            with col_info:
                tec_desig = os_item.get("tecnico_designado")
                tec_desig_nome = TECNICOS.get(tec_desig, {}).get("nome") if tec_desig else None
                tec_desig_html = (
                    f" &nbsp;&middot;&nbsp; 👤 Designada a <strong>{tec_desig_nome}</strong>"
                    if tec_desig_nome else
                    " &nbsp;&middot;&nbsp; <span style='color:#f59e0b;'>⚠️ sem técnico designado</span>"
                )
                st.markdown(
                    f"{_badge(os_item.get('status',''))} &nbsp; "
                    f"<span style='font-size:0.8rem;color:#64748b;'>"
                    f"Aberta por <strong>{os_item.get('aberto_por','?')}</strong> "
                    f"em {str(os_item.get('aberto_em',''))[:10]}"
                    f"{tec_desig_html}"
                    f"</span>",
                    unsafe_allow_html=True
                )
                origem_os = os_item.get("origem", "manual")
                tipo_os   = os_item.get("tipo_os", "—")
                selo = ("🔗 vinda do Checklist" if origem_os == "checklist"
                        else "✍️ aberta manualmente")
                st.markdown(
                    f'<div style="font-size:0.8rem;color:#94a3b8;margin-bottom:0.4rem;">'
                    f'{selo} &nbsp;·&nbsp; OS <strong>{tipo_os}</strong></div>',
                    unsafe_allow_html=True)

                if os_item.get("cc_pendente"):
                    st.warning(
                        f"⚠️ Centro de custo não reconhecido: “{os_item.get('cliente','—')}”. "
                        "Corrija antes de aprovar esta OS."
                    )

                if os_item.get("avaliacao"):
                    with st.expander("🩺 Avaliação (problema relatado)", expanded=False):
                        st.text(os_item["avaliacao"])
                if os_item.get("observacao_validacao"):
                    st.caption(f"💬 Obs. validação: {os_item['observacao_validacao']}")
            with col_totais:
                if not eh_tecnico:
                    st.markdown(f"""
                    <div class="comissao-box">
                        <div class="label">Total OS ({n_servs} serv.)</div>
                        <div class="valor" style="font-size:1.2rem;">R$ {total_serv:,.2f}</div>
                        <div style="margin-top:0.4rem;padding-top:0.4rem;border-top:1px solid #166534;">
                            <div class="label">Comissão Total</div>
                            <div class="valor">R$ {total_com:,.2f}</div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    total_h = sum(float(p.get("horas_trabalhadas", 0)) for p in servs)
                    st.markdown(f"""
                    <div class="comissao-box" style="background:linear-gradient(135deg,#0c1a2e,#1e2738);border-color:#3d4a5c;">
                        <div class="label" style="color:#94a3b8!important;">Total Horas</div>
                        <div class="valor" style="color:#60a5fa!important;font-size:1.3rem;">{total_h:.2f}h</div>
                    </div>
                    """, unsafe_allow_html=True)

            st.markdown("---")

            # Serviços
            pode_editar_serv = os_item.get("status") in ("aberta", "em_andamento")
            _render_servicos(os_item, ud, pode_editar=pode_editar_serv)

            # ── Botão concluir OS — visível para TODOS (técnico e admin)
            status_os = os_item.get("status", "")
            if status_os == "em_andamento" and n_servs > 0:
                st.markdown("---")
                st.markdown("""
                <div class="concluir-box">
                    <div style="font-size:0.9rem;font-weight:700;color:#a5b4fc;">🏁 Concluir esta OS?</div>
                    <div style="font-size:0.82rem;color:#818cf8;margin-top:0.25rem;">
                        Se todos os serviços foram registrados, clique abaixo para enviar para aprovação do supervisor.
                    </div>
                </div>
                """, unsafe_allow_html=True)
                if st.button("🏁 Enviar para Aprovação",
                             key=f"fechar_todos_{os_item.get('numero_os','')}",
                             type="primary", use_container_width=True):
                    fechar_os_para_aprovacao(os_item.get("numero_os"), st.session_state["usuario"])
                    st.success("✅ OS enviada para aprovação!")
                    st.rerun()

            # Ações extras (somente supervisor/admin)
            if not eh_tecnico:
                st.markdown("---")
                status = os_item.get("status", "")

                # Definir / alterar técnico designado
                if status in ("aberta", "em_andamento"):
                    lista_tec_nomes = {v["nome"]: k for k, v in TECNICOS.items()}
                    nomes_tec = list(lista_tec_nomes.keys())
                    idx_atual = (
                        nomes_tec.index(tec_desig_nome)
                        if tec_desig_nome in nomes_tec else 0
                    )
                    col_sel, col_btn2 = st.columns([3, 1])
                    with col_sel:
                        novo_nome_tec = st.selectbox(
                            "Técnico designado",
                            nomes_tec, index=idx_atual,
                            key=f"tec_desig_{os_item.get('numero_os','')}",
                        )
                    with col_btn2:
                        st.markdown("<div style='height:1.6rem'></div>", unsafe_allow_html=True)
                        if st.button("Salvar", key=f"salvar_tec_{os_item.get('numero_os','')}",
                                     use_container_width=True):
                            definir_tecnico_designado(
                                os_item.get("numero_os"),
                                lista_tec_nomes[novo_nome_tec],
                                st.session_state["usuario"],
                            )
                            st.success(f"Técnico designado: {novo_nome_tec}")
                            st.rerun()

                # Adicionar serviço (supervisor)
                if status in ("aberta", "em_andamento"):
                    if st.button("➕ Adicionar Serviço", key=f"add_serv_{os_item.get('numero_os','')}"):
                        st.session_state["_serv_numero_os"]    = os_item.get("numero_os")
                        st.session_state["_serv_os_carregada"] = True
                        st.session_state["page"] = "🔧 Adicionar Serviço"
                        st.rerun()

                # Reabrir (aguardando aprovação → em_andamento)
                if status == "aguardando_aprovacao":
                    if st.button("🔄 Reabrir OS (adicionar mais serviços)",
                                 key=f"reabrir_{os_item.get('numero_os','')}"):
                        reabrir_os(os_item.get("numero_os"), st.session_state["usuario"])
                        st.success("OS reaberta.")
                        st.rerun()

# ─────────────────────────────────────────────
#  PÁGINA: VALIDAÇÕES
# ─────────────────────────────────────────────

def pagina_validacoes():
    st.markdown("""<div class="section-header">
        <span style="font-size:1.3rem">✅</span><h3>Validação de Ordens de Serviço</h3>
    </div>""", unsafe_allow_html=True)

    todas     = carregar_os()
    pendentes = [o for o in todas if o.get("status") == "aguardando_aprovacao"]

    if not pendentes:
        st.success("✅ Nenhuma OS aguardando validação!")
        return

    st.info(f"**{len(pendentes)} OS aguardando aprovação**")

    for os_item in pendentes:
        servs      = os_item.get("procedimentos", [])
        total_com  = _total_comissao_os(os_item)
        total_serv = _total_servico_os(os_item)

        with st.expander(
            f"🔔 {os_item['numero_os']} · Frota {os_item['frota']} · "
            f"{os_item.get('cliente','-')} · {os_item.get('data_abertura','-')} · "
            f"{len(servs)} serviço(s) · Comissão Total: R$ {total_com:,.2f}"
        ):
            col_info, col_val = st.columns([3, 1])
            with col_info:
                st.markdown(f"""
**OS:** {os_item['numero_os']} &nbsp;·&nbsp;
**Frota:** {os_item['frota']} &nbsp;·&nbsp;
**Equipamento:** {os_item.get('equipamento','-')}

**Centro de Custo:** {os_item.get('cod_cc','?')} — {os_item.get('cliente','?')}

**Aberta por:** {os_item.get('aberto_por','?')} em {str(os_item.get('aberto_em',''))[:10]}

**Fechada para aprovação por:** {os_item.get('fechado_para_aprovacao_por','?')} em {str(os_item.get('fechado_para_aprovacao_em',''))[:16].replace('T',' ')}

**Técnicos envolvidos:** {_resumo_tecnicos(os_item)}
                """)
            with col_val:
                st.markdown(f"""
                <div class="comissao-box">
                    <div class="label">Total Serviços ({len(servs)} serv.)</div>
                    <div class="valor" style="font-size:1.2rem;">R$ {total_serv:,.2f}</div>
                    <div style="margin-top:0.5rem;padding-top:0.5rem;border-top:1px solid #166534;">
                        <div class="label">Comissão Total</div>
                        <div class="valor">R$ {total_com:,.2f}</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("**Serviços:**")
            _render_servicos(os_item, st.session_state["user_data"], pode_editar=False)

            st.markdown("---")
            col_ok, col_rej, col_obs = st.columns([1, 1, 2])
            obs = col_obs.text_input("Observação", key=f"obs_{os_item['numero_os']}",
                                     placeholder="Motivo da rejeição ou observação...")
            with col_ok:
                if st.button("✅ Aprovar", key=f"apr_{os_item['numero_os']}",
                             type="primary", use_container_width=True):
                    validar_os(os_item["numero_os"], "aprovada",
                               st.session_state["usuario"], obs)
                    st.success("OS aprovada!")
                    st.rerun()
            with col_rej:
                if st.button("❌ Rejeitar", key=f"rej_{os_item['numero_os']}",
                             use_container_width=True):
                    if not obs.strip():
                        st.warning("Informe o motivo da rejeição.")
                    else:
                        validar_os(os_item["numero_os"], "rejeitada",
                                   st.session_state["usuario"], obs)
                        st.error("OS rejeitada.")
                        st.rerun()

# ─────────────────────────────────────────────
#  PÁGINA: DASHBOARD
# ─────────────────────────────────────────────

def pagina_dashboard():
    st.markdown("""<div class="section-header">
        <span style="font-size:1.3rem">📊</span><h3>Dashboard Geral</h3>
    </div>""", unsafe_allow_html=True)

    todas = carregar_os()
    if not todas:
        st.info("Nenhuma OS registrada.")
        return

    # Métricas gerais
    def _count(s):  return sum(1 for o in todas if o.get("status") == s)
    def _com(s):    return sum(_total_comissao_os(o) for o in todas if o.get("status") == s)
    def _serv(s):   return sum(_total_servico_os(o)  for o in todas if o.get("status") == s)

    c1, c2, c3, c4, c5 = st.columns(5)
    cards = [
        (c1, "yellow",  "Abertas",          _count("aberta"),               ""),
        (c2, "yellow",  "Em Andamento",      _count("em_andamento"),         ""),
        (c3, "",        "Aguard. Aprovação", _count("aguardando_aprovacao"), ""),
        (c4, "green",   "Aprovadas",         _count("aprovada"),             ""),
        (c5, "purple",  "Comissões Aprov.",  f"R$ {_com('aprovada'):,.0f}",  "a pagar"),
    ]
    for col, cor, label, valor, sub in cards:
        with col:
            st.markdown(f"""<div class="metric-card {cor}">
                <div class="metric-label">{label}</div>
                <div class="metric-value">{valor}</div>
                <div class="metric-sub">{sub}</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("---")

    # Tabela por técnico (aprovadas)
    aprovadas = [o for o in todas if o.get("status") == "aprovada"]
    if aprovadas:
        rows = []
        for o in aprovadas:
            for p in o.get("procedimentos", []):
                rows.append({
                    "OS":       o["numero_os"],
                    "Técnico":  p.get("nome_tecnico","?"),
                    "Horas":    float(p.get("horas_trabalhadas", 0)),
                    "KM":       float(p.get("km_rodado", 0)),
                    "Serviço":  float(p.get("valor_servico", 0)),
                    "Comissão": float(p.get("comissao", 0)),
                })
        if rows:
            df = pd.DataFrame(rows)
            c5, c6 = st.columns(2)
            with c5:
                st.markdown("**💰 Comissão por Técnico (aprovadas)**")
                st.bar_chart(df.groupby("Técnico")["Comissão"].sum().sort_values())
            with c6:
                st.markdown("**⏱ Horas por Técnico (aprovadas)**")
                st.bar_chart(df.groupby("Técnico")["Horas"].sum().sort_values())

# ─────────────────────────────────────────────
#  PÁGINA: COMISSÕES
# ─────────────────────────────────────────────

def pagina_comissoes():
    st.markdown("""<div class="section-header">
        <span style="font-size:1.3rem">💰</span><h3>Comissões</h3>
    </div>""", unsafe_allow_html=True)

    todas     = carregar_os()
    aprovadas = [o for o in todas if o.get("status") == "aprovada"]
    if not aprovadas:
        st.info("Nenhuma OS aprovada ainda.")
        return

    c1, c2 = st.columns(2)
    with c1: data_ini = st.date_input("De",  value=date.today().replace(day=1))
    with c2: data_fim = st.date_input("Até", value=date.today())

    rows = []
    for o in aprovadas:
        dt = o.get("data_abertura", "")
        if not (str(data_ini) <= dt <= str(data_fim)):
            continue
        for p in o.get("procedimentos", []):
            rows.append({
                "OS":          o["numero_os"],
                "Frota":       o.get("frota",""),
                "Data":        p.get("data",""),
                "Técnico":     p.get("nome_tecnico","?"),
                "Nível":       p.get("nivel_tecnico","?"),
                "Tipo":        p.get("tipo_servico","?"),
                "Local":       p.get("local_servico","?"),
                "Cliente":     o.get("cliente","?"),
                "Horas":       float(p.get("horas_trabalhadas",0)),
                "KM":          float(p.get("km_rodado",0)),
                "Vlr Serviço": float(p.get("valor_servico",0)),
                "% Comissão":  float(p.get("percentual_comissao",0)),
                "Comissão":    float(p.get("comissao",0)),
            })

    if not rows:
        st.info("Nenhuma OS aprovada no período.")
        return

    df = pd.DataFrame(rows)
    st.markdown("**Resumo por Técnico**")
    resumo = df.groupby("Técnico").agg(
        Servicos=("Comissão","count"),
        Horas=("Horas","sum"),
        KM=("KM","sum"),
        Vlr_Servico=("Vlr Serviço","sum"),
        Comissao=("Comissão","sum"),
    ).reset_index()
    resumo = resumo.rename(columns={
        "Servicos":   "Serviços",
        "Vlr_Servico":"Vlr Serviço",
        "Comissao":   "Comissão",
    })
    st.dataframe(
        resumo.style.format({
            "Vlr Serviço":"R$ {:.2f}", "Comissão":"R$ {:.2f}",
            "Horas":"{:.2f}", "KM":"{:.1f}",
        }),
        use_container_width=True, hide_index=True
    )

    st.markdown("---")
    st.markdown("**Detalhamento por Serviço**")
    st.dataframe(
        df.style.format({
            "Vlr Serviço":"R$ {:.2f}", "Comissão":"R$ {:.2f}",
            "Horas":"{:.2f}", "KM":"{:.1f}", "% Comissão":"{:.0f}%",
        }),
        use_container_width=True, hide_index=True
    )

# ─────────────────────────────────────────────
#  PÁGINA: EXPORTAR RELATÓRIO
# ─────────────────────────────────────────────

def pagina_exportar():
    st.markdown("""<div class="section-header">
        <span style="font-size:1.3rem">📤</span><h3>Exportar Relatório para Excel</h3>
    </div>""", unsafe_allow_html=True)

    todas = carregar_os()
    if not todas:
        st.info("Nenhuma OS registrada.")
        return

    cf1, cf2, cf3 = st.columns(3)
    with cf1: data_ini = st.date_input("De",  value=date.today().replace(day=1))
    with cf2: data_fim = st.date_input("Até", value=date.today())
    with cf3:
        status_opts = ["Todos","aprovada","aguardando_aprovacao","em_andamento","aberta","rejeitada"]
        status_sel  = st.selectbox("Status", status_opts)

    rows = []
    for o in todas:
        dt = o.get("data_abertura","")
        if not (str(data_ini) <= dt <= str(data_fim)):
            continue
        if status_sel != "Todos" and o.get("status") != status_sel:
            continue
        for p in o.get("procedimentos", []):
            rows.append({
                "OS":             o["numero_os"],
                "Data Abertura":  o.get("data_abertura",""),
                "Status OS":      o.get("status",""),
                "Frota":          o.get("frota",""),
                "Equipamento":    o.get("equipamento",""),
                "Cod CC":         o.get("cod_cc",""),
                "Cliente":        o.get("cliente",""),
                "Seq Servico":    p.get("seq",""),
                "Data Servico":   p.get("data",""),
                "Tecnico":        p.get("nome_tecnico",""),
                "Nivel":          p.get("nivel_tecnico",""),
                "Tipo Servico":   p.get("tipo_servico",""),
                "Local":          p.get("local_servico",""),
                "Componente Tempo Fixo": p.get("componente_tempo_fixo","") or "",
                "Tempo Fixo Aplicado":   "Sim" if p.get("tempo_fixo_aplicado") else "Não",
                "Hora Saida":     p.get("hora_saida",""),
                "Hora Chegada":   p.get("hora_chegada",""),
                "Horas Trab":     float(p.get("horas_trabalhadas",0)),
                "KM Ini":         float(p.get("km_inicial",0)),
                "KM Fim":         float(p.get("km_final",0)),
                "KM Rodado":      float(p.get("km_rodado",0)),
                "Munck Ini":      float(p.get("hora_munck_inicial",0)),
                "Munck Fim":      float(p.get("hora_munck_final",0)),
                "Horas Munck":    float(p.get("horas_munck",0)),
                "Vlr Servico":    float(p.get("valor_servico",0)),
                "Vlr KM":         float(p.get("valor_km",0)),
                "Vlr Munck":      float(p.get("valor_munck",0)),
                "Perc Comissao":  float(p.get("percentual_comissao",0)),
                "Comissao":       float(p.get("comissao",0)),
                "Registrado Por": p.get("registrado_por",""),
                "Registrado Em":  p.get("registrado_em",""),
                "Validado Por":   o.get("validado_por",""),
                "Validado Em":    o.get("validado_em",""),
                "Obs Validacao":  o.get("observacao_validacao",""),
            })

    if not rows:
        st.warning("Nenhum dado com os filtros selecionados.")
        return

    df = pd.DataFrame(rows)
    st.markdown(f"**{len(df)} serviços / {len(set(r['OS'] for r in rows))} OS no filtro**")
    st.dataframe(df.head(10), use_container_width=True, hide_index=True)

    st.markdown("---")
    t1, t2, t3, t4 = st.columns(4)
    with t1: st.metric("Serviços",       len(df))
    with t2: st.metric("OS distintas",   len(set(r["OS"] for r in rows)))
    with t3: st.metric("Horas Totais",   f"{df['Horas Trab'].sum():.2f}h")
    with t4: st.metric("Total Comissões",f"R$ {df['Comissao'].sum():,.2f}")

    st.markdown("---")
    nome_arq = (
        f"OS_{status_sel}_{data_ini.strftime('%d-%m-%Y')}_a_{data_fim.strftime('%d-%m-%Y')}.xlsx"
    )
    st.download_button(
        "⬇️ Baixar Excel (.xlsx)", data=df_para_excel(df),
        file_name=nome_arq,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary", use_container_width=True,
    )
    csv_bytes = df.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button(
        "⬇️ Baixar CSV", data=csv_bytes,
        file_name=nome_arq.replace(".xlsx",".csv"),
        mime="text/csv", use_container_width=True,
    )

# ─────────────────────────────────────────────
#  PÁGINA: CONFIGURAÇÕES
# ─────────────────────────────────────────────

def pagina_configuracoes():
    st.markdown("""<div class="section-header">
        <span style="font-size:1.3rem">⚙️</span><h3>Configurações do Sistema</h3>
    </div>""", unsafe_allow_html=True)
    tab1, tab2, tab3 = st.tabs(["Tabela de Valores", "Técnicos", "Tempos Fixos por Componente"])
    with tab1:
        st.markdown("**KM Rodado**")
        st.dataframe(pd.DataFrame([{"Tipo KM":k,"R$/km":v} for k,v in TABELA_KM.items()]), hide_index=True)
        st.markdown("**Valor Hora por Nível**")
        st.dataframe(pd.DataFrame([{"Nível":k,"R$/hora":v} for k,v in TABELA_HORA.items()]), hide_index=True)
        st.markdown("**Comissão por Local**")
        st.dataframe(pd.DataFrame([{"Local":k,"%":v} for k,v in PERCENTUAIS_LOCAL.items()]), hide_index=True)
    with tab2:
        st.dataframe(pd.DataFrame([
            {"Cód":k,"Nome":v["nome"],"Nível":v["nivel"],"Função":v["funcao"],
             "KM":v["tipo_km"],"R$/hora":TABELA_HORA.get(v["nivel"],50)}
            for k,v in TECNICOS.items()
        ]), use_container_width=True, hide_index=True)
    with tab3:
        tabela_tempos, _ = carregar_tabela_tempos()
        if not tabela_tempos:
            st.warning("Planilha TEMPO_SERVIÇO.xlsx não encontrada na pasta do app.")
        else:
            st.caption("1 dia útil = 8,8h (jornada semanal de 44h em 5 dias, 7h–17h com 1h12 de almoço).")
            st.dataframe(pd.DataFrame([
                {"Equipamento": eq, "Componente": comp, "Tempo Fixo (h)": round(h, 2),
                 "Tempo Fixo": formatar_horas(h)}
                for (eq, comp), h in sorted(tabela_tempos.items())
            ]), use_container_width=True, hide_index=True)

# ─────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────

def main():
    if not st.session_state.get("logado"):
        tela_login()
        return

    ud = st.session_state["user_data"]
    st.markdown(f"""
    <div class="main-header">
        <span class="icon">🔧</span>
        <div>
            <h1>Assistência Técnica Teston</h1>
            <p>Sistema de OS e Comissões · Olá, <strong>{ud['nome']}</strong></p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    page = render_sidebar()

    if "Minhas OS" in page:
        pagina_os(somente_minhas=True)
    elif "Adicionar Serviço" in page:
        pagina_adicionar_servico()
    elif "Todas as OS" in page:
        pagina_os(somente_minhas=False)
    elif "Abrir Nova OS" in page:
        pagina_abrir_os()
    elif "Validações" in page:
        pagina_validacoes()
    elif "Dashboard" in page:
        pagina_dashboard()
    elif "Comissões" in page:
        pagina_comissoes()
    elif "Exportar" in page:
        pagina_exportar()
    elif "Configurações" in page:
        pagina_configuracoes()

if __name__ == "__main__":
    main()