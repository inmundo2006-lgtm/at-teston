import io
import streamlit as st
import pandas as pd
import hashlib
import os
from datetime import date, datetime, time

from calculos import calcular_servico, calcular_deslocamento, calcular_lavagem
from tempos_fixos import carregar_tabela_tempos, formatar_horas, NENHUM_COMPONENTE
from medias import avaliar_tempo, ranking_acima_media
from cidades import campo_cidade

# ─────────────────────────────────────────────
#  TABELA DE FROTAS
# ─────────────────────────────────────────────
# Vem da lista KanbanFrotas do SharePoint — a mesma que o app de Checklist
# usa. O CADASTRO.xlsx virou fallback (só descrição), porque só era
# atualizado quando alguém lembrava de reenviar a planilha.
from frotas import carregar_frotas, alerta_status

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
    LOCAIS_OS, LABELS_LOCAIS_OS, tipo_os_do_local,
    NATUREZAS, LABELS_NATUREZA, resolver_cod_cc,
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
    """
    Valor cobrado da OS — equivale à coluna "Serviço" da planilha
    (horas + KM + munck + deslocamento).

    Usa valor_total, que já vem somado do calculos.py. Somar apenas
    valor_servico deixaria os lançamentos de deslocamento de fora, já
    que neles esse campo é zero.
    """
    total = 0.0
    for p in os_item.get("procedimentos", []):
        if p.get("valor_total") is not None:
            total += float(p.get("valor_total") or 0)
        else:  # lançamento antigo, sem o campo
            total += (float(p.get("valor_servico", 0) or 0)
                      + float(p.get("valor_km", 0) or 0)
                      + float(p.get("valor_munck", 0) or 0))
    return total

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
#  NAVEGAÇÃO
# ─────────────────────────────────────────────

PAGES_TECNICO = ["📋 Minhas OS", "➕ Abrir Nova OS", "🔧 Adicionar Serviço"]

PAGES_GESTAO = ["📊 Dashboard", "✅ Validações", "📋 Todas as OS",
                "➕ Abrir Nova OS", "🔧 Adicionar Serviço",
                "💰 Comissões", "📤 Exportar Relatório", "⚙️ Configurações"]

# Rótulos curtos para o menu do topo — o valor continua o mesmo, só o
# texto exibido encolhe, para caber na largura de um celular.
LABELS_CURTOS = {
    "📋 Minhas OS":         "📋 Minhas OS",
    "📋 Todas as OS":       "📋 Todas",
    "➕ Abrir Nova OS":     "➕ Abrir OS",
    "🔧 Adicionar Serviço": "🔧 Lançar",
    "📤 Exportar Relatório": "📤 Exportar",
    "⚙️ Configurações":     "⚙️ Config.",
}

CSS_SEM_SIDEBAR = """
<style>
[data-testid="stSidebar"], [data-testid="collapsedControl"] { display: none !important; }
[data-testid="block-container"] { padding-top: 0.5rem; }
</style>
"""


def _seletor_topo(pages: list[str]) -> str:
    """
    Seletor horizontal do menu. Tenta os widgets mais novos primeiro e cai
    para o radio horizontal, que existe em qualquer versão do Streamlit.
    """
    atual = st.session_state.get("page")
    idx = pages.index(atual) if atual in pages else 0
    rotulo = lambda p: LABELS_CURTOS.get(p, p)

    if hasattr(st, "segmented_control"):
        escolha = st.segmented_control(
            "Navegação", pages, default=pages[idx],
            format_func=rotulo, label_visibility="collapsed", key="_nav_topo")
        return escolha or pages[idx]

    if hasattr(st, "pills"):
        escolha = st.pills(
            "Navegação", pages, default=pages[idx],
            format_func=rotulo, label_visibility="collapsed", key="_nav_topo")
        return escolha or pages[idx]

    return st.radio("Navegação", pages, index=idx, horizontal=True,
                    format_func=rotulo, label_visibility="collapsed",
                    key="_nav_topo")


def _sair():
    for k in list(st.session_state.keys()):
        del st.session_state[k]
    st.rerun()


def _nav_topo(ud) -> str:
    """Menu no topo — técnico e supervisor, pensado para o celular."""
    st.markdown(CSS_SEM_SIDEBAR, unsafe_allow_html=True)

    c_user, c_sair = st.columns([4, 1])
    with c_user:
        st.markdown(
            f'<div style="font-size:0.85rem;color:#94a3b8;padding-top:0.4rem;">'
            f'👤 <strong style="color:#e2e8f0;">{ud["nome"]}</strong>'
            f' &nbsp;·&nbsp; {ud["perfil"].upper()}</div>',
            unsafe_allow_html=True)
    with c_sair:
        if st.button("🚪 Sair", use_container_width=True):
            _sair()

    page = _seletor_topo(PAGES_TECNICO if ud["perfil"] == "tecnico" else PAGES_GESTAO)
    st.session_state["page"] = page
    st.markdown("---")
    return page


def _nav_sidebar(ud) -> str:
    """Menu lateral — admin, que trabalha no PC."""
    with st.sidebar:
        st.markdown(f"""
        <div style="padding:1rem 0;border-bottom:1px solid #3d4a5c;margin-bottom:1rem;">
            <div style="font-size:1.8rem;text-align:center;">👤</div>
            <div style="text-align:center;font-weight:700;font-size:0.95rem;">{ud['nome']}</div>
            <div style="text-align:center;font-size:0.75rem;color:#94a3b8;">{ud['perfil'].upper()}</div>
        </div>
        """, unsafe_allow_html=True)

        pages = PAGES_TECNICO if ud["perfil"] == "tecnico" else PAGES_GESTAO
        page = st.radio("Navegação", pages, label_visibility="collapsed")
        st.session_state["page"] = page
        st.markdown("---")
        if st.button("🚪 Sair", use_container_width=True):
            _sair()
    return page


def render_sidebar() -> str:
    """
    Ponto de entrada da navegação. Mantém o nome antigo porque o main()
    chama por ele.

    A escolha é por PERFIL, não por dispositivo: o Streamlit não informa
    ao servidor de forma confiável se é celular ou PC, e detectar por
    largura de tela quebra ao redimensionar. Como o admin é justamente
    quem só usa PC, a regra por perfil resolve sem gambiarra.
    """
    ud = st.session_state["user_data"]
    if ud["perfil"] == "admin":
        return _nav_sidebar(ud)
    return _nav_topo(ud)

# ─────────────────────────────────────────────
#  FORMULÁRIO DE LANÇAMENTO
#  (o técnico registra um SERVIÇO ou um DESLOCAMENTO numa OS)
# ─────────────────────────────────────────────

def _hhmm(t) -> str:
    return t.strftime("%H:%M") if hasattr(t, "strftime") else str(t)


def _pt(v, default):
    """Converte string 'HH:MM' de volta para time, para reabrir em edição."""
    if isinstance(v, time):
        return v
    if isinstance(v, str) and v:
        try:
            p = v.split(":")
            return time(int(p[0]), int(p[1]))
        except (ValueError, IndexError):
            return default
    return default


def _limpar_form():
    for k in list(st.session_state.keys()):
        if k.startswith("_pr_"):
            st.session_state.pop(k, None)
    st.session_state["_pr_reset"] = True


def _cabecalho_os(os_item, texto_seq):
    """
    Faixa de identificação da OS no topo do formulário.

    Montado numa linha só, sem indentação: com HTML indentado e uma linha
    condicional para a frota, a OS de deslocamento (que não tem frota)
    deixava uma linha vazia e o Markdown passava a tratar o resto como
    bloco de código, mostrando o "&nbsp;" cru na tela.
    """
    partes = [f"📋 <strong>{os_item.get('numero_os','—')}</strong>"]
    if os_item.get("frota"):
        partes.append(f"Frota <strong>{os_item['frota']}</strong>")
    partes.append(os_item.get("equipamento") or "—")
    partes.append(os_item.get("cliente") or "—")
    partes.append(texto_seq)
    corpo = " &nbsp;&middot;&nbsp; ".join(p for p in partes if p)
    st.markdown(f'<div class="nova-os-banner">{corpo}</div>',
                unsafe_allow_html=True)


def _form_servico(ud, os_item: dict, serv_existente: dict | None = None):
    """
    Renderiza o formulário de lançamento e devolve o dict salvo, ou None.

    A natureza (serviço ou deslocamento) é escolhida no topo. Numa OS
    aberta como deslocamento puro, só o deslocamento fica disponível —
    ela não tem frota nem local de serviço.
    """
    editando   = serv_existente is not None
    e          = serv_existente or {}
    eh_tecnico = ud["perfil"] == "tecnico"

    n_lancs   = len(os_item.get("procedimentos", []))
    seq_atual = e.get("seq", n_lancs + 1)
    _cabecalho_os(os_item, ("Editando #" + str(seq_atual)) if editando
                  else f"Novo lançamento (#{seq_atual})")

    # ── Natureza ──
    os_so_desloc = os_item.get("natureza_os") == "deslocamento"
    if os_so_desloc:
        natureza = "deslocamento"
        st.info("Esta OS foi aberta como **deslocamento** — só aceita lançamentos de deslocamento.")
    elif editando:
        natureza = e.get("natureza", "servico")
        st.caption(f"Natureza: **{LABELS_NATUREZA.get(natureza, natureza)}** "
                   "(não muda em edição — exclua e lance de novo se precisar trocar)")
    else:
        natureza = st.radio(
            "O que você vai lançar?", list(NATUREZAS), horizontal=True,
            format_func=lambda n: LABELS_NATUREZA.get(n, n),
            key="_pr_nat",
        )

    st.markdown("---")

    # ── Dados comuns: data e técnico ──
    c1, c2 = st.columns(2)
    with c1:
        data_serv = st.date_input(
            "Data", value=date.fromisoformat(e["data"]) if e.get("data") else date.today())
    with c2:
        if eh_tecnico and ud.get("cod_tecnico"):
            cod_tecnico = ud["cod_tecnico"]
            st.text_input("Técnico", value=TECNICOS[cod_tecnico]["nome"], disabled=True)
        else:
            lista_tec = {v["nome"]: k for k, v in TECNICOS.items()}
            nomes = list(lista_tec.keys())
            idx = nomes.index(e["nome_tecnico"]) if e.get("nome_tecnico") in nomes else 0
            cod_tecnico = lista_tec[st.selectbox("Técnico", nomes, index=idx)]

    tec_info = TECNICOS.get(cod_tecnico, {})
    nivel    = tec_info.get("nivel", "Técnico Um")
    tipo_km  = tec_info.get("tipo_km", "Km Um")

    if natureza == "deslocamento":
        return _form_deslocamento(ud, os_item, e, editando,
                                  data_serv, cod_tecnico, nivel, tipo_km)
    if natureza == "lavagem":
        return _form_lavagem(ud, os_item, e, editando,
                             data_serv, cod_tecnico, nivel, tipo_km)
    return _form_servico_normal(ud, os_item, e, editando,
                                data_serv, cod_tecnico, nivel, tipo_km)


# ─────────────────────────────────────────────
#  SERVIÇO
# ─────────────────────────────────────────────

def _form_servico_normal(ud, os_item, e, editando, data_serv, cod_tecnico, nivel, tipo_km):
    eh_tecnico = ud["perfil"] == "tecnico"

    c1, c2 = st.columns(2)
    with c1:
        idx_tp = TIPOS_SERVICO.index(e["tipo_servico"]) if e.get("tipo_servico") in TIPOS_SERVICO else 0
        tipo_servico = st.selectbox("Tipo de Serviço", TIPOS_SERVICO, index=idx_tp)
    with c2:
        locais = list(LOCAIS_OS)
        padrao = e.get("local_servico") or os_item.get("local_previsto") or locais[0]
        idx_lc = locais.index(padrao) if padrao in locais else 0
        local_servico = st.selectbox("Local do Serviço", locais, index=idx_lc,
                                      format_func=lambda l: LABELS_LOCAIS_OS.get(l, l))

    cc_txt = (f"{os_item['cod_cc']} - {os_item['cliente']}"
              if os_item.get("cod_cc") is not None
              else f"⚠️ pendente - {os_item.get('cliente','—')}")
    st.text_input("Centro de Custo", value=cc_txt, disabled=True)
    st.caption("Herdado da OS — altere na OS se necessário")

    st.markdown("---")

    # ── Tempo fixo por componente ──
    st.markdown("##### ⏱️ Tempo Fixo por Componente (opcional)")
    st.caption("Se o serviço é a troca/reparo de um componente da tabela, o tempo já "
               "vem pronto da planilha e os horários abaixo viram só registro.")
    tabela_tempos, por_equipamento = carregar_tabela_tempos()

    if not por_equipamento:
        st.warning("⚠️ TEMPO_SERVIÇO.xlsx não encontrada — tempo fixo indisponível.")
        equip_sel = comp_sel = NENHUM_COMPONENTE
    else:
        equipamentos_tabela = [NENHUM_COMPONENTE] + sorted(por_equipamento.keys())
        if st.session_state.get("_pr_equip_fixo") not in equipamentos_tabela:
            st.session_state["_pr_equip_fixo"] = (
                e.get("equipamento_tempo_fixo")
                if e.get("equipamento_tempo_fixo") in equipamentos_tabela
                else NENHUM_COMPONENTE)
        c_eq, c_comp = st.columns(2)
        equip_sel = c_eq.selectbox("Equipamento (tabela de tempos)",
                                    equipamentos_tabela, key="_pr_equip_fixo")
        comps = ([NENHUM_COMPONENTE] + por_equipamento.get(equip_sel, [])
                 if equip_sel != NENHUM_COMPONENTE else [NENHUM_COMPONENTE])
        if st.session_state.get("_pr_componente_fixo") not in comps:
            st.session_state["_pr_componente_fixo"] = (
                e.get("componente_tempo_fixo")
                if e.get("componente_tempo_fixo") in comps else NENHUM_COMPONENTE)
        comp_sel = c_comp.selectbox("Componente", comps, key="_pr_componente_fixo",
                                     disabled=(equip_sel == NENHUM_COMPONENTE))

    tempo_fixo = (equip_sel != NENHUM_COMPONENTE and comp_sel != NENHUM_COMPONENTE)
    horas_fixas = tabela_tempos.get((equip_sel, comp_sel), 0.0) if tempo_fixo else None
    if tempo_fixo:
        st.markdown(f"""
        <div class="tempo-fixo-box">
            ⏱️ <strong>Tempo fixo: {horas_fixas:.2f}h ({formatar_horas(horas_fixas)})</strong><br>
            Não há desconto de intervalo — o tempo da planilha já é o de execução.
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")

    # ── Horários, KM e Munck ──
    c4, c5 = st.columns(2)
    with c4:
        hora_saida = st.time_input("Hora Saída",
            value=_pt(e.get("hora_saida"), time(7, 30)), key="_pr_hora_saida")
        hora_chegada = st.time_input("Hora Chegada",
            value=_pt(e.get("hora_chegada"), time(17, 0)), key="_pr_hora_cheg")
    with c5:
        km_ini = st.number_input("KM Inicial", min_value=0.0, step=1.0, format="%.1f",
                                  value=float(e.get("km_inicial", 0)), key="_pr_km_ini")
        km_fim = st.number_input("KM Final", min_value=0.0, step=1.0, format="%.1f",
                                  value=float(e.get("km_final", 0)), key="_pr_km_fim")
        mun_ini = st.number_input("Hora Munck Inicial", min_value=0.0, step=0.1, format="%.2f",
                                   value=float(e.get("hora_munck_inicial", 0)), key="_pr_mun_ini")
        mun_fim = st.number_input("Hora Munck Final", min_value=0.0, step=0.1, format="%.2f",
                                   value=float(e.get("hora_munck_final", 0)), key="_pr_mun_fim")

    descricao = st.text_area("Descrição do Serviço Executado",
        value=e.get("descricao", ""), height=100,
        placeholder="Descreva detalhadamente o serviço realizado...")

    km_rodado   = max(0.0, km_fim - km_ini)
    horas_munck = max(0.0, mun_fim - mun_ini)

    calc = calcular_servico(
        local_servico=local_servico, nivel_tecnico=nivel,
        hora_saida=hora_saida, hora_chegada=hora_chegada,
        horas_trabalhadas=horas_fixas if tempo_fixo else None,
        km_rodado=km_rodado, tipo_km=tipo_km, horas_munck=horas_munck,
    )

    st.markdown("---")
    if calc["intervalo_maior_que_jornada"]:
        st.error("⚠️ O intervalo padrão (1h12 de almoço + 15min de café) é maior que a "
                 "jornada informada. Confira os horários.")
    elif calc["intervalo"] > 0:
        st.caption(f"ℹ️ Atravessa o meio-dia: descontados {calc['intervalo']:.2f}h de "
                   f"intervalo ({calc['horas_brutas']:.2f}h brutas → "
                   f"{calc['horas_trabalhadas']:.2f}h trabalhadas).")

    if eh_tecnico:
        st.markdown(f"""
        <div class="preview-box">
            <div class="title">📊 Resumo</div>
            <div class="preview-grid-2">
                <div class="preview-item">
                    <div class="lbl">{"Horas (tempo fixo)" if tempo_fixo else "Horas Trabalhadas"}</div>
                    <div class="val">{calc['horas_trabalhadas']:.2f}h</div>
                    <div class="sub">{f"bruto {calc['horas_brutas']:.2f}h" if calc['intervalo'] else ""}</div>
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
            <div class="title">📊 Preview</div>
            <div class="preview-grid">
                <div class="preview-item">
                    <div class="lbl">Horas</div>
                    <div class="val">{calc['horas_trabalhadas']:.2f}h</div>
                    <div class="sub">R$ {calc['valor_hora']:.0f}/h</div>
                </div>
                <div class="preview-item">
                    <div class="lbl">Valor Serviço</div>
                    <div class="val">R$ {calc['valor_servico']:,.2f}</div>
                    <div class="sub">base da comissão</div>
                </div>
                <div class="preview-item">
                    <div class="lbl">Total Cobrado</div>
                    <div class="val">R$ {calc['valor_total']:,.2f}</div>
                    <div class="sub">+ KM R$ {calc['valor_km']:,.2f}</div>
                </div>
                <div class="preview-item">
                    <div class="lbl">Comissão ({calc['percentual']:.0f}%)</div>
                    <div class="val" style="color:#4ade80;">R$ {calc['comissao']:,.2f}</div>
                    <div class="sub">{local_servico}</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        st.caption("A comissão incide só sobre o valor das horas — KM e Munck ficam fora da base.")

    if st.button("💾 Salvar Alterações" if editando else "💾 Salvar Serviço",
                 type="primary", use_container_width=True):
        if not descricao.strip():
            st.error("A descrição do serviço é obrigatória.")
            return None
        if calc["horas_trabalhadas"] <= 0:
            st.error("Tempo de serviço inválido. Confira os horários ou o componente.")
            return None

        # Carimbo do alerta de tempo: avalia o lançamento contra a média
        # histórica viva e guarda o resultado. "disparou_original" registra a
        # 1ª versão e NÃO é apagado por uma correção posterior — é o histórico
        # de alertas do mecânico. "disparou" reflete a versão atual (após ajuste
        # da gestão). Tempo fixo não é avaliado (não varia).
        _eh_gestao  = ud["perfil"] in ("admin", "supervisor")
        _av_carimbo = None if tempo_fixo else avaliar_tempo(
            tipo_servico, os_item.get("equipamento", ""), calc["horas_trabalhadas"])
        _ant = e.get("alerta_tempo") if isinstance(e.get("alerta_tempo"), dict) else None
        if _av_carimbo is None:
            _alerta_tempo = {"avaliavel": False}
        else:
            if editando and _ant and _ant.get("avaliavel"):
                _orig_disp  = _ant.get("disparou_original", _ant.get("disparou", False))
                _orig_horas = _ant.get("horas_originais", _ant.get("horas"))
            else:
                _orig_disp, _orig_horas = _av_carimbo["acima"], calc["horas_trabalhadas"]
            _corrig_agora = bool(editando and _eh_gestao and _orig_disp)
            _ja_corrig    = bool(_ant and _ant.get("corrigido"))
            _alerta_tempo = {
                "avaliavel":         True,
                "disparou":          _av_carimbo["acima"],
                "disparou_original": bool(_orig_disp),
                "horas_originais":   round(float(_orig_horas), 2) if _orig_horas is not None else None,
                "media":             _av_carimbo["media"],
                "limite":            _av_carimbo["limite"],
                "excedente_pct":     _av_carimbo["excedente_pct"],
                "corrigido":         _corrig_agora or _ja_corrig,
                "corrigido_por":     (st.session_state["usuario"] if _corrig_agora
                                      else (_ant.get("corrigido_por") if _ant else None)),
                "corrigido_em":      (datetime.now().isoformat() if _corrig_agora
                                      else (_ant.get("corrigido_em") if _ant else None)),
            }

        serv = {
            "natureza":            "servico",
            "data":                str(data_serv),
            "cod_tecnico":         cod_tecnico,
            "nome_tecnico":        TECNICOS[cod_tecnico]["nome"],
            "nivel_tecnico":       nivel,
            "tipo_servico":        tipo_servico,
            "local_servico":       local_servico,
            "hora_saida":          _hhmm(hora_saida),
            "hora_chegada":        _hhmm(hora_chegada),
            "horas_brutas":        calc["horas_brutas"],
            "intervalo":           calc["intervalo"],
            "horas_trabalhadas":   calc["horas_trabalhadas"],
            "equipamento_tempo_fixo": equip_sel if equip_sel != NENHUM_COMPONENTE else None,
            "componente_tempo_fixo":  comp_sel  if comp_sel  != NENHUM_COMPONENTE else None,
            "tempo_fixo_aplicado":    tempo_fixo,
            "alerta_tempo":        _alerta_tempo,
            "km_inicial":          km_ini,
            "km_final":            km_fim,
            "km_rodado":           round(km_rodado, 2),
            "hora_munck_inicial":  mun_ini,
            "hora_munck_final":    mun_fim,
            "horas_munck":         round(horas_munck, 4),
            "descricao":           descricao,
            "valor_hora":          calc["valor_hora"],
            "valor_servico":       round(calc["valor_servico"], 2),
            "valor_km":            round(calc["valor_km"], 2),
            "valor_munck":         round(calc["valor_munck"], 2),
            "valor_deslocamento":  0.0,
            "valor_total":         round(calc["valor_total"], 2),
            "base_comissao":       round(calc["base_comissao"], 2),
            "percentual_comissao": calc["percentual"],
            "comissao":            round(calc["comissao"], 2),
            "registrado_por":      st.session_state["usuario"],
            "registrado_em":       datetime.now().isoformat(),
            "editado_por":         None,
            "editado_em":          None,
        }
        return _persistir(os_item, e, serv, editando)
    return None


# ─────────────────────────────────────────────
#  DESLOCAMENTO
# ─────────────────────────────────────────────

def _form_deslocamento(ud, os_item, e, editando, data_serv, cod_tecnico, nivel, tipo_km):
    eh_tecnico = ud["perfil"] == "tecnico"

    st.markdown("##### 📍 Trajeto")
    c1, c2 = st.columns(2)
    with c1:
        cidade_origem = campo_cidade("Cidade de Origem", "_pr_cid_orig",
                                      e.get("cidade_origem", ""))
    with c2:
        cidade_destino = campo_cidade("Cidade de Destino", "_pr_cid_dest",
                                       e.get("cidade_destino", ""))

    st.markdown("---")
    st.markdown("##### 🕐 Horários")
    st.caption("Ida e retorno são lançados separados, como na planilha. "
               "Trajeto não desconta almoço.")

    c3, c4 = st.columns(2)
    with c3:
        st.markdown("**Ida**")
        ida_sai = st.time_input("Saída (ida)",
            value=_pt(e.get("ida_saida"), time(6, 0)), key="_pr_ida_sai")
        ida_cheg = st.time_input("Chegada (ida)",
            value=_pt(e.get("ida_chegada"), time(10, 0)), key="_pr_ida_cheg")
    with c4:
        st.markdown("**Retorno**")
        ret_sai = st.time_input("Saída (retorno)",
            value=_pt(e.get("retorno_saida"), time(16, 0)), key="_pr_ret_sai")
        ret_cheg = st.time_input("Chegada (retorno)",
            value=_pt(e.get("retorno_chegada"), time(20, 0)), key="_pr_ret_cheg")

    st.markdown("---")
    st.markdown("##### 🛣️ Quilometragem")
    c5, c6 = st.columns(2)
    with c5:
        km_ini = st.number_input("KM Inicial", min_value=0.0, step=1.0, format="%.1f",
                                  value=float(e.get("km_inicial", 0)), key="_pr_km_ini")
    with c6:
        km_fim = st.number_input("KM Final", min_value=0.0, step=1.0, format="%.1f",
                                  value=float(e.get("km_final", 0)), key="_pr_km_fim")

    descricao = st.text_area("Observação (opcional)", value=e.get("descricao", ""),
        height=80, placeholder="Motivo do deslocamento, ocorrências no trajeto...")

    km_rodado = max(0.0, km_fim - km_ini)
    calc = calcular_deslocamento(
        ida_saida=ida_sai, ida_chegada=ida_cheg,
        retorno_saida=ret_sai, retorno_chegada=ret_cheg,
        km_rodado=km_rodado, tipo_km=tipo_km,
    )

    st.markdown("---")

    # Velocidade média — conferência de digitação
    vm = calc["velocidade_media"]
    if calc["horas_trabalhadas"] > 0 and km_rodado > 0:
        if vm > 110:
            st.error(f"⚠️ Velocidade média de {vm:.0f} km/h. Confira os horários ou o KM.")
        elif vm < 20:
            st.warning(f"⚠️ Velocidade média de {vm:.0f} km/h — baixa para um trajeto. "
                       "Confira os horários ou o KM.")

    if eh_tecnico:
        st.markdown(f"""
        <div class="preview-box">
            <div class="title">📊 Resumo do Deslocamento</div>
            <div class="preview-grid">
                <div class="preview-item">
                    <div class="lbl">Ida</div><div class="val">{calc['horas_ida']:.2f}h</div>
                </div>
                <div class="preview-item">
                    <div class="lbl">Retorno</div><div class="val">{calc['horas_retorno']:.2f}h</div>
                </div>
                <div class="preview-item">
                    <div class="lbl">Total</div><div class="val">{calc['horas_trabalhadas']:.2f}h</div>
                </div>
                <div class="preview-item">
                    <div class="lbl">Velocidade Média</div>
                    <div class="val">{vm:.0f}</div><div class="sub">km/h</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown(f"""
        <div class="preview-box">
            <div class="title">📊 Preview do Deslocamento</div>
            <div class="preview-grid">
                <div class="preview-item">
                    <div class="lbl">Horas de Trajeto</div>
                    <div class="val">{calc['horas_trabalhadas']:.2f}h</div>
                    <div class="sub">ida {calc['horas_ida']:.2f}h + volta {calc['horas_retorno']:.2f}h</div>
                </div>
                <div class="preview-item">
                    <div class="lbl">Valor Deslocamento</div>
                    <div class="val">R$ {calc['valor_deslocamento']:,.2f}</div>
                    <div class="sub">R$ 60/h fixo</div>
                </div>
                <div class="preview-item">
                    <div class="lbl">Velocidade Média</div>
                    <div class="val">{vm:.0f}</div>
                    <div class="sub">{km_rodado:.0f} km</div>
                </div>
                <div class="preview-item">
                    <div class="lbl">Comissão (6%)</div>
                    <div class="val" style="color:#4ade80;">R$ {calc['comissao']:,.2f}</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    if st.button("💾 Salvar Alterações" if editando else "💾 Salvar Deslocamento",
                 type="primary", use_container_width=True):
        if not cidade_origem or not cidade_destino:
            st.error("Informe a cidade de origem e a de destino.")
            return None
        if calc["horas_trabalhadas"] <= 0:
            st.error("Horários inválidos — a chegada deve ser posterior à saída "
                     "em pelo menos um dos trajetos.")
            return None

        desl = {
            "natureza":            "deslocamento",
            "data":                str(data_serv),
            "cod_tecnico":         cod_tecnico,
            "nome_tecnico":        TECNICOS[cod_tecnico]["nome"],
            "nivel_tecnico":       nivel,
            "tipo_servico":        "Deslocamento",
            "local_servico":       "",
            "cidade_origem":       cidade_origem,
            "cidade_destino":      cidade_destino,
            "ida_saida":           _hhmm(ida_sai),
            "ida_chegada":         _hhmm(ida_cheg),
            "retorno_saida":       _hhmm(ret_sai),
            "retorno_chegada":     _hhmm(ret_cheg),
            "horas_ida":           calc["horas_ida"],
            "horas_retorno":       calc["horas_retorno"],
            "horas_trabalhadas":   calc["horas_trabalhadas"],
            "km_inicial":          km_ini,
            "km_final":            km_fim,
            "km_rodado":           round(km_rodado, 2),
            "velocidade_media":    vm,
            "descricao":           descricao,
            "valor_hora":          calc["valor_hora"],
            "valor_servico":       0.0,
            "valor_km":            round(calc["valor_km"], 2),
            "valor_munck":         0.0,
            "valor_deslocamento":  round(calc["valor_deslocamento"], 2),
            "valor_total":         round(calc["valor_total"], 2),
            "base_comissao":       round(calc["base_comissao"], 2),
            "percentual_comissao": calc["percentual"],
            "comissao":            round(calc["comissao"], 2),
            "registrado_por":      st.session_state["usuario"],
            "registrado_em":       datetime.now().isoformat(),
            "editado_por":         None,
            "editado_em":          None,
        }
        return _persistir(os_item, e, desl, editando)
    return None


# ─────────────────────────────────────────────
#  LAVAGEM
# ─────────────────────────────────────────────

def _form_lavagem(ud, os_item, e, editando, data_serv, cod_tecnico, nivel, tipo_km):
    eh_tecnico = ud["perfil"] == "tecnico"

    st.markdown("##### 🧼 Lavagem")
    TIPOS_VEIC = ["carro", "onibus", "outros"]
    LBL_VEIC = {"carro": "🚗 Carro (R$80)", "onibus": "🚌 Ônibus (R$250)",
                "outros": "🚜 Outros (por tempo)"}
    idx_v = TIPOS_VEIC.index(e.get("tipo_veiculo")) if e.get("tipo_veiculo") in TIPOS_VEIC else 0
    tipo_veiculo = st.radio("Tipo de veículo", TIPOS_VEIC, index=idx_v, horizontal=True,
                            format_func=lambda t: LBL_VEIC.get(t, t), key="_pr_lav_veic")

    locais = list(LOCAIS_OS)
    padrao = e.get("local_servico") or os_item.get("local_previsto") or "interno barracão"
    idx_lc = locais.index(padrao) if padrao in locais else 0
    local_servico = st.selectbox("Local", locais, index=idx_lc,
                                 format_func=lambda l: LABELS_LOCAIS_OS.get(l, l),
                                 key="_pr_lav_local")

    eh_outros = tipo_veiculo == "outros"
    hora_saida = hora_chegada = None
    if eh_outros:
        st.caption("Outros é medido por horário (horas × valor-hora do nível), com "
                   "desconto de almoço/café se atravessar o meio-dia.")
        c4, c5 = st.columns(2)
        with c4:
            hora_saida = st.time_input("Hora Saída",
                value=_pt(e.get("hora_saida"), time(7, 30)), key="_pr_lav_saida")
        with c5:
            hora_chegada = st.time_input("Hora Chegada",
                value=_pt(e.get("hora_chegada"), time(17, 0)), key="_pr_lav_cheg")

    descricao = st.text_area("Descrição / Observação", value=e.get("descricao", ""),
        height=80, placeholder="Placa, frota, o que foi lavado...")

    calc = calcular_lavagem(
        tipo_veiculo=tipo_veiculo, local_servico=local_servico,
        nivel_tecnico=nivel, hora_saida=hora_saida, hora_chegada=hora_chegada,
    )
    _veic_nome = {"carro": "Carro", "onibus": "Ônibus"}.get(tipo_veiculo, "Outros")

    st.markdown("---")
    if eh_outros and calc["intervalo_maior_que_jornada"]:
        st.error("⚠️ O intervalo padrão (1h12 + 15min) é maior que a jornada. Confira os horários.")
    elif eh_outros and calc["intervalo"] > 0:
        st.caption(f"ℹ️ Atravessa o meio-dia: descontados {calc['intervalo']:.2f}h de intervalo.")

    if eh_tecnico:
        st.markdown(f"""
        <div class="preview-box">
            <div class="title">📊 Resumo da Lavagem</div>
            <div class="preview-grid-2">
                <div class="preview-item">
                    <div class="lbl">{"Horas" if eh_outros else "Veículo"}</div>
                    <div class="val">{(f"{calc['horas_trabalhadas']:.2f}h") if eh_outros else _veic_nome}</div>
                </div>
                <div class="preview-item">
                    <div class="lbl">Valor</div>
                    <div class="val">R$ {calc['valor_servico']:,.2f}</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown(f"""
        <div class="preview-box">
            <div class="title">📊 Preview da Lavagem</div>
            <div class="preview-grid">
                <div class="preview-item">
                    <div class="lbl">Veículo</div>
                    <div class="val" style="font-size:1.05rem;">{_veic_nome}</div>
                    <div class="sub">{(f"{calc['horas_trabalhadas']:.2f}h × R$ {calc['valor_hora']:.0f}") if eh_outros else "valor fixo"}</div>
                </div>
                <div class="preview-item">
                    <div class="lbl">Valor Serviço</div>
                    <div class="val">R$ {calc['valor_servico']:,.2f}</div>
                    <div class="sub">base da comissão</div>
                </div>
                <div class="preview-item">
                    <div class="lbl">Total Cobrado</div>
                    <div class="val">R$ {calc['valor_total']:,.2f}</div>
                </div>
                <div class="preview-item">
                    <div class="lbl">Comissão ({calc['percentual']:.0f}%)</div>
                    <div class="val" style="color:#4ade80;">R$ {calc['comissao']:,.2f}</div>
                    <div class="sub">{local_servico}</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    if st.button("💾 Salvar Alterações" if editando else "💾 Salvar Lavagem",
                 type="primary", use_container_width=True):
        if not descricao.strip():
            st.error("Informe uma descrição (placa/frota/o que foi lavado).")
            return None
        if eh_outros and calc["horas_trabalhadas"] <= 0:
            st.error("Horário inválido. Confira a hora de saída e chegada.")
            return None

        lav = {
            "natureza":            "lavagem",
            "tipo_veiculo":        tipo_veiculo,
            "data":                str(data_serv),
            "cod_tecnico":         cod_tecnico,
            "nome_tecnico":        TECNICOS[cod_tecnico]["nome"],
            "nivel_tecnico":       nivel,
            "tipo_servico":        "Lavagem",
            "local_servico":       local_servico,
            "hora_saida":          _hhmm(hora_saida) if eh_outros else "",
            "hora_chegada":        _hhmm(hora_chegada) if eh_outros else "",
            "horas_brutas":        calc["horas_brutas"],
            "intervalo":           calc["intervalo"],
            "horas_trabalhadas":   calc["horas_trabalhadas"],
            "km_rodado":           0.0,
            "horas_munck":         0.0,
            "descricao":           descricao,
            "valor_hora":          calc["valor_hora"],
            "valor_servico":       round(calc["valor_servico"], 2),
            "valor_km":            0.0,
            "valor_munck":         0.0,
            "valor_deslocamento":  0.0,
            "valor_total":         round(calc["valor_total"], 2),
            "base_comissao":       round(calc["base_comissao"], 2),
            "percentual_comissao": calc["percentual"],
            "comissao":            round(calc["comissao"], 2),
            "registrado_por":      st.session_state["usuario"],
            "registrado_em":       datetime.now().isoformat(),
            "editado_por":         None,
            "editado_em":          None,
        }
        return _persistir(os_item, e, lav, editando)
    return None


def _persistir(os_item, e, registro, editando):
    """Grava o lançamento (novo ou editado) e limpa o formulário."""
    if editando:
        registro["editado_por"]    = st.session_state["usuario"]
        registro["editado_em"]     = datetime.now().isoformat()
        registro["registrado_por"] = e.get("registrado_por", st.session_state["usuario"])
        registro["registrado_em"]  = e.get("registrado_em", datetime.now().isoformat())
        ok = editar_servico(os_item["numero_os"], e["proc_id"], registro)
        if not ok:
            st.error("Não foi possível editar. A OS pode já ter sido enviada para aprovação.")
            return None
    else:
        ok = adicionar_servico(os_item["numero_os"], registro)
        if not ok:
            st.error("Não foi possível adicionar. Verifique o status da OS.")
            return None

    _limpar_form()
    return registro

# ─────────────────────────────────────────────
#  COMPONENTE: timeline de lançamentos de uma OS
# ─────────────────────────────────────────────

def _render_servicos(os_item: dict, ud: dict, pode_editar: bool = False):
    lancs = os_item.get("procedimentos", [])
    if not lancs:
        st.info("Nenhum lançamento registrado ainda.")
        return

    eh_tecnico = ud["perfil"] == "tecnico"

    for p in lancs:
        eh_desloc  = p.get("natureza") == "deslocamento"
        eh_lavagem = p.get("natureza") == "lavagem"

        if eh_desloc:
            titulo  = "🚚 Deslocamento"
            sub     = (f"{p.get('cidade_origem','?')} → {p.get('cidade_destino','?')}"
                       f" &nbsp;·&nbsp; {p.get('data','?')}")
            detalhe = (
                f"<strong>Ida:</strong> {p.get('ida_saida','?')} → {p.get('ida_chegada','?')}"
                f" ({float(p.get('horas_ida',0)):.2f}h)"
                f" &nbsp;|&nbsp; <strong>Volta:</strong> {p.get('retorno_saida','?')} → "
                f"{p.get('retorno_chegada','?')} ({float(p.get('horas_retorno',0)):.2f}h)"
                f"<br><strong>KM:</strong> {float(p.get('km_rodado',0)):.1f} km"
                f" &nbsp;|&nbsp; <strong>Vel. média:</strong> "
                f"{float(p.get('velocidade_media',0)):.0f} km/h"
            )
            borda = "#f59e0b"
        elif eh_lavagem:
            _vn = {"carro": "Carro", "onibus": "Ônibus"}.get(
                (p.get("tipo_veiculo") or "").lower(), "Outros")
            titulo  = f"🧼 {p.get('nome_tecnico','?')}"
            sub     = (f"Lavagem &nbsp;·&nbsp; {_vn} &nbsp;·&nbsp; "
                       f"{p.get('local_servico','?')} &nbsp;·&nbsp; {p.get('data','?')}")
            if (p.get("tipo_veiculo") or "").lower() in ("carro", "onibus"):
                detalhe = (f"<strong>Veículo:</strong> {_vn}"
                           f" &nbsp;|&nbsp; <strong>Valor fixo:</strong> "
                           f"R$ {float(p.get('valor_servico',0)):,.2f}")
            else:
                detalhe = (f"<strong>Veículo:</strong> Outros"
                           f" &nbsp;|&nbsp; <strong>Horário:</strong> "
                           f"{p.get('hora_saida','?')} → {p.get('hora_chegada','?')}"
                           f" &nbsp;|&nbsp; <strong>Horas:</strong> "
                           f"{float(p.get('horas_trabalhadas',0)):.2f}h")
            borda = "#22d3ee"
        else:
            tf = (f" &nbsp;·&nbsp; ⏱️ Tempo fixo: <strong>{p.get('componente_tempo_fixo','?')}</strong>"
                  if p.get("tempo_fixo_aplicado") else "")
            titulo  = f"🔧 {p.get('nome_tecnico','?')}"
            sub     = (f"{p.get('tipo_servico','?')} &nbsp;·&nbsp; "
                       f"{p.get('local_servico','?')} &nbsp;·&nbsp; {p.get('data','?')}{tf}")
            intervalo = float(p.get("intervalo", 0) or 0)
            det_int = (f" (bruto {float(p.get('horas_brutas',0)):.2f}h &minus; "
                       f"{intervalo:.2f}h de intervalo)" if intervalo else "")
            detalhe = (
                f"<strong>Horário:</strong> {p.get('hora_saida','?')} → {p.get('hora_chegada','?')}"
                f" &nbsp;|&nbsp; <strong>Horas:</strong> "
                f"{float(p.get('horas_trabalhadas',0)):.2f}h{det_int}"
                f" &nbsp;|&nbsp; <strong>KM:</strong> {float(p.get('km_rodado',0)):.1f} km"
            )
            borda = "#3b82f6"

        c_left, c_right = st.columns([3, 1])
        with c_left:
            st.markdown(f"""
            <div class="serv-card" style="border-left-color:{borda};">
                <div class="serv-seq">Lançamento #{p.get('seq','?')}</div>
                <div class="serv-tec">{titulo} <span style="color:#64748b;font-size:0.78rem;">— {p.get('nivel_tecnico','?')}</span></div>
                <div class="serv-tipo">{sub}</div>
                <div style="margin-top:0.5rem;font-size:0.82rem;color:#cbd5e1;">{detalhe}</div>
                <div style="margin-top:0.4rem;font-size:0.82rem;color:#94a3b8;">{p.get('descricao','-') or '-'}</div>
                {"<div style='margin-top:0.3rem;font-size:0.72rem;color:#4a5568;'>✏️ Editado por " + str(p.get('editado_por','')) + " em " + str(p.get('editado_em',''))[:16].replace('T',' ') + "</div>" if p.get('editado_em') else ""}
            </div>
            """, unsafe_allow_html=True)
            # Alerta de tempo acima da média histórica — só adm/supervisor,
            # só serviço medido por horário (tempo fixo não varia). Usa o carimbo
            # gravado no lançamento; se faltar (lançamento antigo), avalia ao vivo.
            if (not eh_tecnico and not eh_desloc and not eh_lavagem
                    and not p.get("tempo_fixo_aplicado")):
                _at = p.get("alerta_tempo") if isinstance(p.get("alerta_tempo"), dict) else None
                if _at is None or not _at.get("avaliavel", False):
                    _live = avaliar_tempo(p.get("tipo_servico", ""),
                                          os_item.get("equipamento", ""),
                                          float(p.get("horas_trabalhadas", 0) or 0))
                    _at = ({"avaliavel": True, "disparou": _live["acima"],
                            "disparou_original": _live["acima"],
                            "horas_originais": _live["horas"], "media": _live["media"],
                            "limite": _live["limite"], "excedente_pct": _live["excedente_pct"],
                            "corrigido": False} if _live else {"avaliavel": False})
                _horas_atual = float(p.get("horas_trabalhadas", 0) or 0)
                if _at.get("avaliavel") and _at.get("disparou"):
                    # Ainda acima da média (valor atual estourou o limite).
                    st.markdown(f"""
                    <div style="margin-top:-0.4rem;margin-bottom:0.8rem;padding:0.6rem 0.9rem;
                                border-radius:8px;background:#3a1d1d;border:1px solid #b91c1c;
                                border-left:4px solid #ef4444;font-size:0.82rem;color:#fecaca;">
                        ⚠️ <strong>Tempo acima da média histórica.</strong>
                        {_horas_atual:.2f}h neste lançamento &nbsp;·&nbsp;
                        média {_at["media"]:.2f}h &nbsp;·&nbsp;
                        limite {_at["limite"]:.2f}h (média + 1 desvio)
                        <br><span style="color:#f87171;font-size:0.75rem;">
                        Confirme com o mecânico. Se o tempo não se justificar, ajuste o
                        lançamento no botão Editar — use o tempo real informado, não um
                        número só para zerar o alerta.
                        </span>
                    </div>
                    """, unsafe_allow_html=True)
                elif _at.get("avaliavel") and _at.get("disparou_original") and _at.get("corrigido"):
                    # Estava acima e foi ajustado pela gestão para dentro da faixa.
                    _ho = _at.get("horas_originais")
                    _ho_txt = f"{_ho:.2f}h" if isinstance(_ho, (int, float)) else "—"
                    st.markdown(f"""
                    <div style="margin-top:-0.4rem;margin-bottom:0.8rem;padding:0.5rem 0.9rem;
                                border-radius:8px;background:#2a2410;border:1px solid #a16207;
                                border-left:4px solid #ca8a04;font-size:0.78rem;color:#fde68a;">
                        ✏️ Estava acima da média ({_ho_txt}) e foi ajustado por
                        {_at.get("corrigido_por","gestão")} para {_horas_atual:.2f}h
                        (dentro da faixa, limite {_at["limite"]:.2f}h).
                    </div>
                    """, unsafe_allow_html=True)
        with c_right:
            if eh_tecnico and eh_lavagem:
                _tv_nome = {"carro": "Carro", "onibus": "Ônibus"}.get(
                    (p.get("tipo_veiculo") or "").lower(), "Outros")
                st.markdown(f"""
                <div class="comissao-box" style="background:linear-gradient(135deg,#0c1a2e,#1e2738);border-color:#3d4a5c;">
                    <div class="label" style="color:#94a3b8!important;">Lavagem</div>
                    <div class="valor" style="color:#60a5fa!important;font-size:1.2rem;">R$ {float(p.get('valor_servico',0)):,.2f}</div>
                    <div style="margin-top:0.5rem;"><span class="info-pill">{_tv_nome}</span></div>
                </div>
                """, unsafe_allow_html=True)
            elif eh_tecnico:
                st.markdown(f"""
                <div class="comissao-box" style="background:linear-gradient(135deg,#0c1a2e,#1e2738);border-color:#3d4a5c;">
                    <div class="label" style="color:#94a3b8!important;">Horas</div>
                    <div class="valor" style="color:#60a5fa!important;font-size:1.3rem;">{float(p.get('horas_trabalhadas',0)):.2f}h</div>
                    <div style="margin-top:0.5rem;"><span class="info-pill">KM {float(p.get('km_rodado',0)):.1f}</span></div>
                </div>
                """, unsafe_allow_html=True)
            else:
                base = float(p.get("base_comissao", p.get("valor_servico", 0)) or 0)
                st.markdown(f"""
                <div class="comissao-box">
                    <div class="label">{"Lavagem" if eh_lavagem else ("Deslocamento" if eh_desloc else "Serviço")}</div>
                    <div class="valor" style="font-size:1.2rem;">R$ {base:,.2f}</div>
                    <div style="margin-top:0.4rem;"><span class="info-pill">KM R$ {float(p.get('valor_km',0)):,.2f}</span></div>
                    <div style="margin-top:0.5rem;padding-top:0.5rem;border-top:1px solid #166534;">
                        <div class="label">Comissão ({p.get('percentual_comissao','?')}%)</div>
                        <div class="valor">R$ {float(p.get('comissao',0)):,.2f}</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

            if pode_editar:
                eh_meu = p.get("registrado_por") == st.session_state["usuario"]
                if (not eh_tecnico) or eh_meu:
                    bc1, bc2 = st.columns(2) if ud["perfil"] == "admin" else (st.container(), None)
                    with bc1:
                        if st.button("✏️ Editar", key=f"edit_serv_{p.get('proc_id','')}"):
                            st.session_state["_editando_serv"] = {
                                "numero_os": os_item["numero_os"], "serv": p}
                            _limpar_form()
                            st.rerun()
                    if ud["perfil"] == "admin":
                        ck = f"_confirm_excl_{p.get('proc_id','')}"
                        with bc2:
                            if st.session_state.get(ck):
                                if st.button("✅ Confirmar", key=f"conf_excl_{p.get('proc_id','')}"):
                                    ok = excluir_servico(os_item["numero_os"],
                                                          p.get("proc_id", ""),
                                                          st.session_state["usuario"])
                                    st.session_state.pop(ck, None)
                                    st.success("Lançamento excluído!") if ok else st.error("Não foi possível excluir.")
                                    st.rerun()
                            else:
                                if st.button("🗑️ Excluir", key=f"del_serv_{p.get('proc_id','')}"):
                                    st.session_state[ck] = True
                                    st.rerun()
                        if st.session_state.get(ck):
                            st.caption("⚠️ Confirma a exclusão? Não pode ser desfeito.")
                            if st.button("✖️ Cancelar", key=f"canc_excl_{p.get('proc_id','')}"):
                                st.session_state.pop(ck, None)
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

    # Locais que este usuário pode abrir. O tipo (interna/externa) é
    # derivado do local — a regra vive no dados.py, não aqui. Passa o login
    # porque a OS interna tem exceção nominal (TECNICOS_OS_INTERNA).
    usuario_login = st.session_state["usuario"]
    locais_permitidos = [l for l in LOCAIS_OS
                         if pode_abrir_os(perfil, tipo_os_do_local(l),
                                          "manual", usuario_login)]
    if not locais_permitidos:
        st.error("Seu perfil não pode abrir OS.")
        return

    if perfil == "tecnico" and "interno barracão" not in locais_permitidos:
        st.info("Você pode abrir OS de **campo**, **M.S** e de **deslocamento**. "
                "OS interna (barracão) é aberta pelo técnico autorizado ou pelo "
                "supervisor, ou automaticamente pelo app de Checklist de Veículos.")

    FROTAS = carregar_frotas()

    # Sem st.form de propósito: dentro de um form o Streamlit não re-executa
    # enquanto se digita, e o Equipamento nunca preencheria a partir da frota.
    with st.container():
        natureza_os = st.radio(
            "Tipo de OS", ["servico", "deslocamento"], horizontal=True,
            format_func=lambda n: {"servico": "🔧 Serviço (com frota)",
                                    "deslocamento": "🚚 Deslocamento (sem frota)"}.get(n, n),
            key="_os_nat",
        )
        eh_desloc_os = natureza_os == "deslocamento"

        if eh_desloc_os:
            st.caption("OS de deslocamento não tem frota nem local de serviço — só "
                       "centro de custo. Os trajetos são lançados depois.")
            local_previsto = ""
        else:
            local_previsto = st.radio(
                "Local do Serviço", locais_permitidos, horizontal=True,
                format_func=lambda l: LABELS_LOCAIS_OS.get(l, l), key="_os_local")

        c1, c2 = st.columns(2)
        with c1:
            data_ab = st.date_input("Data de Abertura", value=date.today())
            frota_val = ("" if eh_desloc_os else
                         st.text_input("Frota", placeholder="Ex: 1201", key="_os_frota"))

        frota = (frota_val or "").strip()
        info_frota = FROTAS.get(frota, {})
        desc_frota = info_frota.get("descricao", "")

        # O centro de custo da frota vem do cadastro. É o campo que mais
        # errava sendo escolhido na mão — daí a OS nascer com CC pendente.
        lista_cc = {f"{k} - {v}": k for k, v in CENTROS_CUSTO.items()}
        rotulos_cc = list(lista_cc.keys())
        idx_cc = 0
        cc_da_frota = info_frota.get("cc_nome", "")
        cod_cc_sugerido, _ = resolver_cod_cc(cc_da_frota) if cc_da_frota else (None, "")
        if cod_cc_sugerido is not None:
            alvo = f"{cod_cc_sugerido} - {CENTROS_CUSTO[cod_cc_sugerido]}"
            if alvo in rotulos_cc:
                idx_cc = rotulos_cc.index(alvo)

        with c2:
            cod_cc = lista_cc[st.selectbox("Centro de Custo / Cliente",
                                            rotulos_cc, index=idx_cc)]
            if cc_da_frota and cod_cc_sugerido is not None:
                st.caption(f"↩️ Sugerido pelo cadastro da frota: {cc_da_frota}")
            elif cc_da_frota:
                st.caption(f"⚠️ Cadastro da frota diz “{cc_da_frota}”, que não existe "
                           "na tabela de centros de custo. Escolha manualmente.")

        if eh_desloc_os:
            equipamento = "Deslocamento"
        else:
            # Só sobrescreve o Equipamento quando a frota MUDA e existe no
            # cadastro, para não apagar o que o usuário digitou à mão.
            st.session_state.setdefault("_os_equip", "")
            if desc_frota and st.session_state.get("_os_frota_ant") != frota:
                st.session_state["_os_equip"]     = desc_frota
                st.session_state["_os_frota_ant"] = frota
            equipamento = st.text_input(
                "Equipamento / Descrição", key="_os_equip",
                placeholder="Preenchido automaticamente ao digitar a frota")
            if frota and not desc_frota:
                st.caption(f"⚠️ Frota {frota} não encontrada no cadastro — "
                           "descreva o equipamento manualmente.")
            aviso = alerta_status(info_frota.get("status", ""))
            if aviso:
                st.warning(aviso)

        avaliacao = st.text_area(
            "Avaliação",
            placeholder=("Descreva o problema apresentado pelo veículo — ex: para-brisa "
                         "trincado, motor com fumaça excessiva, vazamento hidráulico."),
            height=110,
            help="Obrigatório. É o que orienta quem vai executar o serviço.")

        if perfil == "tecnico" and ud.get("cod_tecnico"):
            cod_tecnico_sel = ud["cod_tecnico"]
            nome_tec_sel    = TECNICOS[cod_tecnico_sel]["nome"]
            st.text_input("Técnico Responsável", value=nome_tec_sel, disabled=True)
        else:
            lista_tec_nomes = {v["nome"]: k for k, v in TECNICOS.items()}
            nome_tec_sel = st.selectbox("Técnico Responsável", list(lista_tec_nomes.keys()),
                help="Apenas este técnico verá esta OS para lançar e encerrá-la.")
            cod_tecnico_sel = lista_tec_nomes[nome_tec_sel]

        if st.button("🚀 Abrir OS", type="primary", use_container_width=True):
            if not eh_desloc_os and not frota:
                st.error("Informe a frota.")
            elif not avaliacao.strip():
                st.error("Preencha a Avaliação — descreva o motivo da OS.")
            else:
                try:
                    nova = criar_os(
                        frota=frota,
                        equipamento=equipamento or desc_frota or frota or "Deslocamento",
                        cod_cc=cod_cc,
                        aberto_por=st.session_state["usuario"],
                        data_abertura=str(data_ab),
                        tecnico_designado=cod_tecnico_sel,
                        perfil_usuario=perfil,
                        local_previsto=local_previsto,
                        natureza_os=natureza_os,
                        avaliacao=avaliacao,
                        origem="manual",
                    )
                except PermissaoNegada as e:
                    st.error(f"🚫 {e}")
                    return

                st.success(f"✅ OS **{nova['numero_os']}** criada com sucesso!")
                st.balloons()
                rotulo = ("🚚 Deslocamento" if eh_desloc_os
                          else LABELS_LOCAIS_OS.get(local_previsto, local_previsto))
                st.markdown(f"""
                <div style="background:#052e16;border:2px solid #16a34a;border-radius:12px;
                     padding:1.5rem;text-align:center;margin-top:1rem;">
                    <div style="font-size:0.8rem;color:#86efac;font-weight:600;text-transform:uppercase;
                         letter-spacing:0.08em;">Número da OS</div>
                    <div style="font-size:2.5rem;font-weight:900;color:#4ade80;letter-spacing:0.1em;">
                        {nova['numero_os']}
                    </div>
                    <div style="font-size:0.85rem;color:#6ee7b7;margin-top:0.25rem;">
                        {("Frota " + frota + " · ") if frota else ""}{CENTROS_CUSTO[cod_cc]} · {rotulo}
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
            rotulo_lanc = ("Deslocamento" if res.get("natureza") == "deslocamento"
                            else "Serviço")
            st.success(f"✅ {rotulo_lanc} adicionado à OS {numero_os_sel}!")

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

    # Edição de serviço disparada na própria tela de aprovação (ajuste de
    # tempo quando o alerta não se justifica): delega o formulário de edição.
    if st.session_state.get("_editando_serv"):
        pagina_adicionar_servico()
        return

    ud_val    = st.session_state["user_data"]
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

**Centro de Custo:** {os_item.get('cod_cc') if os_item.get('cod_cc') is not None else '⚠️ pendente'} — {os_item.get('cliente','?')}

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
            # Gestão pode ajustar o tempo aqui mesmo (adm e supervisor). A OS
            # permanece na fila de aprovação após o ajuste.
            _pode_edit_val = ud_val.get("perfil") in ("admin", "supervisor")
            _render_servicos(os_item, ud_val, pode_editar=_pode_edit_val)

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

    # ── Tempo acima da média histórica: ranking por técnico ──
    st.markdown("---")
    st.markdown("**⚠️ Tempo acima da média histórica — por técnico (OS aprovadas)**")
    st.caption("Serviços medidos por horário cujo tempo passou de média + 1 desvio-padrão. "
               "Tempo fixo e deslocamento não entram. Base: histórico da planilha + aprovados do app.")
    _rk = ranking_acima_media(todas)
    if not _rk or all(d["acima"] == 0 for d in _rk):
        st.info("Nenhum serviço acima da média até agora (ou ainda sem histórico comparável).")
    else:
        _rk = [d for d in _rk if d["avaliados"] > 0][:10]
        _df_rk = pd.DataFrame([{
            "Técnico":        d["tecnico"],
            "Acima da média": d["acima"],
            "Corrigidos":     d["corrigidos"],
            "Avaliados":      d["avaliados"],
            "% acima":        f'{d["pct"]:.0f}%',
        } for d in _rk])
        _lider = _rk[0]
        st.markdown(
            f"<div class='metric-card yellow' style='margin-bottom:0.8rem;'>"
            f"<div class='metric-label'>Mais lançamentos acima da média</div>"
            f"<div class='metric-value'>{_lider['tecnico']}</div>"
            f"<div class='metric-sub'>{_lider['acima']} de {_lider['avaliados']} serviços "
            f"({_lider['pct']:.0f}%)</div></div>",
            unsafe_allow_html=True)
        st.dataframe(_df_rk, use_container_width=True, hide_index=True)

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
                "Local":       ("deslocamento" if p.get("natureza") == "deslocamento"
                                else p.get("local_servico","?")),
                "Cliente":     o.get("cliente","?"),
                "Horas":       float(p.get("horas_trabalhadas",0)),
                "KM":          float(p.get("km_rodado",0)),
                "Vlr Serviço": float(p.get("base_comissao",
                                            p.get("valor_servico", 0)) or 0),
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
                "Cod CC":         o.get("cod_cc") or "",
                "Cliente":        o.get("cliente",""),
                "Seq Servico":    p.get("seq",""),
                "Data Servico":   p.get("data",""),
                "Tecnico":        p.get("nome_tecnico",""),
                "Nivel":          p.get("nivel_tecnico",""),
                "Natureza":       p.get("natureza", "servico"),
                "Tipo Servico":   p.get("tipo_servico",""),
                "Local":          ("deslocamento" if p.get("natureza") == "deslocamento"
                                    else p.get("local_servico","")),
                "Cidade Origem":  p.get("cidade_origem",""),
                "Cidade Destino": p.get("cidade_destino",""),
                "Vlr Deslocamento": float(p.get("valor_deslocamento",0) or 0),
                "Velocidade Media": float(p.get("velocidade_media",0) or 0),
                "Horas Brutas":   float(p.get("horas_brutas",0) or 0),
                "Intervalo":      float(p.get("intervalo",0) or 0),
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