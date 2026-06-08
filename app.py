import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

from utils.data_loader import load_xlsx, detect_columns
from utils.charts import (
    render_kpis,
    chart_sla_donut,
    chart_status_donut,
    chart_atendentes,
    chart_tempo_por_categoria,
    chart_top_sistemas,
    abrir_chamados_do_sistema,
    chart_finalizados_periodo,
    periodos_finalizados,
    FONT_FAMILY,
)

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Dashboard de Chamados TI",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

/* Fonte padrão de dashboards de tecnologia (Inter) em toda a interface */
html, body, [class*="css"], [class*="st-"], .stApp,
button, input, textarea, select {
    font-family: 'Inter', 'Segoe UI', Roboto, system-ui, -apple-system, sans-serif !important;
}

/* Restaura a fonte de ÍCONES (Material Symbols) — senão aparece o texto
   literal do ícone, ex: 'keyboard_double_arrow_left' no menu lateral */
[data-testid="stIconMaterial"],
span[class*="material-symbols"], span[class*="material-icons"],
.material-icons, .material-symbols-rounded, .material-symbols-outlined {
    font-family: 'Material Symbols Rounded', 'Material Symbols Outlined',
                 'Material Icons' !important;
}

/* Números das métricas com peso forte e tabular (alinhamento de dígitos) */
[data-testid="stMetricValue"] {
    font-weight: 700;
    font-feature-settings: "tnum" 1;
    letter-spacing: -0.02em;
}
[data-testid="stMetricLabel"] { font-weight: 500; }
h1, h2, h3, h4 { font-weight: 700; letter-spacing: -0.01em; }

/* KPIs como cards clicáveis (botões com chave kpicard_* abrem a janela de detalhes) */
[class*="st-key-kpicard"] button {
    background: #1A2332 !important;
    border: 1px solid #243044 !important;
    border-left: 4px solid #3498DB !important;
    border-radius: 10px !important;
    padding: 14px 18px !important;
    width: 100% !important;
    min-height: 96px !important;
    text-align: left !important;
    color: #ECF0F1 !important;
    line-height: 1.25 !important;
    align-items: flex-start !important;
    transition: all .15s ease;
}
/* Rótulo (título) do card */
[class*="st-key-kpicard"] button p {
    font-size: 0.82rem !important;
    color: #9FB3C8 !important;
    font-weight: 600 !important;
    margin: 0 !important;
}
/* O valor (negrito **...** dentro do label) vira o número grande */
[class*="st-key-kpicard"] button strong {
    font-size: 1.7rem !important;
    font-weight: 800 !important;
    color: #ECF0F1 !important;
    letter-spacing: -0.02em;
    font-feature-settings: "tnum" 1;
}
[class*="st-key-kpicard"] button:hover {
    border-left-color: #5DADE2 !important;
    background: #202c3f !important;
    transform: translateY(-1px);
}

[data-testid="stSidebar"] { background-color: #0D1B2A; }
[data-testid="stSidebar"] * { color: #ECF0F1 !important; }

[data-testid="metric-container"] {
    background: #1A2332;
    border-radius: 10px;
    padding: 14px 18px;
    border-left: 4px solid #3498DB;
}

div[data-testid="stFileUploader"] {
    border: 2px dashed #3498DB;
    border-radius: 10px;
    padding: 6px 10px;
}

.stTabs [data-baseweb="tab"] {
    background: #1A2332;
    border-radius: 8px 8px 0 0;
    color: #BDC3C7;
    padding: 8px 16px;
}
.stTabs [aria-selected="true"] {
    background: #3498DB !important;
    color: #fff !important;
}

h1, h2, h3 { color: #ECF0F1; }
</style>
""", unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────────────────────
if 'files_data' not in st.session_state:
    st.session_state.files_data: dict = {}   # {label: {df, col_map}}
if 'page' not in st.session_state:
    st.session_state.page = 'import'
if 'raw_files' not in st.session_state:
    st.session_state.raw_files: dict = {}    # {i: {bytes, name, label}}

MAX_FILES = 6
DEFAULT_LABELS = [
    "[TI] Acessos Novos Colaboradores",
    "Infraestrutura",
    "Sistemas Internos",
    "Redes",
    "[TI] Sistemas",
    "Outros",
]


def go_to(page: str):
    st.session_state.page = page
    st.rerun()


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🔧 TI Dashboard")
    st.markdown("---")

    if st.button("📥  Importar Arquivos", use_container_width=True):
        go_to('import')

    loaded = st.session_state.files_data
    if loaded:
        st.markdown("### Dashboards")
        for name in loaded:
            label = f"📊  {name}"
            if st.button(label, key=f"nav_{name}", use_container_width=True):
                go_to(name)

        if len(loaded) > 1:
            st.markdown("---")
            if st.button("🌐  Visão Geral", use_container_width=True):
                go_to('overview')

    st.markdown("---")
    n = len(loaded)
    total_r = sum(len(d['df']) for d in loaded.values())
    st.caption(f"Arquivos carregados: **{n}/{MAX_FILES}**")
    if total_r:
        st.caption(f"Total de registros: **{total_r:,}**")


# ═══════════════════════════════════════════════════════════════════════════════
# IMPORT PAGE
# ═══════════════════════════════════════════════════════════════════════════════

def render_import():
    st.title("📥 Importação de Dados")
    st.markdown(
        "Faça o upload de até **6 arquivos XLSX** de chamados de TI. "
        "Nomeie cada arquivo conforme o tipo de solicitação."
    )

    # Sample data button
    with st.expander("🧪 Sem arquivos? Gere dados de exemplo para teste", expanded=False):
        st.markdown(
            "Clique no botão para gerar 6 arquivos de exemplo com dados realistas de chamados de TI."
        )
        if st.button("⚡ Carregar Dados de Exemplo", type="primary"):
            with st.spinner("Gerando dados de exemplo..."):
                from utils.sample_data import generate_all
                samples = generate_all()
                for i, (name, raw) in enumerate(samples.items()):
                    st.session_state.raw_files[i] = {
                        'bytes': raw, 'name': f'{name}.xlsx', 'label': name
                    }
                st.success("✅ Dados de exemplo carregados! Clique em **Analisar Dados** abaixo.")
                st.rerun()

    st.markdown("---")

    # 6 upload slots (2 rows × 3 cols)
    row1 = st.columns(3)
    row2 = st.columns(3)
    slots = row1 + row2

    for i, col in enumerate(slots):
        with col:
            raw = st.session_state.raw_files.get(i, {})
            current_label = raw.get('label', DEFAULT_LABELS[i])

            st.markdown(f"**Arquivo {i + 1}**")
            label = st.text_input(
                "Nome", value=current_label, key=f"label_{i}",
                label_visibility="collapsed", placeholder=DEFAULT_LABELS[i]
            )

            uploaded = st.file_uploader(
                "Escolher arquivo", type=["xlsx", "xls", "csv", "txt"],
                key=f"upload_{i}", label_visibility="collapsed"
            )

            if uploaded is not None:
                file_bytes = uploaded.read()
                st.session_state.raw_files[i] = {
                    'bytes': file_bytes,
                    'name': uploaded.name,
                    'label': label or DEFAULT_LABELS[i],
                }
                try:
                    preview = load_xlsx(file_bytes, uploaded.name)
                    st.success(f"✓ {uploaded.name}  ({len(preview.columns)} colunas, "
                               f"{len(preview)} linhas)")
                except Exception as e:
                    st.error(f"Erro ao ler o arquivo: {e}")
            elif raw:
                st.info(f"📄 {raw['name']}")

    st.markdown("---")

    # Update labels in raw_files from current text inputs
    for i in range(MAX_FILES):
        key = f"label_{i}"
        if key in st.session_state and i in st.session_state.raw_files:
            st.session_state.raw_files[i]['label'] = st.session_state[key] or DEFAULT_LABELS[i]

    valid = {i: r for i, r in st.session_state.raw_files.items() if r.get('bytes')}

    if valid:
        st.info(f"✅ **{len(valid)} arquivo(s)** prontos para análise.")
        c1, c2, _ = st.columns([1, 1, 3])

        with c1:
            if st.button("🔍 Analisar Dados", type="primary", use_container_width=True):
                with st.spinner("Processando arquivos..."):
                    for idx, raw in valid.items():
                        try:
                            df = load_xlsx(raw['bytes'], raw['name'])
                            col_map = detect_columns(df)
                            name = raw['label']
                            st.session_state.files_data[name] = {'df': df, 'col_map': col_map}
                        except Exception as e:
                            st.error(f"Erro em '{raw['name']}': {e}")

                if st.session_state.files_data:
                    first = next(iter(st.session_state.files_data))
                    go_to(first)

        with c2:
            if st.session_state.files_data:
                if st.button("🗑 Limpar Tudo", use_container_width=True):
                    st.session_state.files_data = {}
                    st.session_state.raw_files = {}
                    go_to('import')
    else:
        st.warning("⬆️ Faça o upload de pelo menos 1 arquivo XLSX para continuar.")


# ═══════════════════════════════════════════════════════════════════════════════
# INDIVIDUAL DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════════

def _grid(figs, cols=2):
    """Render figures in a grid, skipping any that are None."""
    figs = [f for f in figs if f is not None]
    for i in range(0, len(figs), cols):
        row = figs[i:i + cols]
        columns = st.columns(cols)
        for col, fig in zip(columns, row):
            col.plotly_chart(fig, use_container_width=True)


def render_dashboard(name: str):
    data = st.session_state.files_data[name]
    df_full: pd.DataFrame = data['df']
    col_map: dict = data['col_map']

    # Janela de 6 meses aplicada às métricas e gráficos (KPIs, SLA, status)
    df = _filtrar_6m(df_full, col_map)

    # ── Header ──
    st.title(f"📊 {name}")
    st.caption(f"📅 Dados referentes aos últimos 6 meses · {len(df):,} chamados")

    # Meta de SLA (dias) configurável e lembrada por arquivo
    c_meta, _ = st.columns([1, 3])
    meta = int(c_meta.number_input(
        "🎯 Meta de SLA (dias na fase Em Andamento)",
        min_value=1, max_value=180, value=7, step=1, key=f"meta_{name}",
    ))

    st.markdown("---")

    # ── KPIs (resumo executivo, sempre visível) ──
    render_kpis(df, col_map, meta)

    st.markdown("---")

    # ── Dashboards separados por seções (abas) ──
    # A aba de Sistemas só aparece quando o arquivo tem uma coluna de sistema
    tem_sistema = 'sistema' in col_map
    labels = ["📊 Fases & Status", "⏱ SLA & Tempo"]
    if tem_sistema:
        labels.append("🖥️ Sistemas Mais Solicitados")
    labels.append("👥 Equipe")
    labels.append("📈 Finalizados por Período")
    abas = dict(zip(labels, st.tabs(labels)))

    # Seção: distribuição por fase + conformidade de SLA
    with abas["📊 Fases & Status"]:
        _grid([chart_status_donut(df, col_map), chart_sla_donut(df, col_map)])

    # Seção: tempo médio de atendimento por departamento (vs meta de dias)
    with abas["⏱ SLA & Tempo"]:
        fig_tempo = chart_tempo_por_categoria(df, col_map, meta)
        if fig_tempo:
            st.plotly_chart(fig_tempo, use_container_width=True,
                            config={'staticPlot': True, 'displayModeBar': False})
        else:
            st.info("Sem dados de tempo de atendimento para exibir.")

    # Seção: sistemas mais solicitados (só quando há coluna de sistema)
    if tem_sistema:
        with abas["🖥️ Sistemas Mais Solicitados"]:
            fig_sis = chart_top_sistemas(df, col_map)
            if fig_sis:
                st.caption("💡 Clique em uma barra para ver os chamados do sistema.")
                ev = st.plotly_chart(
                    fig_sis, use_container_width=True,
                    key=f"sis_chart_{name}",
                    on_select="rerun", selection_mode="points",
                    config={'displayModeBar': False},
                )
                # Clique na barra → abre modal com os chamados daquele sistema
                sel_key = f"sis_sel_{name}"
                pts = []
                if ev and getattr(ev, "selection", None):
                    pts = ev.selection.get("points", [])
                if pts:
                    sistema_sel = pts[0].get("y")  # barras horizontais: categoria no eixo Y
                    if sistema_sel and st.session_state.get(sel_key) != sistema_sel:
                        st.session_state[sel_key] = sistema_sel
                        abrir_chamados_do_sistema(sistema_sel, df, col_map)
                else:
                    st.session_state[sel_key] = None
            else:
                st.info("Sem dados de sistemas para exibir.")

    # Seção: carga da equipe (Top 5 atendentes — finalizados nos últimos 6 meses)
    with abas["👥 Equipe"]:
        fig_top = chart_atendentes(df, col_map)
        if fig_top:
            st.plotly_chart(fig_top, use_container_width=True)
        else:
            st.info("Sem dados de atendentes para exibir.")

    # Seção: finalizados por responsável (pizza) com filtro de período
    with abas["📈 Finalizados por Período"]:
        c1, c2 = st.columns([1, 2])
        gran = c1.selectbox("Agrupar por:",
                            ["Últimos 6 meses", "Mensal", "Semanal", "Diário"],
                            key=f"granf_{name}")
        inicio = fim = None
        escopo = "últimos 6 meses"
        if gran != "Últimos 6 meses":
            freq = {"Mensal": "M", "Semanal": "W", "Diário": "D"}[gran]
            ops = periodos_finalizados(df, col_map, freq)
            if ops:
                labels = [o[0] for o in ops]
                sel = c2.selectbox("Período:", labels, key=f"perf_{name}")
                escopo, inicio, fim = next(o for o in ops if o[0] == sel)
        fig_per = chart_finalizados_periodo(df, col_map, inicio, fim, escopo)
        if fig_per:
            st.plotly_chart(fig_per, use_container_width=True)
        else:
            st.info("Sem chamados finalizados no período selecionado.")


# ═══════════════════════════════════════════════════════════════════════════════
# OVERVIEW PAGE
# ═══════════════════════════════════════════════════════════════════════════════

def _filtrar_6m(df, col_map):
    """Filtra o DataFrame para chamados dos últimos 6 meses (por data de criação)."""
    date_col = col_map.get('data_abertura') or col_map.get('data_fechamento')
    if not date_col:
        return df
    dt = pd.to_datetime(df[date_col], errors='coerce')
    cutoff = pd.Timestamp.now() - pd.DateOffset(months=6)
    return df[dt >= cutoff]


def render_overview():
    st.title("🌐 Visão Geral — Todos os Tipos de Chamado")

    files = st.session_state.files_data
    # Aplica a janela de 6 meses uma única vez e reutiliza em todos os gráficos
    recent = {n: (_filtrar_6m(d['df'], d['col_map']), d['col_map'])
              for n, d in files.items()}
    total = sum(len(dfx) for dfx, _ in recent.values())

    c1, c2, c3 = st.columns(3)
    c1.metric("📁 Tipos Carregados", len(files))
    c2.metric("📋 Total de Chamados", f"{total:,}")
    c3.metric("📅 Período", "Últimos 6 meses")
    st.caption("ℹ️ A Visão Geral considera apenas chamados criados nos últimos 6 meses.")

    st.markdown("---")

    # Volume por tipo de chamado (comparativo objetivo entre os arquivos)
    cnt_df = pd.DataFrame(
        [(n, len(dfx)) for n, (dfx, _) in recent.items()],
        columns=['Tipo', 'Chamados']
    )

    fig_bar = px.bar(
        cnt_df.sort_values('Chamados', ascending=False),
        x='Tipo', y='Chamados',
        title='Volume por Tipo de Chamado',
        color='Chamados', color_continuous_scale='Blues', text='Chamados',
    )
    fig_bar.update_traces(texttemplate='%{text:,}', textposition='outside')
    fig_bar.update_layout(
        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#ECF0F1', family=FONT_FAMILY), showlegend=False, coloraxis_showscale=False,
        margin=dict(t=50, b=30),
    )
    st.plotly_chart(fig_bar, use_container_width=True)

    # Aggregate SLA across files
    sla_rows = []
    for name, (df, col_map) in recent.items():
        from utils.data_loader import sla_series
        sla_all = sla_series(df, col_map)
        if sla_all is not None:
            sla = sla_all.dropna()
            if len(sla):
                sla_rows.append({'Tipo': name, '% SLA': round(sla.sum() / len(sla) * 100, 1)})

    if sla_rows:
        st.markdown("---")
        sla_df = pd.DataFrame(sla_rows).sort_values('% SLA')
        fig_sla = px.bar(
            sla_df, x='% SLA', y='Tipo', orientation='h',
            title='% SLA Cumprido por Tipo de Chamado',
            color='% SLA',
            color_continuous_scale=['#E74C3C', '#F39C12', '#2ECC71'],
            range_color=[0, 100], text='% SLA',
        )
        fig_sla.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
        fig_sla.add_vline(x=80, line_dash='dash', line_color='#F39C12',
                          annotation_text=' Meta 80%', annotation_font_color='#F39C12')
        fig_sla.update_layout(
            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
            font=dict(color='#ECF0F1', family=FONT_FAMILY), coloraxis_showscale=False,
            yaxis=dict(categoryorder='total ascending'),
            margin=dict(t=50, b=30, r=60),
        )
        st.plotly_chart(fig_sla, use_container_width=True)

    # Aggregate status
    status_rows = []
    for name, (df, col_map) in recent.items():
        if 'status' in col_map:
            from utils.data_loader import normalize_status
            status = normalize_status(df[col_map['status']])
            atend = (status == 'Atendido').sum()
            status_rows.append({
                'Tipo': name,
                'Atendido': int(atend),
                'Não Atendido': int(len(status) - atend),
            })

    if status_rows:
        st.markdown("---")
        st.subheader("Status por Tipo de Chamado")
        status_df = pd.DataFrame(status_rows)
        status_melt = status_df.melt('Tipo', var_name='Status', value_name='Chamados')

        fig_st = px.bar(
            status_melt, x='Tipo', y='Chamados', color='Status',
            barmode='group', title='Atendidos vs Não Atendidos por Tipo',
            color_discrete_map={'Atendido': '#3498DB', 'Não Atendido': '#E74C3C'},
            text='Chamados',
        )
        fig_st.update_traces(texttemplate='%{text:,}', textposition='outside')
        fig_st.update_layout(
            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
            font=dict(color='#ECF0F1', family=FONT_FAMILY), margin=dict(t=50, b=30),
        )
        st.plotly_chart(fig_st, use_container_width=True)


# ═══════════════════════════════════════════════════════════════════════════════
# ROUTER
# ═══════════════════════════════════════════════════════════════════════════════

page = st.session_state.page

if page == 'import':
    render_import()
elif page == 'overview' and st.session_state.files_data:
    render_overview()
elif page in st.session_state.files_data:
    render_dashboard(page)
else:
    render_import()
