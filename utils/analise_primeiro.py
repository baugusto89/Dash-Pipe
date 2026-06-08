"""Análises detalhadas para o primeiro arquivo de chamados importado."""

from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from utils.data_loader import _normalize, normalize_sla, normalize_vencido, sla_series

# ── Layout padrão Plotly (tema escuro) ────────────────────────────────────────
FONT_FAMILY = "Inter, 'Segoe UI', Roboto, system-ui, sans-serif"
_L = dict(
    plot_bgcolor='rgba(0,0,0,0)',
    paper_bgcolor='rgba(0,0,0,0)',
    font=dict(color='#ECF0F1', family=FONT_FAMILY, size=13),
    title_font=dict(size=15, family=FONT_FAMILY, color='#ECF0F1'),
    margin=dict(t=50, b=30, l=10, r=10),
)

# ── Palavras-chave de status ───────────────────────────────────────────────────
_APROVACAO = ['aprovacao', 'aprovação', 'aguardando aprovacao', 'aguardando aprovação',
              'em aprovacao', 'em aprovação', 'aprovado', 'pendente aprovacao']
_EM_ANDAMENTO = ['em andamento', 'andamento', 'in progress', 'em execucao',
                 'em execução', 'execucao', 'atendimento']
_FINALIZADO = ['finalizado', 'fechado', 'concluido', 'concluído', 'resolvido',
               'resolved', 'closed', 'done', 'encerrado', 'solucionado']


def _match(val: str, keywords: list) -> bool:
    v = _normalize(str(val))
    return any(kw in v or v in kw for kw in keywords)


def _filter(df: pd.DataFrame, col: str, keywords: list) -> pd.DataFrame:
    return df[df[col].apply(lambda x: _match(x, keywords))]


def _timeline_fig(df_slice: pd.DataFrame, date_col: str, value_col: str,
                  title: str, color: str = '#3498DB', chart='bar'):
    """Return a bar or line figure grouped by auto-detected period."""
    tmp = df_slice[[date_col]].copy()
    tmp['_dt'] = pd.to_datetime(tmp[date_col], dayfirst=True, errors='coerce')
    tmp = tmp.dropna(subset=['_dt'])
    if len(tmp) < 2:
        return None
    span = (tmp['_dt'].max() - tmp['_dt'].min()).days
    if span <= 62:
        tmp['_p'] = tmp['_dt'].dt.strftime('%d/%m/%Y')
        plbl = 'Dia'
    elif span <= 730:
        tmp['_p'] = tmp['_dt'].dt.strftime('%m/%Y')
        plbl = 'Mês'
    else:
        tmp['_p'] = tmp['_dt'].dt.strftime('%Y')
        plbl = 'Ano'
    tl = tmp.groupby('_p').size().reset_index(name=value_col)
    tl.columns = [plbl, value_col]
    if chart == 'line':
        fig = px.line(tl, x=plbl, y=value_col, title=title, markers=True, text=value_col)
        fig.update_traces(line_color=color, marker_color=color, textposition='top center')
    else:
        fig = px.bar(tl, x=plbl, y=value_col, title=title,
                     color=value_col, color_continuous_scale='Blues', text=value_col)
        fig.update_traces(texttemplate='%{text:,}', textposition='outside')
        fig.update_layout(coloraxis_showscale=False)
    fig.update_layout(**_L)
    return fig


def _show_missing_status(df: pd.DataFrame, col: str) -> None:
    with st.expander("🔍 Valores encontrados na coluna de status"):
        vc = df[col].value_counts().reset_index()
        vc.columns = ['Status', 'Qtd']
        st.dataframe(vc, use_container_width=True, height=200)


# ══════════════════════════════════════════════════════════════════════════════
# SEÇÃO 1 — Chamados em Aprovação de Solicitação
# ══════════════════════════════════════════════════════════════════════════════

def render_aprovacao(df: pd.DataFrame, col_map: dict) -> None:
    st.subheader("1️⃣  Chamados em Aprovação de Solicitação")

    if 'status' not in col_map:
        st.warning("Coluna de **status** não detectada no arquivo.")
        return

    df_aprov = _filter(df, col_map['status'], _APROVACAO)
    n = len(df_aprov)
    total = len(df)

    if n == 0:
        st.info("Nenhum chamado com status 'Aprovação de Solicitação' encontrado.")
        _show_missing_status(df, col_map['status'])
        return

    # ── KPIs ──
    c1, c2, c3 = st.columns(3)
    c1.metric("📬 Em Aprovação", f"{n:,}", f"{n / total * 100:.1f}% do total")

    if 'categoria' in col_map:
        c2.metric("Categorias distintas", df_aprov[col_map['categoria']].nunique())
    if 'tipo' in col_map:
        c3.metric("Tipos distintos", df_aprov[col_map['tipo']].nunique())

    # ── Gráficos ──
    charts = []

    if 'categoria' in col_map:
        cnt = df_aprov[col_map['categoria']].value_counts().head(12).reset_index()
        cnt.columns = ['Categoria', 'Chamados']
        fig = px.bar(cnt, x='Chamados', y='Categoria', orientation='h',
                     title='Em Aprovação — por Categoria',
                     color='Chamados', color_continuous_scale='Oranges', text='Chamados')
        fig.update_traces(texttemplate='%{text:,}', textposition='outside')
        fig.update_layout(coloraxis_showscale=False, showlegend=False,
                          yaxis=dict(categoryorder='total ascending'), **_L)
        charts.append(fig)

    if 'tipo' in col_map:
        cnt = df_aprov[col_map['tipo']].value_counts().reset_index()
        cnt.columns = ['Tipo', 'Chamados']
        fig = px.bar(cnt, x='Tipo', y='Chamados', title='Em Aprovação — por Tipo',
                     color='Tipo', text='Chamados',
                     color_discrete_sequence=['#E67E22', '#F39C12', '#D35400', '#E74C3C'])
        fig.update_traces(texttemplate='%{text:,}', textposition='outside')
        fig.update_layout(showlegend=False, **_L)
        charts.append(fig)

    for i in range(0, len(charts), 2):
        pair = charts[i: i + 2]
        cols = st.columns(len(pair))
        for col_st, fig in zip(cols, pair):
            col_st.plotly_chart(fig, use_container_width=True)

    # Abertura por período (se tiver data)
    date_col = col_map.get('data_abertura')
    if date_col:
        fig_t = _timeline_fig(df_aprov, date_col, 'Chamados',
                              'Em Aprovação — Evolução por Período', '#E67E22')
        if fig_t:
            st.plotly_chart(fig_t, use_container_width=True)

    with st.expander(f"📄 Ver {n} chamados em Aprovação de Solicitação", expanded=False):
        st.dataframe(df_aprov, use_container_width=True, height=300)


# ══════════════════════════════════════════════════════════════════════════════
# SEÇÃO 2 — Chamados Em Andamento
# ══════════════════════════════════════════════════════════════════════════════

def render_em_andamento(df: pd.DataFrame, col_map: dict) -> None:
    st.subheader("2️⃣  Chamados Em Andamento")

    if 'status' not in col_map:
        st.warning("Coluna de **status** não detectada no arquivo.")
        return

    df_and = _filter(df, col_map['status'], _EM_ANDAMENTO)
    n = len(df_and)

    if n == 0:
        st.info("Nenhum chamado com status 'Em Andamento' encontrado.")
        _show_missing_status(df, col_map['status'])
        return

    # ── Com / Sem Responsável ──
    atend_col = col_map.get('atendente')
    com_resp = sem_resp = 0
    df_com = df_sem = pd.DataFrame()

    if atend_col:
        mask_com = (
            df_and[atend_col].notna()
            & (df_and[atend_col].astype(str).str.strip() != '')
            & (~df_and[atend_col].astype(str).str.lower().isin(['nan', 'none', '-', '']))
        )
        df_com = df_and[mask_com]
        df_sem = df_and[~mask_com]
        com_resp = len(df_com)
        sem_resp = len(df_sem)

    # ── Vencimento de SLA ──
    vcto_col = col_map.get('sla_vencimento')
    dentro_sla = fora_sla = sem_data = 0
    tem_vcto = False

    if 'vencido' in col_map:
        # Coluna Vencido (TRUE/FALSE) é a fonte mais confiável
        sla_norm = normalize_vencido(df_and[col_map['vencido']])
        dentro_sla = int((sla_norm == True).sum())
        fora_sla = int((sla_norm == False).sum())
        sem_data = int(sla_norm.isna().sum())
        tem_vcto = (dentro_sla + fora_sla) > 0
    elif vcto_col:
        now = datetime.now()
        vcto_series = pd.to_datetime(df_and[vcto_col], dayfirst=True, errors='coerce')
        dentro_sla = int((vcto_series >= now).sum())
        fora_sla = int((vcto_series < now).sum())
        sem_data = int(vcto_series.isna().sum())
        tem_vcto = (dentro_sla + fora_sla) > 0
    elif 'sla' in col_map:
        sla_norm = normalize_sla(df_and[col_map['sla']]).dropna()
        dentro_sla = int(sla_norm.sum())
        fora_sla = int((~sla_norm).sum())
        tem_vcto = True

    # ── KPIs ──
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("⚙️ Em Andamento", f"{n:,}")
    if atend_col:
        c2.metric("👤 Com Responsável", f"{com_resp:,}", f"{com_resp/n*100:.0f}%")
        c3.metric("❓ Sem Responsável", f"{sem_resp:,}",
                  f"-{sem_resp/n*100:.0f}%", delta_color="inverse")
    if tem_vcto:
        c4.metric("🔴 Fora do SLA", f"{fora_sla:,}",
                  f"{fora_sla/n*100:.0f}%" if n else "—",
                  delta_color="inverse" if fora_sla > 0 else "normal")

    st.markdown("---")

    # ── Donuts: responsável + SLA ──
    donuts = []

    if atend_col:
        fig_resp = go.Figure(go.Pie(
            labels=['Com Responsável', 'Sem Responsável'],
            values=[com_resp, sem_resp],
            hole=0.62,
            marker_colors=['#3498DB', '#E74C3C'],
            textinfo='label+percent',
        ))
        fig_resp.update_layout(
            title='Com vs Sem Responsável',
            annotations=[dict(text=f'{n:,}', x=0.5, y=0.5,
                              font_size=22, showarrow=False, font_color='#ECF0F1')],
            legend=dict(orientation='h', y=-0.08), **_L,
        )
        donuts.append(fig_resp)

    if tem_vcto:
        lbl_sla = ['Dentro do Prazo', 'Fora do Prazo']
        val_sla = [dentro_sla, fora_sla]
        clr_sla = ['#2ECC71', '#E74C3C']
        if vcto_col and sem_data:
            lbl_sla.append('Sem Data')
            val_sla.append(sem_data)
            clr_sla.append('#95A5A6')
        pct_fora = fora_sla / (dentro_sla + fora_sla) * 100 if (dentro_sla + fora_sla) else 0
        title_sla = ('Vencimento de SLA' if vcto_col else 'Status de SLA')
        fig_sla = go.Figure(go.Pie(
            labels=lbl_sla, values=val_sla, hole=0.62,
            marker_colors=clr_sla, textinfo='label+percent',
        ))
        fig_sla.update_layout(
            title=title_sla,
            annotations=[dict(text=f'{pct_fora:.0f}%\nfora', x=0.5, y=0.5,
                              font_size=18, showarrow=False, font_color='#E74C3C')],
            legend=dict(orientation='h', y=-0.08), **_L,
        )
        donuts.append(fig_sla)

    if donuts:
        cols_d = st.columns(len(donuts))
        for col_st, fig in zip(cols_d, donuts):
            col_st.plotly_chart(fig, use_container_width=True)

    # ── Expanders com tabelas ──
    if atend_col and sem_resp > 0:
        with st.expander(f"⚠️ {sem_resp} chamados SEM responsável", expanded=False):
            st.dataframe(df_sem, use_container_width=True, height=280)
            csv = df_sem.to_csv(index=False).encode('utf-8-sig')
            st.download_button("⬇️ Baixar CSV", csv, "sem_responsavel.csv", "text/csv",
                               key="dl_sem_resp")

    with st.expander(f"📄 Todos os {n} chamados Em Andamento", expanded=False):
        st.dataframe(df_and, use_container_width=True, height=300)


# ══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT — chamado por app.py
# ══════════════════════════════════════════════════════════════════════════════

def render_analise_detalhada(df: pd.DataFrame, col_map: dict, nome_arquivo: str) -> None:
    st.markdown("---")
    st.markdown(f"## 📌 Análise Detalhada — {nome_arquivo}")
    st.caption(
        "Análise específica de chamados por etapa do fluxo: "
        "Aprovação de Solicitação · Em Andamento"
    )
    st.markdown("---")

    render_aprovacao(df, col_map)
    st.markdown("---")
    render_em_andamento(df, col_map)
