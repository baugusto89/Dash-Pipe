import json
import re
from html import escape as _esc
from io import BytesIO

import pandas as pd
import streamlit.components.v1 as components
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from utils.data_loader import (
    normalize_sla, normalize_status, sla_series,
    classify_fase, FASE_CORES, sla_iniciado_mask,
    tempo_atendimento, tempo_dias, SLA_META_DIAS, fmt_data,
    sistema_series, abreviar_nome,
)

# ── Palette ──────────────────────────────────────────────────────────────────
SEQ = ['#3498DB', '#2ECC71', '#E74C3C', '#F39C12', '#9B59B6',
       '#1ABC9C', '#E67E22', '#34495E', '#E91E63', '#00BCD4']
SLA_PALETTE = {'Cumprido': '#2ECC71', 'Violado': '#E74C3C'}
STATUS_PALETTE = {'Atendido': '#3498DB', 'Não Atendido': '#E74C3C'}

# Fonte padrão de dashboards de tecnologia (Inter, com fallbacks de sistema)
FONT_FAMILY = "Inter, 'Segoe UI', Roboto, system-ui, sans-serif"

_LAYOUT = dict(
    plot_bgcolor='rgba(0,0,0,0)',
    paper_bgcolor='rgba(0,0,0,0)',
    font=dict(color='#ECF0F1', family=FONT_FAMILY, size=13),
    title_font=dict(size=15, family=FONT_FAMILY, color='#ECF0F1'),
    margin=dict(t=50, b=30, l=10, r=10),
)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _apply(fig: go.Figure) -> go.Figure:
    fig.update_layout(**_LAYOUT)
    return fig


# ── KPIs ──────────────────────────────────────────────────────────────────────

# ── Tabela de detalhes + exportação Excel ─────────────────────────────────────

def _nome_do_email(email) -> str:
    """Deriva o nome do gestor a partir do e-mail (ex.: eduardo.monesi@nsx.bet → Eduardo Monesi)."""
    e = str(email).strip()
    if not e or '@' not in e:
        return ''
    user = e.split('@')[0]
    return user.replace('.', ' ').replace('_', ' ').replace('-', ' ').strip().title()


def _tabela_detalhes(df_sub: pd.DataFrame, col_map: dict) -> pd.DataFrame:
    """Monta a tabela com as informações importantes para a janela de detalhes."""
    # (campo, título, é_data?) — só entram as colunas que existem no arquivo
    campos = [
        ('solicitante', 'Para quem é esta Solicitação', False),
        ('solicitacao', 'Solicitação', False),
        ('data_inicio', 'Data de Início', True),
        ('data_credenciais', 'Data Envio Credenciais', True),
        ('cargo', 'Cargo', False),
        ('categoria', 'Departamento', False),
    ]
    dados = {}
    for field, titulo, eh_data in campos:
        c = col_map.get(field)
        if c and c in df_sub.columns:
            dados[titulo] = (fmt_data(df_sub[c]) if eh_data else df_sub[c]).values

    # Sistema (combinado: usa o texto livre quando 'Outro')
    sis = sistema_series(df_sub, col_map)
    if sis is not None:
        dados['Sistema'] = sis.values

    tab = pd.DataFrame(dados, index=df_sub.index)

    eg = col_map.get('email_gestor')
    if eg and eg in df_sub.columns:
        tab['Nome do Gestor'] = df_sub[eg].apply(_nome_do_email).values

    return tab.reset_index(drop=True)


def _to_excel(dfx: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w:
        dfx.to_excel(w, index=False, sheet_name='Chamados')
    return buf.getvalue()


@st.dialog("📋 Detalhes dos Chamados", width="large")
def _dialog_detalhes(titulo: str, df_sub: pd.DataFrame, col_map: dict) -> None:
    st.markdown(f"### {titulo}")
    st.caption(f"{len(df_sub):,} chamado(s)")
    tab = _tabela_detalhes(df_sub, col_map)
    if tab.empty or tab.shape[1] == 0:
        st.info("Não há colunas de detalhe disponíveis nesta planilha.")
        return
    st.dataframe(tab, use_container_width=True, hide_index=True, height=420)
    st.download_button(
        "⬇️ Baixar Excel (.xlsx)", data=_to_excel(tab),
        file_name=f"{titulo}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


def _card(col, key: str, titulo: str, valor: str,
          df_sub: pd.DataFrame, col_map: dict, titulo_janela: str = None) -> None:
    """Card clicável: ao clicar, abre uma janela (modal) com a tabela de detalhes."""
    if col.button(f"{titulo}\n\n**{valor}**", key=key, use_container_width=True):
        _dialog_detalhes(titulo_janela or titulo, df_sub, col_map)


def render_kpis(df: pd.DataFrame, col_map: dict, meta: int = SLA_META_DIAS) -> None:
    total = len(df)
    fases = classify_fase(df[col_map['status']]) if 'status' in col_map else None
    sla_all = sla_series(df, col_map)
    tempo = tempo_atendimento(df, col_map)

    # Subconjuntos de chamados por fase
    if fases is not None:
        df_fin = df[fases == 'Finalizado']
        df_em = df[fases == 'Em Andamento']
        df_aprov = df[fases == 'Aprovação de Solicitação']
    else:
        df_fin = df_em = df_aprov = df.iloc[0:0]

    st.caption("💡 Clique em qualquer card para abrir a janela com a tabela detalhada "
               "(exportável para Excel).")

    # ════════════ Linha 1: Operação / Fases ════════════
    o1, o2, o3, o4 = st.columns(4)
    _card(o1, "kpicard_total", "📋 Total de Chamados", f"{total:,}", df, col_map)
    _card(o2, "kpicard_fin", "✅ Finalizados", f"{len(df_fin):,}", df_fin, col_map)
    _card(o3, "kpicard_em", "⚙️ Em Andamento", f"{len(df_em):,}", df_em, col_map)
    _card(o4, "kpicard_aprov", "📬 Aprovação de Solicitação",
          f"{len(df_aprov):,}", df_aprov, col_map)

    # ════════════ Linha 2: SLA (conformidade + tempo) ════════════
    s1, s2, s3 = st.columns(3)

    # SLA Cumprido (% padrão) — a JANELA lista os VIOLADOS (acionáveis)
    if sla_all is not None and sla_all.notna().any():
        sl = sla_all.dropna()
        val_sla = f"{sl.sum() / len(sl) * 100:.1f}%"
        df_viol = df[sla_all == False]  # noqa: E712  (False = violado)
    else:
        val_sla = "—"
        df_viol = df.iloc[0:0]
    _card(s1, "kpicard_sla", "⏱ SLA Cumprido (prazo)", val_sla, df_viol, col_map,
          titulo_janela="⏱ Chamados com SLA VIOLADO")

    # Tempo Médio de Atendimento — janela lista chamados com SLA iniciado
    t = tempo.dropna() if tempo is not None else None
    if t is not None and len(t):
        val_tempo = f"{t.mean():.1f}d"
        df_tempo = df[tempo.notna()]
    else:
        val_tempo = "—"
        df_tempo = df.iloc[0:0]
    _card(s2, "kpicard_tempo", "⏳ Tempo Médio de Atendimento",
          val_tempo, df_tempo, col_map,
          titulo_janela="⏳ Chamados com SLA iniciado")

    # Dentro da Meta (% padrão) — a JANELA lista os que ESTOURARAM a meta
    if t is not None and len(t):
        val_meta = f"{(t <= meta).sum() / len(t) * 100:.1f}%"
        df_acima = df[tempo > meta]
    else:
        val_meta = "—"
        df_acima = df.iloc[0:0]
    _card(s3, "kpicard_meta", f"🎯 Dentro da Meta ({meta}d)",
          val_meta, df_acima, col_map,
          titulo_janela=f"🎯 Chamados ACIMA da meta ({meta}d)")


# ── Donut: SLA ────────────────────────────────────────────────────────────────

def chart_sla_donut(df: pd.DataFrame, col_map: dict):
    sla_all = sla_series(df, col_map)
    if sla_all is None:
        return None
    sla = sla_all.dropna()
    if len(sla) == 0:
        return None
    labels = ['Cumprido', 'Violado']
    values = [int(sla.sum()), int((~sla).sum())]
    pct = values[0] / len(sla) * 100

    fig = go.Figure(go.Pie(
        labels=labels, values=values, hole=0.62,
        marker_colors=[SLA_PALETTE['Cumprido'], SLA_PALETTE['Violado']],
        textinfo='label+percent',
    ))
    fig.update_layout(
        title='Conformidade SLA (a partir de Em Andamento)',
        annotations=[dict(text=f'{pct:.0f}%', x=0.5, y=0.5,
                          font_size=24, showarrow=False, font_color='#ECF0F1')],
        legend=dict(orientation='h', y=-0.08),
        **_LAYOUT,
    )
    return fig


# ── Donut: Status ─────────────────────────────────────────────────────────────

def chart_status_donut(df: pd.DataFrame, col_map: dict):
    if 'status' not in col_map:
        return None
    fases = classify_fase(df[col_map['status']])
    cnt = fases.value_counts()
    total = len(fases)
    pct_fin = cnt.get('Finalizado', 0) / total * 100 if total else 0

    fig = go.Figure(go.Pie(
        labels=cnt.index.tolist(), values=cnt.values.tolist(), hole=0.62,
        marker_colors=[FASE_CORES.get(l, '#95A5A6') for l in cnt.index],
        textinfo='label+percent', sort=False,
    ))
    fig.update_layout(
        title='Distribuição por Fase',
        annotations=[dict(text=f'{pct_fin:.0f}%\nfinalizado', x=0.5, y=0.5,
                          font_size=16, showarrow=False, font_color='#ECF0F1')],
        legend=dict(orientation='h', y=-0.12),
        **_LAYOUT,
    )
    return fig


# ── Line: Evolução temporal ───────────────────────────────────────────────────

def chart_timeline(df: pd.DataFrame, col_map: dict):
    date_col = col_map.get('data_abertura')
    if not date_col:
        return None
    try:
        tmp = df[[date_col]].copy()
        tmp['_dt'] = pd.to_datetime(tmp[date_col], dayfirst=True, errors='coerce')
        tmp = tmp.dropna(subset=['_dt'])
        if len(tmp) < 2:
            return None
        span = (tmp['_dt'].max() - tmp['_dt'].min()).days
        if span <= 62:
            tmp['_p'] = tmp['_dt'].dt.strftime('%d/%m/%Y')
            label = 'Dia'
        elif span <= 730:
            tmp['_p'] = tmp['_dt'].dt.strftime('%m/%Y')
            label = 'Mês'
        else:
            tmp['_p'] = tmp['_dt'].dt.strftime('%Y')
            label = 'Ano'
        tl = tmp.groupby('_p').size().reset_index(name='Chamados')
        tl.columns = [label, 'Chamados']
        fig = px.line(tl, x=label, y='Chamados',
                      title=f'Evolução de Chamados por {label}', markers=True)
        fig.update_traces(line_color='#3498DB', marker_color='#3498DB')
        return _apply(fig)
    except Exception:
        return None


# ── Bar: % SLA por Categoria ─────────────────────────────────────────────────

def chart_sla_por_categoria(df: pd.DataFrame, col_map: dict):
    if 'categoria' not in col_map or sla_series(df, col_map) is None:
        return None
    tmp = df[[col_map['categoria']]].copy()
    # True -> 1.0, False -> 0.0, None -> NaN (garante dtype numérico)
    tmp['_sla'] = sla_series(df, col_map).map({True: 1.0, False: 0.0})
    tmp['_cat'] = tmp[col_map['categoria']]
    tmp = tmp.dropna(subset=['_sla'])
    if tmp.empty:
        return None
    grp = tmp.groupby('_cat')['_sla'].agg(['sum', 'count']).reset_index()
    grp.columns = ['Categoria', 'Cumprido', 'Total']
    grp['% SLA'] = (grp['Cumprido'] / grp['Total'] * 100).round(1)
    grp = grp.sort_values('% SLA').head(15)

    fig = px.bar(grp, x='% SLA', y='Categoria', orientation='h',
                 title='% SLA Cumprido por Categoria',
                 color='% SLA',
                 color_continuous_scale=['#E74C3C', '#F39C12', '#2ECC71'],
                 range_color=[0, 100], text='% SLA')
    fig.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
    fig.add_vline(x=80, line_dash='dash', line_color='#F39C12',
                  annotation_text=' Meta 80%', annotation_font_color='#F39C12')
    fig.update_layout(coloraxis_showscale=False,
                      yaxis=dict(categoryorder='total ascending'))
    return _apply(fig)


# ── Bar: Tempo médio de atendimento por Departamento (com meta) ───────────────

def chart_tempo_por_categoria(df: pd.DataFrame, col_map: dict, meta: int = SLA_META_DIAS):
    if 'categoria' not in col_map:
        return None
    tempo = tempo_atendimento(df, col_map)
    if tempo is None:
        return None

    tmp = pd.DataFrame({'cat': df[col_map['categoria']], 'dias': tempo}).dropna()
    if tmp.empty:
        return None

    import numpy as np
    grp = tmp.groupby('cat')['dias'].mean().reset_index()
    grp.columns = ['Departamento', 'media']
    grp['Dias'] = np.ceil(grp['media']).astype(int)   # arredonda os dias PARA CIMA
    grp = grp.sort_values('Dias', ascending=False).head(12)

    # Verde se dentro da meta, vermelho se acima
    cores = ['#2ECC71' if v <= meta else '#E74C3C' for v in grp['Dias']]
    fig = px.bar(grp, x='Dias', y='Departamento', orientation='h',
                 title=f'Tempo Médio de Atendimento por Departamento (meta {meta}d)',
                 text='Dias')
    fig.update_traces(marker_color=cores, texttemplate='%{text} d',
                      textposition='outside', cliponaxis=False,
                      hoverinfo='skip', hovertemplate=None)  # sem interação/hover
    fig.add_vline(x=meta, line_dash='dash', line_color='#F39C12',
                  annotation_text=f' Meta {meta}d',
                  annotation_font_color='#F39C12')
    fig.update_layout(showlegend=False,
                      yaxis=dict(categoryorder='total ascending'))
    return _apply(fig)


# ── Bar: Sistemas mais solicitados ────────────────────────────────────────────

def chart_top_sistemas(df: pd.DataFrame, col_map: dict, n: int = 12):
    # Considera apenas chamados com Fase atual = 'Em Andamento'
    if 'status' in col_map:
        df = df[classify_fase(df[col_map['status']]) == 'Em Andamento']
    s = sistema_series(df, col_map)
    if s is None:
        return None
    s = s[~s.str.lower().isin(['', 'nan', 'none', '-'])]
    if s.empty:
        return None

    cnt = s.value_counts().reset_index()   # todos os sistemas (sem corte)
    cnt.columns = ['Sistema', 'Solicitações']

    # Barras HORIZONTAIS modernas (cantos arredondados) — SEM texto nativo,
    # pois o Plotly ignora o ângulo no texto interno da barra
    fig = px.bar(cnt, x='Solicitações', y='Sistema', orientation='h',
                 title=f'Sistemas Mais Solicitados (Em Andamento) — {len(cnt)} sistemas')
    fig.update_traces(
        marker=dict(color='#2E86C1', cornerradius=8, line=dict(width=0)),
        hovertemplate='%{y}: %{x} chamado(s)<extra>clique para detalhar</extra>',
    )

    # Números EM PÉ (upright, sem rotação) no fim de cada barra, via anotações
    xmax = cnt['Solicitações'].max() or 1
    for _, r in cnt.iterrows():
        v = int(r['Solicitações'])
        grande = v / xmax > 0.12          # cabe o número dentro, perto da ponta?
        fig.add_annotation(
            x=v, y=r['Sistema'], text=str(v),
            textangle=0, showarrow=False, yanchor='middle',
            xanchor=('right' if grande else 'left'),
            xshift=(-8 if grande else 6),
            font=dict(color=('white' if grande else '#ECF0F1'),
                      size=13, family=FONT_FAMILY),
        )

    # altura proporcional à quantidade de barras (legível mesmo com muitas)
    fig.update_layout(
        height=max(360, 26 * len(cnt) + 90),
        showlegend=False,
        xaxis=dict(visible=False),
        yaxis=dict(categoryorder='total ascending'),
    )
    return _apply(fig)


# ── Modal: chamados de um sistema (ao clicar na barra) ────────────────────────

def _tabela_sistema(df_sub: pd.DataFrame, col_map: dict) -> pd.DataFrame:
    """Tabela de chamados de um sistema, com dados de SLA (mais antigos primeiro)."""
    import numpy as np

    # Ordena pelos mais antigos primeiro (data de criação ascendente; sem data por último)
    date_col = col_map.get('data_abertura')
    if date_col and date_col in df_sub.columns:
        ordem = pd.to_datetime(df_sub[date_col], errors='coerce')
        df_sub = df_sub.loc[ordem.sort_values(kind='stable', na_position='last').index]

    out = pd.DataFrame(index=df_sub.index)

    def add(field, titulo, is_date=False):
        c = col_map.get(field)
        if c and c in df_sub.columns:
            out[titulo] = (fmt_data(df_sub[c]) if is_date else df_sub[c]).values

    add('id', 'Ticket')
    add('solicitante', 'Nome Completo')
    add('email', 'E-mail')
    add('categoria', 'Departamento')
    sis = sistema_series(df_sub, col_map)
    if sis is not None:
        out['Sistema'] = sis.values
    ac = col_map.get('atendente')
    if ac and ac in df_sub.columns:
        out['Responsável'] = df_sub[ac].apply(abreviar_nome).values

    # Dados de SLA
    sla = sla_series(df_sub, col_map)
    if sla is not None:
        out['SLA'] = sla.map({True: '✅ Cumprido', False: '❌ Violado'}).fillna('—').values
    add('sla_vencimento', 'Vencimento', is_date=True)
    tempo = tempo_atendimento(df_sub, col_map)
    if tempo is not None:
        out['Tempo (dias)'] = ['' if pd.isna(v) else int(np.ceil(v)) for v in tempo]

    return out.reset_index(drop=True)


# Colunas que ganham botão de copiar
_COPIAVEIS = {'Ticket', 'E-mail'}


def _tabela_copiavel(tab: pd.DataFrame) -> None:
    """Tabela HTML com grade alinhada e botão de copiar em Código e E-mail."""
    cols = list(tab.columns)
    thead = ''.join(f'<th>{_esc(c)}</th>' for c in cols)

    linhas = []
    for _, row in tab.iterrows():
        tds = []
        for c in cols:
            val = '' if pd.isna(row[c]) else str(row[c])
            if c in _COPIAVEIS and val:
                tds.append(
                    f'<td><span class="v">{_esc(val)}</span>'
                    f'<button class="cp" onclick=\'cp(this,{json.dumps(val)})\' '
                    f'title="Copiar">📋</button></td>')
            else:
                tds.append(f'<td>{_esc(val)}</td>')
        linhas.append('<tr>' + ''.join(tds) + '</tr>')
    corpo = '\n'.join(linhas)

    html = f"""
    <style>
      *{{box-sizing:border-box;font-family:Inter,'Segoe UI',Roboto,sans-serif;}}
      body{{margin:0;background:#0E1117;color:#ECF0F1;}}
      table{{border-collapse:collapse;width:100%;font-size:13px;}}
      th,td{{border:1px solid #2A3A4F;padding:6px 9px;text-align:left;
             vertical-align:middle;white-space:nowrap;}}
      th{{background:#1A2332;position:sticky;top:0;font-weight:600;}}
      tr:nth-child(even) td{{background:#141b27;}}
      td .v{{user-select:all;}}
      button.cp{{margin-left:8px;border:none;background:#243044;color:#cfe3ff;
                border-radius:4px;cursor:pointer;font-size:11px;padding:1px 6px;
                vertical-align:middle;}}
      button.cp:hover{{background:#3498DB;color:#fff;}}
    </style>
    <table><thead><tr>{thead}</tr></thead><tbody>{corpo}</tbody></table>
    <script>
      function cp(btn, text){{
        const done = () => {{ const o=btn.textContent; btn.textContent='✓';
          setTimeout(()=>btn.textContent=o, 1000); }};
        if (navigator.clipboard && navigator.clipboard.writeText) {{
          navigator.clipboard.writeText(text).then(done, ()=>fallback(text, done));
        }} else {{ fallback(text, done); }}
      }}
      function fallback(text, done){{
        const t=document.createElement('textarea'); t.value=text;
        document.body.appendChild(t); t.select();
        try{{document.execCommand('copy');}}catch(e){{}} t.remove(); done();
      }}
    </script>
    """
    altura = min(470, 70 + len(tab) * 39)
    components.html(html, height=altura, scrolling=True)


@st.dialog("🖥️ Chamados do Sistema", width="large")
def _dialog_sistema(sistema_nome: str, df_sub: pd.DataFrame, col_map: dict) -> None:
    st.markdown(f"### {sistema_nome}")
    st.caption(f"{len(df_sub):,} chamado(s) Em Andamento")
    tab = _tabela_sistema(df_sub, col_map)
    if tab.empty:
        st.info("Sem chamados para este sistema.")
        return
    _tabela_copiavel(tab)
    safe = re.sub(r'[^A-Za-z0-9_-]+', '_', str(sistema_nome)).strip('_') or 'sistema'
    st.download_button(
        "⬇️ Baixar Excel (.xlsx)", data=_to_excel(tab),
        file_name=f"chamados_{safe}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


def abrir_chamados_do_sistema(sistema_nome: str, df: pd.DataFrame, col_map: dict) -> None:
    """Filtra os chamados Em Andamento do sistema clicado e abre o modal."""
    sub = df
    if 'status' in col_map:
        sub = sub[classify_fase(sub[col_map['status']]) == 'Em Andamento']
    sis = sistema_series(sub, col_map)
    if sis is not None:
        sub = sub[sis == sistema_nome]
    _dialog_sistema(sistema_nome, sub, col_map)


# ── Bar: Top 5 Atendentes que finalizaram nos últimos 6 meses ─────────────────

def chart_atendentes(df: pd.DataFrame, col_map: dict):
    atend_col = col_map.get('atendente')
    if not atend_col:
        return None

    tmp = df.copy()

    # 1) Somente chamados finalizados
    if 'status' in col_map:
        tmp = tmp[normalize_status(tmp[col_map['status']]) == 'Atendido']

    # 2) Apenas finalizações dos últimos 6 meses (pela data de fechamento)
    date_col = col_map.get('data_fechamento') or col_map.get('data_abertura')
    periodo = False
    if date_col:
        dt = pd.to_datetime(tmp[date_col], errors='coerce')
        cutoff = pd.Timestamp.now() - pd.DateOffset(months=6)
        tmp = tmp[dt >= cutoff]
        periodo = True

    # 3) Remover responsáveis vazios, abreviar nomes e contar
    resp = tmp[atend_col].fillna('').astype(str).str.strip()
    resp = resp[~resp.str.lower().isin(['', 'nan', 'none', '-'])]
    if resp.empty:
        return None
    resp = resp.apply(abreviar_nome)

    cnt = resp.value_counts().head(5).reset_index()
    cnt.columns = ['Atendente', 'Finalizados']

    titulo = ('Top 5 Atendentes — Finalizados (últimos 6 meses)'
              if periodo else 'Top 5 Atendentes — Finalizados')
    # Cores sólidas e escuras — contraste garantido com o texto branco
    BAR_COLORS = ['#1A5276', '#117A65', '#7D3C98', '#B9770E', '#922B21']
    fig = px.bar(cnt, x='Finalizados', y='Atendente', orientation='h',
                 title=titulo, text='Finalizados',
                 color='Atendente',
                 color_discrete_sequence=BAR_COLORS)
    # Números centralizados DENTRO das barras, em branco e negrito
    fig.update_traces(texttemplate='<b>%{text:,}</b>', textposition='inside',
                      insidetextanchor='middle',
                      textfont=dict(size=16, color='white'), cliponaxis=False)
    fig.update_layout(
        showlegend=False,
        title_x=0.5,                              # título centralizado
        yaxis=dict(categoryorder='total ascending'),
        xaxis=dict(visible=False),                # eixo X some (número já está na barra)
        uniformtext_minsize=12, uniformtext_mode='show',
    )
    return _apply(fig)


# ── Pizza (donut): Finalizados por responsável (com filtro de período) ────────

def _finalizados_base(df: pd.DataFrame, col_map: dict):
    """(fin_df, dt) de chamados finalizados nos últimos 6 meses."""
    if 'status' not in col_map:
        return None, None
    atend = col_map.get('atendente')
    dcol = col_map.get('data_fechamento') or col_map.get('data_abertura')
    if not atend or not dcol:
        return None, None
    fin = df[classify_fase(df[col_map['status']]) == 'Finalizado'].copy()
    dt = pd.to_datetime(fin[dcol], errors='coerce')
    m = dt >= (pd.Timestamp.now() - pd.DateOffset(months=6))
    return fin[m], dt[m]


def periodos_finalizados(df: pd.DataFrame, col_map: dict, freq: str):
    """Lista de (rótulo, início, fim) dos períodos com finalizados (mais recente 1º)."""
    fin, dt = _finalizados_base(df, col_map)
    if fin is None or fin.empty:
        return []
    pers = sorted(pd.PeriodIndex(dt.dt.to_period(freq)).unique(), reverse=True)
    out = []
    for p in pers:
        if freq == 'M':
            rotulo = p.start_time.strftime('%m/%Y')
        elif freq == 'W':
            rotulo = 'Semana de ' + p.start_time.strftime('%d/%m/%Y')
        else:
            rotulo = p.start_time.strftime('%d/%m/%Y')
        out.append((rotulo, p.start_time, p.end_time))
    return out


def chart_finalizados_periodo(df: pd.DataFrame, col_map: dict,
                              inicio=None, fim=None,
                              escopo: str = 'últimos 6 meses', top: int = 8):
    """Pizza/donut de finalizados por responsável; opcionalmente filtrado por período."""
    fin, dt = _finalizados_base(df, col_map)
    if fin is None or fin.empty:
        return None
    if inicio is not None and fim is not None:
        m = (dt >= inicio) & (dt <= fim)
        fin = fin[m]
    if fin.empty:
        return None

    atend = col_map['atendente']
    resp = fin[atend].fillna('').astype(str).str.strip()
    resp = resp.where(~resp.str.lower().isin(['', 'nan', 'none', '-']), '—')
    resp = resp.apply(lambda v: abreviar_nome(v) if v != '—' else v)

    cnt = resp.value_counts()
    if len(cnt) > top:                       # agrupa a cauda em "Outros"
        cnt = pd.concat([cnt.head(top), pd.Series({'Outros': cnt.iloc[top:].sum()})])

    total = int(cnt.sum())
    fig = px.pie(values=cnt.values, names=cnt.index, hole=0.5,
                 color_discrete_sequence=px.colors.qualitative.Bold,
                 title=f'Finalizados por Responsável ({escopo}) — {total} no total')
    # Rótulos PARA FORA com conectores (linha de apontamento) em cada fatia.
    # A pizza é empurrada para baixo (domain y) para os rótulos do topo não
    # atropelarem o título.
    fig.update_traces(textposition='outside', textinfo='label+value',
                      automargin=True, sort=True, pull=0.03,
                      domain=dict(x=[0, 1], y=[0.0, 0.82]),
                      marker=dict(line=dict(color='#0D1B2A', width=1)),
                      hovertemplate='%{label}: %{value} (%{percent})<extra></extra>')
    fig.update_layout(
        height=560, showlegend=False,
        title=dict(x=0.5, xanchor='center', y=0.99, yanchor='top'),
        uniformtext_minsize=11, uniformtext_mode='show',
        annotations=[dict(text=f'{total}<br>finalizados', x=0.5, y=0.41,
                          font_size=16, showarrow=False, font_color='#ECF0F1')],
        margin=dict(t=70, b=40, l=70, r=70),
    )
    return _apply(fig)
