import re
import unicodedata
import pandas as pd
from io import BytesIO
from typing import Optional

# Keywords mapped to semantic fields
COLUMN_KEYWORDS = {
    'id': ['id', 'numero', 'número', 'chamado', 'ticket', 'protocolo', 'cod', 'codigo', 'código', 'seq', 'n°', 'n.'],
    'tipo': ['tipo', 'type', 'modalidade', 'natureza', 'classificacao', 'classificação'],
    'categoria': ['categoria', 'category', 'categ', 'subcategoria', 'grupo', 'assunto', 'servico', 'serviço', 'item'],
    'sistema': ['sistema', 'sistemas', 'aplicativo', 'aplicacao', 'aplicação',
                'software', 'ferramenta', 'plataforma', 'qual_sistema', 'qual_o_sistema'],
    # Bare 'sla' removed on purpose: it would match "slack". Use separators or EXACT_MAP.
    'sla': ['dentro_sla', 'fora_sla', 'sla_status', 'status_sla', 'cumprimento_sla',
            'prazo_sla', 'sla_cumprido', 'sla_atendido'],
    'vencido': ['vencido', 'expirado', 'em_atraso', 'atrasado'],
    'status': ['status', 'situacao', 'situação', 'fase_atual', 'fase atual', 'atendido', 'resolvido', 'estado', 'fechado', 'encerrado', 'resolved', 'closed'],
    'data_abertura': ['data_abertura', 'data abertura', 'abertura', 'criacao', 'criação', 'created', 'opened', 'dt_abertura', 'aberto_em', 'data'],
    'data_fechamento': ['data_fechamento', 'fechamento', 'encerramento', 'resolved_at', 'closed_at', 'dt_fechamento'],
    'prioridade': ['prioridade', 'priority', 'urgencia', 'urgência', 'criticidade', 'nivel', 'nível'],
    'atendente': ['atendente', 'responsavel', 'responsável', 'analista', 'agent', 'assignee', 'tecnico', 'técnico'],
    'solicitante': ['solicitante', 'usuario', 'usuário', 'requester', 'colaborador', 'user', 'cliente'],
    'sla_vencimento': ['vencimento', 'vencimento_sla', 'prazo_sla', 'deadline', 'data_vencimento',
                       'data_limite', 'limite_sla', 'sla_date', 'expiracao', 'expiração'],
    'data_em_andamento': ['data_em_andamento', 'dt_em_andamento', 'inicio_andamento',
                          'data_inicio_atend', 'andamento_em', 'data_inicio', 'dt_inicio'],
    'data_aprovacao': ['data_aprovacao', 'dt_aprovacao', 'aprovado_em', 'data_aprov',
                       'aprovacao_em', 'dt_aprov'],
}

# Exact (normalized) column name -> field. Checked BEFORE fuzzy keywords.
# Handles structured exports (e.g. Pipefy) where fuzzy matching is ambiguous.
EXACT_MAP = {
    'codigo': 'id',
    'sla': 'sla',
    'fase_atual': 'status',
    'vencido': 'vencido',
    'responsavel': 'atendente',
    'criado_em': 'data_abertura',
    'finalizado_em': 'data_fechamento',
    'data_de_vencimento': 'sla_vencimento',
    'tipo_de_funcionario:': 'tipo',
    'departamento:': 'categoria',
    'cargo:': 'cargo',
    'para_quem_e_esta_solicitacao:': 'solicitante',
    'primeira_vez_que_entrou_na_fase_em_andamento': 'data_em_andamento',
    'primeira_vez_que_entrou_na_fase_aprovacao_de_solicitacao': 'data_aprovacao',
    'tempo_total_na_fase_em_andamento_(dias)': 'tempo_em_andamento',
    'tempo_total_na_fase_aprovacao_de_solicitacao_(dias)': 'tempo_aprovacao',
    'tempo_total_na_fase_finalizados_(dias)': 'tempo_finalizados',
    'data_de_inicio:': 'data_inicio',
    'data_para_envio_das_credencias:': 'data_credenciais',
    'e_mail_gestor:': 'email_gestor',
    # Pipe [TI] Sistemas
    'sistema:': 'sistema',
    'informe_qual_o_sistema:': 'sistema_outro',
    'solicitacao:': 'solicitacao',
    'nome_completo:': 'solicitante',
    'qual_servico?': 'servico',
}

# Meta de SLA (dias) na fase Em Andamento — prazo-alvo definido pela gestão
SLA_META_DIAS = 7

SLA_TRUE = {'sim', 'yes', 'true', '1', 'dentro', 'cumprido', 'atendido', 'ok', 's', 'y', 'no prazo', 'dentro do prazo'}
SLA_FALSE = {'nao', 'não', 'no', 'false', '0', 'fora', 'violado', 'n', 'fora do prazo', 'atrasado', 'quebrado'}

STATUS_DONE = {
    'sim', 'yes', 'true', '1', 'atendido', 'resolvido', 'fechado', 'encerrado',
    'concluido', 'concluído', 'ok', 's', 'y', 'closed', 'resolved', 'done', 'finalizado',
}


def _normalize(text: str) -> str:
    """Lowercase, strip accents, replace spaces/hyphens with underscores."""
    t = str(text).lower().strip()
    nfkd = unicodedata.normalize('NFKD', t)
    t = ''.join(c for c in nfkd if not unicodedata.combining(c))
    return t.replace(' ', '_').replace('-', '_').replace('.', '_')


def _read_csv(raw: bytes) -> pd.DataFrame:
    """Read a CSV, auto-detecting separator and encoding."""
    # Try encodings common in BR exports (utf-8 with BOM, then latin1)
    text = None
    for enc in ('utf-8-sig', 'utf-8', 'latin1', 'cp1252'):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = raw.decode('latin1', errors='replace')

    # Detect separator from the header line (; vs , vs tab)
    header = text.split('\n', 1)[0]
    sep = max([';', ',', '\t'], key=header.count)

    from io import StringIO
    df = pd.read_csv(StringIO(text), sep=sep, dtype=str, engine='python',
                     on_bad_lines='skip')
    return df


def load_xlsx(source, filename: str = '') -> pd.DataFrame:
    """Load a table from XLSX or CSV (file-like, bytes). Name hints the format."""
    raw = source if isinstance(source, bytes) else source.read()
    name = (filename or getattr(source, 'name', '')).lower()

    if name.endswith('.csv') or name.endswith('.txt'):
        df = _read_csv(raw)
    else:
        try:
            df = pd.read_excel(BytesIO(raw), engine='openpyxl')
        except Exception:
            df = _read_csv(raw)  # fallback: maybe it's a CSV mislabeled

    df.columns = [str(c).strip() for c in df.columns]
    return df


def detect_columns(df: pd.DataFrame) -> dict:
    """Return a dict mapping semantic field names to actual column names."""
    norm_to_orig = {_normalize(c): c for c in df.columns}
    detected: dict = {}
    used: set = set()

    # 1) Exact (normalized) matches first — most reliable for structured exports
    for norm_col, orig_col in norm_to_orig.items():
        field = EXACT_MAP.get(norm_col)
        if field and field not in detected:
            detected[field] = orig_col
            used.add(orig_col)

    # 2) Fuzzy keyword matching for whatever is still missing
    for field, keywords in COLUMN_KEYWORDS.items():
        if field in detected:
            continue
        for norm_col, orig_col in norm_to_orig.items():
            if orig_col in used:
                continue
            for kw in keywords:
                nkw = _normalize(kw)
                if nkw in norm_col or norm_col in nkw:
                    detected[field] = orig_col
                    used.add(orig_col)
                    break
            if field in detected:
                break

    # 3) E-mail do solicitante: detectado pelos VALORES (coluna com mais e-mails),
    #    pois o nome da coluna varia entre pipes (ex.: 'Título' guarda o e-mail).
    if 'email' not in detected:
        pat = re.compile(r'^[^@\s]+@[^@\s]+\.[^@\s]+$')
        melhor, melhor_ratio = None, 0.0
        for col in df.columns:
            if col in used:
                continue
            s = df[col].dropna().astype(str).str.strip()
            s = s[s != '']
            if len(s) < 5:
                continue
            ratio = s.apply(lambda v: bool(pat.match(v))).mean()
            if ratio > melhor_ratio:
                melhor, melhor_ratio = col, ratio
        if melhor and melhor_ratio >= 0.7:
            detected['email'] = melhor
            used.add(melhor)

    return detected


def normalize_sla(series: pd.Series) -> pd.Series:
    """Map SLA column to boolean (True = SLA met). Unknown values become NaN."""
    if series.dtype == bool:
        return series

    normalized = series.astype(str).apply(_normalize)

    def _map(val: str) -> Optional[bool]:
        if val in SLA_TRUE or any(kw in val for kw in ('cumpr', 'dentro', 'sim', 'yes', 'atend')):
            return True
        if val in SLA_FALSE or any(kw in val for kw in ('nao', 'fora', 'viol', 'atras', 'quebr')):
            return False
        try:
            return float(val) > 0
        except ValueError:
            return None

    return normalized.apply(_map).astype(object)


def normalize_vencido(series: pd.Series) -> pd.Series:
    """Map a 'Vencido' column to boolean SLA-met (True = NOT expired).

    Vencido=TRUE means the deadline was missed (SLA violado), so SLA met = False.
    """
    normalized = series.astype(str).apply(_normalize)

    def _map(val: str) -> Optional[bool]:
        if val in ('true', 'sim', '1', 's', 'yes', 'vencido'):
            return False   # vencido -> SLA NÃO cumprido
        if val in ('false', 'nao', '0', 'n', 'no'):
            return True    # não vencido -> SLA cumprido
        return None

    return normalized.apply(_map).astype(object)


def sistema_series(df: pd.DataFrame, col_map: dict) -> Optional[pd.Series]:
    """Nome do sistema solicitado, com resolução de 'Outro'.

    Quando 'Sistema:' = 'Outro' (ou vazio), busca o sistema real em:
      1) 'Informe qual o sistema:'  2) 'Qual Serviço?'
    E normaliza (case/acentos/espaços) para SOMAR ao sistema canônico
    correspondente do dropdown — ex.: 'databricks' → 'Databricks'.
    """
    base_col = col_map.get('sistema')
    if not base_col:
        return None

    # Apelidos: se a chave normalizada CONTÉM o termo, vira o nome canônico.
    # Ex.: 'Claude', 'Claude AI', 'Claude IA', 'Claude Code' → 'Claude Code'.
    SISTEMA_ALIASES = {
        'claude': 'Claude Code',
        'cloude': 'Claude Code',   # erro de grafia comum de "Claude"
        'admin': 'CMS / Admin',          # Admin e CMS são o mesmo sistema
        'cms': 'CMS / Admin',
        'sistemans': 'CMS / Admin',      # betnacional.sistemans.com → CMS/Admin
    }

    base = df[base_col].fillna('').astype(str).str.strip()
    n = len(df)
    outro_col = col_map.get('sistema_outro')
    serv_col = col_map.get('servico')
    outro = (df[outro_col].fillna('').astype(str).str.strip()
             if outro_col else pd.Series([''] * n, index=df.index))
    serv = (df[serv_col].fillna('').astype(str).str.strip()
            if serv_col else pd.Series([''] * n, index=df.index))

    def _key(v: str) -> str:
        v = unicodedata.normalize('NFKD', str(v).strip().lower())
        v = ''.join(c for c in v if not unicodedata.combining(c))
        return re.sub(r'\s+', ' ', v)

    GENERICO = {'outro', 'outros', '', 'nan', 'none', '-'}

    # Termos que indicam descrição/serviço (não é nome de sistema)
    NAO_SISTEMA = ('solicit', 'criaca', 'criacao', 'inclus', 'encaminh', 'transmiss',
                   'assinatura', 'liberac', 'cadastr', 'permiss', 'configurac',
                   'alterac', 'recebid', 'automatic')

    def _limpar(v: str) -> str:
        """Extrai o nome do sistema de um texto livre verboso.
        - URL → domínio (https://x.com/a → x.com)
        - remove detalhes entre parênteses: 'Argo CD (PROD)' → 'Argo CD'
        """
        v = str(v).strip()
        m = re.search(r'https?://([^/\s]+)', v, flags=re.I)
        if m:
            return m.group(1)
        v = re.sub(r'\(.*?\)', '', v)              # remove (...)
        v = re.sub(r'\s+', ' ', v).strip(' -–—:;,/')
        return v

    def _parece_sistema(val: str) -> bool:
        """Heurística: nomes de sistema são curtos; descrições/e-mails não contam."""
        v = val.strip()
        k = _key(v)
        if not v or '@' in v:                       # e-mail não é sistema
            return False
        if len(v) > 30 or len(k.split()) > 4:       # frase longa → não é sistema
            return False
        return not any(t in k for t in NAO_SISTEMA)

    # Nomes canônicos vindos do dropdown (Sistema:), por chave normalizada.
    # Este dicionário CRESCE: textos livres novos são registrados para dedupe.
    canon = {}
    for v in base:
        k = _key(v)
        if k and k not in GENERICO and k not in canon:
            canon[k] = v.strip()

    def _resolver(b: str, o: str, s: str) -> str:
        val = b.strip()
        if _key(val) in GENERICO:                 # é 'Outro' → buscar/limpar o real
            val = _limpar(o if o else s)
        k = _key(val)
        if k in GENERICO or not val:              # sem informação → Outro
            return 'Outro'
        for termo, canonico in SISTEMA_ALIASES.items():  # apelidos têm prioridade
            if termo in k:
                return canonico
        if k in canon:                            # SOMA ao sistema já existente
            return canon[k]
        if _parece_sistema(val):                  # CRIA novo sistema (e registra p/ dedupe)
            canon[k] = val
            return val
        return 'Outro'                            # descrição/serviço/e-mail → Outro

    resolved = [_resolver(b, o, s) for b, o, s in zip(base, outro, serv)]
    return pd.Series(resolved, index=df.index)


_CONECTORES_NOME = {'de', 'da', 'do', 'das', 'dos', 'e'}
# Apelidos: se o nome (minúsculo) CONTÉM a chave, vira o valor
NOME_ALIASES = {
    'hayslan': 'Hayslan Yuri',
    'saulo': 'Saulo Santos',
    'conde': 'Marcel Conde',            # 'marcel sales conde' (evita colidir c/ 'Marcelo')
    'gabriel-soares': 'Gabriel Soares',  # username gabriel-soares-nsx
}


def abreviar_nome(valor) -> str:
    """Mostra 'Primeiro Último' (ex.: 'Bruno Augusto dos Santos' → 'Bruno Santos').
    Aplica apelidos (hayslan → Hayslan Yuri); trata múltiplos nomes (vírgula) e usernames."""
    s = str(valor).strip()
    if not s or s.lower() == 'nan':
        return ''
    out = []
    for parte in s.split(','):
        p = parte.strip()
        if not p:
            continue
        chave = p.lower()
        alias = next((v for k, v in NOME_ALIASES.items() if k in chave), None)
        if alias:
            out.append(alias)
            continue
        toks = [t for t in p.split() if t.lower() not in _CONECTORES_NOME]
        out.append(f'{toks[0]} {toks[-1]}' if len(toks) >= 2 else p)
    return ', '.join(out)


def tempo_dias(series: pd.Series) -> pd.Series:
    """Converte uma coluna de 'dias' (com vírgula decimal, ex: '38,98') em float."""
    s = series.astype(str).str.strip().str.replace(',', '.', regex=False)
    return pd.to_numeric(s, errors='coerce')


def fmt_data(series: pd.Series, com_hora: bool = False) -> pd.Series:
    """Formata datas para o padrão brasileiro DD/MM/AAAA (datas inválidas viram '').

    Usa format='mixed' para lidar com valores mistos (só data e data+hora).
    """
    try:
        dt = pd.to_datetime(series, errors='coerce', format='mixed', dayfirst=False)
    except (ValueError, TypeError):
        dt = pd.to_datetime(series, errors='coerce')
    fmt = '%d/%m/%Y %H:%M' if com_hora else '%d/%m/%Y'
    return dt.dt.strftime(fmt).fillna('')


def tempo_atendimento(df: pd.DataFrame, col_map: dict) -> Optional[pd.Series]:
    """Tempo (dias) na fase Em Andamento = duração real do SLA. None se não houver."""
    col = col_map.get('tempo_em_andamento')
    if not col:
        return None
    return tempo_dias(df[col])


def sla_iniciado_mask(df: pd.DataFrame, col_map: dict) -> pd.Series:
    """Máscara True para chamados cujo SLA já começou (entraram em 'Em andamento').

    O relógio de SLA só inicia quando o chamado cai na coluna 'Em andamento'.
    Chamados ainda em 'Aprovação de Solicitação' não têm essa data e ficam fora.
    """
    col = col_map.get('data_em_andamento')
    if not col:
        return pd.Series(True, index=df.index)  # sem a info, considera todos
    return pd.to_datetime(df[col], errors='coerce').notna()


def sla_series(df: pd.DataFrame, col_map: dict) -> Optional[pd.Series]:
    """Série booleana de SLA cumprido (True), contada SOMENTE a partir de 'Em andamento'.

    Usa 'sla' ou 'vencido'. Chamados que ainda não iniciaram o SLA viram NaN
    (são excluídos do % de cumprimento).
    """
    if 'sla' in col_map:
        s = normalize_sla(df[col_map['sla']])
    elif 'vencido' in col_map:
        s = normalize_vencido(df[col_map['vencido']])
    else:
        return None
    # SLA só vale a partir de quando o chamado entrou em 'Em andamento'
    return s.where(sla_iniciado_mask(df, col_map))


# Rótulos canônicos de fase + cores para os gráficos
FASE_CORES = {
    'Finalizado': '#2ECC71',
    'Em Andamento': '#3498DB',
    'Aprovação de Solicitação': '#F39C12',
    'Acompanhamento': '#9B59B6',
    'Negado': '#E74C3C',
    'Outros': '#95A5A6',
}


def classify_fase(series: pd.Series) -> pd.Series:
    """Classifica a coluna 'Fase atual' em rótulos canônicos e separados."""
    norm = series.astype(str).apply(_normalize)

    def _map(v: str) -> str:
        if 'finaliz' in v or 'conclu' in v or 'resolv' in v or 'fecha' in v:
            return 'Finalizado'
        if 'andamento' in v:
            return 'Em Andamento'
        if 'aprovacao' in v:
            return 'Aprovação de Solicitação'
        if 'acompanhamento' in v:
            return 'Acompanhamento'
        if 'negad' in v:
            return 'Negado'
        return 'Outros'

    return norm.apply(_map)


def normalize_status(series: pd.Series) -> pd.Series:
    """Map status/phase column to 'Atendido' / 'Não Atendido'."""
    normalized = series.astype(str).apply(_normalize)

    def _map(val: str) -> str:
        if val in STATUS_DONE or any(kw in val for kw in
                                     ('resolv', 'fecha', 'encerr', 'conclu', 'finaliz', 'atend', 'done', 'ok')):
            return 'Atendido'
        return 'Não Atendido'

    return normalized.apply(_map)
