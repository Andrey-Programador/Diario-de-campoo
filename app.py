from pathlib import Path
import os
import re
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title='Dashboard ASRO | Demandas LIGHT', page_icon='💡', layout='wide')

st.markdown("""
<style>
:root{--deep:#087466;--teal:#13a892;--mint:#dff3ef;--ink:#092f34;--paper:#f5faf9;--blue:#dcecf9;--line:#dbe7e5;}
.stApp{background:var(--paper);color:var(--ink);}
[data-testid="stHeader"]{background:#ffffff;height:58px;border-bottom:1px solid #edf2f1;}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#086c60 0%,#0d8d7d 56%,#18a994 100%);border-right:0;}
[data-testid="stSidebar"] *{color:#fff!important;}
[data-testid="stSidebar"] [data-baseweb="select"] > div,
[data-testid="stSidebar"] input{background:rgba(255,255,255,.12)!important;border-color:rgba(255,255,255,.25)!important;}
[data-testid="stSidebar"] .stCaption{opacity:.86;}
.block-container{padding-top:2.4rem;padding-bottom:3rem;max-width:1500px;}
h1,h2,h3{color:var(--ink)!important;font-family:Inter,Segoe UI,Arial,sans-serif;letter-spacing:-.02em;}
h1{font-weight:800!important;font-size:2.2rem!important;}
[data-testid="stMetric"]{background:#fff;border:1px solid var(--line);padding:18px 20px;border-radius:14px;box-shadow:0 4px 14px rgba(6,83,75,.06);min-height:118px;}
[data-testid="stMetricLabel"]{color:#516d6b!important;font-weight:650;}
[data-testid="stMetricValue"]{color:var(--ink)!important;font-weight:800;}
.stTabs [data-baseweb="tab-list"]{gap:8px;border-bottom:1px solid var(--line);}
.stTabs [data-baseweb="tab"]{height:46px;border-radius:10px 10px 0 0;padding:0 18px;font-weight:700;color:#45625f;}
.stTabs [aria-selected="true"]{background:#e3f5f1!important;color:#087466!important;}
.stButton>button,.stDownloadButton>button{border:1px solid var(--teal)!important;color:#087466!important;background:#fff!important;border-radius:10px!important;font-weight:700!important;}
.stButton>button:hover,.stDownloadButton>button:hover{background:#e4f7f3!important;}
[data-testid="stDataFrame"]{background:#fff;border:1px solid var(--line);border-radius:12px;overflow:hidden;}
[data-testid="stAlert"]{border-radius:12px;border:0;background:var(--blue);color:#07559a;}
.asro-brand{padding:14px 4px 22px;border-bottom:1px solid rgba(255,255,255,.14);margin-bottom:18px;}
.asro-brand .small{font-size:17px;font-weight:700;letter-spacing:.3px;}
.asro-brand .logo{font-size:38px;font-weight:900;line-height:1;color:#fff;letter-spacing:-1.5px;margin-top:2px;}
.asro-brand .sub{font-size:12px;margin-top:10px;color:#bff3e9!important;}
.hero{background:#fff;border-top:8px solid var(--teal);border-radius:0 0 16px 16px;padding:26px 28px 22px;margin:-10px 0 22px;box-shadow:0 6px 20px rgba(5,87,78,.08);}
.hero .eyebrow{color:#0b8f7d;font-size:12px;font-weight:800;text-transform:uppercase;letter-spacing:.1em;}
.hero .title{color:var(--ink);font-size:31px;font-weight:850;line-height:1.15;margin-top:6px;}
.hero .desc{color:#597571;margin-top:8px;font-size:14px;}
.section-note{background:#dcecf9;border-radius:12px;padding:17px 20px;color:#07559a;margin-bottom:18px;line-height:1.55;}
.section-note b{display:block;margin-bottom:2px;}
</style>
""", unsafe_allow_html=True)

BASE_DIR = Path(__file__).resolve().parent

# Fonte local sincronizada pelo OneDrive/SharePoint no Windows.
# IMPORTANTE: esta versão deve ser executada no computador onde este caminho existe.
DATA_DIR = Path(r"C:\Users\AndreyLuizTrudesDomi\OneDrive - DEEPESSOAS\Diario-de-Campo")

ALIASES = {
    'agente': ['Nome do Agente'],
    'data_registro': ['Data do registro'],
    'hora_registro': ['Hora do registro'],
    'status_backoffice': ['Situação Backoffice - Visita', 'Situação Backoffice do Bloco'],
    'status_campo': ['Situação Campo'],
    'tipo_demanda': ['Tipo de demanda do cliente:'],
    'tema_reclamacao': ['Reclamação - Tema da demanda:'],
    'tema_reclamacao_outro': ['Reclamação - Tema da demanda (Outro):'],
    'orientacao': ['Orientação da equipe:'],
    'descricao_duvida': ['Descrição da dúvida do morador e o que foi falado para ele:'],
    'solicitacao_anterior': ['O morador já havia feito alguma solicitação em algum canal?'],
    'comentario_light': ['O que o morador comentou sobre os canais apresentados da LIGHT?'],
    'direcionamento': ['Direcionamento da solicitação:'],
    'codigo': ['Code Deep'],
    'link': ['Link Backoffice'],
}

def norm(s):
    return re.sub(r'\s+', ' ', str(s)).strip().casefold()

def find_col(df, candidates):
    lookup = {norm(c): c for c in df.columns}
    for c in candidates:
        if norm(c) in lookup:
            return lookup[norm(c)]
    return None

def first_nonempty(row, cols):
    for c in cols:
        v = row.get(c)
        if pd.notna(v) and str(v).strip() not in ('', 'nan', 'None'):
            return str(v).strip()
    return 'Não informado'

@st.cache_data(show_spinner=False)
def load_files(file_signatures):
    frames = []
    for path_str, _mtime in file_signatures:
        path = Path(path_str)
        try:
            df = pd.read_excel(path, sheet_name="Registros", engine="openpyxl")
        except ValueError:
            df = pd.read_excel(path, engine="openpyxl")
        df["_arquivo"] = path.name
        frames.append(df)
    if not frames:
        return pd.DataFrame()
    out = pd.concat(frames, ignore_index=True, sort=False)
    code_col = find_col(out, ALIASES["codigo"])
    link_col = find_col(out, ALIASES["link"])
    if code_col:
        filled = out[code_col].notna() & out[code_col].astype(str).str.strip().ne("")
        out = pd.concat([out[filled].drop_duplicates(subset=[code_col], keep="last"), out[~filled]], ignore_index=True)
    elif link_col:
        out = out.drop_duplicates(subset=[link_col], keep="last")
    else:
        out = out.drop_duplicates(keep="last")
    return out


def signatures():
    if not DATA_DIR.exists():
        return tuple()
    files = sorted(DATA_DIR.glob("*.xlsx"))
    return tuple((str(p), p.stat().st_mtime_ns) for p in files if not p.name.startswith("~$"))

def text_series(df, col, default='Não informado'):
    if not col:
        return pd.Series(default, index=df.index, dtype='object')
    s = df[col].fillna('').astype(str).str.strip()
    return s.mask(s.eq(''), default)

def contains_any(series, terms):
    pattern = '|'.join(re.escape(t) for t in terms)
    return series.fillna('').astype(str).str.contains(pattern, case=False, regex=True)

def resolve_date_col(df, keywords):
    for c in df.columns:
        n = norm(c)
        if all(k in n for k in keywords):
            return c
    return None

st.markdown("""<div class="hero"><div class="eyebrow">Light Controle · Deep Field</div><div class="title">Dashboard ASRO</div><div class="desc">Demandas de campo, encaminhamentos à LIGHT, retornos e acompanhamento operacional.</div></div>""", unsafe_allow_html=True)

sigs = signatures()
if not DATA_DIR.exists():
    st.error(f"Pasta não encontrada no computador: {DATA_DIR}")
    st.info("Confirme se a pasta Diario-de-Campo está sincronizada pelo OneDrive ou ajuste DATA_DIR no início do app.py.")
    st.stop()
if not sigs:
    st.error(f"Nenhum Excel .xlsx encontrado em: {DATA_DIR}")
    st.stop()

df = load_files(sigs)
arquivos_carregados = len(sigs)
fonte = "OneDrive/SharePoint sincronizado localmente"

if df.empty:
    st.warning("A fonte foi encontrada, mas não há registros para exibir.")
    st.stop()

cols = {k: find_col(df, v) for k, v in ALIASES.items()}
community_cols = [c for c in df.columns if ('comunidade' in norm(c) or norm(c) == 'asro:')]
df['_Comunidade'] = df.apply(lambda r: first_nonempty(r, community_cols), axis=1) if community_cols else 'Não informado'
df['_Agente'] = text_series(df, cols['agente'])
df['_Tipo'] = text_series(df, cols['tipo_demanda'])
df['_Tema'] = text_series(df, cols['tema_reclamacao'])
if cols['tema_reclamacao_outro']:
    outro = text_series(df, cols['tema_reclamacao_outro'], default='')
    df['_Tema'] = df['_Tema'].mask(df['_Tema'].eq('Não informado') & outro.ne(''), outro)
df['_Campo'] = text_series(df, cols['status_campo'])
df['_Backoffice'] = text_series(df, cols['status_backoffice'])
df['_Direcionamento'] = text_series(df, cols['direcionamento'], default='')

if cols['data_registro']:
    df['_Data'] = pd.to_datetime(df[cols['data_registro']], errors='coerce', dayfirst=True)
else:
    df['_Data'] = pd.NaT

# Regras transparentes. Ajuste os termos abaixo se a exportação usar outra nomenclatura.
contexto = (df['_Campo'] + ' | ' + df['_Backoffice'] + ' | ' + df['_Direcionamento']).str.lower()
df['_ResolvidoCampo'] = contains_any(df['_Campo'], ['resolvid', 'conclu', 'finaliz'])
df['_EnviadoLight'] = contains_any(contexto, ['light', 'encaminh', 'enviado'])
df['_ResolvidoLight'] = df['_EnviadoLight'] & contains_any(df['_Backoffice'], ['resolvid', 'conclu', 'finaliz'])
df['_Pendente'] = contains_any(contexto, ['pendente', 'aguard', 'aberto']) & ~df['_ResolvidoLight']

# Procura datas específicas sem presumir que Data do registro seja data LIGHT.
data_envio_col = (resolve_date_col(df, ['data', 'envio', 'light']) or
                  resolve_date_col(df, ['data', 'encaminh', 'light']))
data_solucao_col = (resolve_date_col(df, ['data', 'solu', 'light']) or
                    resolve_date_col(df, ['data', 'retorno', 'light']) or
                    resolve_date_col(df, ['data', 'resol', 'light']))
if data_envio_col:
    df['_DataEnvioLight'] = pd.to_datetime(df[data_envio_col], errors='coerce', dayfirst=True)
else:
    df['_DataEnvioLight'] = pd.NaT
if data_solucao_col:
    df['_DataSolucaoLight'] = pd.to_datetime(df[data_solucao_col], errors='coerce', dayfirst=True)
else:
    df['_DataSolucaoLight'] = pd.NaT

with st.sidebar:
    st.markdown("""<div class="asro-brand"><div class="small">Light</div><div class="logo">Controle.</div><div class="sub">Deep Field<br>Automação e Inteligência Operacional</div></div>""", unsafe_allow_html=True)
    st.markdown("### Filtros do painel")
    st.success('OneDrive local conectado')
    st.caption(f'Fonte: {DATA_DIR}')
    agentes = st.multiselect('Agente', sorted(df['_Agente'].dropna().unique()))
    comunidades = st.multiselect('Comunidade ASRO', sorted(df['_Comunidade'].dropna().unique()))
    tipos = st.multiselect('Tipo de demanda', sorted(df['_Tipo'].dropna().unique()))
    temas = st.multiselect('Tema da reclamação', sorted(df['_Tema'].dropna().unique()))
    if df['_Data'].notna().any():
        dmin, dmax = df['_Data'].min().date(), df['_Data'].max().date()
        periodo = st.date_input('Período do registro', value=(dmin, dmax), min_value=dmin, max_value=dmax)
    else:
        periodo = None
    st.divider()
    st.caption(f'{arquivos_carregados} arquivo(s) carregado(s) | {len(df):,} registro(s) únicos'.replace(',', '.'))

f = df.copy()
if agentes: f = f[f['_Agente'].isin(agentes)]
if comunidades: f = f[f['_Comunidade'].isin(comunidades)]
if tipos: f = f[f['_Tipo'].isin(tipos)]
if temas: f = f[f['_Tema'].isin(temas)]
if periodo and isinstance(periodo, (tuple, list)) and len(periodo) == 2:
    ini, fim = pd.Timestamp(periodo[0]), pd.Timestamp(periodo[1]) + pd.Timedelta(days=1)
    f = f[f['_Data'].between(ini, fim, inclusive='left')]
recebidas = len(f)
res_campo = int(f['_ResolvidoCampo'].sum())
enviadas = int(f['_EnviadoLight'].sum())
res_light = int(f['_ResolvidoLight'].sum())
pendentes = int(f['_Pendente'].sum())

# Métricas de tempo: só calculadas quando existe uma data explícita de solução/retorno LIGHT.
# TMA usa Data do registro -> Data de solução LIGHT. Não inferimos datas ausentes.
tma_horas = None
mediana_horas = None
atendidos_com_tempo = 0
if data_solucao_col and f['_Data'].notna().any():
    tempos = (f['_DataSolucaoLight'] - f['_Data']).dt.total_seconds() / 3600
    tempos = tempos[(tempos >= 0) & tempos.notna()]
    if not tempos.empty:
        tma_horas = float(tempos.mean())
        mediana_horas = float(tempos.median())
        atendidos_com_tempo = int(tempos.count())

def fmt_tempo(horas):
    if horas is None:
        return 'N/D'
    if horas < 24:
        return f'{horas:.1f} h'.replace('.', ',')
    return f'{horas/24:.1f} dias'.replace('.', ',')

m1,m2,m3,m4,m5 = st.columns(5)
m1.metric('📥 Casos recebidos', recebidas)
m2.metric('✅ Resolvidos em campo', res_campo)
m3.metric('📤 Enviados à LIGHT', enviadas)
m4.metric('💡 Resolvidos pela LIGHT', res_light)
m5.metric('⏳ Pendentes', pendentes)

t1,t2,t3 = st.columns(3)
t1.metric('⏱️ TMA geral', fmt_tempo(tma_horas), help='Média entre Data do registro e Data de solução/retorno LIGHT, quando disponível.')
t2.metric('🕒 Tempo mediano', fmt_tempo(mediana_horas), help='Mediana do tempo entre registro e solução/retorno LIGHT.')
t3.metric('📌 Casos com tempo calculável', atendidos_com_tempo, help='Casos que possuem Data do registro e Data de solução/retorno LIGHT válidas.')

if not data_envio_col or not data_solucao_col:
    faltam = []
    if not data_envio_col: faltam.append('data de envio à LIGHT')
    if not data_solucao_col: faltam.append('data de solução/retorno da LIGHT')
    st.info('Datas LIGHT: não encontrei coluna explícita para ' + ' e '.join(faltam) + '. O app não utiliza “Data do registro” como substituta para evitar informação incorreta.')

st.markdown("""<div class="section-note"><b>Visão operacional</b>Use os filtros laterais para analisar período, agentes, comunidades e temas. Os cards incluem volume, TMA e tempo mediano quando houver data real de retorno da LIGHT.</div>""", unsafe_allow_html=True)

aba1, aba2, aba3, aba4 = st.tabs(['📊 Visão geral','💡 Controle LIGHT','👥 Agentes e comunidades','🔎 Detalhamento'])

with aba1:
    c1,c2 = st.columns(2)
    with c1:
        status_df = pd.DataFrame({'Status':['Resolvido em campo','Enviado à LIGHT','Resolvido pela LIGHT','Pendente'], 'Casos':[res_campo,enviadas,res_light,pendentes]})
        st.plotly_chart(px.bar(status_df, x='Status', y='Casos', text_auto=True, title='Fluxo das demandas'), use_container_width=True)
    with c2:
        tema = f['_Tema'].value_counts().head(10).reset_index(); tema.columns=['Tema','Casos']
        st.plotly_chart(px.bar(tema, x='Casos', y='Tema', orientation='h', text_auto=True, title='Top 10 temas de reclamação'), use_container_width=True)
    if f['_Data'].notna().any():
        diario = f.dropna(subset=['_Data']).assign(Dia=lambda x:x['_Data'].dt.date).groupby('Dia').size().reset_index(name='Casos')
        st.plotly_chart(px.line(diario, x='Dia', y='Casos', markers=True, title='Casos recebidos por dia'), use_container_width=True)

with aba2:
    light = f[f['_EnviadoLight']].copy()
    l1,l2,l3 = st.columns(3)
    l1.metric('Enviados', len(light)); l2.metric('Resolvidos', int(light['_ResolvidoLight'].sum())); l3.metric('Pendentes', int(light['_Pendente'].sum()))
    if data_envio_col and data_solucao_col:
        light['_DiasSolucao'] = (light['_DataSolucaoLight'] - light['_DataEnvioLight']).dt.total_seconds()/86400
        valid = light.loc[light['_DiasSolucao'] >= 0, '_DiasSolucao'].dropna()
        if not valid.empty:
            st.metric('⏱️ Tempo médio LIGHT (envio → solução)', f'{valid.mean():.1f} dias'.replace('.', ','))
            faixas = pd.cut(valid, bins=[-0.001,1,2,3,5,float('inf')], labels=['Até 1 dia','1–2 dias','2–3 dias','3–5 dias','Mais de 5 dias'])
            dist = faixas.value_counts(sort=False).reset_index(); dist.columns=['Faixa','Casos']
            st.plotly_chart(px.bar(dist, x='Faixa', y='Casos', text_auto=True, title='Distribuição do tempo de atendimento LIGHT'), use_container_width=True)
    display = pd.DataFrame({
        'Agente': light['_Agente'], 'Comunidade': light['_Comunidade'], 'Tema': light['_Tema'],
        'Data registro': light['_Data'].dt.strftime('%d/%m/%Y'),
        'Data envio LIGHT': light['_DataEnvioLight'].dt.strftime('%d/%m/%Y'),
        'Data solução LIGHT': light['_DataSolucaoLight'].dt.strftime('%d/%m/%Y'),
        'Status campo': light['_Campo'], 'Status backoffice': light['_Backoffice']
    })
    st.dataframe(display, use_container_width=True, hide_index=True)

with aba3:
    st.caption('Visão operacional por volume/status. Os dados não são usados para pontuar, ranquear ou avaliar desempenho individual.')
    c1,c2 = st.columns(2)
    ag = f.groupby('_Agente').agg(Recebidas=('_Agente','size'), Resolvidas_campo=('_ResolvidoCampo','sum'), Enviadas_LIGHT=('_EnviadoLight','sum'), Resolvidas_LIGHT=('_ResolvidoLight','sum'), Pendentes=('_Pendente','sum')).reset_index().rename(columns={'_Agente':'Agente'})
    co = f.groupby('_Comunidade').agg(Recebidas=('_Comunidade','size'), Resolvidas_campo=('_ResolvidoCampo','sum'), Enviadas_LIGHT=('_EnviadoLight','sum'), Resolvidas_LIGHT=('_ResolvidoLight','sum'), Pendentes=('_Pendente','sum')).reset_index().rename(columns={'_Comunidade':'Comunidade'})
    with c1:
        st.subheader('Por agente'); st.dataframe(ag, use_container_width=True, hide_index=True)
    with c2:
        st.subheader('Por comunidade'); st.dataframe(co, use_container_width=True, hide_index=True)

with aba4:
    wanted = [
        ('Agente','_Agente'), ('Comunidade','_Comunidade'), ('Data do registro','_Data'),
        ('Reclamação - Tema da demanda', cols['tema_reclamacao']),
        ('Reclamação - Tema da demanda (Outro)', cols['tema_reclamacao_outro']),
        ('Orientação da equipe', cols['orientacao']),
        ('Descrição da dúvida do morador e o que foi falado para ele', cols['descricao_duvida']),
        ('O morador já havia feito alguma solicitação em algum canal?', cols['solicitacao_anterior']),
        ('O que o morador comentou sobre os canais apresentados da LIGHT?', cols['comentario_light']),
    ]
    detail = pd.DataFrame(index=f.index)
    for label, col in wanted:
        if col and col in f.columns:
            detail[label] = f[col]
    if 'Data do registro' in detail:
        detail['Data do registro'] = pd.to_datetime(detail['Data do registro'], errors='coerce').dt.strftime('%d/%m/%Y')
    st.dataframe(detail, use_container_width=True, hide_index=True, height=520)
    st.download_button('⬇️ Baixar dados filtrados (CSV)', detail.to_csv(index=False).encode('utf-8-sig'), 'dashboard_asro_filtrado.csv', 'text/csv')

with st.expander('⚙️ Diagnóstico das colunas'):
    st.write('Fonte:', fonte)
    st.write('Pasta:', str(DATA_DIR))
    st.write('Arquivo(s):', [Path(p).name for p, _ in sigs])
    st.write('Data de envio LIGHT reconhecida:', data_envio_col or 'não encontrada')
    st.write('Data de solução LIGHT reconhecida:', data_solucao_col or 'não encontrada')
    st.write('Colunas disponíveis:', list(df.columns))
