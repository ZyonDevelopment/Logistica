# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re, warnings
from pathlib import Path
warnings.filterwarnings("ignore")

BASE_DIR = Path(__file__).resolve().parent
LOGISTICA_CSV = BASE_DIR / "9720575575b77630c182a6ddcfc0e90a11526c7d.csv"
CONSULTA_PEDIDOS_XLSX = BASE_DIR / "ConsultaPedidos_89ac3e7f-87ef-4bc8-99b9-3996ab119944.xlsx"

st.set_page_config(
    page_title="Dashboard Logística",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .fonte-badge {
        display:inline-block; font-size:.7rem; font-weight:700;
        padding:2px 10px; border-radius:20px; margin-bottom:6px;
    }
    .fonte-log  { background:#0f2944; color:#60a5fa; border:1px solid #1e4a80; }
    .fonte-ped  { background:#0f2e1a; color:#4ade80; border:1px solid #1a5c30; }
    .fonte-both { background:#2a0f44; color:#c084fc; border:1px solid #5b21b6; }
    .secao {
        font-size:1.05rem; font-weight:700; color:#f1f5f9;
        margin:28px 0 6px 0; padding-bottom:7px;
        border-bottom:2px solid #334155; letter-spacing:.03em;
    }
    .narrativa {
        background:#1e293b; border-left:4px solid #6366f1;
        padding:11px 16px; border-radius:0 8px 8px 0;
        color:#cbd5e1; font-size:.88rem; margin-bottom:12px; line-height:1.65;
    }
    .narrativa-warn {
        background:#1e293b; border-left:4px solid #f59e0b;
        padding:11px 16px; border-radius:0 8px 8px 0;
        color:#cbd5e1; font-size:.88rem; margin-bottom:12px; line-height:1.65;
    }
    div[data-testid="metric-container"] {
        background:#1e293b; border-radius:10px;
        padding:12px 16px; border-left:3px solid #6366f1;
    }
    .js-plotly-plot .annotation rect {
        rx: 6px;
        ry: 6px;
    }
</style>
""", unsafe_allow_html=True)

BADGE_LOG  = '<span class="fonte-badge fonte-log">📦 Fonte: Logística</span>'
BADGE_PED  = '<span class="fonte-badge fonte-ped">🛒 Fonte: Pedidos Comerciais</span>'
BADGE_BOTH = '<span class="fonte-badge fonte-both">🔗 Logística + Pedidos</span>'

COR_OK    = "#22c55e"
COR_WARN  = "#f59e0b"
COR_ERR   = "#ef4444"
COR_NEU   = "#6366f1"
COR_CINZA = "#475569"
TODOS_PDVS = ["PDV 22908","PDV 23868","PDV 24117","PDV 24118","PDV 24239","PDV 24270","PDV 24341"]

ORDEM_STACK = [
    "Entregue","No cliente",
    "Aguardando geracao de rota","Aguardando motorista","Em transito",
    "Nao visitado","Destinatario ausente","Estabelecimento fechado",
    "Endereco nao localizado","Destinatario mudou de endereco",
    "Carga recusada pelo destinatario","Devolvido",
]
STATUS_CORES = {
    "Entregue":"#22c55e","No cliente":"#34d399",
    "Em transito":"#6366f1","Aguardando motorista":"#818cf8",
    "Aguardando geracao de rota":"#475569",
    "Destinatario ausente":"#fb923c","Nao visitado":"#f59e0b",
    "Estabelecimento fechado":"#f97316","Endereco nao localizado":"#c084fc",
    "Destinatario mudou de endereco":"#60a5fa",
    "Carga recusada pelo destinatario":"#f43f5e","Devolvido":"#ef4444",
}

CUSTO_TOTAL_MOTORISTAS = 54_297.00
META_CUSTO_MEDIO = 21.00
METAS_ENTREGUES = {
    "Interna": 98.0,
    "Plataforma Logistica": 90.0,
    "Omnichannel": 100.0,
}
METAS_PRAZO = {
    "Interna": 98.0,
    "Plataforma Logistica": 90.0,
    "Omnichannel": 98.0,
}

# ── Formatação Brasileira ──────────────────────────────────────────────────────
def fmt_num(v, dec=0):
    s = f"{float(v):,.{dec}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_pct(v, dec=2):
    return f"{float(v):.{dec}f}".replace(".", ",") + "%"

def fmt_brl(v):
    return "R$ " + fmt_num(v, 2)

def meta_canal(canal_sel, tipo="entrega"):
    metas = METAS_ENTREGUES if tipo == "entrega" else METAS_PRAZO
    if canal_sel == "Omnichannel":
        return metas["Omnichannel"], "Omnichannel"
    if canal_sel == "Logístico":
        return metas["Plataforma Logistica"], "Plataforma Logistica"
    return metas["Interna"], "Interna"

def status_meta_pct(valor, meta):
    if valor >= meta:
        return "OK", COR_OK
    if valor >= meta * 0.95:
        return "Atenção", COR_WARN
    return "Crítico", COR_ERR

def status_custo(valor):
    return ("OK", COR_OK) if valor < META_CUSTO_MEDIO else ("Preocupante", COR_WARN)

def cor_por_meta(valor, meta):
    return COR_OK if valor >= meta else COR_ERR

def texto_cor_por_meta(valor, meta):
    return "#f8fafc" if valor >= meta else "#fee2e2"

def estilo_pct_meta(v, meta):
    try:
        valor = float(v)
    except Exception:
        return ""
    if valor >= meta:
        return "background-color: rgba(34,197,94,.22); color: #dcfce7; font-weight: 700;"
    return "background-color: rgba(239,68,68,.34); color: #fee2e2; font-weight: 700;"

def eh_itabuna(v):
    return str(v or "").strip().lower() == "itabuna"

def meta_tempo_por_municipio(municipio, *, tipo):
    if tipo == "tat":
        return 2 if eh_itabuna(municipio) else 3
    return 1 if eh_itabuna(municipio) else 2

# ── Helpers ────────────────────────────────────────────────────────────────────
def extrair_pdv(v):
    if pd.isna(v): return "Desconhecido"
    m = re.match(r"(\d{5})", str(v).strip())
    return f"PDV {m.group(1)}" if m else str(v)

def find_col(df, *kws):
    for kw in kws:
        for c in df.columns:
            try:
                if kw.lower() in c.lower(): return c
            except: pass
    return None

def ordenar_stack(df_p, col_cat, col_status, col_val):
    ordem_cat = (df_p.groupby(col_cat)[col_val].sum()
                     .sort_values(ascending=False).index.tolist())
    df_p[col_cat] = pd.Categorical(df_p[col_cat], categories=ordem_cat, ordered=True)
    status_atuais = df_p[col_status].dropna().unique().tolist()
    presentes = [s for s in ORDEM_STACK if s in status_atuais]
    presentes += [s for s in status_atuais if s not in presentes]
    df_p[col_status] = pd.Categorical(df_p[col_status], categories=presentes, ordered=True)
    return df_p.sort_values([col_cat, col_status])

def layout_br(**kwargs):
    base = dict(separators=",.", font=dict(family="sans-serif"))
    base.update(kwargs)
    return base

def dias_entre(inicio, fim):
    dias = (fim - inicio).dt.total_seconds() / 86400
    return dias.where(dias >= 0)

def add_hbar_label_boxes(fig, df_plot, *, y_col, x_col, text_col, base_col=None, xanchor="right"):
    for _, row in df_plot.iterrows():
        valor = row.get(x_col, 0)
        texto = str(row.get(text_col, "") or "")
        if pd.isna(valor) or valor <= 0 or not texto:
            continue
        base = row.get(base_col, 0) if base_col else 0
        fig.add_annotation(
            x=base + valor,
            y=row[y_col],
            text=texto,
            showarrow=False,
            xanchor=xanchor,
            yanchor="middle",
            font=dict(color="#f8fafc", size=13),
            bgcolor="rgba(51,65,85,0.96)",
            bordercolor="rgba(148,163,184,0.85)",
            borderwidth=1,
            borderpad=4,
        )

def add_vbar_label_boxes(fig, df_plot, *, x_col, y_col, text_col):
    ymax = pd.to_numeric(df_plot[y_col], errors="coerce").max()
    offset = (ymax or 1) * 0.025
    for _, row in df_plot.iterrows():
        valor = row.get(y_col, 0)
        texto = str(row.get(text_col, "") or "")
        if pd.isna(valor) or valor <= 0 or not texto:
            continue
        fig.add_annotation(
            x=row[x_col],
            y=valor + offset,
            text=texto,
            showarrow=False,
            xanchor="center",
            yanchor="bottom",
            font=dict(color="#f8fafc", size=13),
            bgcolor="rgba(51,65,85,0.96)",
            bordercolor="rgba(148,163,184,0.85)",
            borderwidth=1,
            borderpad=4,
        )
    fig.update_yaxes(range=[0, (ymax or 1) * 1.14])

def add_stacked_hbar_label_boxes(fig, df_plot, *, group_col, stack_col, value_col, text_col, stack_order=None):
    work = df_plot.copy()
    order = stack_order or work[stack_col].dropna().unique().tolist()
    work["_stack_ordem"] = work[stack_col].apply(lambda x: order.index(x) if x in order else len(order))
    work = work.sort_values([group_col, "_stack_ordem"])
    work["_base"] = work.groupby(group_col)[value_col].cumsum() - work[value_col]
    add_hbar_label_boxes(fig, work, y_col=group_col, x_col=value_col, text_col=text_col, base_col="_base")
    fig.update_traces(text=None)

def preparar_resumo_tempo(df_base, *, tempo_col, meta_tipo, visao):
    base = df_base.copy()
    if tempo_col not in base.columns:
        return pd.DataFrame()
    base = base[base[tempo_col].between(0, 30)].copy()
    if base.empty:
        return pd.DataFrame()
    base["Meta"] = base["Municipio"].apply(lambda x: meta_tempo_por_municipio(x, tipo=meta_tipo))
    base["DentroMeta"] = base[tempo_col] <= base["Meta"]
    base["Rota"] = base["Municipio"].apply(lambda x: "Local" if eh_itabuna(x) else "Fora")
    grupo = "Municipio" if visao == "Cidade" else "PDV"
    resumo = (
        base.groupby(grupo)
        .agg(
            Media=(tempo_col, "mean"),
            MetaMedia=("Meta", "mean"),
            DentroMeta=("DentroMeta", "mean"),
            Pedidos=("Pedido", "count"),
            Rota=("Rota", lambda x: "Local" if (x == "Local").all() else ("Fora" if (x == "Fora").all() else "Misto")),
        )
        .reset_index()
        .rename(columns={grupo: "Grupo"})
    )
    resumo["Media"] = resumo["Media"].round(2)
    resumo["MetaMedia"] = resumo["MetaMedia"].round(2)
    resumo["PctMeta"] = (resumo["DentroMeta"] * 100).round(1)
    resumo["Cor"] = resumo.apply(lambda r: COR_OK if r["Media"] <= r["MetaMedia"] else COR_ERR, axis=1)
    resumo["Rotulo"] = resumo.apply(
        lambda r: f"{fmt_num(r['Media'],1)}d | meta {fmt_num(r['MetaMedia'],1)}d | {fmt_pct(r['PctMeta'],1)}",
        axis=1,
    )
    return resumo.sort_values("Media", ascending=True)

def render_tempo_bar(resumo, *, titulo):
    if resumo.empty:
        st.info(f"Sem dados para {titulo}.")
        return
    fig = go.Figure(go.Bar(
        x=resumo["Media"],
        y=resumo["Grupo"],
        orientation="h",
        marker_color=resumo["Cor"],
        customdata=resumo[["MetaMedia", "PctMeta", "Pedidos", "Rota"]],
        hovertemplate=(
            "%{y}<br>Média: %{x:.2f} dias<br>"
            "Meta média: %{customdata[0]:.2f} dias<br>"
            "Dentro da meta: %{customdata[1]:.1f}%<br>"
            "Pedidos: %{customdata[2]}<br>Rota: %{customdata[3]}<extra></extra>"
        ),
    ))
    add_hbar_label_boxes(fig, resumo, y_col="Grupo", x_col="Media", text_col="Rotulo", xanchor="left")
    fig.update_layout(**layout_br(
        title=titulo,
        height=max(360, 34 * len(resumo) + 110),
        margin=dict(t=55,b=10,l=10,r=170),
        xaxis_title="Dias",
        yaxis_title="",
    ))
    st.plotly_chart(fig, use_container_width=True)

def _serie_limpa(s, vazio="Sem ocorrencia/mensagem registrada"):
    out = s.fillna("").astype(str).str.strip()
    return out.where(out.str.len() > 0, vazio)

def preparar_excecoes(df_base, *, pendente=False):
    out = df_base.copy()
    hoje = pd.Timestamp.today().normalize()
    out["Criacao"] = out["_criacao"].dt.strftime("%d/%m/%Y")
    out["Prazo"] = out["_prazo"].dt.strftime("%d/%m/%Y")
    out["Entrega"] = out["_efetuada"].dt.strftime("%d/%m/%Y %H:%M")
    out["Motorista"] = _serie_limpa(out["Motorista"], "Sem motorista atribuido")
    out["Ultima ocorrencia"] = _serie_limpa(out["OcorrStatus"], "Sem ocorrencia registrada")
    msg = _serie_limpa(out["OcorrMensagem"], "")
    detalhe = _serie_limpa(out["StatusDetalhe"], "")
    out["Mensagem / detalhe"] = msg.where(msg.str.len() > 0, detalhe)
    out["Mensagem / detalhe"] = _serie_limpa(out["Mensagem / detalhe"])
    if pendente:
        out["Dias"] = (hoje - out["_prazo"].dt.normalize()).dt.days
        out["Dias"] = out["Dias"].where(out["Dias"] > 0, 0)
    else:
        out["Dias"] = (out["_efetuada"].dt.normalize() - out["_prazo"].dt.normalize()).dt.days
    out["Dias"] = out["Dias"].fillna(0).astype(int)
    return out

def render_excecoes_pivot(df_ex, *, titulo, dias_label, key_prefix, incluir_status=False):
    st.markdown(f"**{titulo}**")
    if df_ex.empty:
        st.info("Sem pedidos para os filtros selecionados.")
        return

    resumo = (
        df_ex.groupby("Ultima ocorrencia", dropna=False)
             .agg(Pedidos=("Pedido", "count"), MaiorDias=("Dias", "max"))
             .reset_index()
             .sort_values(["Pedidos", "MaiorDias"], ascending=[False, False])
    )
    resumo = resumo.rename(columns={"MaiorDias": f"Maior {dias_label.lower()}"})

    c_resumo, c_detalhe = st.columns([1.05, 1.45])

    with c_resumo:
        st.caption("Resumo por ocorrencia")
        resumo_event = st.dataframe(
            resumo.reset_index(drop=True),
            use_container_width=True,
            height=520,
            on_select="rerun",
            selection_mode="single-row",
            key=f"{key_prefix}_resumo",
        )
        selected_rows = resumo_event.selection.rows if resumo_event and resumo_event.selection else []
        selected_idx = selected_rows[0] if selected_rows else 0
        aberta = str(resumo.iloc[selected_idx]["Ultima ocorrencia"])

    with c_detalhe:
        st.caption(f"Detalhes filtrados: {aberta}")
        detalhe_df = df_ex[df_ex["Ultima ocorrencia"].astype(str) == aberta].copy()
        if detalhe_df.empty:
            st.info("Sem pedidos para a ocorrencia selecionada.")
            return

        detalhe_cols = [
            "Mensagem / detalhe", "PDV", "Motorista", "Municipio",
            "Pedido", "Criacao", "Prazo", "Entrega", "Dias", "QtdOcorr"
        ]
        if incluir_status:
            detalhe_cols.insert(4, "Status")
        detalhe = (
            detalhe_df[detalhe_cols]
            .rename(columns={"Dias": dias_label, "QtdOcorr": "Qtd ocorr."})
            .sort_values(["Mensagem / detalhe", "PDV", "Motorista", "Municipio", dias_label],
                         ascending=[True, True, True, True, False])
        )
        st.dataframe(detalhe.reset_index(drop=True), use_container_width=True, height=520)

# ── Carga ──────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Carregando dados...")
def carregar():
    df_log = pd.read_csv(
        LOGISTICA_CSV,
        encoding="utf-8-sig",
    )
    for c in df_log.columns:
        if "criação" in c or "Data de c" in c:
            df_log["_criacao"] = pd.to_datetime(df_log[c], errors="coerce"); break
    for c in df_log.columns:
        if "hora efetuada" in c:
            df_log["_efetuada"] = pd.to_datetime(df_log[c], errors="coerce"); break
    for c in df_log.columns:
        if "aprovação" in c or "aprovacao" in c:
            df_log["_aprov"] = pd.to_datetime(df_log[c], errors="coerce"); break
    for c in df_log.columns:
        if "Data de Coleta" in c or "coleta" in c.lower():
            df_log["_coleta"] = pd.to_datetime(df_log[c], errors="coerce"); break
    for c in df_log.columns:
        if "Prazo Cliente" in c:
            df_log["_prazo"] = pd.to_datetime(df_log[c], errors="coerce"); break

    if "_aprov" not in df_log.columns:
        df_log["_aprov"] = df_log["_criacao"]
    if "_coleta" not in df_log.columns:
        df_log["_coleta"] = pd.NaT

    tat_raw = (df_log["_efetuada"] - df_log["_aprov"]).dt.total_seconds() / 86400
    df_log["TAT_invalido"] = tat_raw < 0
    df_log["TAT_dias"]  = tat_raw.where(tat_raw >= 0)
    df_log["NoPrazo"]   = (df_log["Status"] == "Entregue") & (df_log["_efetuada"].dt.date <= df_log["_prazo"].dt.date)
    df_log["PDV"]       = df_log["Expedidor"].apply(extrair_pdv)
    df_log["Pedido"]    = df_log["Pedido"].astype(str).str.strip()
    col_mun = find_col(df_log, "Munic", "unicipio")
    df_log["Municipio"] = df_log[col_mun].str.title() if col_mun else ""
    col_oc = find_col(df_log, "Ocorr")
    df_log["TemOcorr"]  = df_log[col_oc].str.strip().str.lower() == "sim" if col_oc else False
    col_oc_status = find_col(df_log, "ltima Ocorr", "Ultima Ocorr")
    col_oc_msg = find_col(df_log, "Mensagem")
    col_oc_qtd = find_col(df_log, "Quantidade de Ocorr")
    col_det = find_col(df_log, "Status (detalhe)", "detalhe")
    df_log["OcorrStatus"]   = df_log[col_oc_status].fillna("").astype(str).str.strip() if col_oc_status else ""
    df_log["OcorrMensagem"] = df_log[col_oc_msg].fillna("").astype(str).str.strip() if col_oc_msg else ""
    df_log["QtdOcorr"]      = pd.to_numeric(df_log[col_oc_qtd], errors="coerce").fillna(0).astype(int) if col_oc_qtd else 0
    df_log["StatusDetalhe"] = df_log[col_det].fillna("").astype(str).str.strip() if col_det else ""
    col_vf = find_col(df_log, "Valor do frete")
    df_log["_frete"]    = df_log[col_vf] if col_vf else 0.0

    # Atribui Alex Pereira a pedidos com Pedido ERP e sem motorista
    col_erp = find_col(df_log, "Pedido ERP", "PedidoERP")
    if col_erp:
        sem_mot = df_log["Motorista"].isna()
        tem_erp = df_log[col_erp].notna()
        df_log.loc[sem_mot & tem_erp, "Motorista"] = "alex Pereira de Jesus"

    df_ped = pd.read_excel(
        CONSULTA_PEDIDOS_XLSX,
        engine="openpyxl",
    )
    df_ped["CodigoPedido"] = df_ped["CodigoPedido"].astype(str).str.strip()
    for c in ["DataFaturamento", "Data Aprovação", "DataEntrega"]:
        if c in df_ped.columns:
            df_ped[c] = pd.to_datetime(df_ped[c], errors="coerce", dayfirst=True)
    col_canal = find_col(df_ped, "CanalDistribuicao","Canal")
    if col_canal:
        df_ped["PDV_Ped"] = df_ped[col_canal].apply(extrair_pdv)

    # ── Classificação TipoCanal ────────────────────────────────────────────────
    # Omnichannel  = MeioCaptacao contém "Omni" (prioridade)
    # Logístico    = CodigoPedido tem correspondência no CSV logístico
    # Não Logístico = demais
    ids_log = set(df_log["Pedido"].astype(str).str.strip())
    col_meio_ped = find_col(df_ped, "MeioCaptacao", "eioCaptacao")

    def classificar_canal(row):
        meio = str(row.get(col_meio_ped, "") or "").strip()
        if "omni" in meio.lower():
            return "Omnichannel"
        if row["CodigoPedido"] in ids_log:
            return "Logístico"
        return "Não Logístico"

    df_ped["TipoCanal"] = df_ped.apply(classificar_canal, axis=1)

    df_join = df_log.merge(df_ped, left_on="Pedido", right_on="CodigoPedido",
                           how="left", suffixes=("","_ped"))
    if "DataFaturamento" in df_join.columns:
        df_join["_faturamento"] = pd.to_datetime(df_join["DataFaturamento"], errors="coerce")
    else:
        df_join["_faturamento"] = pd.NaT
    tp_raw = (df_join["_faturamento"] - df_join["_aprov"]).dt.total_seconds() / 86400
    te_raw = (df_join["_efetuada"] - df_join["_coleta"]).dt.total_seconds() / 86400
    df_join["TP_invalido"] = tp_raw < 0
    df_join["TE_invalido"] = te_raw < 0
    df_join["TP_dias"] = tp_raw.where(tp_raw >= 0)
    df_join["TE_dias"] = te_raw.where(te_raw >= 0)

    return df_log, df_ped, df_join, {
        "sit":    find_col(df_ped, "SituacaoComercial","ituacaoComercial"),
        "meio":   col_meio_ped,
        "tp_ent": find_col(df_ped, "Tipo de Entrega"),
        "valor":  find_col(df_ped, "ValorPedido"),
        "ciclo":  find_col(df_ped, "CicloIndicador"),
        "canal":  col_canal,
    }

df_log, df_ped, df_join, COLS = carregar()

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.title("Filtros")
    dmin = df_log["_criacao"].dropna().dt.date.min()
    dmax = df_log["_criacao"].dropna().dt.date.max()
    periodo = st.date_input("Período (criação)", value=(dmin, dmax), min_value=dmin, max_value=dmax)
    dt_ini = pd.Timestamp(periodo[0] if isinstance(periodo,(list,tuple)) else dmin)
    dt_fim = pd.Timestamp(periodo[1] if isinstance(periodo,(list,tuple)) and len(periodo)==2 else dmax)

    pdv_opts = ["Todos"] + sorted(df_log["PDV"].unique().tolist())
    pdv_sel  = st.selectbox("PDV (Expedidor)", pdv_opts)

    # ── Filtro Canal ──────────────────────────────────────────────────────────
    canal_opts = ["Todos", "Logístico", "Não Logístico", "Omnichannel"]
    canal_sel  = st.selectbox(
        "Canal / Tipo de Pedido",
        canal_opts,
        help=(
            "Logístico = pedidos que geraram expedição logística\n"
            "Não Logístico = pedidos comerciais sem movimentação logística\n"
            "Omnichannel = MeioCaptacao 'Pedido Omnichannel'"
        ),
    )

    st.markdown("---")
    st.markdown("**Fontes de dados**")
    st.caption(f"📦 Logística: {fmt_num(len(df_log))} pedidos")
    st.caption(f"🛒 Pedidos comerciais: {fmt_num(len(df_ped))} registros")

    cnt_log  = (df_ped["TipoCanal"] == "Logístico").sum()
    cnt_nlog = (df_ped["TipoCanal"] == "Não Logístico").sum()
    cnt_omni = (df_ped["TipoCanal"] == "Omnichannel").sum()
    st.markdown("**Composição dos canais (total):**")
    st.caption(f"• Logístico: {fmt_num(cnt_log)}")
    st.caption(f"• Não Logístico: {fmt_num(cnt_nlog)}")
    st.caption(f"• Omnichannel: {fmt_num(cnt_omni)}")

# ── Aplicar filtros ────────────────────────────────────────────────────────────
# Logística (CSV) — data + PDV
mask = (df_log["_criacao"] >= dt_ini) & (df_log["_criacao"] <= dt_fim + pd.Timedelta(days=1))
if pdv_sel != "Todos": mask &= df_log["PDV"] == pdv_sel
df = df_log[mask].copy()

# Filtro canal na logística
if canal_sel == "Não Logístico":
    df = df.iloc[0:0].copy()           # não logístico nunca tem registro no CSV
elif canal_sel == "Omnichannel":
    ids_omni = set(df_ped[df_ped["TipoCanal"] == "Omnichannel"]["CodigoPedido"])
    df = df[df["Pedido"].isin(ids_omni)].copy()
# "Logístico" e "Todos" mantêm df intacto

# Pedidos comerciais — filtro canal
if canal_sel == "Todos":
    df_ped_f = df_ped.copy()
else:
    df_ped_f = df_ped[df_ped["TipoCanal"] == canal_sel].copy()

# Join — data + PDV + canal
mask_j = (df_join["_criacao"] >= dt_ini) & (df_join["_criacao"] <= dt_fim + pd.Timedelta(days=1))
if pdv_sel != "Todos": mask_j &= df_join["PDV"] == pdv_sel
if canal_sel != "Todos":
    col_tc_j = find_col(df_join, "TipoCanal")
    if col_tc_j: mask_j &= df_join[col_tc_j] == canal_sel
dj  = df_join[mask_j].copy()
dj_v = dj.dropna(subset=["CodigoPedido"]).copy()

# ── Métricas base ──────────────────────────────────────────────────────────────
total_ped  = len(df_ped_f)
total_log  = len(df)
n_join     = len(dj_v)
sem_log    = total_ped - n_join

entregues  = int((df["Status"] == "Entregue").sum())
devolvidos = int((df["Status"] == "Devolvido").sum())
no_prazo   = int(df["NoPrazo"].sum())
fora_prazo = entregues - no_prazo

tat_vals   = df.loc[(df["Status"]=="Entregue") & df["TAT_dias"].between(0,30), "TAT_dias"]
tat_med    = tat_vals.mean() if len(tat_vals) else 0
tat_invalidos = int(df.get("TAT_invalido", pd.Series(False, index=df.index)).fillna(False).sum())
tp_invalidos = int(dj.get("TP_invalido", pd.Series(False, index=dj.index)).fillna(False).sum())
te_invalidos = int(dj.get("TE_invalido", pd.Series(False, index=dj.index)).fillna(False).sum())

custo_por_pedido = CUSTO_TOTAL_MOTORISTAS / total_log if total_log else 0
df["Custo_Entrega"] = custo_por_pedido if total_log else 0

fretes_inf    = df["_frete"][df["_frete"] > 0]
n_informados  = len(fretes_inf)
pct_informado = n_informados / total_log * 100 if total_log else 0

taxa_dev      = devolvidos / total_log * 100 if total_log else 0
pct_entregues = entregues  / total_log * 100 if total_log else 0
pct_prazo     = no_prazo   / entregues * 100 if entregues else 0
pct_log       = n_join     / total_ped * 100  if total_ped else 0
meta_entrega, meta_label = meta_canal(canal_sel, "entrega")
meta_prazo, _ = meta_canal(canal_sel, "prazo")
status_entrega, cor_meta_entrega = status_meta_pct(pct_entregues, meta_entrega)
status_prazo, cor_meta_prazo = status_meta_pct(pct_prazo, meta_prazo)
status_custo_txt, cor_custo = status_custo(custo_por_pedido)

# -----------------------------------------------------------------------------
# CABEÇALHO
# -----------------------------------------------------------------------------
st.title("📦 Dashboard de Entregas — CP Velanes / GB.Log")

# Indicador do filtro canal ativo
canal_info = {
    "Todos":          ("🔭 Visão completa: todos os canais",                          "#6366f1"),
    "Logístico":      ("📦 Filtro ativo: somente pedidos com expedição logística",    "#22c55e"),
    "Não Logístico":  ("🛒 Filtro ativo: somente pedidos sem movimentação logística", "#f59e0b"),
    "Omnichannel":    ("Filtro ativo: somente pedidos Omnichannel",                "#60a5fa"),
}
ci_label, ci_color = canal_info[canal_sel]
st.markdown(
    f'<div style="background:#1e293b;border-left:4px solid {ci_color};'
    f'padding:7px 14px;border-radius:0 8px 8px 0;color:{ci_color};'
    f'font-size:.82rem;font-weight:600;margin-bottom:8px;">{ci_label}</div>',
    unsafe_allow_html=True,
)

if canal_sel == "Todos":
    resumo = (
        f'No mês de junho, o sistema comercial registrou <b>{fmt_num(total_ped)} pedidos</b> em 7 canais (PDVs). '
        f'Desses, <b>{fmt_num(n_join)} ({fmt_pct(pct_log, 1)}) foram roteados para a logística</b> e atendidos pela gb.log. '
        f'Os demais <b>{fmt_num(sem_log)} permaneceram fora do fluxo logístico</b>. '
        f'Das <b>{fmt_num(total_log)} remessas logísticas</b>, <b>{fmt_num(entregues)} foram entregues ({fmt_pct(pct_entregues, 1)})</b>, '
        f'sendo <b>{fmt_num(no_prazo)} no prazo ({fmt_pct(pct_prazo, 1)})</b>.'
    )
elif canal_sel == "Não Logístico":
    resumo = (
        f'Exibindo apenas os <b>{fmt_num(total_ped)} pedidos comerciais sem movimentação logística</b>. '
        f'Estes pedidos existem no sistema comercial mas <b>não geraram expedição na gb.log</b>. '
        f'As seções de entregas, TAT e motoristas não se aplicam a este canal.'
    )
elif canal_sel == "Omnichannel":
    resumo = (
        f'Exibindo apenas os <b>{fmt_num(total_ped)} pedidos Omnichannel</b>. '
        f'Desses, <b>{fmt_num(n_join)} ({fmt_pct(pct_log, 1)}) geraram expedição logística</b>. ' +
        (
            f'Das <b>{fmt_num(total_log)} remessas</b>, <b>{fmt_num(entregues)} foram entregues ({fmt_pct(pct_entregues, 1)})</b>, '
            f'sendo <b>{fmt_num(no_prazo)} no prazo ({fmt_pct(pct_prazo, 1)})</b>.'
            if total_log > 0
            else 'Nenhum pedido Omnichannel gerou expedição logística no período selecionado.'
        )
    )
else:  # Logístico
    resumo = (
        f'Exibindo apenas os <b>{fmt_num(total_ped)} pedidos com expedição logística</b>. '
        f'Das <b>{fmt_num(total_log)} remessas</b>, <b>{fmt_num(entregues)} foram entregues ({fmt_pct(pct_entregues, 1)})</b>, '
        f'sendo <b>{fmt_num(no_prazo)} no prazo ({fmt_pct(pct_prazo, 1)})</b>. '
        f'Custo médio: <b>{fmt_brl(custo_por_pedido)}/pedido</b>.'
    )

st.markdown(f'<div class="narrativa">{resumo}</div>', unsafe_allow_html=True)
st.markdown(
    f'<div class="narrativa">'
    f'<b>Metas aplicadas ({meta_label}):</b> '
    f'Entregues ≥ <b>{fmt_pct(meta_entrega,0)}</b> · '
    f'Entregues no prazo ≥ <b>{fmt_pct(meta_prazo,0)}</b> · '
    f'Custo médio &lt; <b>{fmt_brl(META_CUSTO_MEDIO)}</b> '
    f'(<span style="color:{cor_custo};font-weight:700">{status_custo_txt}</span>).'
    f'</div>',
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# -----------------------------------------------------------------------------
# -----------------------------------------------------------------------------
st.markdown('<p class="secao">🔭 VISÃO GERAL — Funil de pedidos</p>', unsafe_allow_html=True)
st.markdown(BADGE_BOTH, unsafe_allow_html=True)

f1, f2, f3, f4, f5 = st.columns(5)
f1.metric("Pedidos Comerciais",     fmt_num(total_ped))
f2.metric("Roteados à Logística",   fmt_num(n_join),
          delta=f"{fmt_pct(pct_log, 1)} do total" if total_ped else "–")
f3.metric("Sem Movimentação Log.",  fmt_num(sem_log),
          delta=f"{fmt_pct(100-pct_log, 1)} do total" if total_ped else "–", delta_color="off")
f4.metric("Custo Médio / Pedido",   fmt_brl(custo_por_pedido) if total_log else "–",
          delta=f"Meta < {fmt_brl(META_CUSTO_MEDIO)} | {status_custo_txt}" if total_log else "–",
          delta_color="normal" if custo_por_pedido < META_CUSTO_MEDIO else "inverse")
f5.metric("PDVs Ativos", "7 comerciais · 6 logística",
          help="PDV 24117 só aparece no sistema comercial — sem entregas no período")

col_meio = COLS["meio"]
col_tp   = COLS["tp_ent"]

c_esq, c_dir = st.columns([1.4, 0.7])

with c_esq:
    if canal_sel in ("Todos", "Logístico"):
        # Gráfico PDV: comercial vs logístico
        pdv_com = (df_ped_f[df_ped_f["PDV_Ped"].isin(TODOS_PDVS)]
                     .groupby("PDV_Ped").size().reset_index(name="Comercial"))
        pdv_log_all = dj_v.groupby("PDV_Ped").size().reset_index(name="Logistica")
        pdv_overview = pdv_com.merge(pdv_log_all, on="PDV_Ped", how="left").fillna(0)
        pdv_overview["Logistica"]    = pdv_overview["Logistica"].astype(int)
        pdv_overview["SemLogistica"] = pdv_overview["Comercial"] - pdv_overview["Logistica"]
        pdv_overview = pdv_overview.sort_values("Comercial", ascending=True)

        fig = go.Figure()
        fig.add_trace(go.Bar(
            name="Roteado à Logística",
            y=pdv_overview["PDV_Ped"], x=pdv_overview["Logistica"],
            orientation="h", marker_color=COR_OK,
            text=pdv_overview["Logistica"].apply(fmt_num),
            textposition="inside", textfont_size=11,
        ))
        fig.add_trace(go.Bar(
            name="Sem Movimentação Logística",
            y=pdv_overview["PDV_Ped"], x=pdv_overview["SemLogistica"],
            orientation="h", marker_color=COR_CINZA,
            text=pdv_overview["SemLogistica"].apply(fmt_num),
            textposition="inside", textfont_size=11,
        ))
        p24 = pdv_overview[pdv_overview["PDV_Ped"] == "PDV 24117"]
        if len(p24) > 0:
            fig.add_annotation(
                x=int(p24["Comercial"].values[0]) + 30, y="PDV 24117",
                text="sem logistica", showarrow=False,
                font=dict(color=COR_WARN, size=11), xanchor="left",
            )
        fig.update_layout(**layout_br(
            title="Pedidos por Canal (PDV) — Comercial vs Logística",
            barmode="stack", height=380,
            margin=dict(t=55,b=10,l=10,r=20),
            legend=dict(orientation="h", y=-0.1),
            xaxis_title="Pedidos", yaxis_title="",
        ))
        pdv_overview["RotuloLog"] = pdv_overview["Logistica"].apply(fmt_num)
        pdv_overview["RotuloSemLog"] = pdv_overview["SemLogistica"].apply(fmt_num)
        pdv_overview["BaseSemLog"] = pdv_overview["Logistica"]
        add_hbar_label_boxes(fig, pdv_overview, y_col="PDV_Ped", x_col="Logistica", text_col="RotuloLog")
        add_hbar_label_boxes(fig, pdv_overview, y_col="PDV_Ped", x_col="SemLogistica", text_col="RotuloSemLog", base_col="BaseSemLog")
        fig.update_traces(text=None)
        st.markdown(BADGE_BOTH, unsafe_allow_html=True)
        st.plotly_chart(fig, use_container_width=True)
    else:
        # Para Omnichannel / Não Logístico: breakdown por meio de captação
        if col_meio:
            meio_f = df_ped_f[col_meio].value_counts().reset_index()
            meio_f.columns = ["Canal","Qtd"]
            meio_f["RotuloQtd"] = meio_f["Qtd"].apply(fmt_num)
            fig_m = px.bar(meio_f.sort_values("Qtd"), x="Qtd", y="Canal", orientation="h",
                           title=f"Pedidos por Meio de Captação — {canal_sel}",
                           color_discrete_sequence=[COR_NEU], text="RotuloQtd")
            fig_m.update_traces(textposition="outside")
            fig_m.update_layout(**layout_br(height=380, margin=dict(t=55,b=10,l=10,r=10),
                                            xaxis_title="Pedidos", yaxis_title=""))
            add_hbar_label_boxes(fig_m, meio_f, y_col="Canal", x_col="Qtd", text_col="RotuloQtd", xanchor="left")
            fig_m.update_traces(text=None)
            st.markdown(BADGE_PED, unsafe_allow_html=True)
            st.plotly_chart(fig_m, use_container_width=True)

with c_dir:
    if col_meio and canal_sel == "Todos":
        meio_cnt = df_ped_f[col_meio].value_counts().reset_index()
        meio_cnt.columns = ["Canal","Qtd"]
        fig_m = px.pie(meio_cnt, values="Qtd", names="Canal",
                       title=f"Meio de Captação<br><sup>🛒 {fmt_num(total_ped)} pedidos</sup>",
                       hole=0.50, color_discrete_sequence=px.colors.qualitative.Safe)
        fig_m.update_traces(textinfo="percent+label", textfont_size=11)
        fig_m.update_layout(**layout_br(height=360, margin=dict(t=65,b=10,l=10,r=10), showlegend=False))
        st.plotly_chart(fig_m, use_container_width=True)
    elif col_tp:
        tp_cnt = df_ped_f[col_tp].value_counts().reset_index()
        tp_cnt.columns = ["Tipo","Qtd"]
        tp_cnt["RotuloQtd"] = tp_cnt["Qtd"].apply(fmt_num)
        fig_t = px.bar(tp_cnt.sort_values("Qtd"), x="Qtd", y="Tipo", orientation="h",
                       title=f"Tipo de Entrega<br><sup>{fmt_num(total_ped)} pedidos</sup>",
                       color_discrete_sequence=[COR_NEU], text="RotuloQtd")
        fig_t.update_traces(textposition="outside")
        fig_t.update_layout(**layout_br(height=360, margin=dict(t=65,b=10,l=10,r=10),
                                        yaxis_title="", xaxis_title="Pedidos"))
        add_hbar_label_boxes(fig_t, tp_cnt, y_col="Tipo", x_col="Qtd", text_col="RotuloQtd", xanchor="left")
        fig_t.update_traces(text=None)
        st.plotly_chart(fig_t, use_container_width=True)

if canal_sel == "Não Logístico":
    st.info(
        f"📌 Os {fmt_num(total_ped)} pedidos 'Não Logísticos' existem apenas no sistema comercial. "
        "As seções abaixo (Entregas, TAT, Motoristas) não possuem dados para este canal."
    )

st.markdown("---")

# -----------------------------------------------------------------------------
# -----------------------------------------------------------------------------
# -----------------------------------------------------------------------------
st.markdown('<p class="secao">✅ ENTREGUES — O que chegou, onde e por quem?</p>', unsafe_allow_html=True)
st.markdown(BADGE_LOG, unsafe_allow_html=True)

if total_log == 0:
    st.warning(f"Nenhum pedido log?stico para o canal '{canal_sel}' no per?odo selecionado.")
else:
    l1, l2, l3, l4 = st.columns(4)
    l1.metric("Pedidos Log?sticos",  fmt_num(total_log))
    l2.metric("Entregues",           fmt_num(entregues),
              delta=f"{fmt_pct(pct_entregues, 1)} | meta {fmt_pct(meta_entrega,0)} ({status_entrega})" if total_log else "?")
    l3.metric("No Prazo",            fmt_num(no_prazo),
              delta=f"{fmt_pct(pct_prazo, 1)} | meta {fmt_pct(meta_prazo,0)} ({status_prazo})" if entregues else "?")
    l4.metric("Fora do Prazo",       fmt_num(fora_prazo),
              delta=f"{fmt_pct(100-pct_prazo, 1)} dos entregues" if entregues else "?",
              delta_color="inverse")

    df_ent = df[df["Status"] == "Entregue"].copy()

    if len(df_ent) == 0:
        st.warning("Nenhum pedido entregue no per?odo selecionado.")
    else:
        visao_ent = st.radio(
            "Visao de entregues",
            ["Prazo", "PDV", "Municipios", "Evolucao"],
            index=0,
            horizontal=True,
            key="visao_entregues",
        )

        df_atraso = df_ent[df_ent["NoPrazo"] == False].copy()

        if visao_ent == "Prazo":
            prazo_df = pd.DataFrame({
                "Situacao": ["No prazo", "Fora do prazo"],
                "Qtd": [no_prazo, fora_prazo],
            })
            fig = px.pie(
                prazo_df, values="Qtd", names="Situacao",
                title=f"Prazo de Entrega<br><sup>{fmt_num(entregues)} entregues</sup>",
                hole=0.55, color="Situacao",
                color_discrete_map={"No prazo": COR_OK, "Fora do prazo": COR_WARN},
            )
            fig.update_traces(textinfo="percent+value", textfont_size=14)
            fig.update_layout(**layout_br(height=420, margin=dict(t=65,b=10,l=10,r=10),
                                          legend=dict(orientation="h", y=-0.08)))
            st.plotly_chart(fig, use_container_width=True)

        elif visao_ent == "PDV":
            pdvs_log = ["PDV 22908", "PDV 23868", "PDV 24118", "PDV 24239", "PDV 24270", "PDV 24341"]
            pdv_ent = df_ent.groupby("PDV").agg(
                Total=("Pedido", "count"), NoPrazo=("NoPrazo", "sum"),
            ).reset_index()
            pdv_ent["ForaPrazo"] = pdv_ent["Total"] - pdv_ent["NoPrazo"]
            pdv_ent["PctPrazo"] = (pdv_ent["NoPrazo"] / pdv_ent["Total"].replace(0, 1) * 100).round(2)
            faltando = [p for p in pdvs_log if p not in pdv_ent["PDV"].values]
            if faltando:
                pdv_ent = pd.concat([
                    pdv_ent,
                    pd.DataFrame({"PDV": faltando, "Total": 0, "NoPrazo": 0, "ForaPrazo": 0}),
                ], ignore_index=True)
                pdv_ent["PctPrazo"] = (pdv_ent["NoPrazo"] / pdv_ent["Total"].replace(0, 1) * 100).round(2)
            pdv_ent = pdv_ent.sort_values("Total", ascending=True)
            fig = go.Figure()
            fig.add_trace(go.Bar(
                name="No prazo", y=pdv_ent["PDV"], x=pdv_ent["NoPrazo"],
                orientation="h", marker_color=pdv_ent["PctPrazo"].apply(lambda x: cor_por_meta(x, meta_prazo)),
                text=pdv_ent["NoPrazo"].apply(fmt_num), textposition="inside", textfont_size=10,
            ))
            fig.add_trace(go.Bar(
                name="Fora do prazo", y=pdv_ent["PDV"], x=pdv_ent["ForaPrazo"],
                orientation="h", marker_color=COR_WARN,
                text=pdv_ent["ForaPrazo"].apply(fmt_num), textposition="inside", textfont_size=10,
            ))
            fig.update_layout(**layout_br(
                title="Entregues por PDV", barmode="stack", height=420,
                margin=dict(t=55,b=10,l=10,r=10),
                legend=dict(orientation="h", y=-0.08),
            ))
            pdv_ent["RotuloNoPrazo"] = pdv_ent.apply(lambda r: f"{fmt_num(r['NoPrazo'])} | {fmt_pct(r['PctPrazo'],1)}", axis=1)
            pdv_ent["RotuloForaPrazo"] = pdv_ent["ForaPrazo"].apply(fmt_num)
            pdv_ent["BaseForaPrazo"] = pdv_ent["NoPrazo"]
            add_hbar_label_boxes(fig, pdv_ent, y_col="PDV", x_col="NoPrazo", text_col="RotuloNoPrazo")
            add_hbar_label_boxes(fig, pdv_ent, y_col="PDV", x_col="ForaPrazo", text_col="RotuloForaPrazo", base_col="BaseForaPrazo")
            fig.update_traces(text=None)
            st.plotly_chart(fig, use_container_width=True)

        elif visao_ent == "Municipios":
            muni_ent = df_ent.groupby("Municipio").agg(
                Total=("Pedido", "count"), NoPrazo=("NoPrazo", "sum"),
            ).reset_index()
            muni_ent["PctPrazo"] = (muni_ent["NoPrazo"] / muni_ent["Total"] * 100).round(2)
            muni_ent["Rotulo"] = (
                muni_ent["Total"].apply(fmt_num) + " ped. | " +
                muni_ent["PctPrazo"].apply(lambda x: fmt_pct(x, 1))
            )
            muni_ent = (muni_ent.sort_values("Total", ascending=False)
                                 .head(15).sort_values("Total", ascending=True))
            fig = px.bar(
                muni_ent, x="Total", y="Municipio", orientation="h",
                title="Top 15 Municipios - Entregas<br><sup>Cor = % no prazo | Rotulo = pedidos + % prazo</sup>",
                color_discrete_sequence=[COR_CINZA], text="Rotulo",
                labels={"PctPrazo": "% Prazo", "Total": "Pedidos"},
            )
            fig.update_traces(marker_color=muni_ent["PctPrazo"].apply(lambda x: cor_por_meta(x, meta_prazo)))
            fig.update_traces(textposition="outside", textfont_size=11)
            fig.update_layout(**layout_br(
                height=480, margin=dict(t=65,b=10,l=10,r=170), yaxis_title="",
                coloraxis_colorbar=dict(title="% Prazo", thickness=14, len=0.7),
            ))
            add_hbar_label_boxes(fig, muni_ent, y_col="Municipio", x_col="Total", text_col="Rotulo", xanchor="left")
            fig.update_traces(text=None)
            st.plotly_chart(fig, use_container_width=True)

        elif visao_ent == "Evolucao":
            df_ent["Dia"] = df_ent["_efetuada"].dt.date
            prazo_dia = (df_ent.groupby("Dia")["NoPrazo"].mean() * 100).to_dict()
            evol = df_ent.groupby(["Dia", "NoPrazo"]).size().reset_index(name="Qtd")
            evol["Situacao"] = evol["NoPrazo"].map({True: "No prazo", False: "Fora do prazo"})
            evol["RotuloQtd"] = evol["Qtd"].apply(fmt_num)
            evol["Cor"] = evol.apply(
                lambda r: cor_por_meta(prazo_dia.get(r["Dia"], 0), meta_prazo) if r["NoPrazo"] else COR_WARN,
                axis=1,
            )
            fig_evol = px.bar(
                evol, x="Dia", y="Qtd", color="Situacao",
                title="Evolucao diaria de entregas", barmode="stack",
                text="RotuloQtd",
            )
            for trace in fig_evol.data:
                situacao = trace.name
                cores = []
                for dia in trace.x:
                    if situacao == "No prazo":
                        cores.append(cor_por_meta(prazo_dia.get(dia, 0), meta_prazo))
                    else:
                        cores.append(COR_WARN)
                trace.marker.color = cores
            fig_evol.update_layout(**layout_br(
                height=420, margin=dict(t=55,b=10,l=10,r=10),
                xaxis_title="", yaxis_title="Pedidos entregues",
                legend=dict(orientation="h", y=1.12, x=0),
            ))
            add_vbar_label_boxes(fig_evol, evol, x_col="Dia", y_col="Qtd", text_col="RotuloQtd")
            fig_evol.update_traces(text=None)
            st.plotly_chart(fig_evol, use_container_width=True)

        if len(df_atraso) > 0:
            render_excecoes_pivot(
                preparar_excecoes(df_atraso),
                titulo=f"Entregues com atraso - tabela dinamica ({fmt_num(len(df_atraso))} pedidos)",
                dias_label="Dias de atraso",
                key_prefix="atraso",
            )
        else:
            st.success("Nao ha entregues com atraso no filtro atual.")
st.markdown("---")

# -----------------------------------------------------------------------------
# -----------------------------------------------------------------------------
# -----------------------------------------------------------------------------
st.markdown('<p class="secao">NAO ENTREGUES - Pendencias, devolucoes e falhas</p>', unsafe_allow_html=True)
st.markdown(BADGE_LOG, unsafe_allow_html=True)

if total_log == 0:
    st.info(f"Sem dados logísticos para o canal '{canal_sel}'.")
else:
    df_nent = df[df["Status"] != "Entregue"].copy()
    n_nent  = len(df_nent)

    if n_nent == 0:
        st.success("Todos os pedidos do período foram entregues.")
    else:
        motivo_top = df_nent["Status"].value_counts()
        st.markdown(
            f'<div class="narrativa-warn">'
            f'<b>{fmt_num(n_nent)} pedidos ({fmt_pct(n_nent/total_log*100, 1)} do total logístico)</b> não foram entregues. '
            f'Principal motivo: "<b>{motivo_top.index[0]}</b>" ({fmt_num(int(motivo_top.iloc[0]))} ocorrências). '
            f'<b>{fmt_num(devolvidos)} devoluções</b> e <b>{fmt_num(fora_prazo)} entregas fora do prazo</b>.'
            f'</div>', unsafe_allow_html=True
        )

        ka, kb, kc, kd = st.columns(4)
        n_aberto = int(df_nent[df_nent["Status"].isin(["Em transito","Aguardando geracao de rota","Aguardando motorista"])].shape[0])
        n_ausente= int(df_nent[df_nent["Status"].isin(["Destinatario ausente","Nao visitado"])].shape[0])
        n_falha  = int(df_nent[df_nent["Status"].isin(["Endereco nao localizado","Carga recusada pelo destinatario","Destinatario mudou de endereco","Estabelecimento fechado"])].shape[0])
        ka.metric("Em Trânsito / Aguardando",   fmt_num(n_aberto))
        kb.metric("Destinatário Ausente",        fmt_num(n_ausente))
        kc.metric("Falha de Endereço / Recusa",  fmt_num(n_falha))
        kd.metric("🔴 Devolvidos",              fmt_num(devolvidos),
                  delta=f"{fmt_pct(taxa_dev, 2)} do total", delta_color="inverse")

        visao_nent = st.radio(
            "Visualizar nao entregues por",
            ["Motivo", "PDV", "Municipio"],
            horizontal=True,
            key="visao_nao_entregues",
        )

        if visao_nent == "Motivo":
            graf_nent = df_nent["Status"].value_counts().reset_index()
            graf_nent.columns = ["Grupo", "Qtd"]
            graf_nent["Status"] = graf_nent["Grupo"]
            ordem_grupo = graf_nent.sort_values("Qtd", ascending=True)["Grupo"].tolist()
            titulo_graf = "Nao entregues por motivo"
            mostrar_legenda = False
        elif visao_nent == "PDV":
            graf_nent = (
                df_nent.groupby(["PDV", "Status"])
                .size()
                .reset_index(name="Qtd")
                .rename(columns={"PDV": "Grupo"})
            )
            ordem_grupo = graf_nent.groupby("Grupo")["Qtd"].sum().sort_values(ascending=True).index.tolist()
            titulo_graf = "Nao entregues por PDV"
            mostrar_legenda = True
        else:
            top10 = df_nent["Municipio"].value_counts().head(10).index
            graf_nent = (
                df_nent[df_nent["Municipio"].isin(top10)]
                .groupby(["Municipio", "Status"])
                .size()
                .reset_index(name="Qtd")
                .rename(columns={"Municipio": "Grupo"})
            )
            ordem_grupo = graf_nent.groupby("Grupo")["Qtd"].sum().sort_values(ascending=True).index.tolist()
            titulo_graf = "Top 10 municipios - nao entregues"
            mostrar_legenda = True

        graf_nent = ordenar_stack(graf_nent, "Grupo", "Status", "Qtd")
        graf_nent["RotuloQtd"] = graf_nent["Qtd"].apply(lambda x: fmt_num(x) if x > 0 else "")
        fig = px.bar(
            graf_nent,
            x="Qtd",
            y="Grupo",
            orientation="h",
            title=titulo_graf,
            color="Status",
            barmode="stack",
            color_discrete_map=STATUS_CORES,
            text="RotuloQtd",
            category_orders={
                "Grupo": ordem_grupo,
                "Status": [s for s in ORDEM_STACK if s != "Entregue"],
            },
        )
        fig.update_traces(textposition="outside" if visao_nent == "Motivo" else "inside")
        fig.update_layout(**layout_br(height=430, margin=dict(t=55,b=10,l=10,r=10),
                                      yaxis_title="", xaxis_title="Pedidos",
                                      showlegend=mostrar_legenda,
                                      legend=dict(orientation="h", y=-0.18, x=0)))
        add_stacked_hbar_label_boxes(
            fig,
            graf_nent,
            group_col="Grupo",
            stack_col="Status",
            value_col="Qtd",
            text_col="RotuloQtd",
            stack_order=[s for s in ORDEM_STACK if s != "Entregue"],
        )
        st.plotly_chart(fig, use_container_width=True)

        col_ult = find_col(df_log, "ltima Ocorr","Ultima")
        col_msg = find_col(df_log, "Mensagem")
        cols_t  = ["Pedido","PDV","Municipio","Status","_prazo"]
        if col_ult: cols_t.append(col_ult)
        if col_msg: cols_t.append(col_msg)
        df_tab  = df_nent[[c for c in cols_t if c in df_nent.columns]].copy()
        df_tab["_prazo"] = df_tab["_prazo"].dt.date
        ren = {"Municipio":"Município","_prazo":"Prazo"}
        if col_ult: ren[col_ult] = "Ocorrência"
        if col_msg: ren[col_msg] = "Mensagem"
        # A tabela simples foi substituida pela visao dinamica abaixo.
        # Tabela simples removida; o drill-down abaixo substitui esta visao.

        render_excecoes_pivot(
            preparar_excecoes(df_nent, pendente=True),
            titulo=f"Pedidos com pendencia - tabela dinamica ({fmt_num(n_nent)} pedidos)",
            dias_label="Dias vencido",
            key_prefix="pendencia",
            incluir_status=True,
        )

st.markdown("---")

# -----------------------------------------------------------------------------
# -----------------------------------------------------------------------------
# -----------------------------------------------------------------------------
st.markdown('<p class="secao">TEMPO DE ATENDIMENTO - TAT e TP por cidade/PDV</p>', unsafe_allow_html=True)
st.markdown(BADGE_LOG, unsafe_allow_html=True)
if tat_invalidos:
    st.caption(f"{fmt_num(tat_invalidos)} registro(s) com TAT negativo foram desconsiderados das medias.")
if tp_invalidos:
    st.caption(f"{fmt_num(tp_invalidos)} registro(s) com TP negativo foram desconsiderados das medias.")

if total_log == 0:
    st.info(f"Sem dados logisticos para o canal '{canal_sel}'.")
else:
    df_tat = df[(df["Status"] == "Entregue") & df["TAT_dias"].between(0, 30)].copy()
    df_tp = dj[dj["TP_dias"].between(0, 30)].copy() if "TP_dias" in dj.columns else pd.DataFrame()

    if df_tat.empty and df_tp.empty:
        st.info("Sem dados de TAT/TP para os filtros selecionados.")
    else:
        visao_tempo = st.radio(
            "Visualizar tempos por",
            ["Cidade", "PDV"],
            index=0,
            horizontal=True,
            key="visao_tempo_atendimento",
        )

        tat_resumo = preparar_resumo_tempo(df_tat, tempo_col="TAT_dias", meta_tipo="tat", visao=visao_tempo)
        tp_resumo = preparar_resumo_tempo(df_tp, tempo_col="TP_dias", meta_tipo="tp", visao=visao_tempo)

        pct_tat_ok = tat_resumo["PctMeta"].mean() if not tat_resumo.empty else 0
        pct_tp_ok = tp_resumo["PctMeta"].mean() if not tp_resumo.empty else 0
        tat_media = df_tat["TAT_dias"].mean() if not df_tat.empty else 0
        tp_media = df_tp["TP_dias"].mean() if not df_tp.empty else 0

        st.markdown(
            f'<div class="narrativa">'
            f'<b>TAT</b> = aprovacao ate entrega; meta <b>2 dias local / 3 dias fora</b>. '
            f'Media atual: <b>{fmt_num(tat_media, 1)} dias</b>; dentro da meta: <b>{fmt_pct(pct_tat_ok, 1)}</b>.<br>'
            f'<b>TP</b> = aprovacao ate faturamento; meta <b>1 dia local / 2 dias fora</b>. '
            f'Media atual: <b>{fmt_num(tp_media, 1)} dias</b>; dentro da meta: <b>{fmt_pct(pct_tp_ok, 1)}</b>.'
            f'</div>',
            unsafe_allow_html=True,
        )

        c_tat, c_tp = st.columns(2)
        with c_tat:
            render_tempo_bar(tat_resumo, titulo=f"TAT por {visao_tempo.lower()} - meta local/fora")
        with c_tp:
            render_tempo_bar(tp_resumo, titulo=f"TP por {visao_tempo.lower()} - meta local/fora")
st.markdown("---")

# -----------------------------------------------------------------------------
# -----------------------------------------------------------------------------
# -----------------------------------------------------------------------------
st.markdown('<p class="secao">🚚 MOTORISTAS — Desempenho individual de cada entregador</p>', unsafe_allow_html=True)
st.markdown(BADGE_LOG, unsafe_allow_html=True)
if te_invalidos:
    st.caption(
        f"{fmt_num(te_invalidos)} TE negativo(s) foram desconsiderados das medias."
    )

if total_log == 0:
    st.info(f"Sem dados logísticos para o canal '{canal_sel}'.")
else:
    col_mot = find_col(dj, "Motorista") or find_col(df, "Motorista")
    df_mot_base = dj.copy() if len(dj) else df.copy()
    for col_tempo in ["TE_dias"]:
        if col_tempo not in df_mot_base.columns:
            df_mot_base[col_tempo] = pd.NA
    df_mot_base["Motorista"] = df_mot_base[col_mot].fillna("Sem motorista atribuído") if col_mot else "N/D"
    df_mot_base["MetaTE"] = df_mot_base["Municipio"].apply(lambda x: meta_tempo_por_municipio(x, tipo="te"))
    df_mot_base["TE_OK"] = df_mot_base["TE_dias"].between(0, 30) & (df_mot_base["TE_dias"] <= df_mot_base["MetaTE"]) if "TE_dias" in df_mot_base.columns else False

    mot = df_mot_base.groupby("Motorista").agg(
        Total       = ("Pedido",  "count"),
        Entregues   = ("Status",  lambda x: (x == "Entregue").sum()),
        NoPrazo     = ("NoPrazo", "sum"),
        Devolvidos  = ("Status",  lambda x: (x == "Devolvido").sum()),
        TAT_med     = ("TAT_dias","mean"),
        TE_med      = ("TE_dias", "mean"),
        TE_OK       = ("TE_OK", "sum"),
        Ocorrencias = ("TemOcorr","sum"),
    ).reset_index()
    mot["PctEntrega"] = (mot["Entregues"] / mot["Total"]                  * 100).round(2)
    mot["PctPrazo"]   = (mot["NoPrazo"]   / mot["Entregues"].replace(0,1) * 100).round(2)
    mot["ForaPrazo"]  = mot["Entregues"] - mot["NoPrazo"]
    mot["TAT_med"]    = mot["TAT_med"].round(2)
    mot["TE_med"]     = mot["TE_med"].round(2)
    mot["PctTE"]      = (mot["TE_OK"] / mot["Entregues"].replace(0, 1) * 100).round(2)
    mot_ativos = mot[mot["Motorista"] != "Sem motorista atribuído"].sort_values("Total", ascending=False)
    n_sem_mot  = int(mot.loc[mot["Motorista"] == "Sem motorista atribuído", "Total"].sum())

    alertas = []
    for _, r in mot_ativos.iterrows():
        if r["PctPrazo"] < 60 and r["Entregues"] > 5:
            alertas.append(f"**{r['Motorista']}**: {fmt_pct(r['PctPrazo'],0)} no prazo ({fmt_num(int(r['Entregues']))} entregas)")
        if r["PctEntrega"] < 90 and r["Total"] > 10:
            alertas.append(f"**{r['Motorista']}**: taxa de entrega {fmt_pct(r['PctEntrega'],0)} ({fmt_num(int(r['Total']))} pedidos)")

    narrativa_mot = (
        f"<b>{fmt_num(mot_ativos['Motorista'].nunique())} motoristas</b> ativos, "
        f"responsáveis por <b>{fmt_num(int(mot_ativos['Total'].sum()))} pedidos</b>. "
        f"<b>{fmt_num(n_sem_mot)} pedidos</b> sem motorista atribuído. "
        f"Custo rateado: <b>{fmt_brl(custo_por_pedido)}/pedido</b>."
    )
    if alertas:
        narrativa_mot += "<br><br>🔴 <b>Atenção:</b> " + " · ".join(alertas)

    badge_cor = "narrativa-warn" if alertas else "narrativa"
    st.markdown(f'<div class="{badge_cor}">{narrativa_mot}</div>', unsafe_allow_html=True)

    visao_mot = st.radio(
        "Visão dos motoristas",
        ["Volume / status", "% no prazo", "TE"],
        index=0,
        horizontal=True,
        key="visao_motoristas",
    )

    if visao_mot == "Volume / status":
        status_mot = (df_mot_base[df_mot_base["Motorista"] != "Sem motorista atribuído"]
                      .groupby(["Motorista","Status"]).size().reset_index(name="Qtd"))
        status_mot = ordenar_stack(status_mot, "Motorista", "Status", "Qtd")
        total_mot_status = (
            status_mot.groupby("Motorista")
            .agg(Total=("Qtd", "sum"))
            .reset_index()
        )
        detalhes_status = (
            status_mot[status_mot["Qtd"] > 0]
            .assign(Detalhe=lambda d: d["Status"].astype(str) + ": " + d["Qtd"].apply(fmt_num))
            .groupby("Motorista")["Detalhe"]
            .apply(lambda x: "<br>".join(x))
            .reset_index(name="DetalheHover")
        )
        total_mot_status = total_mot_status.merge(detalhes_status, on="Motorista", how="left")
        total_mot_status["RotuloTotal"] = total_mot_status["Total"].apply(fmt_num)
        ordem_m = total_mot_status.sort_values("Total", ascending=True)["Motorista"].tolist()
        status_mot = status_mot.merge(total_mot_status, on="Motorista", how="left")
        fig = px.bar(status_mot, x="Qtd", y="Motorista", color="Status",
                     orientation="h", barmode="stack",
                     title="Status por Motorista",
                     color_discrete_map=STATUS_CORES,
                     custom_data=["Total", "DetalheHover"],
                     category_orders={"Motorista": ordem_m[::-1], "Status": ORDEM_STACK})
        fig.update_traces(
            hovertemplate=(
                "<b>%{y}</b><br>"
                "Total: %{customdata[0]}<br><br>"
                "%{customdata[1]}<extra></extra>"
            ),
            text=None,
        )
        fig.update_layout(**layout_br(height=420, margin=dict(t=55,b=10,l=10,r=10),
                                      xaxis_title="Pedidos", yaxis_title="",
                                      legend=dict(orientation="h", y=-0.2, x=0)))
        add_hbar_label_boxes(
            fig,
            total_mot_status,
            y_col="Motorista",
            x_col="Total",
            text_col="RotuloTotal",
            xanchor="left",
        )
        st.plotly_chart(fig, use_container_width=True)

    elif visao_mot == "% no prazo":
        mot_prazo = mot_ativos.sort_values("PctPrazo", ascending=True)
        cores_prazo = [cor_por_meta(v, meta_prazo) for v in mot_prazo["PctPrazo"]]
        fig = go.Figure(go.Bar(
            x=mot_prazo["PctPrazo"], y=mot_prazo["Motorista"],
            orientation="h", marker_color=cores_prazo,
            text=mot_prazo["PctPrazo"].apply(lambda x: fmt_pct(x, 2)),
            textposition="outside",
        ))
        fig.add_vline(x=meta_prazo, line_dash="dash", line_color="#94a3b8",
                      annotation_text=f"Meta {fmt_pct(meta_prazo,0)}", annotation_position="top right",
                      annotation_font_color="#94a3b8")
        fig.update_layout(**layout_br(
            title="% No Prazo por Motorista",
            height=420, margin=dict(t=55,b=10,l=10,r=80),
            xaxis=dict(title="% no prazo", range=[0, 115]),
            yaxis_title="",
        ))
        mot_prazo["RotuloPrazo"] = mot_prazo["PctPrazo"].apply(lambda x: fmt_pct(x, 2))
        add_hbar_label_boxes(fig, mot_prazo, y_col="Motorista", x_col="PctPrazo", text_col="RotuloPrazo", xanchor="left")
        fig.update_traces(text=None)
        st.plotly_chart(fig, use_container_width=True)

    elif visao_mot == "TE":
        col_tempo = "TE_med"
        col_pct = "PctTE"
        titulo = "TE - coleta BG até entrega"
        dados = mot_ativos.sort_values(col_tempo, ascending=True).copy()
        cores = [COR_ERR if v > 2 else COR_WARN if v > 1 else COR_OK for v in dados[col_tempo].fillna(99)]
        fig = go.Figure(go.Bar(
            x=dados[col_tempo], y=dados["Motorista"], orientation="h",
            marker_color=cores,
            text=dados[col_tempo].apply(lambda x: "—" if pd.isna(x) else f"{fmt_num(x,1)}d"),
            textposition="outside",
            customdata=dados[col_pct],
            hovertemplate="%{y}<br>Média: %{x:.2f} dias<br>Dentro da meta: %{customdata:.1f}%<extra></extra>",
        ))
        fig.add_vline(x=1, line_dash="dash", line_color=COR_OK,
                      annotation_text="Meta Itabuna 1d", annotation_font_color=COR_OK)
        fig.add_vline(x=2, line_dash="dash", line_color=COR_WARN,
                      annotation_text="Meta fora 2d", annotation_font_color=COR_WARN)
        fig.update_layout(**layout_br(
            title=titulo,
            height=420, margin=dict(t=55,b=10,l=10,r=80),
            xaxis_title="Dias", yaxis_title="",
        ))
        dados["RotuloTempo"] = dados[col_tempo].apply(lambda x: "" if pd.isna(x) else f"{fmt_num(x,1)}d")
        add_hbar_label_boxes(fig, dados, y_col="Motorista", x_col=col_tempo, text_col="RotuloTempo", xanchor="left")
        fig.update_traces(text=None)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("**Resumo por motorista:**")
    tabela_mot = mot_ativos[[
        "Motorista","Total","Entregues","PctEntrega","NoPrazo","PctPrazo",
        "ForaPrazo","Devolvidos","TAT_med","TE_med","PctTE","Ocorrencias"
    ]].rename(columns={
        "Total":"Pedidos","PctEntrega":"% Entrega",
        "NoPrazo":"No Prazo","PctPrazo":"% Prazo","ForaPrazo":"Fora Prazo",
        "TE_med":"TE Medio (d)","PctTE":"% TE Meta",
        "TAT_med":"TAT Médio (d)","Ocorrencias":"Ocorrências",
    }).reset_index(drop=True)

    st.dataframe(
        tabela_mot.style
            .format({
                "Pedidos":       lambda x: fmt_num(x),
                "Entregues":     lambda x: fmt_num(x),
                "% Entrega":     lambda x: fmt_pct(x, 2),
                "No Prazo":      lambda x: fmt_num(x),
                "% Prazo":       lambda x: fmt_pct(x, 2),
                "Fora Prazo":    lambda x: fmt_num(x),
                "Devolvidos":    lambda x: fmt_num(x),
                "TE Medio (d)":   lambda x: fmt_num(x, 2),
                "% TE Meta":      lambda x: fmt_pct(x, 2),
                "TAT Médio (d)": lambda x: fmt_num(x, 2),
                "Ocorrências":   lambda x: fmt_num(x),
            })
            .map(lambda v: estilo_pct_meta(v, meta_prazo), subset=["% Prazo"])
            .map(lambda v: estilo_pct_meta(v, meta_entrega), subset=["% Entrega"])
            .background_gradient(subset=["% TE Meta"],     cmap="RdYlGn", vmin=0,  vmax=100)
            .background_gradient(subset=["TAT Médio (d)"], cmap="RdYlGn_r", vmin=1, vmax=7)
            .background_gradient(subset=["TE Medio (d)"],  cmap="RdYlGn_r", vmin=1, vmax=5),
        use_container_width=True, height=320,
    )
    st.markdown("**Resumo por PDV:**")
    pdv_resumo = df.groupby("PDV").agg(
        Total       = ("Pedido",  "count"),
        Entregues   = ("Status",  lambda x: (x == "Entregue").sum()),
        NoPrazo     = ("NoPrazo", "sum"),
        Devolvidos  = ("Status",  lambda x: (x == "Devolvido").sum()),
        TAT_med     = ("TAT_dias","mean"),
        Ocorrencias = ("TemOcorr","sum"),
    ).reset_index()
    pdv_resumo["PctEntrega"] = (pdv_resumo["Entregues"] / pdv_resumo["Total"]                  * 100).round(2)
    pdv_resumo["PctPrazo"]   = (pdv_resumo["NoPrazo"]   / pdv_resumo["Entregues"].replace(0,1) * 100).round(2)
    pdv_resumo["ForaPrazo"]  = pdv_resumo["Entregues"] - pdv_resumo["NoPrazo"]
    pdv_resumo["TAT_med"]    = pdv_resumo["TAT_med"].round(2)
    pdv_resumo["CustoRateado"] = (pdv_resumo["Total"] * custo_por_pedido).round(2)
    pdv_resumo = pdv_resumo.sort_values("Total", ascending=False)

    tabela_pdv = pdv_resumo[[
        "PDV","Total","Entregues","PctEntrega","NoPrazo","PctPrazo",
        "ForaPrazo","Devolvidos","TAT_med","Ocorrencias","CustoRateado"
    ]].rename(columns={
        "Total":"Pedidos","PctEntrega":"% Entrega",
        "NoPrazo":"No Prazo","PctPrazo":"% Prazo","ForaPrazo":"Fora Prazo",
        "TAT_med":"TAT Medio (d)","Ocorrencias":"Ocorrencias",
        "CustoRateado":"Custo Rateado",
    }).reset_index(drop=True)

    st.dataframe(
        tabela_pdv.style
            .format({
                "Pedidos":       lambda x: fmt_num(x),
                "Entregues":     lambda x: fmt_num(x),
                "% Entrega":     lambda x: fmt_pct(x, 2),
                "No Prazo":      lambda x: fmt_num(x),
                "% Prazo":       lambda x: fmt_pct(x, 2),
                "Fora Prazo":    lambda x: fmt_num(x),
                "Devolvidos":    lambda x: fmt_num(x),
                "TAT Medio (d)": lambda x: fmt_num(x, 2),
                "Ocorrencias":   lambda x: fmt_num(x),
                "Custo Rateado": lambda x: fmt_brl(x),
            })
            .map(lambda v: estilo_pct_meta(v, meta_prazo), subset=["% Prazo"])
            .map(lambda v: estilo_pct_meta(v, meta_entrega), subset=["% Entrega"])
            .background_gradient(subset=["TAT Medio (d)"], cmap="RdYlGn_r", vmin=1, vmax=7)
            .background_gradient(subset=["Custo Rateado"], cmap="YlOrRd"),
        use_container_width=True, height=260,
    )

    if n_sem_mot > 0:
        st.caption(f"Atencao: {fmt_num(n_sem_mot)} pedidos sem motorista atribuido excluidos da tabela.")

st.markdown("---")

if total_log > 0:
    st.warning(
        f"Atencao: **Coluna 'Valor do frete' com entrada incorreta:** "
        f"{fmt_num(n_informados)} de {fmt_num(total_log)} pedidos "
        f"({fmt_pct(pct_informado, 1)}) têm valor registrado. "
        f"O custo exibido usa o total pago a motoristas ({fmt_brl(CUSTO_TOTAL_MOTORISTAS)}) rateado igualmente. "
        f"**Ação necessária:** corrigir o lançamento do frete na plataforma logística."
    )

st.caption("📦 Logística = GB.Log · 🛒 Pedidos = Sistema Comercial · 🔗 Cruzamento por CodigoPedido · 2026-07-03")
