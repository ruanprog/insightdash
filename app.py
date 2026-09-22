import io
import os
import re
import html
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
from pypdf import PdfReader
from docx import Document

st.set_page_config(page_title="InsightDash", page_icon="📊", layout="wide")

st.markdown("""
<style>
.block-container {max-width: 1400px; padding-top: 2rem;}
.hero {padding: 28px; border-radius: 22px; background: linear-gradient(135deg,#0b1220,#172554);
        color:white; margin-bottom:22px;}
.hero h1 {font-size: 42px; margin:0;}
.hero p {font-size:18px; opacity:.85;}
.kpi {padding:20px; border-radius:18px; background:#fff; border:1px solid #e7eaf0;
      box-shadow:0 6px 22px rgba(15,23,42,.06);}
.kpi-label {color:#667085; font-size:14px;}
.kpi-value {font-size:27px; font-weight:700; margin-top:5px;}
.insight {padding:16px 18px; border-left:4px solid #2563eb; background:#f8fafc;
          border-radius:12px; margin:8px 0;}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
<h1>📊 InsightDash</h1>
<p>Transforme seus arquivos empresariais em dashboards executivos.</p>
</div>
""", unsafe_allow_html=True)

@st.cache_data(show_spinner=False)
def parse_file(name, data):
    ext = os.path.splitext(name.lower())[1]
    if ext in [".xlsx", ".xls"]:
        sheets = pd.read_excel(io.BytesIO(data), sheet_name=None)
        frames = []
        for sheet, frame in sheets.items():
            if not frame.empty:
                frame = frame.copy()
                frame["_origem_aba"] = sheet
                frames.append(frame)
        return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    if ext == ".csv":
        for enc in ["utf-8-sig", "utf-8", "latin1"]:
            try:
                return pd.read_csv(io.BytesIO(data), encoding=enc, sep=None, engine="python")
            except Exception:
                pass
        raise ValueError("CSV inválido.")
    if ext == ".pdf":
        # Extração simples de tabelas/texto
        reader = PdfReader(io.BytesIO(data))
        text = "\n".join((p.extract_text() or "") for p in reader.pages)
        return pd.DataFrame({"conteudo_extraido": [x.strip() for x in text.splitlines() if x.strip()]})
    if ext == ".docx":
        doc = Document(io.BytesIO(data))
        rows = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                rows.append(" | ".join(c.text.strip() for c in row.cells))
        return pd.DataFrame({"conteudo_extraido": rows})
    raise ValueError("Formato não suportado.")

def normalize_col(c):
    c = str(c).strip().lower()
    c = re.sub(r"[^\wÀ-ÿ]+", "_", c, flags=re.UNICODE)
    return re.sub(r"_+", "_", c).strip("_") or "campo"

def clean_data(df):
    df = df.copy()
    df.columns = [normalize_col(c) for c in df.columns]
    df = df.dropna(how="all").drop_duplicates().reset_index(drop=True)

    for c in df.columns:
        if df[c].dtype == "object":
            x = df[c].astype(str).str.strip()
            x2 = x.str.replace(r"R\$\s*", "", regex=True)
            br = x2.str.match(r"^-?\d{1,3}(\.\d{3})+,\d+$|^-?\d+,\d+$", na=False)
            y = x2.copy()
            y.loc[br] = y.loc[br].str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
            y.loc[~br] = y.loc[~br].str.replace(",", "", regex=False)
            n = pd.to_numeric(y, errors="coerce")
            if n.notna().mean() >= .75:
                df[c] = n
                continue
            d = pd.to_datetime(x, errors="coerce", dayfirst=True)
            if d.notna().mean() >= .80:
                df[c] = d
    return df

def money(v):
    if pd.isna(v): return "—"
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def number(v):
    if pd.isna(v): return "—"
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

with st.sidebar:
    st.header("📁 Sua análise")
    uploads = st.file_uploader(
        "Envie Excel, CSV, PDF ou Word",
        type=["xlsx","xls","csv","pdf","docx"],
        accept_multiple_files=True
    )
    st.caption("Arquivos são processados na sessão. Para produção, use armazenamento privado e autenticação.")

if not uploads:
    st.info("👆 Envie um arquivo na barra lateral para começar.")
    st.stop()

frames = []
errors = []
for f in uploads:
    try:
        x = parse_file(f.name, f.getvalue())
        if not x.empty:
            x["_arquivo"] = f.name
            frames.append(x)
    except Exception as e:
        errors.append(f"{f.name}: {e}")

if errors:
    st.warning("Alguns arquivos não puderam ser processados: " + " | ".join(errors))
if not frames:
    st.error("Nenhum dado foi extraído.")
    st.stop()

df = clean_data(pd.concat(frames, ignore_index=True, sort=False))

num_cols = df.select_dtypes(include=np.number).columns.tolist()
date_cols = df.select_dtypes(include=["datetime64[ns]", "datetimetz"]).columns.tolist()
cat_cols = [c for c in df.columns if c not in num_cols + date_cols]

if not num_cols:
    st.warning("Não encontrei métricas numéricas. Você ainda pode visualizar a base tratada.")
    st.dataframe(df, use_container_width=True, height=500)
    st.download_button("⬇️ Baixar CSV tratado", df.to_csv(index=False).encode("utf-8-sig"),
                       "insightdash_tratado.csv", "text/csv")
    st.stop()

# Sugestões inteligentes
words = ["venda","fatur","receita","valor","preco","preço","lucro","quant","qtd","total","custo"]
metric_scores = {c: sum(w in c for w in words) for c in num_cols}
metric_default = max(num_cols, key=lambda c: (metric_scores[c], df[c].notna().sum()))

dim_words = ["produto","cliente","categoria","regiao","região","vendedor","cidade","estado","canal","marca"]
if cat_cols:
    dim_default = max(cat_cols, key=lambda c: (sum(w in c for w in dim_words), -df[c].nunique(dropna=True)))
else:
    dim_default = None

with st.sidebar:
    metric = st.selectbox("Métrica principal", num_cols, index=num_cols.index(metric_default))
    if cat_cols:
        dim = st.selectbox("Dimensão", cat_cols, index=cat_cols.index(dim_default))
    else:
        dim = None
    date = st.selectbox("Data", ["Nenhuma"] + date_cols)
    date = None if date == "Nenhuma" else date

    st.divider()
    if dim:
        vals = sorted(df[dim].dropna().astype(str).unique().tolist())
        selected = st.multiselect(f"Filtrar {dim}", vals, default=[])
    else:
        selected = []

view = df.copy()
if dim and selected:
    view = view[view[dim].astype(str).isin(selected)]

st.caption(f"Base analisada: **{len(view):,} registros** · {len(view.columns)} campos")

c1,c2,c3,c4 = st.columns(4)
with c1:
    st.markdown(f'<div class="kpi"><div class="kpi-label">Total de {metric}</div><div class="kpi-value">{number(view[metric].sum())}</div></div>', unsafe_allow_html=True)
with c2:
    st.markdown(f'<div class="kpi"><div class="kpi-label">Média</div><div class="kpi-value">{number(view[metric].mean())}</div></div>', unsafe_allow_html=True)
with c3:
    st.markdown(f'<div class="kpi"><div class="kpi-label">Máximo</div><div class="kpi-value">{number(view[metric].max())}</div></div>', unsafe_allow_html=True)
with c4:
    st.markdown(f'<div class="kpi"><div class="kpi-label">Registros</div><div class="kpi-value">{len(view):,}</div></div>', unsafe_allow_html=True)

st.divider()

if date:
    t = view[[date, metric]].dropna().copy()
    t["periodo"] = t[date].dt.to_period("M").dt.to_timestamp()
    monthly = t.groupby("periodo", as_index=False)[metric].sum()
    fig = px.area(monthly, x="periodo", y=metric, markers=True, title="📈 Evolução")
    fig.update_layout(template="plotly_white", height=420)
    st.plotly_chart(fig, use_container_width=True)

a,b = st.columns(2)
if dim:
    rank = view.groupby(dim, dropna=False)[metric].sum().sort_values(ascending=False).head(15).reset_index()
    with a:
        fig = px.bar(rank, x=metric, y=dim, orientation="h", title=f"🏆 Top {dim}")
        fig.update_layout(template="plotly_white", height=500, yaxis={"categoryorder":"total ascending"})
        st.plotly_chart(fig, use_container_width=True)
    with b:
        mix = view.groupby(dim, dropna=False)[metric].sum().sort_values(ascending=False).head(10).reset_index()
        fig = px.pie(mix, names=dim, values=metric, hole=.5, title=f"🍩 Participação")
        fig.update_layout(template="plotly_white", height=500)
        st.plotly_chart(fig, use_container_width=True)
else:
    st.dataframe(view, use_container_width=True)

st.subheader("🧠 Insights automáticos")
insights = []
total = view[metric].sum()
if dim and total:
    g = view.groupby(dim)[metric].sum().sort_values(ascending=False)
    if len(g):
        share = g.iloc[0]/total*100
        insights.append(f"O grupo **{g.index[0]}** concentra aproximadamente **{share:.1f}%** do total de `{metric}`.")
if date and len(view) >= 2:
    s = view[[date,metric]].dropna().sort_values(date)
    if len(s) >= 2 and s.iloc[0][metric] != 0:
        ch = (s.iloc[-1][metric]-s.iloc[0][metric])/abs(s.iloc[0][metric])*100
        insights.append(f"A variação entre o primeiro e o último registro disponível é de aproximadamente **{ch:.1f}%**.")
q1,q3 = view[metric].quantile([.25,.75])
iqr=q3-q1
out=view[(view[metric] < q1-1.5*iqr) | (view[metric] > q3+1.5*iqr)]
insights.append(f"Foram identificados **{len(out):,} possíveis outliers** na métrica selecionada pelo método do IQR.")
for x in insights:
    st.markdown(f'<div class="insight">💡 {x}</div>', unsafe_allow_html=True)

with st.expander("🔎 Ver dados tratados"):
    st.dataframe(view, use_container_width=True, height=500)

st.download_button(
    "⬇️ Baixar base tratada",
    view.to_csv(index=False).encode("utf-8-sig"),
    "insightdash_tratado.csv",
    "text/csv"
)

st.caption("InsightDash • Protótipo web. Para dados sensíveis, configure autenticação, armazenamento privado, limites de upload e políticas de retenção antes de uso comercial.")
