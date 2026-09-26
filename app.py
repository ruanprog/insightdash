import os,io,sqlite3,hashlib,hmac,re,json
from pathlib import Path
from datetime import datetime
import pandas as pd
from fastapi import FastAPI,Request,UploadFile,File,Form,HTTPException
from fastapi.responses import HTMLResponse,RedirectResponse,StreamingResponse
from starlette.middleware.sessions import SessionMiddleware
from pypdf import PdfReader
from docx import Document

ROOT=Path(__file__).parent; DB=ROOT/'data/insightdash.db'; DB.parent.mkdir(exist_ok=True)
app=FastAPI(title='InsightDash V2'); app.add_middleware(SessionMiddleware,secret_key=os.getenv('SESSION_SECRET','CHANGE_ME'),max_age=28800)

def conn():
 c=sqlite3.connect(DB); c.row_factory=sqlite3.Row; return c
c=conn(); c.executescript('''CREATE TABLE IF NOT EXISTS tenants(id INTEGER PRIMARY KEY,name TEXT,created_at TEXT);CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY,tenant_id INTEGER,name TEXT,email TEXT UNIQUE,password TEXT,role TEXT,created_at TEXT);CREATE TABLE IF NOT EXISTS dashboards(id INTEGER PRIMARY KEY,tenant_id INTEGER,user_id INTEGER,name TEXT,metric TEXT,dimension TEXT,date_col TEXT,data_json TEXT,created_at TEXT);CREATE TABLE IF NOT EXISTS audit(id INTEGER PRIMARY KEY,tenant_id INTEGER,user_id INTEGER,action TEXT,created_at TEXT);'''); c.commit(); c.close()

def hp(x): return hashlib.sha256(x.encode()).hexdigest()
def esc(x): import html; return html.escape(str(x))
CSS='''body{margin:0;background:#f5f7fb;color:#0f172a;font-family:Inter,system-ui,sans-serif}nav{height:64px;background:white;border-bottom:1px solid #e2e8f0;padding:0 5%;display:flex;align-items:center;justify-content:space-between}.logo{font-size:22px;font-weight:800}.logo b{color:#2563eb}a{text-decoration:none;color:#2563eb}.btn{background:#2563eb;color:white;padding:11px 16px;border-radius:11px;border:0;font-weight:700}.container{max-width:1250px;margin:auto;padding:30px 5%}.hero{padding:55px;border-radius:26px;background:linear-gradient(135deg,#0b1220,#2563eb);color:white}.hero h1{font-size:46px;margin:0 0 12px}.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin:22px 0}.card{background:white;border:1px solid #e2e8f0;border-radius:18px;padding:20px;box-shadow:0 8px 25px #0f172a0b}.big{font-size:28px;font-weight:800}.muted{color:#64748b}.form{max-width:500px;margin:35px auto}.form input{width:100%;padding:12px;border:1px solid #cbd5e1;border-radius:10px;margin:7px 0 16px;box-sizing:border-box}.chart{height:380px}@media(max-width:800px){.grid{grid-template-columns:1fr 1fr}.hero h1{font-size:34px}}'''
def page(title,body,u=None):
   nav=f'<nav><a href="/" class="logo">Insight<b>Dash</b></a><div>'+(f'<a href="/dashboard">Visão geral</a> &nbsp; <a href="/upload">Importar</a> &nbsp; <a href="/logout">Sair</a>' if u else '<a href="/login">Entrar</a> &nbsp; <a class="btn" href="/register">Começar</a>')+'</div></nav>'
   nav=f'<nav><a href="/" class="logo">Insight<b>Dash</b></a><div>'+(f'<a href="/dashboard">Visão geral</a> &nbsp; <a href="/upload">Importar</a> &nbsp; <a href="/logout">Sair</a>' if u else '<a href="/login">Entrar</a> &nbsp; <a class="btn" href="/register">Começar</a>')+'</div></nav>'
return f'<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{esc(title)}</title><script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script><style>{CSS}</style></head><body>{nav}<main class="container">{body}</main></body></html>'
def user(req):
 uid=req.session.get('uid');
 if not uid:return None
 c=conn(); u=c.execute('SELECT * FROM users WHERE id=?',(uid,)).fetchone(); c.close(); return u
def auth(req):
 u=user(req)
 if not u: raise HTTPException(303,headers={'Location':'/login'})
 return u
def parse(name,data):
 e=Path(name).suffix.lower()
 if e in ('.xlsx','.xls'):
  ss=pd.read_excel(io.BytesIO(data),sheet_name=None); fs=[f for f in ss.values() if not f.empty]; return pd.concat(fs,ignore_index=True) if fs else pd.DataFrame()
 if e=='.csv': return pd.read_csv(io.BytesIO(data),sep=None,engine='python')
 if e=='.pdf':
  r=PdfReader(io.BytesIO(data)); return pd.DataFrame({'conteudo_extraido':[x for x in '\n'.join((p.extract_text() or '') for p in r.pages).splitlines() if x.strip()]})
 if e=='.docx':
  d=Document(io.BytesIO(data)); return pd.DataFrame({'conteudo_extraido':[p.text for p in d.paragraphs if p.text.strip()]})
 raise ValueError('Formato não suportado')
def clean(df):
 df=df.dropna(how='all').drop_duplicates().copy(); df.columns=[re.sub('_+','_',re.sub(r'[^\wÀ-ÿ]+','_',str(x).lower())).strip('_') or 'campo' for x in df.columns]
 for c in df.columns:
  if df[c].dtype=='object':
   x=df[c].astype(str).str.strip(); n=pd.to_numeric(x.str.replace('R$','',regex=False).str.replace('.','',regex=False).str.replace(',','.',regex=False),errors='coerce')
   if n.notna().mean()>.75: df[c]=n; continue
   d=pd.to_datetime(x,errors='coerce',dayfirst=True)
   if d.notna().mean()>.8: df[c]=d
 return df

@app.get('/',response_class=HTMLResponse)
def home(req:Request):
 if user(req): return RedirectResponse('/dashboard',303)
 return page('Início','''<section class="hero"><h1>Transforme dados em decisões.</h1><p>Envie Excel, CSV, PDF ou Word e crie dashboards executivos em minutos.</p><br><a class="btn" href="/register">Começar agora →</a></section><div class="grid"><div class="card"><h3>📤 Importar</h3><p class="muted">Suas bases empresariais.</p></div><div class="card"><h3>🧹 Tratar</h3><p class="muted">Limpeza automática.</p></div><div class="card"><h3>📊 Analisar</h3><p class="muted">KPIs, rankings e tendências.</p></div><div class="card"><h3>⬇ Exportar</h3><p class="muted">Baixe os resultados.</p></div></div>''')
@app.get('/register',response_class=HTMLResponse)
def rg(req:Request): return page('Cadastro','''<div class="form card"><h1>Criar empresa</h1><form method="post"><label>Empresa</label><input name="company" required><label>Nome</label><input name="name" required><label>E-mail</label><input name="email" type="email" required><label>Senha</label><input name="password" type="password" minlength="8" required><button class="btn">Criar conta</button></form></div>''')
@app.post('/register')
def rp(request:Request,company:str=Form(...),name:str=Form(...),email:str=Form(...),password:str=Form(...)):
 c=conn(); now=datetime.utcnow().isoformat()
 try:
  tid=c.execute('INSERT INTO tenants(name,created_at) VALUES(?,?)',(company,now)).lastrowid; uid=c.execute('INSERT INTO users(tenant_id,name,email,password,role,created_at) VALUES(?,?,?,?,?,?)',(tid,name,email.lower(),hp(password),'admin',now)).lastrowid;c.commit();request.session['uid']=uid;return RedirectResponse('/dashboard',303)
 except sqlite3.IntegrityError:return HTMLResponse(page('Erro','<div class="card"><h2>E-mail já cadastrado.</h2></div>'),400)
 finally:c.close()
@app.get('/login',response_class=HTMLResponse)
def lg(req:Request): return page('Login','''<div class="form card"><h1>Entrar</h1><form method="post"><label>E-mail</label><input name="email" type="email" required><label>Senha</label><input name="password" type="password" required><button class="btn">Entrar</button></form></div>''')
@app.post('/login')
def lp(request:Request,email:str=Form(...),password:str=Form(...)):
 c=conn();u=c.execute('SELECT * FROM users WHERE email=?',(email.lower(),)).fetchone();c.close()
 if not u or hp(password)!=u['password']: return HTMLResponse(page('Erro','<div class="card"><h2>Login inválido.</h2></div>'),401)
 request.session['uid']=u['id'];return RedirectResponse('/dashboard',303)
@app.get('/logout')
def lo(req:Request):req.session.clear();return RedirectResponse('/',303)
@app.get('/dashboard',response_class=HTMLResponse)
def dash(req:Request):
 u=auth(req);c=conn();ds=c.execute('SELECT * FROM dashboards WHERE tenant_id=? ORDER BY id DESC',(u['tenant_id'],)).fetchall();c.close();cards=''.join(f'<div class="card"><h3>📊 {esc(x["name"])}</h3><p class="muted">{x["created_at"][:10]}</p><a class="btn" href="/dashboards/{x["id"]}">Abrir</a> <a href="/download/{x["id"]}">Baixar</a></div>' for x in ds) or '<div class="card"><h3>Nenhum dashboard ainda</h3><p>Envie sua primeira base.</p></div>'
 return page('Visão geral',f'<h1>Olá, {esc(u["name"])} 👋</h1><p class="muted">Ambiente privado da sua empresa.</p><p><a class="btn" href="/upload">＋ Novo dashboard</a></p><div class="grid">{cards}</div>',u)
@app.get('/upload',response_class=HTMLResponse)
def ug(req:Request):
 u=auth(req);return page('Importar','''<div class="form card"><h1>📤 Novo dashboard</h1><form method="post" enctype="multipart/form-data"><label>Nome</label><input name="dash_name" required placeholder="Vendas 2026"><label>Arquivo</label><input type="file" name="file" accept=".xlsx,.xls,.csv,.pdf,.docx" required><button class="btn">Processar →</button></form></div>''',u)
@app.post('/upload')
async def up(req:Request,file:UploadFile=File(...),dash_name:str=Form(...)):
 u=auth(req);data=await file.read()
 if len(data)>20*1024*1024:raise HTTPException(413,'Arquivo maior que 20 MB')
 try:df=clean(parse(file.filename,data))
 except Exception as e:raise HTTPException(400,str(e))
 nums=df.select_dtypes('number').columns.tolist();dates=df.select_dtypes('datetime').columns.tolist();cats=[c for c in df.columns if c not in nums+dates]
 if not nums:raise HTTPException(400,'Nenhuma métrica numérica identificada.')
 words=['venda','fatur','receita','valor','preco','preço','lucro','quant','qtd','total','custo'];metric=max(nums,key=lambda c:(sum(w in c for w in words),df[c].notna().sum()));dim=max(cats,key=lambda c:(sum(w in c for w in ['produto','cliente','categoria','regiao','região','vendedor','marca']),-df[c].nunique())) if cats else None;date=dates[0] if dates else None
 for c in dates:df[c]=df[c].astype(str)
 now=datetime.utcnow().isoformat();c=conn();did=c.execute('INSERT INTO dashboards(tenant_id,user_id,name,metric,dimension,date_col,data_json,created_at) VALUES(?,?,?,?,?,?,?,?)',(u['tenant_id'],u['id'],dash_name,metric,dim,date,df.to_json(orient='records'),now)).lastrowid;c.execute('INSERT INTO audit(tenant_id,user_id,action,created_at) VALUES(?,?,?,?)',(u['tenant_id'],u['id'],f'created:{did}',now));c.commit();c.close();return RedirectResponse(f'/dashboards/{did}',303)
@app.get('/dashboards/{did}',response_class=HTMLResponse)
def view(did:int,req:Request):
 u=auth(req);c=conn();r=c.execute('SELECT * FROM dashboards WHERE id=? AND tenant_id=?',(did,u['tenant_id'])).fetchone();c.close()
 if not r:raise HTTPException(404,'Dashboard não encontrado')
 df=pd.read_json(io.StringIO(r['data_json']));m=r['metric'];d=r['dimension'];date=r['date_col'];total=df[m].sum();avg=df[m].mean();mx=df[m].max();body=f'<h1>{esc(r["name"])}</h1><p class="muted">Dashboard privado</p><div class="grid"><div class="card"><span class="muted">Total</span><div class="big">{total:,.2f}</div></div><div class="card"><span class="muted">Média</span><div class="big">{avg:,.2f}</div></div><div class="card"><span class="muted">Máximo</span><div class="big">{mx:,.2f}</div></div><div class="card"><span class="muted">Registros</span><div class="big">{len(df):,}</div></div></div><p><a class="btn" href="/download/{did}">⬇ Baixar dados</a></p>'
 if date:
  t=df[[date,m]].dropna();t[date]=pd.to_datetime(t[date]);t['p']=t[date].dt.to_period('M').dt.to_timestamp();g=t.groupby('p',as_index=False)[m].sum();body+=f'<div class="card"><h2>📈 Evolução</h2><div id="c1" class="chart"></div></div><script>Plotly.newPlot("c1",[{{x:{json.dumps(g.p.astype(str).tolist())},y:{json.dumps(g[m].tolist())},mode:"lines+markers",fill:"tozeroy"}}],{{template:"plotly_white"}})</script>'
 if d:
  g=df.groupby(d)[m].sum().sort_values(ascending=False).head(12).reset_index();body+=f'<div class="card"><h2>🏆 Ranking</h2><div id="c2" class="chart"></div></div><script>Plotly.newPlot("c2",[{{x:{json.dumps(g[m].tolist())},y:{json.dumps(g[d].astype(str).tolist())},type:"bar",orientation:"h"}}],{{template:"plotly_white",yaxis:{{autorange:"reversed"}}}})</script>'
 return page(r['name'],body,u)
@app.get('/download/{did}')
def dl(did:int,req:Request):
 u=auth(req);c=conn();r=c.execute('SELECT * FROM dashboards WHERE id=? AND tenant_id=?',(did,u['tenant_id'])).fetchone();c.close()
 if not r:raise HTTPException(404,'Dashboard não encontrado')
 df=pd.read_json(io.StringIO(r['data_json']));b=df.to_csv(index=False).encode('utf-8-sig');return StreamingResponse(io.BytesIO(b),media_type='text/csv',headers={'Content-Disposition':f'attachment; filename="insightdash_{did}.csv"'})
