# InsightDash V2

SaaS web multiempresa para transformar Excel/CSV/PDF/Word em dashboards.

## Incluído
- Landing page e UX responsiva
- Cadastro de empresa e login
- Isolamento de dados por tenant_id
- Upload e tratamento de dados
- Detecção de métrica/dimensão/data
- KPIs, evolução e ranking
- Meus dashboards
- Download
- Auditoria básica
- Docker

## Rodar
python -m venv .venv
pip install -r requirements.txt
uvicorn app:app --reload

Abra http://127.0.0.1:8000

## Produção
Antes de dados empresariais reais: PostgreSQL + Row Level Security, storage privado, HTTPS, secrets reais, CSRF, rate limiting, RBAC, validação MIME/assinatura, antivírus/sandbox, migrations, backups e política de retenção. O SQLite desta V2 é para protótipo.
