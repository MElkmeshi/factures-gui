# Générateur de factures — web app

Upload the drivers workbook, pick options, download the generated invoices
(PDF per driver and/or a combined Excel). FastAPI backend + React (Vite +
Tailwind) frontend, packaged in one Docker image with LibreOffice for PDF
rendering. Public — no login.

## Run with Docker (production-like)

```bash
docker compose up --build
# open http://localhost:8000
```

The container bundles the built frontend, the API, and LibreOffice.

## Local development (hot reload)

Two terminals:

```bash
# 1) backend
python3 -m venv .venv && source .venv/bin/activate
pip install -r backend/requirements.txt
uvicorn backend.app:app --reload          # http://localhost:8000
```

```bash
# 2) frontend
cd frontend
npm install
npm run dev                               # http://localhost:5173 (proxies /api -> :8000)
```

LibreOffice must be installed locally for PDF output (`brew install --cask libreoffice`).
The combined-Excel output works without it.

## How it works

- `generate_factures.py` — the engine. `generate(...)` does the work; also usable as a CLI.
- The invoice template (the `Exemple` sheet, including the Presto logo) is hardcoded in
  `factures_template.py`, so an uploaded workbook only needs a `Sheet1`. Regenerate it with
  `python tools/build_template.py "Factures Livreurs.xlsx"` (needs Pillow) if the design changes.
- `backend/app.py` — wraps `generate()`: `POST /api/jobs` (upload) → background job →
  `GET /api/jobs/{id}/stream` (SSE live log) → `GET /api/jobs/{id}/download` (zip).
- Each job gets an unguessable id; uploaded files + output are deleted after one hour.

## Deploying

Any host that runs a container works (Fly.io, Render, a VPS with Docker, etc.).
Put it behind HTTPS at the proxy/host level.
