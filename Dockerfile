# --- Stage 1: build the React frontend ---
FROM node:22-slim AS frontend
WORKDIR /app/frontend
COPY frontend/package.json ./
RUN npm install
COPY frontend/ ./
RUN npm run build

# --- Stage 2: Python runtime with LibreOffice (for PDF rendering) ---
FROM python:3.12-slim
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-calc \
        fonts-dejavu \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY backend/requirements.txt backend/requirements.txt
RUN pip install --no-cache-dir -r backend/requirements.txt

COPY generate_factures.py ./generate_factures.py
COPY backend/ ./backend/
COPY --from=frontend /app/frontend/dist ./frontend/dist

EXPOSE 8000
CMD ["uvicorn", "backend.app:app", "--host", "0.0.0.0", "--port", "8000"]
