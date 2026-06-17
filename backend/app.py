"""FastAPI backend for the invoice (facture) generator.

Wraps generate_factures.generate() behind a small JSON/SSE API and (in
production) serves the built React frontend. No authentication — the service is
public — but each job gets an unguessable id and its uploaded file + output are
deleted after a TTL, so people can't reach each other's data by guessing URLs.
"""
import asyncio
import io
import json
import secrets
import shutil
import sys
import threading
import time
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import Response, StreamingResponse
from fastapi.staticfiles import StaticFiles

# The engine lives at the repo root, one level above this file.
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import generate_factures  # noqa: E402

JOBS_DIR = Path(__file__).resolve().parent / "_jobs"
JOBS_DIR.mkdir(exist_ok=True)
JOB_TTL_SECONDS = 60 * 60  # uploaded data + output wiped after one hour


@dataclass
class Job:
    id: str
    dir: Path
    status: str = "running"          # running | done | error
    error: str | None = None
    out_dir: Path | None = None
    created: float = field(default_factory=time.time)
    logs: list[str] = field(default_factory=list)


JOBS: dict[str, Job] = {}


def cleanup_old_jobs():
    now = time.time()
    for jid, job in list(JOBS.items()):
        if now - job.created > JOB_TTL_SECONDS:
            shutil.rmtree(job.dir, ignore_errors=True)
            JOBS.pop(jid, None)


def has_output(job: Job) -> bool:
    return bool(job.out_dir and job.out_dir.exists() and any(job.out_dir.iterdir()))


def run_job(job: Job, year: int, no_pdf: bool, no_excel: bool, keep_xlsx: bool):
    def log(msg):
        job.logs.append(str(msg))

    try:
        out = generate_factures.generate(
            workbook=str(job.dir / "input.xlsx"),
            year=year,
            out=str(job.dir / "output"),
            keep_xlsx=keep_xlsx,
            no_pdf=no_pdf,
            no_excel=no_excel,
            log=log,
        )
        job.out_dir = Path(out)
        job.status = "done"
        log("✅ Terminé.")
    except Exception as exc:  # noqa: BLE001
        job.status = "error"
        job.error = str(exc)
        log(f"❌ ERREUR: {exc}")


def sse(obj: dict) -> str:
    return f"data: {json.dumps(obj)}\n\n"


app = FastAPI(title="Générateur de factures")


@app.post("/api/jobs")
async def create_job(
    file: UploadFile = File(...),
    year: int = Form(2025),
    pdf: bool = Form(True),
    excel: bool = Form(True),
    keep_xlsx: bool = Form(False),
):
    cleanup_old_jobs()
    if not (pdf or excel):
        raise HTTPException(400, "Choose at least one output (PDF or Excel).")

    jid = secrets.token_urlsafe(16)
    jdir = JOBS_DIR / jid
    jdir.mkdir(parents=True)
    with open(jdir / "input.xlsx", "wb") as f:
        shutil.copyfileobj(file.file, f)

    job = Job(id=jid, dir=jdir)
    JOBS[jid] = job
    threading.Thread(
        target=run_job, args=(job, year, not pdf, not excel, keep_xlsx), daemon=True
    ).start()
    return {"job_id": jid}


@app.get("/api/jobs/{jid}/stream")
async def stream(jid: str):
    job = JOBS.get(jid)
    if not job:
        raise HTTPException(404, "job not found")

    async def gen():
        idx = 0
        while True:
            while idx < len(job.logs):
                yield sse({"type": "log", "line": job.logs[idx]})
                idx += 1
            if job.status != "running":
                if job.status == "error":
                    yield sse({"type": "error", "message": job.error})
                else:
                    yield sse({"type": "done", "download": has_output(job)})
                return
            await asyncio.sleep(0.3)

    return StreamingResponse(gen(), media_type="text/event-stream")


@app.get("/api/jobs/{jid}/download")
async def download(jid: str):
    job = JOBS.get(jid)
    if not job or job.status != "done" or not has_output(job):
        raise HTTPException(404, "no output available")

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for p in job.out_dir.rglob("*"):
            if p.is_file():
                z.write(p, p.relative_to(job.out_dir))
    buf.seek(0)
    return Response(
        buf.getvalue(),
        media_type="application/zip",
        headers={"Content-Disposition": f'attachment; filename="factures_{jid}.zip"'},
    )


# In production the built React app sits in frontend/dist and is served here.
# Mounted last so it never shadows the /api routes above.
DIST = ROOT / "frontend" / "dist"
if DIST.exists():
    app.mount("/", StaticFiles(directory=str(DIST), html=True), name="static")
