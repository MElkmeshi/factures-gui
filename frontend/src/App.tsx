import { useEffect, useRef, useState, type FormEvent } from "react";

type Phase = "idle" | "running" | "done" | "error";

export default function App() {
  const [file, setFile] = useState<File | null>(null);
  const [year, setYear] = useState(2025);
  const [pdf, setPdf] = useState(true);
  const [excel, setExcel] = useState(true);
  const [keepXlsx, setKeepXlsx] = useState(false);

  const [phase, setPhase] = useState<Phase>("idle");
  const [logs, setLogs] = useState<string[]>([]);
  const [jobId, setJobId] = useState<string | null>(null);
  const [canDownload, setCanDownload] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const logRef = useRef<HTMLDivElement>(null);
  const esRef = useRef<EventSource | null>(null);

  useEffect(() => {
    logRef.current?.scrollTo(0, logRef.current.scrollHeight);
  }, [logs]);

  // Tidy up the SSE connection if the component unmounts mid-run.
  useEffect(() => () => esRef.current?.close(), []);

  async function submit(e: FormEvent) {
    e.preventDefault();
    if (!file) {
      setError("Veuillez choisir un classeur .xlsx.");
      return;
    }
    if (!pdf && !excel) {
      setError("Choisissez au moins un type de sortie (PDF ou Excel).");
      return;
    }

    setPhase("running");
    setLogs([]);
    setError(null);
    setCanDownload(false);

    const fd = new FormData();
    fd.append("file", file);
    fd.append("year", String(year));
    fd.append("pdf", String(pdf));
    fd.append("excel", String(excel));
    fd.append("keep_xlsx", String(keepXlsx));

    let id: string;
    try {
      const res = await fetch("/api/jobs", { method: "POST", body: fd });
      if (!res.ok) {
        const body = await res.json().catch(() => ({}));
        throw new Error(body.detail || res.statusText);
      }
      id = (await res.json()).job_id;
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setPhase("error");
      return;
    }

    setJobId(id);
    const es = new EventSource(`/api/jobs/${id}/stream`);
    esRef.current = es;
    es.onmessage = (ev) => {
      const msg = JSON.parse(ev.data);
      if (msg.type === "log") {
        setLogs((prev) => [...prev, msg.line]);
      } else if (msg.type === "done") {
        setCanDownload(Boolean(msg.download));
        setPhase("done");
        es.close();
      } else if (msg.type === "error") {
        setError(msg.message);
        setPhase("error");
        es.close();
      }
    };
    es.onerror = () => es.close();
  }

  function reset() {
    esRef.current?.close();
    setPhase("idle");
    setLogs([]);
    setJobId(null);
    setCanDownload(false);
    setError(null);
  }

  const busy = phase === "running";

  return (
    <div className="min-h-screen bg-slate-100 text-slate-800 flex items-center justify-center p-4">
      <div className="w-full max-w-2xl bg-white rounded-2xl shadow-lg ring-1 ring-slate-200 overflow-hidden">
        <header className="bg-slate-900 px-6 py-5">
          <h1 className="text-xl font-semibold text-white">Générateur de factures</h1>
          <p className="text-sm text-slate-300 mt-1">
            Importez le classeur des livreurs, choisissez les options, téléchargez les factures.
          </p>
        </header>

        <form onSubmit={submit} className="px-6 py-6 space-y-5">
          {/* File */}
          <div>
            <label className="block text-sm font-medium mb-1">Classeur (.xlsx)</label>
            <input
              type="file"
              accept=".xlsx"
              disabled={busy}
              onChange={(e) => setFile(e.target.files?.[0] ?? null)}
              className="block w-full text-sm text-slate-600 file:mr-4 file:rounded-lg file:border-0
                         file:bg-slate-900 file:px-4 file:py-2 file:text-white file:cursor-pointer
                         hover:file:bg-slate-700 disabled:opacity-50"
            />
          </div>

          {/* Year */}
          <div>
            <label className="block text-sm font-medium mb-1">Année</label>
            <input
              type="number"
              min={2020}
              max={2099}
              value={year}
              disabled={busy}
              onChange={(e) => setYear(Number(e.target.value))}
              className="w-32 rounded-lg border border-slate-300 px-3 py-2 text-sm
                         focus:outline-none focus:ring-2 focus:ring-slate-400 disabled:opacity-50"
            />
          </div>

          {/* Options */}
          <fieldset className="space-y-2" disabled={busy}>
            <legend className="text-sm font-medium mb-1">Options</legend>
            <Checkbox label="Générer les PDF (un par livreur)" checked={pdf} onChange={setPdf} />
            <Checkbox label="Générer le classeur Excel combiné" checked={excel} onChange={setExcel} />
            <Checkbox label="Conserver les .xlsx intermédiaires" checked={keepXlsx} onChange={setKeepXlsx} />
          </fieldset>

          <div className="flex items-center gap-3 pt-1">
            <button
              type="submit"
              disabled={busy}
              className="rounded-lg bg-slate-900 px-5 py-2.5 text-sm font-medium text-white
                         hover:bg-slate-700 disabled:opacity-50 disabled:cursor-not-allowed"
            >
              {busy ? "Génération…" : "Générer"}
            </button>
            {phase !== "idle" && !busy && (
              <button
                type="button"
                onClick={reset}
                className="rounded-lg border border-slate-300 px-5 py-2.5 text-sm font-medium
                           text-slate-700 hover:bg-slate-50"
              >
                Nouveau
              </button>
            )}
            {canDownload && jobId && (
              <a
                href={`/api/jobs/${jobId}/download`}
                className="ml-auto rounded-lg bg-emerald-600 px-5 py-2.5 text-sm font-medium text-white
                           hover:bg-emerald-500"
              >
                Télécharger (.zip)
              </a>
            )}
          </div>

          {error && (
            <p className="rounded-lg bg-red-50 px-4 py-3 text-sm text-red-700 ring-1 ring-red-200">
              {error}
            </p>
          )}

          {(phase === "running" || logs.length > 0) && (
            <div
              ref={logRef}
              className="h-56 overflow-auto rounded-lg bg-slate-900 px-4 py-3 font-mono text-xs
                         leading-relaxed text-slate-100"
            >
              {logs.map((line, i) => (
                <div key={i} className="whitespace-pre-wrap">{line}</div>
              ))}
              {busy && <div className="animate-pulse text-slate-400">…</div>}
            </div>
          )}
        </form>
      </div>
    </div>
  );
}

function Checkbox({
  label,
  checked,
  onChange,
}: {
  label: string;
  checked: boolean;
  onChange: (v: boolean) => void;
}) {
  return (
    <label className="flex items-center gap-2 text-sm cursor-pointer select-none">
      <input
        type="checkbox"
        checked={checked}
        onChange={(e) => onChange(e.target.checked)}
        className="h-4 w-4 rounded border-slate-300 text-slate-900 focus:ring-slate-400"
      />
      {label}
    </label>
  );
}
