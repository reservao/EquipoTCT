"use client";

import { useRef, useState } from "react";
import Link from "next/link";

interface ActivitySummary {
  title: string;
  criteriosCount: number;
}

interface Summary {
  nombre: string;
  perfiles: string;
  gerencia: string;
  superintendencia: string;
  area: string;
  codigo: string;
  activities: ActivitySummary[];
}

type Status = "idle" | "loading" | "done" | "error";

export default function SotPage() {
  const [status, setStatus] = useState<Status>("idle");
  const [error, setError] = useState<string | null>(null);
  const [summary, setSummary] = useState<Summary | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const downloadRef = useRef<{ base64: string; fileName: string } | null>(null);

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    if (!selectedFile) return;

    setStatus("loading");
    setError(null);
    setSummary(null);

    try {
      const formData = new FormData();
      formData.append("file", selectedFile);

      const res = await fetch("/api/sot/generate", { method: "POST", body: formData });
      const data = await res.json();

      if (!res.ok) {
        throw new Error(data.error ?? "No se pudo generar el SOT.");
      }

      downloadRef.current = { base64: data.fileBase64, fileName: data.fileName };
      setFileName(data.fileName);
      setSummary(data.summary);
      setStatus("done");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error desconocido.");
      setStatus("error");
    }
  }

  function handleDownload() {
    if (!downloadRef.current) return;
    const { base64, fileName } = downloadRef.current;
    const byteChars = atob(base64);
    const byteNumbers = new Array(byteChars.length);
    for (let i = 0; i < byteChars.length; i++) byteNumbers[i] = byteChars.charCodeAt(i);
    const byteArray = new Uint8Array(byteNumbers);
    const blob = new Blob([byteArray], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  return (
    <div className="mx-auto max-w-3xl px-6 py-10">
      <Link href="/" className="text-sm text-slate-500 hover:text-slate-700">
        ← Volver a herramientas
      </Link>

      <h1 className="mt-3 text-2xl font-semibold tracking-tight">
        SOT — Set de Observación en Terreno
      </h1>
      <p className="mt-2 text-sm text-slate-500">
        Sube el Word de la Competencia (UCL) y genera automáticamente el Excel del SOT, con una
        hoja por cada Actividad Clave y sus Criterios de Desempeño.
      </p>

      <form
        onSubmit={handleSubmit}
        className="mt-8 rounded-xl border border-slate-200 bg-white p-6 shadow-sm"
      >
        <label className="block text-sm font-medium text-slate-700">
          Documento Word de la Competencia (.docx)
        </label>
        <input
          type="file"
          accept=".docx"
          onChange={(e) => setSelectedFile(e.target.files?.[0] ?? null)}
          className="mt-2 block w-full text-sm text-slate-600 file:mr-4 file:rounded-md file:border-0 file:bg-slate-900 file:px-4 file:py-2 file:text-sm file:font-medium file:text-white hover:file:bg-slate-700"
        />

        <button
          type="submit"
          disabled={!selectedFile || status === "loading"}
          className="mt-4 rounded-md bg-slate-900 px-4 py-2 text-sm font-medium text-white disabled:cursor-not-allowed disabled:opacity-40"
        >
          {status === "loading" ? "Generando…" : "Generar SOT"}
        </button>

        {error && (
          <p className="mt-4 rounded-md bg-red-50 px-3 py-2 text-sm text-red-700">{error}</p>
        )}
      </form>

      {status === "done" && summary && (
        <div className="mt-6 rounded-xl border border-slate-200 bg-white p-6 shadow-sm">
          <div className="flex items-center justify-between">
            <h2 className="font-medium">Resumen extraído</h2>
            <button
              onClick={handleDownload}
              className="rounded-md bg-emerald-600 px-4 py-2 text-sm font-medium text-white hover:bg-emerald-700"
            >
              Descargar {fileName}
            </button>
          </div>

          <dl className="mt-4 grid grid-cols-1 gap-x-6 gap-y-2 text-sm sm:grid-cols-2">
            <div>
              <dt className="text-slate-400">Competencia</dt>
              <dd className="text-slate-800">{summary.nombre || "—"}</dd>
            </div>
            <div>
              <dt className="text-slate-400">Perfil(es)</dt>
              <dd className="text-slate-800">{summary.perfiles || "—"}</dd>
            </div>
            <div>
              <dt className="text-slate-400">Gerencia / Superintendencia / Área</dt>
              <dd className="text-slate-800">
                {[summary.gerencia, summary.superintendencia, summary.area]
                  .filter(Boolean)
                  .join(" / ") || "—"}
              </dd>
            </div>
            <div>
              <dt className="text-slate-400">Código</dt>
              <dd className="text-slate-800">{summary.codigo || "—"}</dd>
            </div>
          </dl>

          <h3 className="mt-5 text-sm font-medium text-slate-700">
            Actividades Clave ({summary.activities.length})
          </h3>
          <ul className="mt-2 space-y-2">
            {summary.activities.map((a, i) => (
              <li key={i} className="rounded-md bg-slate-50 px-3 py-2 text-sm">
                <span className="font-medium">
                  {i + 1}. {a.title}
                </span>
                <span className="ml-2 text-slate-400">
                  ({a.criteriosCount} criterio{a.criteriosCount === 1 ? "" : "s"} de desempeño)
                </span>
              </li>
            ))}
          </ul>
        </div>
      )}
    </div>
  );
}
