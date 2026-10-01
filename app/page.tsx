"use client";

import { FormEvent, useState } from "react";

export default function Home() {
  const [status, setStatus] = useState("");
  const [error, setError] = useState(false);
  const [busy, setBusy] = useState(false);
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);

  async function processFiles(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (downloadUrl) URL.revokeObjectURL(downloadUrl);
    setDownloadUrl(null);
    setError(false);
    setBusy(true);
    setStatus("Uploading and processing the supplier reports…");
    const form = new FormData(event.currentTarget);

    try {
      const response = await fetch("/api/process", { method: "POST", body: form });
      if (!response.ok) {
        const raw = await response.text();
        let message = `Processing failed (${response.status}).`;
        try {
          const parsed = JSON.parse(raw);
          message = parsed.detail || message;
        } catch {
          if (raw.trim()) message = raw.slice(0, 500);
        }
        throw new Error(message);
      }
      const blob = await response.blob();
      if (!blob.size) throw new Error("The server returned an empty Excel file.");
      const url = URL.createObjectURL(blob);
      setDownloadUrl(url);
      setStatus("Done. Your updated tracker is ready.");
      const a = document.createElement("a");
      a.href = url;
      a.download = "Uniform Daily Sales Tracker - Updated.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();
    } catch (err) {
      setError(true);
      setStatus(err instanceof Error ? err.message : "Processing failed.");
    } finally {
      setBusy(false);
    }
  }

  return (
    <main>
      <header className="topbar">
        <div className="brand">AutoTask</div>
        <div className="pill">Automation Workspace</div>
      </header>

      <section className="workspace">
        <div className="pageHeading">
          <div>
            <p className="eyebrow">UNIFORM SALES</p>
            <h1>Zona Daily Sales</h1>
          </div>
          <p className="intro">Update the Uniform Daily Sales Tracker from Zona reports.</p>
        </div>

        <section className="grid">
          <form className="card" onSubmit={processFiles}>
            <div className="fieldBlock">
              <div className="fieldTitle"><span className="step">1</span><div><h2>Tracker</h2><p>Select the current Uniform Daily Sales Tracker.</p></div></div>
              <input name="target" type="file" accept=".xlsx" required disabled={busy} />
            </div>

            <div className="divider" />

            <div className="fieldBlock">
              <div className="fieldTitle"><span className="step">2</span><div><h2>Zona report(s)</h2><p>Combined Sales Report, or the four school reports.</p></div></div>
              <input name="sources" type="file" accept=".xlsx" multiple required disabled={busy} />
            </div>

            <button className="processButton" disabled={busy} type="submit">
              {busy && <span className="spinner" aria-hidden="true" />}
              <span>{busy ? "Processing…" : "Process & Download"}</span>
            </button>

            {status && <div className={`status ${error ? "statusError" : downloadUrl ? "statusSuccess" : ""}`}>{status}</div>}
            {downloadUrl && !busy && <a className="downloadButton" href={downloadUrl} download="Uniform Daily Sales Tracker - Updated.xlsx">Download Updated Tracker</a>}
          </form>

          <aside className="card side">
            <div className="statusRow"><span className="dot"/><div><strong>Microsoft 365</strong><small>Not connected yet</small></div></div>
            <div className="divider compact" />
            <h3>Next: automatic mode</h3>
            <p>Connect Outlook, find the latest Zona sales email and update the tracker automatically.</p>
          </aside>
        </section>
      </section>
    </main>
  );
}
