"use client";

import { FormEvent, useState } from "react";

export default function Home() {
  const [status, setStatus] = useState("");
  const [busy, setBusy] = useState(false);

  async function processFiles(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setBusy(true);
    setStatus("Processing supplier reports…");
    const form = new FormData(event.currentTarget);

    try {
      const response = await fetch("/api/process", { method: "POST", body: form });
      if (!response.ok) {
        const error = await response.json().catch(() => ({ detail: "Processing failed." }));
        throw new Error(error.detail || "Processing failed.");
      }
      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "Uniform Daily Sales Tracker - Updated.xlsx";
      a.click();
      URL.revokeObjectURL(url);
      setStatus("Completed. The updated tracker has been downloaded.");
    } catch (error) {
      setStatus(error instanceof Error ? error.message : "Processing failed.");
    } finally {
      setBusy(false);
    }
  }

  return (
    <main>
      <header className="topbar"><div className="brand">AutoTask</div><div className="pill">Automation Workspace</div></header>
      <section className="hero">
        <p className="eyebrow">UNIFORM SALES</p>
        <h1>Zona Daily Sales Tracker</h1>
        <p className="intro">Process Zona supplier sales reports and update the matching dates in your Uniform Daily Sales Tracker.</p>
      </section>

      <section className="grid">
        <form className="card" onSubmit={processFiles}>
          <div className="step">1</div>
          <h2>Upload tracker</h2>
          <p>Your current Uniform Daily Sales Tracker workbook.</p>
          <input name="target" type="file" accept=".xlsx" required />

          <div className="divider" />
          <div className="step">2</div>
          <h2>Upload supplier report(s)</h2>
          <p>Use either the combined Sales Report, or the individual RDXB, RAB, FRY and ROSE files.</p>
          <input name="sources" type="file" accept=".xlsx" multiple required />

          <button disabled={busy} type="submit">{busy ? "Processing…" : "Process & Download"}</button>
          {status && <div className="status">{status}</div>}
        </form>

        <aside className="card side">
          <div className="statusRow"><span className="dot pending"/><div><strong>Microsoft 365</strong><small>Connection comes next</small></div></div>
          <div className="divider" />
          <h3>Processing rules</h3>
          <ul>
            <li>Matches dates to the correct monthly tracker sheet.</li>
            <li>Writes RDXB, RAB, FRY and ROSE daily sales.</li>
            <li>Combines Exchange / Shipping Value across all four schools.</li>
            <li>Uses a combined report when it contains all four campuses.</li>
            <li>Does not overwrite the original workbook on your computer.</li>
          </ul>
        </aside>
      </section>
    </main>
  );
}
