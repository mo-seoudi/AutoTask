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
    setDownloadUrl(null); setError(false); setBusy(true);
    setStatus("Uploading and processing the supplier reports…");
    const form = new FormData(event.currentTarget);
    try {
      const response = await fetch("/api/process", { method: "POST", body: form });
      if (!response.ok) {
        const raw = await response.text();
        let message = `Processing failed (${response.status}).`;
        try { message = JSON.parse(raw).detail || message; } catch { if (raw.trim()) message = raw.slice(0, 500); }
        throw new Error(message);
      }
      const blob = await response.blob();
      if (!blob.size) throw new Error("The server returned an empty Excel file.");
      const url = URL.createObjectURL(blob); setDownloadUrl(url); setStatus("Done. Your updated tracker is ready.");
      const a = document.createElement("a"); a.href = url; a.download = "Uniform Daily Sales Tracker - Updated.xlsx";
      document.body.appendChild(a); a.click(); a.remove();
    } catch (err) { setError(true); setStatus(err instanceof Error ? err.message : "Processing failed."); }
    finally { setBusy(false); }
  }

  return (
    <main>
      <header className="topbar">
        <div className="navInner">
          <div className="brandMark"><span className="logo">A</span><span>AutoTask</span></div>
          <span className="workspaceLabel">Workspace</span>
        </div>
      </header>

      <div className="shell">
        <section className="pageHeader">
          <div>
            <div className="breadcrumb">AUTOMATIONS <span>/</span> UNIFORM SALES</div>
            <h1>Zona Daily Sales</h1>
            <p>Update the Uniform Daily Sales Tracker from supplier reports.</p>
          </div>
          <div className="modeBadge"><span className="modeDot" /> Manual mode</div>
        </section>

        <section className="contentGrid">
          <form className="panel primaryPanel" onSubmit={processFiles}>
            <div className="panelHeader">
              <div><h2>Process reports</h2><p>Select the tracker and the latest Zona report files.</p></div>
            </div>

            <div className="uploadRow">
              <div className="number">1</div>
              <div className="uploadContent">
                <label>Daily Sales Tracker</label>
                <span className="hint">Current master workbook</span>
                <input name="target" type="file" accept=".xlsx" required disabled={busy} />
              </div>
            </div>

            <div className="rowDivider" />

            <div className="uploadRow">
              <div className="number">2</div>
              <div className="uploadContent">
                <label>Zona report(s)</label>
                <span className="hint">Combined report or the four school reports</span>
                <input name="sources" type="file" accept=".xlsx" multiple required disabled={busy} />
              </div>
            </div>

            <div className="actions">
              <button className="processButton" disabled={busy} type="submit">
                {busy && <span className="spinner" aria-hidden="true" />}
                <span>{busy ? "Processing…" : "Process & Download"}</span>
              </button>
              <span className="actionHint">Your original workbook stays unchanged.</span>
            </div>

            {status && <div className={`status ${error ? "statusError" : downloadUrl ? "statusSuccess" : ""}`}>{status}</div>}
            {downloadUrl && !busy && <a className="downloadButton" href={downloadUrl} download="Uniform Daily Sales Tracker - Updated.xlsx">Download again</a>}
          </form>

          <aside className="panel sidePanel">
            <div className="sideTop">
              <span className="msIcon">M</span>
              <div><h3>Microsoft 365</h3><p>Not connected</p></div>
              <span className="statusDot" />
            </div>
            <div className="sideDivider" />
            <div className="comingSoon">AUTOMATIC MODE</div>
            <h4>Remove the uploads</h4>
            <p className="sideCopy">Connect Outlook and OneDrive so AutoTask can find the latest Zona email and update the master tracker directly.</p>
            <button type="button" className="secondaryButton" disabled>Connect Microsoft 365</button>
          </aside>
        </section>
      </div>
    </main>
  );
}
