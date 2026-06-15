import React, { useState, useEffect } from "react";

const API = ""; // proxied by Vite to backend

function formatINR(n) {
  if (n == null || isNaN(n)) return "—";
  return "₹" + Number(n).toLocaleString("en-IN", { maximumFractionDigits: 2 });
}

async function fetchQuoteRef() {
  const r = await fetch(`${API}/api/quote-ref`);
  if (!r.ok) throw new Error("Failed to get quotation ID");
  const j = await r.json();
  return j.quoteRef;
}

// ---------------------------------------------------------------------------
// Static options
// ---------------------------------------------------------------------------

const SYSTEM_OPTIONS = [
  {
    id: "CPVC + UPVC + SWR",
    label: "CPVC + UPVC + SWR",
    tag: "Hot + cold piping with drainage",
    imgs: ["Sintex-Hotx.jpg", "Sintex_coolx.jpg", "Sintex_SWR.png"],
  },
  {
    id: "CPVC + SWR",
    label: "CPVC + SWR",
    tag: "Hot water piping with drainage",
    imgs: ["Sintex-Hotx.jpg", "Sintex_SWR.png"],
  },
  {
    id: "UPVC + SWR",
    label: "UPVC + SWR",
    tag: "Cold water piping with drainage",
    imgs: ["Sintex_coolx.jpg", "Sintex_SWR.png"],
  },
];

const PIPING_OPTIONS = [
  { id: "Single Lever Diverter / Flush Valve", label: "Single Lever Diverter / Flush Valve", img: "Single_Lever_Diverter.jpg" },
  { id: "Wall Mixer",                          label: "Wall Mixer Tap",                       img: "Wall_Mixer.jpg" },
];

const SHAPE_OPTIONS = ["L Shape", "C Shape"];

function groupImageUrl(group) {
  return `${API}/api/group-image/${encodeURIComponent(group)}`;
}
function systemImageUrl(file) {
  return `${API}/api/asset/${encodeURIComponent(file)}`;
}

// ---------------------------------------------------------------------------
// NumberStepper — empty-state safe (first +/- click sets to min)
// ---------------------------------------------------------------------------

function NumberStepper({ value, onChange, min = 1, max = 999 }) {
  const currentNum = value === "" || value == null ? null : parseInt(value, 10);
  const dec = () => {
    if (currentNum == null) { onChange(String(min)); return; }
    onChange(String(Math.max(min, currentNum - 1)));
  };
  const inc = () => {
    if (currentNum == null) { onChange(String(min)); return; }
    onChange(String(Math.min(max, currentNum + 1)));
  };
  return (
    <div className="stepper">
      <button type="button" className="step-btn" onClick={dec} aria-label="decrease">−</button>
      <input
        type="number"
        className="step-input"
        value={value}
        placeholder="—"
        onChange={(e) => {
          const raw = e.target.value;
          if (raw === "") { onChange(""); return; }
          const v = parseInt(raw, 10);
          if (!isNaN(v) && v >= min && v <= max) onChange(String(v));
        }}
        min={min}
        max={max}
      />
      <button type="button" className="step-btn" onClick={inc} aria-label="increase">+</button>
    </div>
  );
}

// ---------------------------------------------------------------------------
// SelectionForm
// ---------------------------------------------------------------------------

function emptyBathroom() {
  return { shape: "", length: "", width: "" };
}

function SelectionForm({ onCalculate, error }) {
  const [system, setSystem] = useState(null);
  // Bathrooms: a list of objects. Length always = bathroomCount.
  // We use a separate `bathroomCount` so the user can change it without
  // immediately losing in-progress data; we sync on edit.
  const [bathroomCount, setBathroomCount] = useState("1"); // kitchens default 1, bathrooms also default 1
  const [bathrooms, setBathrooms] = useState([emptyBathroom()]);
  const [kitchens, setKitchens] = useState("1");
  const [pipingSystem, setPipingSystem] = useState(null);
  const [localError, setLocalError] = useState(null);

  // Keep bathroom array length synced with the bathroomCount value
  const syncBathroomCount = (raw) => {
    setBathroomCount(raw);
    const n = parseInt(raw, 10);
    if (!isNaN(n) && n >= 1) {
      setBathrooms((prev) => {
        if (n === prev.length) return prev;
        if (n > prev.length) return [...prev, ...Array.from({ length: n - prev.length }, emptyBathroom)];
        return prev.slice(0, n);
      });
    }
  };

  const setBathroomField = (idx, field, value) => {
    setBathrooms((prev) => prev.map((b, i) => i === idx ? { ...b, [field]: value } : b));
  };

  const submit = (e) => {
    e.preventDefault();
    setLocalError(null);

    if (!system) return setLocalError("Please select a system.");
    if (!pipingSystem) return setLocalError("Please select a piping system.");

    const B = parseInt(bathroomCount, 10);
    if (!Number.isInteger(B) || B < 1) return setLocalError("Please enter the number of bathrooms (at least 1).");

    for (let i = 0; i < bathrooms.length; i++) {
      const b = bathrooms[i];
      if (!b.shape) return setLocalError(`Bathroom ${i + 1}: please select a shape.`);
      const L = Number(b.length), W = Number(b.width);
      if (!(L > 0)) return setLocalError(`Bathroom ${i + 1}: please enter a length in feet.`);
      if (!(W > 0)) return setLocalError(`Bathroom ${i + 1}: please enter a width in feet.`);
    }

    const K = parseInt(kitchens, 10);
    if (!Number.isInteger(K) || K < 1) return setLocalError("Please enter the number of kitchens (at least 1).");

    onCalculate({
      system,
      pipingSystem,
      bathrooms: bathrooms.map((b) => ({
        shape: b.shape,
        length: Number(b.length),
        width: Number(b.width),
      })),
      kitchens: K,
    });
  };

  return (
    <form onSubmit={submit}>
      <div className="hero">
        <h1>Plan Your Plumbing.<br /><span className="hero-accent">Get an Instant Estimate.</span></h1>
        <p className="hero-sub">
          Enter your requirements and select a plumbing system to get your
          estimated plumbing value and pipe requirement.
        </p>
      </div>

      {(error || localError) && <div className="error">{error || localError}</div>}

      {/* Step 1 — System */}
      <div className="panel">
        <div className="step-row">
          <div className="step-num">1</div>
          <div>
            <h2>Select System <span className="req">*</span></h2>
            <div className="step-help">Choose the plumbing system that best suits your needs.</div>
          </div>
        </div>
        <div className="system-grid">
          {SYSTEM_OPTIONS.map((o) => (
            <label key={o.id} className={`system-card ${system === o.id ? "selected" : ""}`}>
              <input
                type="radio"
                name="system"
                value={o.id}
                checked={system === o.id}
                onChange={() => setSystem(o.id)}
              />
              <div className="system-card-radio" />
              <div className="system-card-title">{o.label}</div>
              <div className={`system-card-imgs count-${o.imgs.length}`}>
                {o.imgs.map((file, i) => (
                  <img
                    key={file + i}
                    className="system-card-img"
                    src={systemImageUrl(file)}
                    alt=""
                    onError={(e) => { e.target.style.visibility = "hidden"; }}
                  />
                ))}
              </div>
              <div className="system-card-tag">{o.tag}</div>
            </label>
          ))}
        </div>
      </div>

      {/* Step 2 — No. of Bathrooms */}
      <div className="panel">
        <div className="step-row">
          <div className="step-num">2</div>
          <div>
            <h2>No. of Bathrooms <span className="req">*</span></h2>
            <div className="step-help">Each bathroom is sized separately below.</div>
          </div>
          <div className="step-right">
            <NumberStepper value={bathroomCount} onChange={syncBathroomCount} min={1} />
          </div>
        </div>
      </div>

      {/* Step 3 — Per-bathroom shape + dimensions */}
      <div className="panel">
        <div className="step-row">
          <div className="step-num">3</div>
          <div>
            <h2>Bathroom Details <span className="req">*</span></h2>
            <div className="step-help">Shape, length and width for each bathroom (reference room is 8 × 4 ft).</div>
          </div>
        </div>
        <div className="bathroom-list">
          {bathrooms.map((b, idx) => (
            <div key={idx} className="bathroom-card">
              <div className="bathroom-card-head">Bathroom {idx + 1}</div>
              <div className="bathroom-card-body">
                <div className="bathroom-field">
                  <label>Shape</label>
                  <select
                    className="shape-select"
                    value={b.shape}
                    onChange={(e) => setBathroomField(idx, "shape", e.target.value)}
                  >
                    <option value="">Select shape…</option>
                    {SHAPE_OPTIONS.map((s) => <option key={s} value={s}>{s}</option>)}
                  </select>
                </div>
                <div className="bathroom-field">
                  <label>Length (ft)</label>
                  <input
                    type="number"
                    min="1"
                    step="0.5"
                    placeholder="e.g. 8"
                    value={b.length}
                    onChange={(e) => setBathroomField(idx, "length", e.target.value)}
                  />
                </div>
                <div className="bathroom-field">
                  <label>Width (ft)</label>
                  <input
                    type="number"
                    min="1"
                    step="0.5"
                    placeholder="e.g. 4"
                    value={b.width}
                    onChange={(e) => setBathroomField(idx, "width", e.target.value)}
                  />
                </div>
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* Step 4 — No. of Kitchens */}
      <div className="panel">
        <div className="step-row">
          <div className="step-num">4</div>
          <div>
            <h2>No. of Kitchens <span className="req">*</span></h2>
            <div className="step-help">
              Each kitchen is priced as one base 8 × 4 ft BoQ. No per-kitchen sizing is taken.
            </div>
          </div>
          <div className="step-right">
            <NumberStepper value={kitchens} onChange={setKitchens} min={1} />
          </div>
        </div>
      </div>

      {/* Step 5 — Piping system */}
      <div className="panel">
        <div className="step-row">
          <div className="step-num">5</div>
          <div>
            <h2>Select Hot &amp; Cold Piping System <span className="req">*</span></h2>
            <div className="step-help">Choose the piping system for your bathrooms.</div>
          </div>
        </div>
        <div className="piping-grid">
          {PIPING_OPTIONS.map((o) => (
            <label key={o.id} className={`piping-card ${pipingSystem === o.id ? "selected" : ""}`}>
              <input
                type="radio"
                name="piping"
                value={o.id}
                checked={pipingSystem === o.id}
                onChange={() => setPipingSystem(o.id)}
              />
              <div className="piping-card-radio" />
              <div className="piping-card-title">{o.label}</div>
              <img
                className="piping-card-img"
                src={systemImageUrl(o.img)}
                alt={o.label}
                onError={(e) => { e.target.style.visibility = "hidden"; }}
              />
            </label>
          ))}
        </div>
      </div>

      <button type="submit" className="btn-cta">
        <span className="calc-icon">▦</span> Calculate My Plumbing Value
      </button>
      <p className="security-note">🔒 Your information is secure and will not be shared.</p>
    </form>
  );
}

// ---------------------------------------------------------------------------
// Result screen
// ---------------------------------------------------------------------------

function ResultScreen({ quote, quoteRef, onBack, downloadBoq, downloadSummary, previewBoq, previewSummary, busy }) {
  const sel = quote.selection;

  return (
    <div>
      <div className="result-head">
        <div className="result-check">✓</div>
        <div>
          <h1>Here's Your Plumbing Estimate!</h1>
          <p className="hero-sub">
            Based on your input, here's the material summary and estimated value.
          </p>
        </div>
      </div>

      {/* Selection summary chips */}
      <div className="chips">
        <div className="chip"><span>System</span><b>{sel.system}</b></div>
        <div className="chip"><span>Piping</span><b>{sel.pipingSystem}</b></div>
        <div className="chip"><span>Bathrooms</span><b>{sel.bathrooms.length}</b></div>
        <div className="chip"><span>Kitchens</span><b>{sel.kitchens}</b></div>
        <div className="chip"><span>BOQ ID</span><b className="qid">{quoteRef}</b></div>
      </div>

      {/* Summary table — mirrors the second PDF */}
      <div className="groups-panel">
        <div className="groups-head">
          <div className="cell-mat">Item</div>
          <div className="cell-cost">Estimated Cost (₹)</div>
        </div>
        {quote.summaryRows.map((row, i) => {
          const label = row.kind === "Bathroom"
            ? `Bathroom — ${row.shape}, ${row.dimensionsLabel}`
            : "Kitchen";
          return (
            <div key={i} className="group-row" style={{ cursor: "default" }}>
              <div className="group-left">
                <div className="group-img-pair">
                  <span className="count-pill">×{row.count}</span>
                </div>
                <div>
                  <div className="group-name">{label}</div>
                  <div className="group-sub">{formatINR(row.perCost)} each</div>
                </div>
              </div>
              <div className="group-cost">{formatINR(row.total)}</div>
            </div>
          );
        })}
        <div className="group-total">
          <div>Total Plumbing Value</div>
          <div className="group-total-amt">{formatINR(quote.grandTotal)}</div>
        </div>
      </div>

      {/* Per-bathroom detail — expand to see each individual room */}
      <details className="bath-detail">
        <summary>Per-bathroom detail ({quote.bathroomSections.length} {quote.bathroomSections.length === 1 ? "room" : "rooms"})</summary>
        <table className="bath-detail-table">
          <thead>
            <tr>
              <th>#</th>
              <th>Shape</th>
              <th>Dimensions</th>
              <th className="num">Uplift</th>
              <th className="num">Subtotal</th>
            </tr>
          </thead>
          <tbody>
            {quote.bathroomSections.map((b) => (
              <tr key={b.index}>
                <td>{b.index}</td>
                <td>{b.shape}</td>
                <td>{b.dimensionsLabel}</td>
                <td className="num">× {b.multiplier}</td>
                <td className="num">{formatINR(b.subtotal)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </details>

      <div className="result-actions">
        <button className="btn-secondary" onClick={onBack}>← Back &amp; Edit</button>
        <div className="doc-pair">
          <button className="btn-preview" onClick={previewBoq} disabled={busy} type="button">
            <span className="btn-preview-icon" aria-hidden="true">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12Z"/><circle cx="12" cy="12" r="3"/>
              </svg>
            </span>
            Preview BOQ
          </button>
          <button className="btn-cta" onClick={downloadBoq} disabled={busy}>
            {busy ? "Generating…" : "Download BOQ (PDF)"}
          </button>
        </div>
        <div className="doc-pair">
          <button className="btn-preview" onClick={previewSummary} disabled={busy} type="button">
            <span className="btn-preview-icon" aria-hidden="true">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12Z"/><circle cx="12" cy="12" r="3"/>
              </svg>
            </span>
            Preview Summary
          </button>
          <button className="btn-cta btn-cta-alt" onClick={downloadSummary} disabled={busy}>
            {busy ? "Generating…" : "Download Summary (PDF)"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// PDF Preview Modal — shows the generated PDF inline using an <iframe>.
// Works on desktop and mobile browsers. For browsers that don't render PDFs
// inline (some older mobile browsers), an "Open in new tab" link is provided
// as a reliable fallback.
// ---------------------------------------------------------------------------

function PdfPreviewModal({ preview, onClose }) {
  // Close on Escape key
  useEffect(() => {
    if (!preview) return;
    const onKey = (e) => { if (e.key === "Escape") onClose(); };
    window.addEventListener("keydown", onKey);
    // Lock background scrolling while open
    const prevOverflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => {
      window.removeEventListener("keydown", onKey);
      document.body.style.overflow = prevOverflow;
    };
  }, [preview, onClose]);

  if (!preview) return null;

  const title = preview.kind === "boq" ? "BOQ Preview" : "Summary Preview";

  return (
    <div
      className="pdf-preview-overlay"
      onClick={onClose}
      role="dialog"
      aria-modal="true"
      aria-label={title}
    >
      <div className="pdf-preview-modal" onClick={(e) => e.stopPropagation()}>
        <div className="pdf-preview-head">
          <div className="pdf-preview-title">{title}</div>
          <div className="pdf-preview-actions">
            <a
              href={preview.url}
              target="_blank"
              rel="noopener noreferrer"
              className="pdf-preview-link"
            >
              Open in new tab
            </a>
            <a
              href={preview.url}
              download={preview.filename}
              className="pdf-preview-link pdf-preview-link-primary"
            >
              Download
            </a>
            <button
              className="pdf-preview-close"
              onClick={onClose}
              aria-label="Close preview"
              type="button"
            >
              ×
            </button>
          </div>
        </div>
        <iframe
          src={preview.url}
          className="pdf-preview-iframe"
          title={title}
        />
        <div className="pdf-preview-foot">
          <span>If the document doesn't appear above, tap “Open in new tab”.</span>
        </div>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Root
// ---------------------------------------------------------------------------

export default function App() {
  const [step, setStep] = useState("form");
  const [error, setError] = useState(null);
  const [quote, setQuote] = useState(null);
  const [quoteRef, setQuoteRef] = useState(null);
  const [payload, setPayload] = useState(null);
  const [busy, setBusy] = useState(false);
  const [preview, setPreview] = useState(null); // { url, kind, filename } | null

  // Revoke the blob URL whenever the preview changes (or unmounts) so we
  // don't leak memory across multiple opens.
  useEffect(() => {
    return () => {
      if (preview?.url) URL.revokeObjectURL(preview.url);
    };
  }, [preview]);

  const closePreview = () => setPreview(null);

  const onCalculate = async (body) => {
    setError(null);
    try {
      const r = await fetch(`${API}/api/quote`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });
      const j = await r.json();
      if (!r.ok) { setError(j.error || "Calculation failed"); return; }
      const ref = await fetchQuoteRef();
      setQuote(j);
      setQuoteRef(ref);
      setPayload(body);
      setStep("result");
      window.scrollTo({ top: 0, behavior: "smooth" });
    } catch (e) {
      setError(String(e.message || e));
    }
  };

  const download = async (kind) => {
    setBusy(true);
    setError(null);
    try {
      const url = kind === "boq" ? "/api/quote/boq-pdf" : "/api/quote/summary-pdf";
      const r = await fetch(`${API}${url}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ...payload, quoteRef }),
      });
      if (!r.ok) {
        const j = await r.json().catch(() => ({}));
        setError(j.error || "PDF export failed");
        return;
      }
      const blob = await r.blob();
      const u = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = u;
      const dispo = r.headers.get("Content-Disposition") || "";
      const m = dispo.match(/filename="([^"]+)"/);
      a.download = m ? m[1] : `${kind === "boq" ? "BOQ" : "Estimate"}_${quoteRef}.pdf`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(u);
    } finally {
      setBusy(false);
    }
  };

  // Fetch the PDF and open it in the in-app preview modal. Uses the same
  // backend endpoints as `download`, but presents the result inline instead
  // of triggering a save. The previous preview's blob URL (if any) is
  // revoked by the cleanup effect on the `preview` state.
  const previewDoc = async (kind) => {
    setBusy(true);
    setError(null);
    try {
      const url = kind === "boq" ? "/api/quote/boq-pdf" : "/api/quote/summary-pdf";
      const r = await fetch(`${API}${url}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ...payload, quoteRef }),
      });
      if (!r.ok) {
        const j = await r.json().catch(() => ({}));
        setError(j.error || "PDF preview failed");
        return;
      }
      const blob = await r.blob();
      const u = URL.createObjectURL(blob);
      const dispo = r.headers.get("Content-Disposition") || "";
      const m = dispo.match(/filename="([^"]+)"/);
      const filename = m ? m[1] : `${kind === "boq" ? "BOQ" : "Estimate"}_${quoteRef}.pdf`;
      setPreview({ url: u, kind, filename });
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="app">
      <div className="topbar">
        <div className="brand">Sintex<span> by Welspun</span></div>
        <div className="brand-tag">Plumbing Contractor Calculator</div>
      </div>

      {step === "form" && <SelectionForm onCalculate={onCalculate} error={error} />}
      {step === "result" && (
        <ResultScreen
          quote={quote}
          quoteRef={quoteRef}
          onBack={() => { setStep("form"); setError(null); }}
          downloadBoq={() => download("boq")}
          downloadSummary={() => download("summary")}
          previewBoq={() => previewDoc("boq")}
          previewSummary={() => previewDoc("summary")}
          busy={busy}
        />
      )}

      <PdfPreviewModal preview={preview} onClose={closePreview} />
    </div>
  );
}
