import { useEffect, useRef, useState, useCallback } from "react";
import { BrowserMultiFormatReader } from "@zxing/browser";
import * as XLSX from "xlsx";
import type { CellValue } from "./types";

/* ── BarcodeDetector type ─────────────────────────────────────────── */
interface BarcodeDetectorResult { rawValue: string; }
interface BarcodeDetectorConstructor {
  new (opts?: { formats: string[] }): { detect(src: HTMLVideoElement): Promise<BarcodeDetectorResult[]> };
  getSupportedFormats(): Promise<string[]>;
}
declare const BarcodeDetector: BarcodeDetectorConstructor;

/* ── LS helper ────────────────────────────────────────────────────── */
const LS = {
  get<T>(key: string, fallback: T): T {
    try { const v = localStorage.getItem(key); return v ? (JSON.parse(v) as T) : fallback; } catch { return fallback; }
  },
  set<T>(key: string, val: T) { try { localStorage.setItem(key, JSON.stringify(val)); } catch {} },
};

/* ── Types ────────────────────────────────────────────────────────── */
interface HangarItem  { code: string; weight: number | null; fromExcel: boolean; }
interface HangarLine  { id: string; name: string; items: HangarItem[]; }
interface HangarProps { dark: boolean; }

const MONO  = "'IBM Plex Mono', 'Fira Mono', monospace";
const newId = () => `hl_${Date.now()}_${Math.random().toString(36).slice(2,6)}`;
const fmtW  = (w: number | null) =>
  w === null ? "—" : w.toLocaleString("fr-FR", { minimumFractionDigits: 2, maximumFractionDigits: 3 });

/* ── Responsive hook ───────────────────────────────────────── */
function useWindowWidth() {
  const [width, setWidth] = useState(() => window.innerWidth);
  useEffect(() => {
    const handler = () => setWidth(window.innerWidth);
    window.addEventListener("resize", handler);
    return () => window.removeEventListener("resize", handler);
  }, []);
  return width;
}

export default function Hangar({ dark }: HangarProps) {
  const bg      = dark ? "#0d0d0d" : "#f5f5f5";
  const surface = dark ? "#141414" : "#ffffff";
  const text    = dark ? "#e8e8e0" : "#1a1a1a";
  const accent  = dark ? "#6ee7b7" : "#059669";
  const muted   = dark ? "#555"    : "#888";
  const border  = dark ? "#2a2a2a" : "#d0d0d0";
  const danger  = dark ? "#f87171" : "#dc2626";
  const amber   = "#d97706";

  /* ── Breakpoints ─────────────────────────────────────────────── */
  const vw        = useWindowWidth();
  const isMobile  = vw < 640;
  const isTablet  = vw >= 640 && vw < 1024;
  const isDesktop = vw >= 1024;

  /* ── Excel data ──────────────────────────────────────────────── */
  const [xlHeaders,  setXlHeaders]  = useState<string[]>([]);
  const [codeColIdx, setCodeColIdx] = useState<number>(() => LS.get("hgr_codeCol", 0));
  const [wtColIdx,   setWtColIdx]   = useState<number>(() => LS.get("hgr_wtCol",   1));

  const reloadXl = useCallback(() => {
    setXlHeaders(LS.get<string[]>("ptg_headers", []));
  }, []);
  useEffect(() => { reloadXl(); }, [reloadXl]);
  useEffect(() => { LS.set("hgr_codeCol", codeColIdx); }, [codeColIdx]);
  useEffect(() => { LS.set("hgr_wtCol",   wtColIdx);   }, [wtColIdx]);

  /* ── Lines ───────────────────────────────────────────────────── */
  const [lines,       setLines]       = useState<HangarLine[]>(() => LS.get("hgr_lines2", []));
  const [selectedId,  setSelectedId]  = useState<string | null>(() => LS.get("hgr_selectedId", null));
  const [newLineName, setNewLineName] = useState("");
  const [editingId,   setEditingId]   = useState<string | null>(null);
  const [editingName, setEditingName] = useState("");
  useEffect(() => { LS.set("hgr_lines2",     lines);      }, [lines]);
  useEffect(() => { LS.set("hgr_selectedId", selectedId); }, [selectedId]);

  /* ── Pending (unknown code → manual weight entry) ───────────── */
  const [pending, setPending] = useState<{ code: string; weight: string } | null>(null);
  const pendingRef = useRef<boolean>(false);
  useEffect(() => { pendingRef.current = pending !== null; }, [pending]);

  /* ── Scan delay ───────────────────────────────────────────────── */
  const [scanDelay, setScanDelay] = useState<number>(() => LS.get("hgr_scanDelay", 1800));
  const scanDelayRef = useRef<number>(scanDelay);
  useEffect(() => { scanDelayRef.current = scanDelay; LS.set("hgr_scanDelay", scanDelay); }, [scanDelay]);

  /* ── Camera ──────────────────────────────────────────────────── */
  const videoRef      = useRef<HTMLVideoElement>(null);
  const streamRef     = useRef<MediaStream | null>(null);
  const detectorRef   = useRef<{ detect(src: HTMLVideoElement): Promise<BarcodeDetectorResult[]> } | null>(null);
  const rafRef        = useRef<number | null>(null);
  const zxingReader   = useRef<BrowserMultiFormatReader | null>(null);
  const zxingControls = useRef<{ stop(): void } | null>(null);
  const assignRef     = useRef<(raw: string) => void>(() => {});
  const [cameraOn,    setCameraOn]    = useState(false);
  const [cameraError, setCameraError] = useState<string | null>(null);
  const [useNative,   setUseNative]   = useState(false);

  /* ── Audio feedback ────────────────────────────────────────── */
  const audioCtxRef = useRef<AudioContext | null>(null);
  const playTone = useCallback((type: "scan" | "found" | "error") => {
    try {
      if (!audioCtxRef.current) audioCtxRef.current = new AudioContext();
      const ctx = audioCtxRef.current;
      if (ctx.state === "suspended") void ctx.resume();
      const now = ctx.currentTime;
      if (type === "scan") {
        // Court bip agréable : onde sinusoïdale 1046 Hz, 70 ms
        const osc = ctx.createOscillator(); const g = ctx.createGain();
        osc.connect(g); g.connect(ctx.destination);
        osc.type = "sine"; osc.frequency.value = 1046;
        g.gain.setValueAtTime(0.25, now); g.gain.exponentialRampToValueAtTime(0.001, now + 0.07);
        osc.start(now); osc.stop(now + 0.07);
      } else if (type === "found") {
        // Double bip montant : 880 Hz puis 1320 Hz
        ([[880, 0], [1320, 0.11]] as [number, number][]).forEach(([freq, delay]) => {
          const osc = ctx.createOscillator(); const g = ctx.createGain();
          osc.connect(g); g.connect(ctx.destination);
          osc.type = "sine"; osc.frequency.value = freq;
          const t = now + delay;
          g.gain.setValueAtTime(0.22, t); g.gain.exponentialRampToValueAtTime(0.001, t + 0.18);
          osc.start(t); osc.stop(t + 0.18);
        });
      } else {
        // Buzz descendant : dents de scie 380→120 Hz, 300 ms
        const osc = ctx.createOscillator(); const g = ctx.createGain();
        osc.connect(g); g.connect(ctx.destination);
        osc.type = "sawtooth";
        osc.frequency.setValueAtTime(380, now); osc.frequency.exponentialRampToValueAtTime(120, now + 0.3);
        g.gain.setValueAtTime(0.2, now); g.gain.exponentialRampToValueAtTime(0.001, now + 0.3);
        osc.start(now); osc.stop(now + 0.3);
      }
    } catch { /* audio non disponible */ }
  }, []);

  /* ── Feedback ────────────────────────────────────────────────── */
  const [lastScan,   setLastScan]   = useState<{ code: string; status: "added"|"duplicate"|"noline"|"unknown" } | null>(null);
  const [flashColor, setFlashColor] = useState<string | null>(null);

  /* ── Manual input ────────────────────────────────────────────── */
  const [manualCode, setManualCode] = useState("");

  const selectedLine = lines.find((l) => l.id === selectedId) ?? null;

  /* returns true = found in Excel, false = not found */
  const lookupResult = useCallback((code: string): { found: boolean; weight: number | null } => {
    const rows    = LS.get<CellValue[][]>("ptg_rows", []);
    const codeCol = LS.get<number>("hgr_codeCol", 0);
    const wtCol   = LS.get<number>("hgr_wtCol",   1);
    for (const row of rows) {
      if (String(row[codeCol] ?? "").trim() === code.trim()) {
        const w = parseFloat(String(row[wtCol] ?? "").replace(",", ".").replace(/\s/g, ""));
        return { found: true, weight: isNaN(w) ? null : w };
      }
    }
    return { found: false, weight: null };
  }, []);

  /* ── Add item to selected line ───────────────────────────────── */
  const addItem = useCallback((code: string, weight: number | null, fromExcel: boolean) => {
    const selId = LS.get<string | null>("hgr_selectedId", null);
    if (!selId) { setLastScan({ code, status: "noline" }); setFlashColor(amber); setTimeout(() => setFlashColor(null), 700); return; }
    setLines((prev) => {
      const idx = prev.findIndex((l) => l.id === selId);
      if (idx === -1) return prev;
      const line = prev[idx];
      if (line.items.some((it) => it.code === code)) {
        setLastScan({ code, status: "duplicate" }); setFlashColor(amber); setTimeout(() => setFlashColor(null), 700);
        return prev;
      }
      setLastScan({ code, status: "added" }); setFlashColor(accent); setTimeout(() => setFlashColor(null), 700);
      const n = [...prev]; n[idx] = { ...line, items: [...line.items, { code, weight, fromExcel }] }; return n;
    });
  }, [accent]);

  /* ── Handle scanned code ─────────────────────────────────────── */
  const handleCode = useCallback((raw: string) => {
    const code = raw.trim();
    if (!code) return;
    // Son "code lu" dès la détection
    playTone("scan");
    const selId = LS.get<string | null>("hgr_selectedId", null);
    if (!selId) { setLastScan({ code, status: "noline" }); setFlashColor(amber); setTimeout(() => setFlashColor(null), 700); setTimeout(() => playTone("error"), 120); return; }
    // check duplicate
    const freshLines = LS.get<HangarLine[]>("hgr_lines2", []);
    const line = freshLines.find((l) => l.id === selId);
    if (line?.items.some((it) => it.code === code)) {
      setLastScan({ code, status: "duplicate" }); setFlashColor(amber); setTimeout(() => setFlashColor(null), 700); setTimeout(() => playTone("error"), 120); return;
    }
    const { found, weight } = lookupResult(code);
    if (found) {
      setTimeout(() => playTone("found"), 120);
      addItem(code, weight, true);
    } else {
      setLastScan({ code, status: "unknown" });
      setFlashColor(amber); setTimeout(() => setFlashColor(null), 700);
      setTimeout(() => playTone("error"), 120);
      setPending({ code, weight: "" });
    }
  }, [lookupResult, addItem, playTone]);

  useEffect(() => { assignRef.current = handleCode; }, [handleCode]);

  /* ── Confirm pending manual weight ──────────────────────────── */
  const confirmPending = () => {
    if (!pending) return;
    const w = parseFloat(pending.weight.replace(",", "."));
    addItem(pending.code, isNaN(w) ? null : w, false);
    setPending(null);
  };

  /* ── Camera start ────────────────────────────────────────────── */
  const startCamera = useCallback(async () => {
    setCameraError(null);
    const native = typeof BarcodeDetector !== "undefined";
    setUseNative(native);
    try {
      if (native) {
        const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: "environment" } } });
        streamRef.current = stream;
        if (videoRef.current) { videoRef.current.srcObject = stream; await videoRef.current.play(); }
        detectorRef.current = new BarcodeDetector({ formats: await BarcodeDetector.getSupportedFormats() });
        setCameraOn(true);
      } else {
        if (!videoRef.current) return;
        const reader = new BrowserMultiFormatReader();
        zxingReader.current = reader;
        const lastTimes: Record<string, number> = {};
        zxingControls.current = await reader.decodeFromConstraints(
          { video: { facingMode: { ideal: "environment" } } }, videoRef.current,
          (r) => { if (!r) return; if (pendingRef.current) return; const raw = r.getText(); const now = Date.now();
            if ((now - (lastTimes[raw] ?? 0)) > scanDelayRef.current) { lastTimes[raw] = now; assignRef.current(raw); } }
        );
        setCameraOn(true);
      }
    } catch (e) { setCameraError(e instanceof Error ? e.message : "Erreur caméra"); }
  }, []);

  const stopCamera = useCallback(() => {
    if (rafRef.current) cancelAnimationFrame(rafRef.current); rafRef.current = null;
    streamRef.current?.getTracks().forEach((t) => t.stop()); streamRef.current = null;
    if (videoRef.current) videoRef.current.srcObject = null;
    zxingControls.current?.stop(); zxingControls.current = null; zxingReader.current = null;
    setCameraOn(false);
  }, []);

  useEffect(() => {
    if (!cameraOn || !useNative || !detectorRef.current) return;
    let alive = true; const lt: Record<string, number> = {};
    const tick = async () => {
      if (!alive) return;
      const vid = videoRef.current;
      if (vid && vid.readyState >= 2 && detectorRef.current) {
        try { for (const r of await detectorRef.current.detect(vid)) {
          if (pendingRef.current) continue;
          const now = Date.now(); if ((now-(lt[r.rawValue]??0))>scanDelayRef.current){lt[r.rawValue]=now;assignRef.current(r.rawValue);}
        }} catch {/**/}
      }
      rafRef.current = requestAnimationFrame(tick);
    };
    rafRef.current = requestAnimationFrame(tick);
    return () => { alive = false; if (rafRef.current) cancelAnimationFrame(rafRef.current); };
  }, [cameraOn, useNative]);

  useEffect(() => () => { stopCamera(); }, [stopCamera]);

  /* ── Line management ─────────────────────────────────────────── */
  const createLine = () => {
    const name = newLineName.trim() || `Ligne ${lines.length + 1}`;
    const l: HangarLine = { id: newId(), name, items: [] };
    setLines((p) => [...p, l]); setSelectedId(l.id); setNewLineName("");
  };
  const deleteLine = (id: string) => {
    setLines((p) => p.filter((l) => l.id !== id));
    if (selectedId === id) setSelectedId(null);
  };
  const clearLine = (id: string) => setLines((p) => p.map((l) => l.id === id ? { ...l, items: [] } : l));
  const removeItem = (lineId: string, code: string) =>
    setLines((p) => p.map((l) => l.id === lineId ? { ...l, items: l.items.filter((it) => it.code !== code) } : l));
  const commitRename = () => {
    if (!editingId) return;
    const name = editingName.trim();
    if (name) setLines((p) => p.map((l) => l.id === editingId ? { ...l, name } : l));
    setEditingId(null); setEditingName("");
  };

  const submitManual = () => {
    const v = manualCode.trim(); if (!v) return;
    handleCode(v); setManualCode("");
  };

  /* ── Summary computations ────────────────────────────────────── */
  const lineSummary = lines.map((l) => ({
    id: l.id, name: l.name,
    qty: l.items.length,
    weight: l.items.reduce((s, it) => s + (it.weight ?? 0), 0),
    hasNull: l.items.some((it) => it.weight === null),
  }));
  const totalQty    = lineSummary.reduce((s, l) => s + l.qty, 0);
  const totalWeight = lineSummary.reduce((s, l) => s + l.weight, 0);
  const anyNull     = lineSummary.some((l) => l.hasNull);

  /* ── Export ──────────────────────────────────────────────────── */
  const exportXLSX = () => {
    const wb = XLSX.utils.book_new();
    // Summary sheet
    const summaryData: (string | number)[][] = [
      ["Ligne", "Quantité", "Poids total (kg)"],
      ...lineSummary.map((l) => [l.name, l.qty, l.weight]),
      ["TOTAL", totalQty, totalWeight],
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summaryData), "Résumé");
    // Detail sheet per line
    for (const line of lines) {
      const data: (string | number | null)[][] = [
        ["#", "Code", "Poids (kg)", "Source"],
        ...line.items.map((it, i) => [i + 1, it.code, it.weight, it.fromExcel ? "Excel" : "Manuel"]),
      ];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(data), line.name.slice(0, 31));
    }
    XLSX.writeFile(wb, `hangar_${new Date().toISOString().slice(0,10)}.xlsx`);
  };

  /* ── Styles ──────────────────────────────────────────────────── */
  const pad     = isMobile ? "12px 10px" : isTablet ? "16px 18px" : "20px 28px";
  const btnBase: React.CSSProperties = { fontFamily: MONO, fontSize: isMobile ? 12 : 11, letterSpacing: "0.08em",
    padding: isMobile ? "7px 14px" : "4px 10px", borderRadius: 4, cursor: "pointer", border: "none" };
  const thS: React.CSSProperties = { padding: isMobile ? "6px 10px" : "5px 12px", textAlign: "left", fontSize: isMobile ? 11 : 10,
    color: muted, textTransform: "uppercase" as const, letterSpacing: "0.1em", fontWeight: 400,
    borderBottom: `1px solid ${border}`, whiteSpace: "nowrap" };
  const tdS = (right?: boolean): React.CSSProperties => ({
    padding: isMobile ? "7px 10px" : "5px 12px", fontSize: isMobile ? 12 : 11, color: text, textAlign: right ? "right" : "left",
    borderBottom: `1px solid ${border}`, whiteSpace: "nowrap" });

  return (
    <div style={{ padding: pad, fontFamily: MONO, color: text, background: bg, minHeight: "calc(100vh - 44px)", boxSizing: "border-box", maxWidth: "100%", overflowX: "hidden" }}>

      {/* ── Header ──────────────────────────────────────────────── */}
      <div style={{ display: "flex", alignItems: "center", gap: 16, marginBottom: 20, flexWrap: "wrap" }}>
        <h2 style={{ fontSize: 13, letterSpacing: "0.18em", textTransform: "uppercase", color: accent, margin: 0 }}>
          🏭 Hangar
        </h2>
        <button onClick={reloadXl}
          style={{ ...btnBase, background: "transparent", border: `1px solid ${border}`, color: muted }}>
          ↺ Recharger Excel
        </button>
        {xlHeaders.length > 0 && (
          <>
            <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <span style={{ fontSize: 10, color: muted }}>Code :</span>
              <select value={codeColIdx} onChange={(e) => setCodeColIdx(Number(e.target.value))}
                style={{ fontFamily: MONO, fontSize: 10, padding: "3px 6px", borderRadius: 3,
                  background: surface, border: `1px solid ${border}`, color: text }}>
                {xlHeaders.map((h, i) => <option key={i} value={i}>{h || `Col ${i+1}`}</option>)}
              </select>
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <span style={{ fontSize: 10, color: muted }}>Poids :</span>
              <select value={wtColIdx} onChange={(e) => setWtColIdx(Number(e.target.value))}
                style={{ fontFamily: MONO, fontSize: 10, padding: "3px 6px", borderRadius: 3,
                  background: surface, border: `1px solid ${border}`, color: text }}>
                {xlHeaders.map((h, i) => <option key={i} value={i}>{h || `Col ${i+1}`}</option>)}
              </select>
            </div>
          </>
        )}
        {xlHeaders.length === 0 && (
          <span style={{ fontSize: 10, color: amber }}>⚠ Aucun fichier chargé dans Pointage</span>
        )}
        <button onClick={exportXLSX}
          style={{ ...btnBase, marginLeft: "auto", background: accent+"22", border: `1px solid ${accent}`, color: accent, fontWeight: 700 }}>
          ↓ Exporter XLSX
        </button>
      </div>

      <div style={{ display: "flex", gap: 16, alignItems: "flex-start",
        flexDirection: isDesktop ? "row" : "column", flexWrap: isDesktop ? "nowrap" : "wrap" }}>

        {/* ── Lines panel ─────────────────────────────────────── */}
        <div style={{ width: isDesktop ? 260 : "100%", flexShrink: 0, display: "flex", flexDirection: "column", gap: 10 }}>
          <div style={{ fontSize: 10, color: muted, textTransform: "uppercase", letterSpacing: "0.12em" }}>
            Lignes ({lines.length})
          </div>
          <div style={{ display: "flex", gap: 6 }}>
            <input value={newLineName} onChange={(e) => setNewLineName(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && createLine()}
              placeholder="Nom de la ligne…"
              style={{ flex: 1, fontFamily: MONO, fontSize: 11, padding: "5px 8px", borderRadius: 4,
                background: surface, border: `1px solid ${border}`, color: text, outline: "none" }} />
            <button onClick={createLine}
              style={{ ...btnBase, background: accent+"22", border: `1px solid ${accent}`, color: accent, fontWeight: 700 }}>+</button>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 4, maxHeight: isDesktop ? 520 : 200, overflowY: "auto" }}>
            {lines.length === 0 && <div style={{ fontSize: 11, color: muted }}>Aucune ligne créée</div>}
            {lines.map((l) => {
              const isActive = l.id === selectedId;
              const sum = lineSummary.find((s) => s.id === l.id);
              return (
                <div key={l.id} onClick={() => setSelectedId(l.id)}
                  style={{ display: "flex", alignItems: "center", gap: 5, padding: "6px 10px", borderRadius: 5,
                    background: isActive ? accent+"18" : (dark?"#1a1a1a":"#f0f0f0"),
                    border: `1px solid ${isActive ? accent : border}`, cursor: "pointer" }}>
                  {editingId === l.id ? (
                    <input autoFocus value={editingName}
                      onChange={(e) => setEditingName(e.target.value)}
                      onKeyDown={(e) => { if (e.key==="Enter") commitRename(); if (e.key==="Escape") setEditingId(null); }}
                      onBlur={commitRename} onClick={(e) => e.stopPropagation()}
                      style={{ flex:1, fontFamily:MONO, fontSize:11, padding:"2px 4px",
                        background:"transparent", border:`1px solid ${accent}`, color:text, borderRadius:3, outline:"none" }} />
                  ) : (
                    <span style={{ flex:1, fontSize:11, fontWeight:isActive?700:400,
                      color:isActive?accent:text, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>
                      {l.name}
                    </span>
                  )}
                  <span style={{ fontSize:9, color:muted, whiteSpace:"nowrap" }}>{sum?.qty} · {fmtW(sum?.weight??0)}</span>
                  <button onClick={(e)=>{e.stopPropagation();setEditingId(l.id);setEditingName(l.name);}}
                    style={{...btnBase,padding:"1px 5px",background:"transparent",border:"none",color:muted,fontSize:12}}>✎</button>
                  <button onClick={(e)=>{e.stopPropagation();deleteLine(l.id);}}
                    style={{...btnBase,padding:"1px 5px",background:"transparent",border:"none",color:danger,fontSize:12}}>✕</button>
                </div>
              );
            })}
          </div>
        </div>

        {/* ── Camera column ───────────────────────────────────── */}
        <div style={{ display:"flex", flexDirection:"column", gap:10, alignItems:"stretch",
          width: isDesktop ? 420 : "100%", flexShrink: 0 }}>
          <div style={{ padding:"5px 14px", borderRadius:4, fontSize:11, fontWeight:700,
            background: selectedLine ? accent+"18" : (dark?"#1a1a1a":"#f0f0f0"),
            border:`1px solid ${selectedLine ? accent : border}`,
            color: selectedLine ? accent : muted }}>
            {selectedLine ? `▶ ${selectedLine.name}` : "— Sélectionner une ligne —"}
          </div>

          {/* Video */}
          <div style={{ position:"relative", borderRadius:8, overflow:"hidden",
            border:`2px solid ${flashColor ?? border}`,
            boxShadow: flashColor ? `0 0 18px ${flashColor}55` : "none",
            transition:"border-color 0.15s, box-shadow 0.15s",
            width:"100%", aspectRatio:"4/3", background:dark?"#0a0a0a":"#ddd",
            display:"flex", alignItems:"center", justifyContent:"center" }}>
            <video ref={videoRef} playsInline muted
              style={{ width:"100%", height:"100%", display:cameraOn?"block":"none", objectFit:"cover" }} />
            {!cameraOn && <div style={{ textAlign:"center", color:muted, fontSize:12 }}><div style={{fontSize:32}}>📷</div><div>Caméra arrêtée</div></div>}
            {cameraOn && (
              <div style={{ position:"absolute", inset:0, pointerEvents:"none" }}>
                <div style={{ position:"absolute", top:"50%", left:"50%", transform:"translate(-50%,-50%)", width:170, height:170 }}>
                  {(["tl","tr","bl","br"] as const).map((c,i) => (
                    <div key={i} style={{ position:"absolute", width:18, height:18,
                      ...(c==="tl"?{top:-2,left:-2,borderTop:`3px solid ${accent}`,borderLeft:`3px solid ${accent}`,borderTopLeftRadius:5}:{}),
                      ...(c==="tr"?{top:-2,right:-2,borderTop:`3px solid ${accent}`,borderRight:`3px solid ${accent}`,borderTopRightRadius:5}:{}),
                      ...(c==="bl"?{bottom:-2,left:-2,borderBottom:`3px solid ${accent}`,borderLeft:`3px solid ${accent}`,borderBottomLeftRadius:5}:{}),
                      ...(c==="br"?{bottom:-2,right:-2,borderBottom:`3px solid ${accent}`,borderRight:`3px solid ${accent}`,borderBottomRightRadius:5}:{}),
                    }} />
                  ))}
                </div>
              </div>
            )}
          </div>

          <div style={{ display:"flex", gap:8, alignItems:"center" }}>
            <button onClick={cameraOn ? stopCamera : startCamera}
              style={{ ...btnBase, fontSize:12, padding:"8px 16px",
                background: cameraOn?(dark?"#2d0a0a":"#fee2e2"):(dark?"#0a200f":"#dcfce7"),
                border:`1px solid ${cameraOn?"#ef4444":accent}`,
                color: cameraOn?"#ef4444":accent, fontWeight:700 }}>
              {cameraOn ? "⏹ Arrêter" : "▶ Caméra"}
            </button>
            {cameraOn && (
              <div style={{ fontSize:10, padding:"3px 8px", borderRadius:4,
                background: useNative?accent+"18":(dark?"#0a0f20":"#e0e7ff"),
                border:`1px solid ${useNative?accent:(dark?"#818cf8":"#4f46e5")}`,
                color: useNative?accent:(dark?"#818cf8":"#4f46e5") }}>
                {useNative?"⚡ Natif":"⚙ ZXing"}
              </div>
            )}
          </div>

          {cameraError && (
            <div style={{ fontSize:11, color:"#ef4444", padding:"7px 12px",
              background:dark?"#2d0a0a":"#fee2e2", borderRadius:4, border:"1px solid #ef4444" }}>⚠ {cameraError}</div>
          )}

          {/* Scan delay slider */}
          <div style={{ display:"flex", alignItems:"center", gap:10 }}>
            <span style={{ fontSize:10, color:muted, whiteSpace:"nowrap" }}>Délai :</span>
            <input type="range" min={300} max={5000} step={100} value={scanDelay}
              onChange={(e) => setScanDelay(Number(e.target.value))}
              style={{ flex:1, accentColor:accent }} />
            <span style={{ fontSize:10, color:text, whiteSpace:"nowrap", minWidth:42, textAlign:"right" }}>{(scanDelay/1000).toFixed(1)} s</span>
          </div>

          {/* Manual entry */}
          <div style={{ display:"flex", gap:6, width:"100%" }}>
            <input value={manualCode} onChange={(e)=>setManualCode(e.target.value)}
              onKeyDown={(e)=>e.key==="Enter"&&submitManual()}
              placeholder="Saisie manuelle d'un code…"
              style={{ flex:1, fontFamily:MONO, fontSize:11, padding:"6px 10px", borderRadius:4,
                background:surface, border:`1px solid ${border}`, color:text, outline:"none" }} />
            <button onClick={submitManual}
              style={{ ...btnBase, background:accent+"22", border:`1px solid ${accent}`, color:accent, fontWeight:700 }}>↵</button>
          </div>

          {/* Scan feedback */}
          {lastScan && (
            <div style={{ padding:"9px 12px", borderRadius:6, width:"100%", boxSizing:"border-box" as const,
              background: lastScan.status==="added"?(dark?"#0a1f0f":"#f0fdf4"):(dark?"#1c1000":"#fffbeb"),
              border:`1px solid ${lastScan.status==="added"?accent:amber}` }}>
              <div style={{fontSize:10,color:muted,textTransform:"uppercase",letterSpacing:"0.1em",marginBottom:3}}>Dernier code</div>
              <div style={{fontSize:12,fontWeight:700,color:text,wordBreak:"break-all",marginBottom:3}}>{lastScan.code}</div>
              {lastScan.status==="added"    && <div style={{fontSize:11,color:accent}}>✓ Ajouté</div>}
              {lastScan.status==="duplicate"&& <div style={{fontSize:11,color:amber}}>⚠ Code déjà présent dans cette ligne</div>}
              {lastScan.status==="noline"   && <div style={{fontSize:11,color:amber}}>⚠ Aucune ligne sélectionnée</div>}
              {lastScan.status==="unknown"  && <div style={{fontSize:11,color:amber}}>? Code absent du fichier Excel → saisie manuelle</div>}
            </div>
          )}
        </div>

        {/* ── Detail of selected line ──────────────────────────── */}
        {selectedLine && (
          <div style={{ flex:1, minWidth:0 }}>
            <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:10 }}>
              <div style={{ fontSize:12, fontWeight:700, color:accent }}>
                {selectedLine.name}
                <span style={{ marginLeft:8, fontSize:10, color:muted, fontWeight:400 }}>
                  {selectedLine.items.length} art. · {fmtW(selectedLine.items.reduce((s,it)=>s+(it.weight??0),0))} kg
                </span>
              </div>
              <button onClick={()=>clearLine(selectedLine.id)}
                style={{ ...btnBase, background:"transparent", border:`1px solid ${danger}`, color:danger }}>Vider</button>
            </div>
            {selectedLine.items.length === 0 ? (
              <div style={{ fontSize:11, color:muted }}>Aucun article — scannez ou saisissez un code</div>
            ) : (
              <div style={{ maxHeight:380, overflowY:"auto" }}>
                <table style={{ borderCollapse:"collapse", width:"100%" }}>
                  <thead>
                    <tr>
                      <th style={thS}>#</th>
                      <th style={thS}>Code</th>
                      <th style={{...thS, textAlign:"right"}}>Poids (kg)</th>
                      <th style={thS}>Source</th>
                      <th style={thS}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {selectedLine.items.map((it, i) => (
                      <tr key={it.code}>
                        <td style={tdS()}><span style={{color:muted}}>{i+1}</span></td>
                        <td style={tdS()}>{it.code}</td>
                        <td style={{...tdS(true), color: it.weight===null?amber:text}}>{fmtW(it.weight)}</td>
                        <td style={tdS()}>
                          <span style={{ fontSize:9, padding:"1px 6px", borderRadius:8,
                            background: it.fromExcel?(dark?"#0a200f":"#dcfce7"):(dark?"#1a100a":"#fef3c7"),
                            border:`1px solid ${it.fromExcel?accent:amber}`,
                            color: it.fromExcel?accent:amber }}>
                            {it.fromExcel?"Excel":"Manuel"}
                          </span>
                        </td>
                        <td style={tdS()}>
                          <button onClick={()=>removeItem(selectedLine.id, it.code)}
                            style={{...btnBase, padding:"1px 6px", background:"transparent", border:"none", color:muted, fontSize:12}}>×</button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}
      </div>

      {/* ── Summary table ───────────────────────────────────────── */}
      {lines.length > 0 && (
        <div style={{ marginTop:24, maxWidth:"100%" }}>
          <div style={{ fontSize:10, color:muted, textTransform:"uppercase", letterSpacing:"0.12em", marginBottom:10 }}>
            Récapitulatif
          </div>
          <table style={{ borderCollapse:"collapse", width:"100%", background:surface, borderRadius:6, overflow:"hidden" }}>
            <thead>
              <tr style={{ background:dark?"#1a1a1a":"#f0f0f0" }}>
                <th style={thS}>Ligne</th>
                <th style={{...thS, textAlign:"right"}}>Qté</th>
                <th style={{...thS, textAlign:"right"}}>Poids total (kg)</th>
              </tr>
            </thead>
            <tbody>
              {lineSummary.map((l) => (
                <tr key={l.id} style={{ background: l.id===selectedId?(accent+"0d"):"transparent" }}>
                  <td style={tdS()}>{l.name}</td>
                  <td style={tdS(true)}>{l.qty}</td>
                  <td style={{...tdS(true), color: l.hasNull?amber:text}}>
                    {fmtW(l.weight)}{l.hasNull?" *":""}
                  </td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr style={{ background:dark?"#0e2016":"#e8f5ee" }}>
                <td style={{...tdS(), fontWeight:700, color:accent, fontSize:10, textTransform:"uppercase", letterSpacing:"0.1em"}}>Total</td>
                <td style={{...tdS(true), fontWeight:700, color:accent}}>{totalQty}</td>
                <td style={{...tdS(true), fontWeight:700, color:anyNull?amber:accent}}>
                  {fmtW(totalWeight)}{anyNull?" *":""}
                </td>
              </tr>
            </tfoot>
          </table>
          {anyNull && <div style={{ fontSize:10, color:amber, marginTop:6 }}>* Poids inconnu pour certains articles (non saisi)</div>}
        </div>
      )}

      {/* ── Pending weight modal ─────────────────────────────────── */}
      {pending && (
        <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.6)", display:"flex",
          alignItems:"center", justifyContent:"center", zIndex:1000 }}
          onClick={(e) => { if (e.target === e.currentTarget) setPending(null); }}>
          <div style={{ background:surface, border:`1px solid ${border}`, borderRadius:10,
            padding:"24px 20px", width:"min(360px, calc(100vw - 32px))", fontFamily:MONO, boxShadow:"0 8px 40px rgba(0,0,0,0.4)", boxSizing:"border-box" as const }}>
            <div style={{ fontSize:10, color:muted, textTransform:"uppercase", letterSpacing:"0.12em", marginBottom:12 }}>
              Code non trouvé dans Excel
            </div>
            <div style={{ fontSize:13, fontWeight:700, color:text, wordBreak:"break-all", marginBottom:16 }}>
              {pending.code}
            </div>
            <div style={{ fontSize:11, color:muted, marginBottom:6 }}>Poids (kg) :</div>
            <input autoFocus value={pending.weight}
              onChange={(e) => setPending({ ...pending, weight: e.target.value })}
              onKeyDown={(e) => { if (e.key==="Enter") confirmPending(); if (e.key==="Escape") setPending(null); }}
              placeholder="ex: 12.5"
              style={{ width:"100%", fontFamily:MONO, fontSize:13, padding:"8px 10px",
                borderRadius:5, background:bg, border:`1px solid ${accent}`, color:text, outline:"none",
                boxSizing:"border-box", marginBottom:16 }} />
            <div style={{ display:"flex", gap:10 }}>
              <button onClick={confirmPending}
                style={{ ...btnBase, flex:1, padding:"9px", background:accent+"22",
                  border:`1px solid ${accent}`, color:accent, fontWeight:700, fontSize:12 }}>
                ✓ Ajouter
              </button>
              <button onClick={() => setPending(null)}
                style={{ ...btnBase, padding:"9px 14px", background:"transparent",
                  border:`1px solid ${border}`, color:muted }}>
                Annuler
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
