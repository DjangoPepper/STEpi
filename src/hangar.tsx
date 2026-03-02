import { useEffect, useRef, useState, useCallback } from "react";
import { BrowserMultiFormatReader } from "@zxing/browser";
import * as XLSX from "xlsx";
import type { CellValue } from "./types";
import { useWindowWidth } from "./useWindowWidth";

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
interface HangarItem  { code: string; weight: number | null; fromExcel: boolean; position?: string; }
interface HangarLine  { id: string; name: string; items: HangarItem[]; }
interface HangarProps { dark: boolean; }
interface ForceWarning {
  code: string;
  weight: number;
  position: string;
  offenders: { pos: string; weight: number; pct: number }[];
}
interface PosConflict {
  code: string;
  weight: number | null;
  fromExcel: boolean;
  position: string;
  originalPosition: string;
  conflictLine: string;
  conflictCode: string;
}

const MONO  = "'IBM Plex Mono', 'Fira Mono', monospace";
const newId = () => `hl_${Date.now()}_${Math.random().toString(36).slice(2,6)}`;
const fmtW  = (w: number | null) =>
  w === null ? "—" : w.toLocaleString("fr-FR", { minimumFractionDigits: 2, maximumFractionDigits: 3 });

/* line background palette */
const LINE_BG_DARK  = ["#1c1305","#051c09","#06091c","#1a0512","#051a1a","#14190a","#190a05"];
const LINE_BG_LIGHT = ["#fff7ed","#f0fdf4","#eff6ff","#fdf4ff","#f0fdfa","#fefce8","#fff1f2"];

export default function Hangar({ dark }: HangarProps) {
  const bg      = dark ? "#0d0d0d" : "#f5f5f5";
  const surface = dark ? "#141414" : "#ffffff";
  const text    = dark ? "#e8e8e0" : "#1a1a1a";
  const muted   = dark ? "#555"    : "#888";
  const border  = dark ? "#2a2a2a" : "#d0d0d0";
  const danger  = dark ? "#f87171" : "#dc2626";
  const amber   = "#d97706";

  /* ── Breakpoints ─────────────────────────────────────────────── */
  const vw        = useWindowWidth();
  const isMobile  = vw < 640;
  const isTablet  = vw >= 640 && vw < 1024;

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
  const [pending, setPending] = useState<{ code: string; weight: string; position: string } | null>(null);
  const pendingRef = useRef<boolean>(false);
  useEffect(() => { pendingRef.current = pending !== null; }, [pending]);

  /* ── Force-position warning ──────────────────────────────────── */
  const [forceWarning, setForceWarning] = useState<ForceWarning | null>(null);
  const [posConflict,  setPosConflict]  = useState<PosConflict  | null>(null);

  /* ── Scan delay ───────────────────────────────────────────────── */
  const [scanDelay, setScanDelay] = useState<number>(() => LS.get("hgr_scanDelay", 1800));
  const scanDelayRef = useRef<number>(scanDelay);
  useEffect(() => { scanDelayRef.current = scanDelay; LS.set("hgr_scanDelay", scanDelay); }, [scanDelay]);

  /* ── F12 / QAA mode ─────────────────────────────────────────── */
  const [f12Mode, setF12Mode] = useState<boolean>(() => LS.get("hgr_f12mode", false));
  const f12ModeRef = useRef<boolean>(false);
  useEffect(() => { f12ModeRef.current = f12Mode; LS.set("hgr_f12mode", f12Mode); }, [f12Mode]);
  /* mode colour: orange (QAA) | red-pink (F12) */
  const accent = f12Mode ? (dark ? "#fca5a5" : "#ef4444") : "#f97316";

  /* ── QAA position — per line ─────────────────────────────── */
  const [linePositions, setLinePositions] = useState<Record<string,string>>(
    () => LS.get<Record<string,string>>("hgr_linepos", {})
  );
  const qaaPositionRef = useRef<string>("");
  // derived: position of the currently selected line
  const qaaPosition = linePositions[selectedId ?? ""] ?? "";
  qaaPositionRef.current = qaaPosition; // keep ref synchronously in sync
  const setQaaPosition = (pos: string) => {
    const id = selectedId;
    if (!id) return;
    setLinePositions((prev) => {
      const n = { ...prev };
      if (pos) n[id] = pos; else delete n[id];
      return n;
    });
  };
  useEffect(() => { LS.set("hgr_linepos", linePositions); }, [linePositions]);

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
  const [lastScan,   setLastScan]   = useState<{ code: string; status: "added"|"duplicate"|"noline"|"unknown"|"posdup" } | null>(null);
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
    const pos = qaaPositionRef.current || undefined;
    if (pos) {
      const fresh = LS.get<HangarLine[]>("hgr_lines2", []);
      const currLine = fresh.find((l) => l.id === selId);
      const hit = currLine?.items.find((it) => it.position === pos);
      if (hit) {
        setLastScan({ code, status: "posdup" }); setFlashColor(amber); setTimeout(() => setFlashColor(null), 700);
        setPosConflict({ code, weight, fromExcel, position: pos, originalPosition: pos, conflictLine: currLine!.name, conflictCode: hit.code });
        return;
      }
    }
    setLines((prev) => {
      const idx = prev.findIndex((l) => l.id === selId);
      if (idx === -1) return prev;
      const line = prev[idx];
      if (line.items.some((it) => it.code === code)) {
        setLastScan({ code, status: "duplicate" }); setFlashColor(amber); setTimeout(() => setFlashColor(null), 700);
        return prev;
      }
      setLastScan({ code, status: "added" }); setFlashColor(accent); setTimeout(() => setFlashColor(null), 700);
      const n = [...prev]; n[idx] = { ...line, items: [...line.items, { code, weight, fromExcel, position: pos }] }; return n;
    });
  }, [accent]);

  /* ── Null scan (empty slot) ──────────────────────────────────── */
  const addNullItem = useCallback(() => {
    playTone("scan");
    const selId = LS.get<string | null>("hgr_selectedId", null);
    if (!selId) { setLastScan({ code: "∅", status: "noline" }); setFlashColor(amber); setTimeout(() => setFlashColor(null), 700); return; }
    const code = `∅${Date.now().toString(36)}`;
    const pos = qaaPositionRef.current || undefined;
    if (pos) {
      const fresh = LS.get<HangarLine[]>("hgr_lines2", []);
      const currLine = fresh.find((l) => l.id === selId);
      const hit = currLine?.items.find((it) => it.position === pos);
      if (hit) {
        setLastScan({ code: "∅ vide", status: "posdup" }); setFlashColor(amber); setTimeout(() => setFlashColor(null), 700);
        setTimeout(() => playTone("error"), 120);
        setPosConflict({ code, weight: null, fromExcel: false, position: pos, originalPosition: pos, conflictLine: currLine!.name, conflictCode: hit.code });
        return;
      }
    }
    setLines((prev) => {
      const idx = prev.findIndex((l) => l.id === selId);
      if (idx === -1) return prev;
      const line = prev[idx];
      const n = [...prev]; n[idx] = { ...line, items: [...line.items, { code, weight: null, fromExcel: false, position: pos }] }; return n;
    });
    setLastScan({ code: "∅ vide", status: "added" }); setFlashColor(accent); setTimeout(() => setFlashColor(null), 700);
    setTimeout(() => playTone("found"), 120);
  }, [accent, playTone]);

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
      setPending({ code, weight: "", position: qaaPositionRef.current });
    }
  }, [lookupResult, addItem, playTone]);

  useEffect(() => { assignRef.current = handleCode; }, [handleCode]);

  /* ── Confirm pending manual weight ──────────────────────────── */
  const confirmPending = () => {
    if (!pending) return;
    if (pending.position === "" || pending.weight.trim() === "") return;
    const w = parseFloat(pending.weight.replace(",", "."));
    /* even-position weight check: warn if >15% heavier than an adjacent odd neighbor */
    const posNum = parseInt(pending.position, 10);
    if (!isNaN(w) && posNum % 2 === 0) {
      const allItems = lines.flatMap((l) => l.items);
      const offenders = [String(posNum - 1), String(posNum + 1)]
        .map((npos) => {
          const nb = allItems.find((it) => it.position === npos && it.weight !== null);
          if (nb && nb.weight !== null && w > nb.weight * 1.15)
            return { pos: npos, weight: nb.weight, pct: Math.round((w / nb.weight - 1) * 100) };
          return null;
        })
        .filter((x): x is { pos: string; weight: number; pct: number } => x !== null);
      if (offenders.length > 0) {
        setForceWarning({ code: pending.code, weight: w, position: pending.position, offenders });
        return;
      }
    }
    /* sync position back so addItem reads the right value from the ref */
    qaaPositionRef.current = pending.position;
    setQaaPosition(pending.position);
    addItem(pending.code, isNaN(w) ? null : w, false);
    setPending(null);
  };

  const forceConfirm = () => {
    if (!forceWarning) return;
    qaaPositionRef.current = forceWarning.position;
    setQaaPosition(forceWarning.position);
    addItem(forceWarning.code, forceWarning.weight, false);
    setForceWarning(null);
    setPending(null);
  };

  const forcePosConfirm = () => {
    if (!posConflict) return;
    const selId = LS.get<string | null>("hgr_selectedId", null);
    if (!selId) { setPosConflict(null); return; }
    setLines((prev) => {
      const idx = prev.findIndex((l) => l.id === selId);
      if (idx === -1) return prev;
      const line = prev[idx];
      if (line.items.some((it) => it.code === posConflict.code)) {
        setLastScan({ code: posConflict.code, status: "duplicate" });
        setFlashColor(amber); setTimeout(() => setFlashColor(null), 700);
        return prev;
      }
      setLastScan({ code: posConflict.code, status: "added" });
      setFlashColor(accent); setTimeout(() => setFlashColor(null), 700);
      const n = [...prev];
      n[idx] = { ...line, items: [...line.items, { code: posConflict.code, weight: posConflict.weight, fromExcel: posConflict.fromExcel, position: posConflict.position }] };
      return n;
    });
    setPosConflict(null);
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

  /* ── Summary computations ────────────────────────────────────── */  const occupiedPositions = new Set(
    (selectedLine?.items ?? []).map((it) => it.position).filter(Boolean)
  ) as Set<string>;
  const videPositions = new Set(
    (selectedLine?.items ?? []).filter((it) => it.code.startsWith("∅")).map((it) => it.position).filter(Boolean)
  ) as Set<string>;
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
      const hasPos = line.items.some((it) => it.position);
      const data: (string | number | null)[][] = [
        hasPos ? ["#", "Réf.", "Poids (kg)", "Position", "Source"] : ["#", "Réf.", "Poids (kg)", "Source"],
        ...line.items.map((it, i) => hasPos
          ? [i + 1, it.code, it.weight, it.position ?? "", it.fromExcel ? "Excel" : "Manuel"]
          : [i + 1, it.code, it.weight, it.fromExcel ? "Excel" : "Manuel"]),
      ];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(data), line.name.slice(0, 31));
    }
    XLSX.writeFile(wb, `hangar_${new Date().toISOString().slice(0,10)}.xlsx`);
  };

  /* ── Styles ──────────────────────────────────────────────────── */
  const pad     = isMobile ? "12px 10px" : isTablet ? "16px 18px" : "20px 28px";
  const btnBase: React.CSSProperties = { fontFamily: MONO, fontSize: isMobile ? 15 : 11, letterSpacing: isMobile ? 0 : "0.08em",
    padding: isMobile ? "12px 16px" : "4px 10px", borderRadius: 4, cursor: "pointer", border: "none" };
  const thS: React.CSSProperties = { padding: isMobile ? "10px 12px" : "5px 12px", textAlign: "left", fontSize: isMobile ? 13 : 10,
    color: muted, textTransform: isMobile ? "none" : "uppercase" as const, letterSpacing: isMobile ? 0 : "0.1em", fontWeight: 400,
    borderBottom: `1px solid ${border}`, whiteSpace: "nowrap" };
  const tdS = (right?: boolean): React.CSSProperties => ({
    padding: isMobile ? "10px 12px" : "5px 12px", fontSize: isMobile ? 14 : 11, color: text, textAlign: right ? "right" : "left",
    borderBottom: `1px solid ${border}`, whiteSpace: "nowrap" });

  return (
    <div style={{ padding: pad, fontFamily: MONO, color: text, background: bg, minHeight: "calc(100vh - 44px)", boxSizing: "border-box", maxWidth: "100%", overflowX: "hidden" }}>

      {/* ── Header ──────────────────────────────────────────────── */}
      <div style={{ display: "flex", alignItems: "center", gap: 16, marginBottom: 20, flexWrap: "wrap" }}>
        <h2 style={{ fontSize: isMobile ? 17 : 13, letterSpacing: isMobile ? 0 : "0.18em", textTransform: isMobile ? "none" : "uppercase", color: accent, margin: 0 }}>
          🏭 Hangar
        </h2>

        {/* F12 / QAA toggle */}
        <button
          onClick={() => setF12Mode((v) => !v)}
          title={f12Mode ? "Mode F12 : chaque scan → {nom} up + {nom} down" : "Mode QAA : scan normal"}
          style={{
            display: "flex", alignItems: "center", gap: 0,
            padding: 0, border: `1px solid ${f12Mode ? accent : border}`,
            borderRadius: 20, overflow: "hidden", cursor: "pointer",
            background: "transparent", flexShrink: 0,
            fontSize: isMobile ? 12 : 10, fontFamily: MONO,
          }}
        >
          <span style={{
            padding: isMobile ? "7px 11px" : "4px 9px",
            background: !f12Mode ? (dark ? accent + "33" : accent + "22") : "transparent",
            color: !f12Mode ? accent : muted, fontWeight: !f12Mode ? 700 : 400,
            transition: "background 0.2s, color 0.2s",
          }}>QAA</span>
          <span style={{
            padding: isMobile ? "7px 11px" : "4px 9px",
            background: f12Mode ? (dark ? accent + "33" : accent + "22") : "transparent",
            color: f12Mode ? accent : muted, fontWeight: f12Mode ? 700 : 400,
            transition: "background 0.2s, color 0.2s",
          }}>F12</span>
        </button>

        <button onClick={reloadXl}
          style={{ ...btnBase, background: "transparent", border: `1px solid ${border}`, color: muted }}>
          ↺ recharger
        </button>
        {xlHeaders.length > 0 && (
          <>
            <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <span style={{ fontSize: isMobile ? 13 : 10, color: muted }}>Réf. :</span>
              <select value={codeColIdx} onChange={(e) => setCodeColIdx(Number(e.target.value))}
                style={{ fontFamily: MONO, fontSize: isMobile ? 14 : 10, padding: isMobile ? "10px 8px" : "3px 6px", borderRadius: 3,
                  background: surface, border: `1px solid ${border}`, color: text }}>
                {xlHeaders.map((h, i) => <option key={i} value={i}>{h || `Col ${i+1}`}</option>)}
              </select>
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <span style={{ fontSize: isMobile ? 13 : 10, color: muted }}>Poids :</span>
              <select value={wtColIdx} onChange={(e) => setWtColIdx(Number(e.target.value))}
                style={{ fontFamily: MONO, fontSize: isMobile ? 14 : 10, padding: isMobile ? "10px 8px" : "3px 6px", borderRadius: 3,
                  background: surface, border: `1px solid ${border}`, color: text }}>
                {xlHeaders.map((h, i) => <option key={i} value={i}>{h || `Col ${i+1}`}</option>)}
              </select>
            </div>
          </>
        )}
        {xlHeaders.length === 0 && (
          <span style={{ fontSize: isMobile ? 13 : 10, color: amber }}>⚠ Aucun fichier chargé dans Pointage</span>
        )}
        <button onClick={exportXLSX}
          style={{ ...btnBase, marginLeft: "auto", background: accent+"22", border: `1px solid ${accent}`, color: accent, fontWeight: 700 }}>
          ↓ Export
        </button>
      </div>

      <div style={{ display: "flex", gap: 16, alignItems: "flex-start", flexDirection: "column" }}>

        {/* ── Lines panel ─────────────────────────────────────── */}
        <div style={{ width: "100%", flexShrink: 0, display: "flex", flexDirection: "column", gap: 10 }}>
          <div style={{ fontSize: isMobile ? 13 : 10, color: muted, textTransform: isMobile ? "none" : "uppercase", letterSpacing: isMobile ? 0 : "0.12em" }}>
            Lignes ({lines.length})
          </div>
          <div style={{ display: "flex", gap: 6 }}>
            <input value={newLineName} onChange={(e) => setNewLineName(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && createLine()}
              placeholder="Nom de la ligne…"
              style={{ flex: 1, fontFamily: MONO, fontSize: isMobile ? 15 : 11, padding: isMobile ? "10px 10px" : "5px 8px", borderRadius: 4,
                background: surface, border: `1px solid ${border}`, color: text, outline: "none" }} />
            <button onClick={createLine}
              style={{ ...btnBase, background: accent+"22", border: `1px solid ${accent}`, color: accent, fontWeight: 700 }}>+</button>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 4, maxHeight: 160, overflowY: "auto" }}>
            {lines.length === 0 && <div style={{ fontSize: isMobile ? 14 : 11, color: muted }}>Aucune ligne créée</div>}
            {lines.map((l, lIdx) => {
              const isActive = l.id === selectedId;
              const sum = lineSummary.find((s) => s.id === l.id);
              const lineBg = dark ? LINE_BG_DARK[lIdx % LINE_BG_DARK.length] : LINE_BG_LIGHT[lIdx % LINE_BG_LIGHT.length];
              return (
                <div key={l.id} onClick={() => setSelectedId(l.id)}
                  style={{ display: "flex", alignItems: "center", gap: 5, padding: "6px 10px", borderRadius: 5,
                    background: isActive ? accent+"40" : lineBg,
                    border: `1px solid ${isActive ? accent : border}`, cursor: "pointer" }}>
                  {editingId === l.id ? (
                    <input autoFocus value={editingName}
                      onChange={(e) => setEditingName(e.target.value)}
                      onKeyDown={(e) => { if (e.key==="Enter") commitRename(); if (e.key==="Escape") setEditingId(null); }}
                      onBlur={commitRename} onClick={(e) => e.stopPropagation()}
                      style={{ flex:1, fontFamily:MONO, fontSize: isMobile ? 14 : 11, padding: isMobile ? "6px 6px" : "2px 4px",
                        background:"transparent", border:`1px solid ${accent}`, color:text, borderRadius:3, outline:"none" }} />
                  ) : (
                    <span style={{ flex:1, fontSize: isMobile ? 14 : 11, fontWeight:isActive?700:400,
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
        <div style={{ display:"flex", flexDirection:"column", gap:10, alignItems:"stretch", width:"100%", flexShrink: 0 }}>
          <div style={{ padding:"5px 14px", borderRadius:4, fontSize: isMobile ? 15 : 11, fontWeight:700,
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

          {/* Row 1 : all controls on one line */}
          <div style={{ display:"flex", gap:8, alignItems:"center", flexWrap:"wrap" }}>
            <button onClick={cameraOn ? stopCamera : startCamera}
              style={{ ...btnBase, fontSize: isMobile ? 15 : 12, padding: isMobile ? "12px 20px" : "8px 16px",
                flexShrink:0,
                background: cameraOn?(dark?"#2d0a0a":"#fee2e2"):(dark?"#0a200f":"#dcfce7"),
                border:`1px solid ${cameraOn?"#ef4444":accent}`,
                color: cameraOn?"#ef4444":accent, fontWeight:700 }}>
              {cameraOn ? "⏹ Arrêter" : "▶ Caméra"}
            </button>
            <button onClick={addNullItem}
              title="Ajouter un emplacement vide (sans référence)"
              style={{ ...btnBase, fontSize: isMobile ? 14 : 11, padding: isMobile ? "12px 16px" : "8px 13px",
                flexShrink:0,
                background: "transparent", border:`1px solid ${muted}`, color: muted }}>
              ∅ vide
            </button>
            {cameraOn && (
              <div style={{ fontSize: isMobile ? 11 : 10, padding:"3px 6px", borderRadius:4, flexShrink:0,
                background: useNative?accent+"18":(dark?"#0a0f20":"#e0e7ff"),
                border:`1px solid ${useNative?accent:(dark?"#818cf8":"#4f46e5")}`,
                color: useNative?accent:(dark?"#818cf8":"#4f46e5") }}>
                {useNative?"⚡":"⚙"}
              </div>
            )}
            <span style={{ fontSize: isMobile ? 12 : 10, color:muted, whiteSpace:"nowrap", flexShrink:0 }}>Délai</span>
            <input type="range" min={300} max={5000} step={100} value={scanDelay}
              onChange={(e) => setScanDelay(Number(e.target.value))}
              style={{ width: isMobile ? 90 : 70, flexShrink:0, accentColor:accent }} />
            <span style={{ fontSize: isMobile ? 12 : 10, color:text, whiteSpace:"nowrap", flexShrink:0 }}>{(scanDelay/1000).toFixed(1)}s</span>
            <input value={manualCode} onChange={(e)=>setManualCode(e.target.value)}
              onKeyDown={(e)=>e.key==="Enter"&&submitManual()}
              placeholder="Saisie manuelle…"
              style={{ flex:1, minWidth: isMobile ? 120 : 80, fontFamily:MONO, fontSize: isMobile ? 15 : 11,
                padding: isMobile ? "10px 12px" : "6px 10px", borderRadius:4,
                background:surface, border:`1px solid ${border}`, color:text, outline:"none" }} />
            <button onClick={submitManual}
              style={{ ...btnBase, padding: isMobile ? "12px 14px" : "6px 10px",
                background:accent+"22", border:`1px solid ${accent}`, color:accent, fontWeight:700, flexShrink:0 }}>↵</button>
          </div>

          {/* Row 3 : position slots 1–51 */}
          {(() => {
            return (
              <div>
                <div style={{ fontSize: isMobile ? 12 : 10, color:muted, marginBottom:5 }}>
                  Emplacement
                  {qaaPosition && <span style={{color:accent, fontWeight:700, marginLeft:6}}>→ {qaaPosition}</span>}
                </div>
{(() => {
                  const CHIP = isMobile ? 34 : 26;
                  const GAP  = 3;
                  const half = (CHIP + GAP) / 2;
                  const odds  = Array.from({length:26},(_,i)=>String(i*2+1));
                  const evens = Array.from({length:25},(_,i)=>String(i*2+2));
                  const chip = (s: string) => {
                    const isSel = qaaPosition === s;
                    const isOcc = occupiedPositions.has(s);
                    const isVide = videPositions.has(s);
                    return (
                      <button key={s} disabled={isOcc && !isSel}
                        onClick={() => setQaaPosition(isSel ? "" : s)}
                        title={isOcc && !isSel ? `Emplacement ${s} occupé${isVide ? " (∅ vide)" : ""}` : `Emplacement ${s}`}
                        style={{
                          fontFamily:MONO, fontSize: isMobile ? 12 : 9,
                          width:CHIP, height:CHIP, flexShrink:0,
                          borderRadius:4, padding:0, lineHeight:1,
                          cursor: isOcc && !isSel ? "not-allowed" : "pointer",
                          border: isSel ? `2px solid ${accent}` : `1px solid ${isOcc ? "transparent" : border}`,
                          background: isSel ? accent+"33"
                            : isVide ? (dark ? "#3b0764" : "#ede9fe")
                            : isOcc  ? (dark ? "#2a1a00" : "#fde68a")
                            : (dark ? "#1a1a1a" : "#f0f0f0"),
                          color: isSel ? accent
                            : isVide ? (dark ? "#ddd6fe" : "#7c3aed")
                            : isOcc  ? "#92400e"
                            : muted,
                          fontWeight: isSel || isVide ? 700 : 400,
                          opacity: isOcc && !isSel ? 0.75 : 1,
                        }}>{s}</button>
                    );
                  };
                  return (
                    <div style={{ overflowX:"auto", paddingBottom:4 }}>
                      <div style={{ display:"inline-block", minWidth:"max-content" }}>
                        <div style={{ display:"flex", gap:GAP, marginLeft:half, marginBottom:GAP }}>
                          {evens.map(chip)}
                        </div>
                        <div style={{ display:"flex", gap:GAP }}>
                          {odds.map(chip)}
                        </div>
                      </div>
                    </div>
                  );
                })()}
              </div>
            );
          })()}

          {cameraError && (
            <div style={{ fontSize: isMobile ? 14 : 11, color:"#ef4444", padding:"7px 12px",
              background:dark?"#2d0a0a":"#fee2e2", borderRadius:4, border:"1px solid #ef4444" }}>⚠ {cameraError}</div>
          )}

          {/* Scan feedback */}
          {lastScan && (
            <div style={{ padding:"9px 12px", borderRadius:6, width:"100%", boxSizing:"border-box" as const,
              background: lastScan.status==="added"?(dark?"#0a1f0f":"#f0fdf4"):(dark?"#1c1000":"#fffbeb"),
              border:`1px solid ${lastScan.status==="added"?accent:amber}` }}>
              <div style={{fontSize:10,color:muted,textTransform:"uppercase",letterSpacing:"0.1em",marginBottom:3}}>Dernière réf.</div>
              <div style={{fontSize:12,fontWeight:700,color:text,wordBreak:"break-all",marginBottom:3}}>{lastScan.code}</div>
              {lastScan.status==="added"    && <div style={{fontSize:11,color:accent}}>✓ Ajouté{qaaPosition ? ` → ${qaaPosition}` : ""}</div>}
              {lastScan.status==="duplicate"&& <div style={{fontSize:11,color:amber}}>⚠ Code déjà présent dans cette ligne</div>}
              {lastScan.status==="posdup"   && <div style={{fontSize:11,color:amber}}>⚠ Position "{qaaPosition}" déjà occupée</div>}
              {lastScan.status==="noline"   && <div style={{fontSize:11,color:amber}}>⚠ Aucune ligne sélectionnée</div>}
              {lastScan.status==="unknown"  && <div style={{fontSize:11,color:amber}}>? Réf. absente du fichier Excel → saisie manuelle</div>}
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
              <div style={{ fontSize:11, color:muted }}>Aucun article — scannez ou saisissez une réf.</div>
            ) : (
              <div style={{ maxHeight:380, overflowY:"auto" }}>
                <table style={{ borderCollapse:"collapse", width:"100%" }}>
                  {(() => {
                    const showPos = selectedLine.items.some((it) => it.position);
                    return (
                      <>
                      <thead>
                        <tr>
                          <th style={thS}>#</th>
                          <th style={thS}>Réf.</th>
                          <th style={{...thS, textAlign:"right"}}>Poids (kg)</th>
                          {showPos && <th style={thS}>Pos.</th>}
                          <th style={thS}>Source</th>
                          <th style={thS}></th>
                        </tr>
                      </thead>
                      <tbody>
                        {selectedLine.items.map((it, i) => {
                          const isVide = it.code.startsWith("∅");
                          const lavBg  = "#7c3aed";
                          const lavFg  = "#ffffff";
                          return (
                          <tr key={it.code} style={{ background: isVide ? lavBg : undefined }}>
                            <td style={tdS()}><span style={{color: isVide ? "#ddd6fe" : muted}}>{i+1}</span></td>
                            <td style={{...tdS(), color: isVide ? lavFg : text}}>{isVide ? "∅ vide" : it.code}</td>
                            <td style={{...tdS(true), color: isVide ? "#ddd6fe" : (it.weight===null?amber:text)}}>{fmtW(it.weight)}</td>
                            {showPos && <td style={{...tdS(), fontWeight:700, color: isVide ? "#ddd6fe" : (it.position?accent:muted)}}>{it.position ?? "—"}</td>}
                            <td style={tdS()}>
                              <span style={{ fontSize:9, padding:"1px 6px", borderRadius:8,
                                background: isVide ? "#5b21b6" : (it.fromExcel?(dark?"#0a200f":"#dcfce7"):(dark?"#1a100a":"#fef3c7")),
                                border:`1px solid ${isVide ? "#8b5cf6" : (it.fromExcel?accent:amber)}`,
                                color: isVide ? lavFg : (it.fromExcel?accent:amber) }}>
                                {isVide ? "∅" : (it.fromExcel?"Excel":"Manuel")}
                              </span>
                            </td>
                            <td style={tdS()}>
                              <button onClick={()=>removeItem(selectedLine.id, it.code)}
                                style={{...btnBase, padding:"1px 6px", background:"transparent", border:"none", color: isVide ? "#ddd6fe" : muted, fontSize:12}}>×</button>
                            </td>
                          </tr>
                          );
                        })}
                      </tbody>
                      </>
                    );
                  })()}
                </table>
              </div>
            )}
          </div>
        )}
      </div>

      {/* ── Summary table ───────────────────────────────────────── */}
      {lines.length > 0 && (
        <div style={{ marginTop:24, maxWidth:"100%" }}>
          <div style={{ fontSize: isMobile ? 13 : 10, color:muted, textTransform: isMobile ? "none" : "uppercase", letterSpacing: isMobile ? 0 : "0.12em", marginBottom:10 }}>
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
              {lineSummary.map((l, lIdx) => {
              const lineBg = dark ? LINE_BG_DARK[lIdx % LINE_BG_DARK.length] : LINE_BG_LIGHT[lIdx % LINE_BG_LIGHT.length];
              return (
                <tr key={l.id} style={{ background: l.id===selectedId?(accent+"28"):lineBg }}>
                  <td style={tdS()}>{l.name}</td>
                  <td style={tdS(true)}>{l.qty}</td>
                  <td style={{...tdS(true), color: l.hasNull?amber:text}}>
                    {fmtW(l.weight)}{l.hasNull?" *":""}
                  </td>
                </tr>
              );
            })}
            </tbody>
            <tfoot>
              <tr style={{ background:dark?"#0e2016":"#e8f5ee" }}>
                <td style={{...tdS(), fontWeight:700, color:accent, fontSize: isMobile ? 13 : 10, textTransform: isMobile ? "none" : "uppercase", letterSpacing: isMobile ? 0 : "0.1em"}}>Total</td>
                <td style={{...tdS(true), fontWeight:700, color:accent}}>{totalQty}</td>
                <td style={{...tdS(true), fontWeight:700, color:anyNull?amber:accent}}>
                  {fmtW(totalWeight)}{anyNull?" *":""}
                </td>
              </tr>
            </tfoot>
          </table>
          {anyNull && <div style={{ fontSize: isMobile ? 13 : 10, color:amber, marginTop:6 }}>* Poids inconnu pour certains articles (non saisi)</div>}
        </div>
      )}

      {/* ── Pending weight modal ─────────────────────────────────── */}
      {pending && (
        <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.6)", display:"flex",
          alignItems:"center", justifyContent:"center", zIndex:1000 }}
          onClick={(e) => { if (e.target === e.currentTarget) setPending(null); }}>
          <div style={{ background:surface, border:`1px solid ${border}`, borderRadius:10,
            padding:"24px 20px", width:"min(420px, calc(100vw - 32px))", fontFamily:MONO, boxShadow:"0 8px 40px rgba(0,0,0,0.4)", boxSizing:"border-box" as const }}>
            <div style={{ fontSize: isMobile ? 14 : 10, color:muted, textTransform: isMobile ? "none" : "uppercase", letterSpacing: isMobile ? 0 : "0.12em", marginBottom:12 }}>
              Réf. non trouvée dans Excel
            </div>
            <div style={{ fontSize: isMobile ? 15 : 13, fontWeight:700, color:text, wordBreak:"break-all", marginBottom:14 }}>
              {pending.code}
            </div>
            {/* Position slot grid */}
            <div style={{ marginBottom:14 }}>
              <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:6 }}>
                <span style={{ fontSize: isMobile ? 13 : 10, color:muted }}>Emplacement</span>
                {pending.position
                  ? <span style={{color:accent, fontWeight:700}}>→ {pending.position}</span>
                  : <span style={{fontSize: isMobile ? 11 : 9, color:"#ef4444"}}>obligatoire</span>}
              </div>
{(() => {
                const CHIP = isMobile ? 34 : 26;
                const GAP  = 3;
                const half = (CHIP + GAP) / 2;
                const odds  = Array.from({length:26},(_,i)=>String(i*2+1));
                const evens = Array.from({length:25},(_,i)=>String(i*2+2));
                const chip = (s: string) => {
                  const isSel = pending.position === s;
                  const isOcc = occupiedPositions.has(s);
                  const isVide = videPositions.has(s);
                  return (
                    <button key={s} disabled={isOcc && !isSel}
                      onClick={() => setPending({ ...pending, position: isSel ? "" : s })}
                      title={isOcc && !isSel ? `Emplacement ${s} occupé${isVide ? " (∅ vide)" : ""}` : `Emplacement ${s}`}
                      style={{
                        fontFamily:MONO, fontSize: isMobile ? 12 : 9,
                        width:CHIP, height:CHIP, flexShrink:0,
                        borderRadius:4, padding:0, lineHeight:1,
                        cursor: isOcc && !isSel ? "not-allowed" : "pointer",
                        border: isSel ? `2px solid ${accent}` : `1px solid ${isOcc ? "transparent" : border}`,
                        background: isSel ? accent+"33"
                          : isVide ? (dark ? "#3b0764" : "#ede9fe")
                          : isOcc  ? (dark ? "#2a1a00" : "#fde68a")
                          : (dark ? "#1c1c1c" : "#f0f0f0"),
                        color: isSel ? accent
                          : isVide ? (dark ? "#ddd6fe" : "#7c3aed")
                          : isOcc  ? "#92400e"
                          : muted,
                        fontWeight: isSel || isVide ? 700 : 400,
                        opacity: isOcc && !isSel ? 0.75 : 1,
                      }}>{s}</button>
                  );
                };
                return (
                  <div style={{ overflowX:"auto", paddingBottom:4 }}>
                    <div style={{ display:"inline-block", minWidth:"max-content" }}>
                      <div style={{ display:"flex", gap:GAP, marginLeft:half, marginBottom:GAP }}>
                        {evens.map(chip)}
                      </div>
                      <div style={{ display:"flex", gap:GAP }}>
                        {odds.map(chip)}
                      </div>
                    </div>
                  </div>
                );
              })()}
            </div>
            {/* Poids */}
            <div style={{ marginBottom:14 }}>
              <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:5 }}>
                <span style={{ fontSize: isMobile ? 13 : 10, color:muted }}>Poids (kg) :</span>
                {pending.weight.trim() === "" && <span style={{ fontSize: isMobile ? 11 : 9, color:"#ef4444" }}>obligatoire</span>}
              </div>
              <input autoFocus value={pending.weight} inputMode="decimal"
                onChange={(e) => setPending({ ...pending, weight: e.target.value })}
                onKeyDown={(e) => { if (e.key==="Enter") confirmPending(); if (e.key==="Escape") setPending(null); }}
                placeholder="12.5"
                style={{ width:"100%", fontFamily:MONO, fontSize: isMobile ? 16 : 13,
                  padding: isMobile ? "12px 10px" : "8px 10px", borderRadius:5,
                  background:bg, border:`1px solid ${pending.weight.trim()==="" ? "#ef4444" : border}`, color:text, outline:"none",
                  boxSizing:"border-box" as const }} />
            </div>
            {(() => {
              const missingPos = pending.position === "";
              const missingWt  = pending.weight.trim() === "";
              const canAdd = !missingPos && !missingWt;
              return (
                <div style={{ display:"flex", gap:10 }}>
                  <button onClick={confirmPending} disabled={!canAdd}
                    style={{ ...btnBase, flex:1, padding: isMobile ? "13px" : "9px",
                      background: canAdd ? accent+"22" : (dark?"#2a2a2a":"#e5e7eb"),
                      border:`1px solid ${canAdd ? accent : border}`,
                      color: canAdd ? accent : muted,
                      fontWeight:700, fontSize: isMobile ? 15 : 12,
                      cursor: canAdd ? "pointer" : "not-allowed",
                      opacity: canAdd ? 1 : 0.5 }}>
                    ✓ Ajouter
                  </button>
                  <button onClick={() => setPending(null)}
                    style={{ ...btnBase, padding:"9px 14px", background:"transparent",
                      border:`1px solid ${border}`, color:muted }}>
                    Annuler
                  </button>
                </div>
              );
            })()}
          </div>
        </div>
      )}

      {/* ── Position-conflict force modal ─────────────────────────── */}
      {posConflict && (
        <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.7)", display:"flex",
          alignItems:"center", justifyContent:"center", zIndex:1050 }}
          onClick={(e) => { if (e.target === e.currentTarget) setPosConflict(null); }}>
          <div style={{ background:surface, border:`2px solid ${amber}`, borderRadius:10,
            padding:"24px 20px", width:"min(400px, calc(100vw - 32px))", fontFamily:MONO,
            boxShadow:"0 8px 40px rgba(0,0,0,0.5)", boxSizing:"border-box" as const }}>
            <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:14 }}>
              <span style={{ fontSize:20 }}>📌</span>
              <span style={{ fontSize: isMobile ? 14 : 12, fontWeight:700, color:amber }}>Emplacement déjà occupé</span>
            </div>
            <div style={{ fontSize: isMobile ? 13 : 11, color:text, marginBottom:4 }}>
              <span style={{ color:muted }}>Nouvelle réf. </span>
              <span style={{ fontWeight:700, wordBreak:"break-all" }}>{posConflict.code}</span>
            </div>
            <div style={{ fontSize: isMobile ? 13 : 11, color:text, marginBottom:14 }}>
              <span style={{ color:muted }}>Pos. demandée </span>
              <span style={{ fontWeight:700, color:accent }}>{posConflict.position}</span>
            </div>
            <div style={{ marginBottom:16, padding:"10px 12px", borderRadius:6,
              background:dark?"#1c1000":"#fffbeb", border:`1px solid ${amber}` }}>
              <div style={{ fontSize: isMobile ? 11 : 9, color:amber, textTransform:"uppercase", letterSpacing:"0.1em", marginBottom:4 }}>Occupée par</div>
              <div style={{ fontSize: isMobile ? 13 : 11, color:text, fontWeight:700, wordBreak:"break-all", marginBottom:2 }}>{posConflict.conflictCode}</div>
              <div style={{ fontSize: isMobile ? 11 : 9, color:muted }}>ligne : {posConflict.conflictLine}</div>
            </div>
            {/* slot grid to pick a different position */}
            {(() => {
              const CHIP = isMobile ? 32 : 24;
              const GAP  = 3;
              const half = (CHIP + GAP) / 2;
              const odds  = Array.from({length:26},(_,i)=>String(i*2+1));
              const evens = Array.from({length:25},(_,i)=>String(i*2+2));
              const isOrig = (s: string) => s === posConflict.originalPosition;
              const chip = (s: string) => {
                const isSel = posConflict.position === s;
                const isOcc = occupiedPositions.has(s);
                const disabled = isOcc && !isSel && !isOrig(s);
                return (
                  <button key={s} disabled={disabled}
                    onClick={() => setPosConflict({ ...posConflict, position: isSel && !isOrig(s) ? posConflict.originalPosition : s })}
                    title={isOrig(s) ? `Emplacement ${s} (conflit)` : isOcc && !isSel ? `Emplacement ${s} occupé` : `Emplacement ${s}`}
                    style={{
                      fontFamily:MONO, fontSize: isMobile ? 11 : 9,
                      width:CHIP, height:CHIP, flexShrink:0,
                      borderRadius:4, padding:0, lineHeight:1,
                      cursor: disabled ? "not-allowed" : "pointer",
                      border: isSel
                        ? `2px solid ${isOrig(s) ? amber : accent}`
                        : `1px solid ${isOcc ? "transparent" : border}`,
                      background: isSel
                        ? (isOrig(s) ? amber+"33" : accent+"33")
                        : isOcc ? (dark?"#2a1a00":"#fde68a") : (dark?"#1c1c1c":"#f0f0f0"),
                      color: isSel
                        ? (isOrig(s) ? amber : accent)
                        : isOcc ? "#92400e" : muted,
                      fontWeight: isSel ? 700 : 400,
                      opacity: disabled ? 0.4 : 1,
                    }}>{s}</button>
                );
              };
              return (
                <div style={{ marginBottom:16 }}>
                  <div style={{ fontSize: isMobile ? 12 : 10, color:muted, marginBottom:6 }}>
                    Choisir un autre emplacement
                    {posConflict.position !== posConflict.originalPosition &&
                      <span style={{color:accent, fontWeight:700, marginLeft:6}}>→ {posConflict.position}</span>}
                  </div>
                  <div style={{ overflowX:"auto", paddingBottom:4 }}>
                    <div style={{ display:"inline-block", minWidth:"max-content" }}>
                      <div style={{ display:"flex", gap:GAP, marginLeft:half, marginBottom:GAP }}>{evens.map(chip)}</div>
                      <div style={{ display:"flex", gap:GAP }}>{odds.map(chip)}</div>
                    </div>
                  </div>
                </div>
              );
            })()}
            {/* actions */}
            {(() => {
              const isStillConflict = posConflict.position === posConflict.originalPosition;
              return (
                <div style={{ display:"flex", gap:10 }}>
                  <button onClick={forcePosConfirm}
                    style={{ ...btnBase, flex:1, padding: isMobile ? "13px" : "9px",
                      background: isStillConflict ? amber+"22" : accent+"22",
                      border:`1px solid ${isStillConflict ? amber : accent}`,
                      color: isStillConflict ? amber : accent,
                      fontWeight:700, fontSize: isMobile ? 14 : 12, cursor:"pointer" }}>
                    {isStillConflict ? "⚠ Forcer quand même" : `✓ Confirmer → ${posConflict.position}`}
                  </button>
                  <button onClick={() => setPosConflict(null)}
                    style={{ ...btnBase, padding:"9px 14px", background:"transparent",
                      border:`1px solid ${border}`, color:muted }}>
                    Annuler
                  </button>
                </div>
              );
            })()}
          </div>
        </div>
      )}

      {/* ── Force-position warning modal ─────────────────────────── */}
      {forceWarning && (
        <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.7)", display:"flex",
          alignItems:"center", justifyContent:"center", zIndex:1100 }}
          onClick={(e) => { if (e.target === e.currentTarget) setForceWarning(null); }}>
          <div style={{ background:surface, border:`2px solid ${danger}`, borderRadius:10,
            padding:"24px 20px", width:"min(400px, calc(100vw - 32px))", fontFamily:MONO,
            boxShadow:"0 8px 40px rgba(0,0,0,0.5)", boxSizing:"border-box" as const }}>
            {/* header */}
            <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:14 }}>
              <span style={{ fontSize:20 }}>⚠️</span>
              <span style={{ fontSize: isMobile ? 14 : 12, fontWeight:700, color:danger }}>Poids excessif en hauteur</span>
            </div>
            {/* coil info */}
            <div style={{ fontSize: isMobile ? 13 : 11, color:text, marginBottom:4 }}>
              <span style={{ color:muted }}>Réf. </span>
              <span style={{ fontWeight:700, wordBreak:"break-all" }}>{forceWarning.code}</span>
            </div>
            <div style={{ fontSize: isMobile ? 13 : 11, color:text, marginBottom:14 }}>
              <span style={{ color:muted }}>Pos. </span>
              <span style={{ fontWeight:700, color:accent }}>{forceWarning.position}</span>
              <span style={{ color:muted }}> — poids </span>
              <span style={{ fontWeight:700 }}>{fmtW(forceWarning.weight)} kg</span>
            </div>
            {/* offenders */}
            {forceWarning.offenders.map((o) => (
              <div key={o.pos} style={{ marginBottom:8, padding:"9px 12px", borderRadius:6,
                background:dark?"#2d0a0a":"#fee2e2", border:`1px solid ${danger}` }}>
                <div style={{ fontSize: isMobile ? 12 : 10, color:danger, marginBottom:2 }}>
                  Pos. <strong>{o.pos}</strong> (dessous) — {fmtW(o.weight)} kg
                </div>
                <div style={{ fontSize: isMobile ? 13 : 11, color:danger, fontWeight:700 }}>
                  +{o.pct}% plus lourd que cette bobine voisine
                </div>
              </div>
            ))}
            <div style={{ fontSize: isMobile ? 12 : 10, color:muted, marginBottom:16, marginTop:4 }}>
              Une bobine lourde en position haute sur une plus légère présente un risque de stabilité.
            </div>
            {/* actions */}
            <div style={{ display:"flex", gap:10 }}>
              <button onClick={forceConfirm}
                style={{ ...btnBase, flex:1, padding: isMobile ? "13px" : "9px",
                  background:danger+"22", border:`1px solid ${danger}`, color:danger,
                  fontWeight:700, fontSize: isMobile ? 14 : 12 }}>
                ⚠ Forcer quand même
              </button>
              <button onClick={() => setForceWarning(null)}
                style={{ ...btnBase, padding:"9px 14px", background:"transparent",
                  border:`1px solid ${border}`, color:muted }}>
                Modifier
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
