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

import { dechargementColor } from "./dechargementColor";

/* ── Types ────────────────────────────────────────────────────────── */
interface HangarItem  { code: string; weight: number | null; fromExcel: boolean; position?: string; wagonId?: string; dechargement?: string; }
interface HangarLine  { id: string; name: string; items: HangarItem[]; }
interface HangarProps { dark: boolean; }
interface Wagon { id: string; serial: string; coilCount: number; port?: string; }
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
  const bg      = dark ? "#0f0800" : "#fff5f2";
  const surface = dark ? "#1a0e00" : "#ffffff";
  const text    = dark ? "#fed7aa" : "#9f1239";
  const muted   = dark ? "#a16207" : "#be185d";
  const border  = dark ? "#3d1a00" : "#fecdd3";
  const danger  = dark ? "#f87171" : "#dc2626";
  const amber   = "#d97706";

  /* ── Breakpoints ─────────────────────────────────────────────── */
  const vw        = useWindowWidth();
  const isMobile  = vw < 640;

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
  useEffect(() => { LS.set("hgr_lines2",     lines);      }, [lines]);
  useEffect(() => { LS.set("hgr_selectedId", selectedId); }, [selectedId]);

  /* ── Wagons ───────────────────────────────────────────────────── */
  const [wagons,          setWagons]          = useState<Wagon[]>(() => LS.get("hgr_wagons", []));
  const [selectedWagonId, setSelectedWagonId] = useState<string | null>(() => LS.get("hgr_wagonId", null));
  const [wagonSerial,     setWagonSerial]     = useState("");
  const [wagonCoils,      setWagonCoils]      = useState("");
  const [editingWagonId,  setEditingWagonId]  = useState<string | null>(null);
  const [addingWagon,     setAddingWagon]     = useState(false);
  const [wagonPort,       setWagonPort]       = useState("");
  const [recapSortAsc,    setRecapSortAsc]    = useState<boolean | null>(null);
  const [wagonSortAsc,    setWagonSortAsc]    = useState<boolean | null>(null);
  const [clearConfirmId,  setClearConfirmId]  = useState<string | null>(null);
  const [hoveredWagonBtn, setHoveredWagonBtn] = useState<string | null>(null);
  const [hoveredLetter,   setHoveredLetter]   = useState<string | null>(null);
  const [hoveredChip,     setHoveredChip]     = useState<string | null>(null);
  const [itemsOpen,       setItemsOpen]       = useState(true);
  const [recapOpen,       setRecapOpen]       = useState(true);
  const [wagonSumOpen,    setWagonSumOpen]    = useState(true);
  const selectedWagonIdRef = useRef<string | null>(null);
  selectedWagonIdRef.current = selectedWagonId;
  useEffect(() => { LS.set("hgr_wagons",  wagons);          }, [wagons]);
  useEffect(() => { LS.set("hgr_wagonId", selectedWagonId); }, [selectedWagonId]);

  const addWagon = () => {
    const serial = wagonSerial.trim();
    const count  = parseInt(wagonCoils.trim(), 10);
    if (!serial || isNaN(count) || count < 0) return;
    const port = wagonPort.trim() || undefined;
    if (editingWagonId) {
      setWagons((p) => p.map((w) => w.id === editingWagonId ? { ...w, serial, coilCount: count, port } : w));
      setEditingWagonId(null);
    } else {
      const w: Wagon = { id: newId(), serial, coilCount: count, port };
      setWagons((p) => [...p, w]);
      setSelectedWagonId(w.id);
    }
    setAddingWagon(false);
    setWagonSerial(""); setWagonCoils(""); setWagonPort("");
  };
  const deleteWagon = (id: string) => {
    setWagons((p) => p.filter((w) => w.id !== id));
    if (selectedWagonId === id) setSelectedWagonId(null);
  };
  const startEditWagon = (w: Wagon) => {
    setEditingWagonId(w.id); setWagonSerial(w.serial); setWagonCoils(String(w.coilCount)); setWagonPort(w.port ?? "");
  };
  const cancelEditWagon = () => { setEditingWagonId(null); setAddingWagon(false); setWagonSerial(""); setWagonCoils(""); setWagonPort(""); };

  /* ── Pending (unknown code → manual weight entry) ───────────── */
  const [pending, setPending] = useState<{ code: string; weight: string; position: string; dechargement: string } | null>(null);
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
  const addItem = useCallback((code: string, weight: number | null, fromExcel: boolean, dechargement?: string) => {
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
      const n = [...prev]; n[idx] = { ...line, items: [...line.items, { code, weight, fromExcel, position: pos, wagonId: LS.get<string|null>("hgr_wagonId", null) ?? undefined, dechargement: dechargement || undefined }] }; return n;
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
      const n = [...prev]; n[idx] = { ...line, items: [...line.items, { code, weight: null, fromExcel: false, position: pos, wagonId: LS.get<string|null>("hgr_wagonId", null) ?? undefined }] }; return n;
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
      setPending({ code, weight: "", position: qaaPositionRef.current, dechargement: "" });
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
    addItem(pending.code, isNaN(w) ? null : w, false, pending.dechargement || undefined);
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
      n[idx] = { ...line, items: [...line.items, { code: posConflict.code, weight: posConflict.weight, fromExcel: posConflict.fromExcel, position: posConflict.position, wagonId: LS.get<string|null>("hgr_wagonId", null) ?? undefined }] };
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
  const selectOrCreateLetter = (letter: string) => {
    const existing = lines.find((l) => l.name === letter);
    if (existing) {
      setSelectedId(existing.id);
    } else {
      const l: HangarLine = { id: newId(), name: letter, items: [] };
      setLines((p) => [...p, l]);
      setSelectedId(l.id);
    }
  };
  const clearLine = (id: string) => setLines((p) => p.map((l) => l.id === id ? { ...l, items: [] } : l));
  const removeItem = (lineId: string, code: string) =>
    setLines((p) => p.map((l) => l.id === lineId ? { ...l, items: l.items.filter((it) => it.code !== code) } : l));

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
  const lineSummaryRaw = lines.map((l) => ({
    id: l.id, name: l.name,
    qty: l.items.length,
    weight: l.items.reduce((s, it) => s + (it.weight ?? 0), 0),
    hasNull: l.items.some((it) => it.weight === null),
  }));
  const lineSummary = recapSortAsc === null ? lineSummaryRaw
    : [...lineSummaryRaw].sort((a,b) => recapSortAsc
        ? a.name.localeCompare(b.name)
        : b.name.localeCompare(a.name));
  const totalQty    = lineSummary.reduce((s, l) => s + l.qty, 0);
  const totalWeight = lineSummary.reduce((s, l) => s + l.weight, 0);
  const anyNull     = lineSummary.some((l) => l.hasNull);

  const allItems = lines.flatMap((l) => l.items);
  const wagonSummaryRaw = wagons.map((w) => {
    const items = allItems.filter((it) => it.wagonId === w.id);
    return {
      ...w,
      scanned: items.length,
      weight:  items.reduce((s, it) => s + (it.weight ?? 0), 0),
      hasNull: items.some((it) => it.weight === null),
      done:    w.coilCount > 0 && items.length >= w.coilCount,
    };
  });
  const wagonSummary = wagonSortAsc === null ? wagonSummaryRaw
    : [...wagonSummaryRaw].sort((a,b) => wagonSortAsc
        ? a.serial.localeCompare(b.serial)
        : b.serial.localeCompare(a.serial));
  const existingPorts = [...new Set(wagons.map((w) => w.port).filter((p): p is string => !!p))];
  const existingDechargements = [...new Set(allItems.map((it) => it.dechargement).filter((d): d is string => !!d))];

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
  const pad     = isMobile ? "14px 10px" : "24px 28px";
  const btnBase: React.CSSProperties = { fontFamily: MONO, fontSize: isMobile ? 15 : 11, letterSpacing: isMobile ? 0 : "0.08em",
    padding: isMobile ? "12px 16px" : "4px 10px", borderRadius: 4, cursor: "pointer", border: "none" };
  const thS: React.CSSProperties = { padding: isMobile ? "10px 12px" : "5px 12px", textAlign: "left", fontSize: isMobile ? 13 : 10,
    color: muted, textTransform: isMobile ? "none" : "uppercase" as const, letterSpacing: isMobile ? 0 : "0.1em", fontWeight: 400,
    borderBottom: `1px solid ${border}`, whiteSpace: "nowrap" };
  const tdS = (right?: boolean): React.CSSProperties => ({
    padding: isMobile ? "10px 12px" : "5px 12px", fontSize: isMobile ? 14 : 11, color: text, textAlign: right ? "right" : "left",
    borderBottom: `1px solid ${border}`, whiteSpace: "nowrap" });

  return (
    <div style={{ fontFamily: MONO, color: text, background: bg, minHeight: "calc(100vh - 44px)" }}>
      <div style={{ maxWidth: 675, margin: "0 auto", padding: pad, boxSizing: "border-box", overflowX: "hidden" }}>

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

        {/* ── Camera column ───────────────────────────────────── */}
        <div style={{ display:"flex", flexDirection:"column", gap:10, alignItems:"stretch", width:"100%", flexShrink: 0 }}>

          {/* Video + A–Z letter sidebar */}
          <div style={{ display:"flex", gap:4, alignItems:"stretch", width:"100%" }}>
            {/* A–Z column */}
            <div style={{ position:"relative", display:"flex", flexDirection:"column", gap:2, flexShrink:0, overflow:"visible" }}>
              {Array.from({length:26},(_,i)=>String.fromCharCode(65+i)).map((letter) => {
                const isActiveLetter = selectedLine?.name === letter;
                const lineObj = lines.find((l) => l.name === letter);
                const hasLine = !!lineObj;
                const hasContent = hasLine && (lineObj?.items ?? []).length > 0;
                const isHov = hoveredLetter === letter;
                return (
                  <button key={letter}
                    onClick={() => selectOrCreateLetter(letter)}
                    onMouseEnter={() => setHoveredLetter(letter)}
                    onMouseLeave={() => setHoveredLetter(null)}
                    title={hasLine ? `Sélectionner la ligne "${letter}"` : `Créer et sélectionner la ligne "${letter}"`}
                    style={{
                      fontFamily:MONO, fontSize: isMobile ? 13 : 9, fontWeight: isActiveLetter ? 700 : 600,
                      width: isMobile ? 32 : 22, flex:1,
                      padding:0, lineHeight:1, borderRadius:3, cursor:"pointer",
                      position:"relative",
                      zIndex: isHov ? 20 : 1,
                      transform: isHov ? "scale(2.5) translateX(28%)" : "scale(1)",
                      transition:"transform 0.12s",
                      background: isActiveLetter ? accent+"33" : (hasContent ? (isHov ? surface : (dark?"#1e1e1e":"#e8e8e8")) : (isHov ? surface : "transparent")),
                      border: isActiveLetter ? `1px solid ${accent}` : `1px solid ${border}`,
                      color: isActiveLetter ? accent : (hasContent ? text : muted),
                    }}>
                    {letter}
                  </button>
                );
              })}
              {/* Déchargement hover overlay */}
              {hoveredLetter && (() => {
                const hovLine = lines.find((l) => l.name === hoveredLetter);
                const dechs = [...new Set((hovLine?.items ?? []).map((it) => it.dechargement).filter((d): d is string => !!d))];
                if (dechs.length === 0) return null;
                return (
                  <div style={{ position:"absolute", left:"calc(100% + 8px)", top:0,
                    background:surface, border:`1px solid ${dark?"#4ade80":"#16a34a"}`,
                    borderRadius:6, padding:"7px 10px", zIndex:30,
                    minWidth:130, pointerEvents:"none",
                    boxShadow:"0 4px 20px rgba(0,0,0,0.35)" }}>
                    <div style={{ fontSize:9, color:muted, textTransform:"uppercase",
                      letterSpacing:"0.1em", marginBottom:5 }}>Déchargements · {hoveredLetter}</div>
                    {dechs.map((d) => (
                      <div key={d} style={{ fontFamily:MONO, fontSize: isMobile ? 12 : 10,
                        color:dark?"#86efac":"#166534", fontWeight:700 }}>{d}</div>
                    ))}
                  </div>
                );
              })()}
            </div>
            {/* Camera frame */}
            <div style={{ position:"relative", borderRadius:8, overflow:"hidden",
              border:`2px solid ${flashColor ?? border}`,
              boxShadow: flashColor ? `0 0 20px ${flashColor}55` : "none",
              transition:"border-color 0.15s, box-shadow 0.15s",
              flex:1, aspectRatio:"4/3", background:dark?"#0a0a0a":"#ddd",
              display:"flex", alignItems:"center", justifyContent:"center" }}>
              <video ref={videoRef} playsInline muted
                style={{ width:"100%", height:"100%", display:cameraOn?"block":"none", objectFit:"cover" }} />
              {!cameraOn && <div style={{ textAlign:"center", color:muted, fontSize:12, padding:20 }}><div style={{fontSize:40, marginBottom:8}}>📷</div><div>Caméra arrêtée</div></div>}
              {cameraOn && (
                <div style={{ position:"absolute", inset:0, pointerEvents:"none" }}>
                  <div style={{ position:"absolute", top:"50%", left:"50%", transform:"translate(-50%,-50%)", width:200, height:200 }}>
                    {[{t:-2,l:-2,bt:"borderTop",bl:"borderLeft",br:"borderTopLeftRadius"},
                      {t:-2,r:-2,bt:"borderTop",bl:"borderRight",br:"borderTopRightRadius"},
                      {b:-2,l:-2,bt:"borderBottom",bl:"borderLeft",br:"borderBottomLeftRadius"},
                      {b:-2,r:-2,bt:"borderBottom",bl:"borderRight",br:"borderBottomRightRadius"}
                    ].map((c,i) => (
                      <div key={i} style={{ position:"absolute", width:22, height:22,
                        top:c.t, left:c.l, bottom:c.b, right:c.r,
                        [c.bt]:`3px solid ${accent}`, [c.bl]:`3px solid ${accent}`,
                        [c.br]:6 } as React.CSSProperties} />
                    ))}
                  </div>
                </div>
              )}
            </div>
            {/* Wagon column (right) */}
            <div style={{ display:"flex", flexDirection:"column", gap:2, flexShrink:0, width: isMobile ? 44 : 32, overflow:"visible" }}>
              {Array.from({length:26}).map((_, idx) => {
                const w = wagons[idx] as (typeof wagons)[number] | undefined;
                if (!w) {
                  if (idx === wagons.length) {
                    return (
                      <button key="add"
                        onClick={() => { setAddingWagon(true); setEditingWagonId(null); setWagonSerial(""); setWagonCoils(""); }}
                        onMouseEnter={() => setHoveredWagonBtn("add")}
                        onMouseLeave={() => setHoveredWagonBtn(null)}
                        title="Ajouter un wagon"
                        style={{ fontFamily:MONO, fontSize: isMobile ? 14 : 10, flex:1,
                          padding:0, lineHeight:1, borderRadius:3, cursor:"pointer", position:"relative",
                          zIndex: hoveredWagonBtn==="add" ? 20 : 1,
                          transform: hoveredWagonBtn==="add" ? "scale(2.6) translateX(-28%)" : "scale(1)",
                          transition:"transform 0.12s",
                          background: hoveredWagonBtn==="add" ? surface : "transparent",
                          border:`1px dashed ${border}`, color:muted }}>+</button>
                    );
                  }
                  return <div key={`empty-${idx}`} style={{ flex:1, minHeight:0 }} />;
                }
                const isSel = w.id === selectedWagonId;
                const sc = allItems.filter((it) => it.wagonId === w.id).length;
                const done = w.coilCount > 0 ? sc >= w.coilCount : false;
                const shortSerial = w.serial.length > 5 ? w.serial.slice(0,4)+"…" : w.serial;
                return (
                  <button key={w.id}
                    onClick={() => {
                      if (isSel) { startEditWagon(w); setAddingWagon(false); }
                      else { setSelectedWagonId(w.id); setEditingWagonId(null); setAddingWagon(false); }
                    }}
                    onMouseEnter={() => setHoveredWagonBtn(w.id)}
                    onMouseLeave={() => setHoveredWagonBtn(null)}
                    title={`${w.serial} — ${sc}${w.coilCount > 0 ? `/${w.coilCount}` : ""} coil(s)${isSel ? " · cliquer pour modifier" : ""}`}
                    style={{
                      fontFamily:MONO, fontWeight: isSel ? 700 : 600,
                      width:"100%", flex:1,
                      padding:"1px 2px", lineHeight:1.1, borderRadius:3, cursor:"pointer",
                      display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", gap:1,
                      position:"relative",
                      zIndex: hoveredWagonBtn===w.id ? 20 : 1,
                      transform: hoveredWagonBtn===w.id ? "scale(2.5) translateX(-26%)" : "scale(1)",
                      transition:"transform 0.12s",
                      background: isSel ? accent+"33" : done ? (dark?"#0a200f":"#dcfce7") : (hoveredWagonBtn===w.id ? surface : (dark?"#1e1e1e":"#e8e8e8")),
                      border: isSel ? `1px solid ${accent}` : done ? `1px solid ${dark?"#4ade80":"#16a34a"}` : `1px solid ${border}`,
                      color: isSel ? accent : done ? (dark?"#86efac":"#166534") : text,
                    }}>
                    <span style={{ overflow:"hidden", textOverflow:"ellipsis", maxWidth:"100%",
                      whiteSpace:"nowrap", display:"block", fontSize: isMobile ? 11 : 8 }}>
                      {shortSerial}
                    </span>
                    <span style={{ fontSize: isMobile ? 10 : 7,
                      color: done?(dark?"#86efac":"#166534"):(sc>0?accent:muted),
                      fontWeight:done?700:400, lineHeight:1 }}>
                      {sc}{w.coilCount > 0 ? `/${w.coilCount}` : ""}{done?"✓":""}
                    </span>
                  </button>
                );
              })}
            </div>
          </div>

          {/* Row 1 : all controls on one line */}
          {(() => {
            const canScan = !!selectedId && !!qaaPosition;
            const disabledStyle: React.CSSProperties = { opacity: 0.35, pointerEvents:"none" as const, cursor:"not-allowed" as const };
            return (
          <div style={{ display:"flex", gap:8, alignItems:"center", flexWrap:"wrap" }}>
            <button onClick={cameraOn ? stopCamera : startCamera}
              disabled={!canScan && !cameraOn}
              style={{ ...btnBase, fontSize: isMobile ? 15 : 12, padding: isMobile ? "12px 20px" : "8px 16px",
                flexShrink:0,
                background: cameraOn?(dark?"#2d0a0a":"#fee2e2"):(dark?"#0a200f":"#dcfce7"),
                border:`1px solid ${cameraOn?"#ef4444":accent}`,
                color: cameraOn?"#ef4444":accent, fontWeight:700,
                ...(!canScan && !cameraOn ? disabledStyle : {}) }}>
              {cameraOn ? "⏹ Arrêter" : "▶ Caméra"}
            </button>
            <button onClick={addNullItem}
              disabled={!canScan}
              title="Ajouter un emplacement vide (sans référence)"
              style={{ ...btnBase, fontSize: isMobile ? 14 : 11, padding: isMobile ? "12px 16px" : "8px 13px",
                flexShrink:0,
                background: "transparent", border:`1px solid ${muted}`, color: muted,
                ...(!canScan ? disabledStyle : {}) }}>
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
              onKeyDown={(e)=>e.key==="Enter"&&canScan&&submitManual()}
              disabled={!canScan}
              placeholder={canScan ? "Saisie manuelle…" : "Sélectionner ligne + emplacement"}
              style={{ flex:1, minWidth: isMobile ? 120 : 80, fontFamily:MONO, fontSize: isMobile ? 15 : 11,
                padding: isMobile ? "10px 12px" : "6px 10px", borderRadius:4,
                background:surface, border:`1px solid ${border}`, color: canScan ? text : muted, outline:"none",
                ...(!canScan ? { opacity:0.45 } : {}) }} />
            <button onClick={submitManual} disabled={!canScan}
              style={{ ...btnBase, padding: isMobile ? "12px 14px" : "6px 10px",
                background:accent+"22", border:`1px solid ${accent}`, color:accent, fontWeight:700, flexShrink:0,
                ...(!canScan ? disabledStyle : {}) }}>↵</button>
          </div>
            );
          })()}

          {/* Scan feedback */}
          {lastScan && (() => {
            const scanItem = lastScan.status==="added" ? (selectedLine?.items ?? []).find((it) => it.code === lastScan.code) : undefined;
            const scanWagon = scanItem?.wagonId ? wagons.find((w) => w.id === scanItem.wagonId) : undefined;
            return (
            <div style={{ padding:"6px 10px", borderRadius:6, width:"100%", boxSizing:"border-box" as const,
              display:"flex", alignItems:"center", gap:8, flexWrap:"wrap",
              background: lastScan.status==="added"?(dark?"#0a1f0f":"#f0fdf4"):(dark?"#1c1000":"#fffbeb"),
              border:`1px solid ${lastScan.status==="added"?accent:amber}` }}>
              <span style={{fontSize:9,color:muted,textTransform:"uppercase",letterSpacing:"0.1em",flexShrink:0}}>Dernière réf.</span>
              <span style={{fontFamily:MONO,fontSize: isMobile ? 13 : 11,fontWeight:700,color:text,wordBreak:"break-all",flex:1,minWidth:0}}>{lastScan.code}</span>
              {lastScan.status==="added"    && <span style={{fontSize: isMobile ? 12 : 10,color:accent,flexShrink:0}}>✓ Ajouté{qaaPosition ? ` → ${qaaPosition}` : ""}</span>}
              {lastScan.status==="duplicate"&& <span style={{fontSize: isMobile ? 12 : 10,color:amber,flexShrink:0}}>⚠ déjà présent</span>}
              {lastScan.status==="posdup"   && <span style={{fontSize: isMobile ? 12 : 10,color:amber,flexShrink:0}}>⚠ pos. "{qaaPosition}" occupée</span>}
              {lastScan.status==="noline"   && <span style={{fontSize: isMobile ? 12 : 10,color:amber,flexShrink:0}}>⚠ aucune ligne</span>}
              {lastScan.status==="unknown"  && <span style={{fontSize: isMobile ? 12 : 10,color:amber,flexShrink:0}}>? absent du fichier Excel</span>}
              {scanItem && scanItem.weight !== null && <span style={{fontSize: isMobile ? 11 : 9,color:text,flexShrink:0,fontFamily:MONO}}>{fmtW(scanItem.weight)} kg</span>}
              {scanItem?.position && <span style={{fontSize: isMobile ? 11 : 9,color:accent,flexShrink:0,fontFamily:MONO,fontWeight:700}}>{selectedLine?.name}{String(parseInt(scanItem.position)).padStart(2,"0")}</span>}
              {scanItem?.dechargement && <span style={{fontSize: isMobile ? 11 : 9,color:dark?"#86efac":"#166534",flexShrink:0,fontWeight:700}}>{scanItem.dechargement}</span>}
              {scanWagon && <span style={{fontSize: isMobile ? 11 : 9,fontFamily:MONO,padding:"1px 6px",borderRadius:8,flexShrink:0,background:accent+"22",border:`1px solid ${accent}`,color:accent}}>{scanWagon.serial}</span>}
            </div>
            );
          })()}

          {/* Row 3 : position slots 1–51 + slider */}
          {(() => {
            const CHIP = isMobile ? 34 : 26;
            const GAP  = 3;
            const half = (CHIP + GAP) / 2;
            const odds  = Array.from({length:26},(_,i)=>String(i*2+1));  /* 1,3,5…51 */
            const evens = Array.from({length:25},(_,i)=>String(i*2+2));  /* 2,4,6…50 */
            const chip = (s: string) => {
              const isSel  = qaaPosition === s;
              const isOcc  = occupiedPositions.has(s);
              const isVide = videPositions.has(s);
              const item   = isOcc ? (selectedLine?.items ?? []).find((it) => it.position === s) : undefined;
              const dech   = item?.dechargement ?? "";
              const dc     = isOcc && !isVide && dech ? dechargementColor(dech, dark) : null;
              return (
                <div key={s} style={{ position:"relative", display:"inline-block" }}
                  onMouseEnter={() => setHoveredChip(s)}
                  onMouseLeave={() => setHoveredChip(null)}>
                <button disabled={isOcc && !isSel}
                  onClick={() => setQaaPosition(isSel ? "" : s)}
                  title={isOcc && !isSel ? `Emplacement ${s} occupé${isVide ? " (∅ vide)" : ""}${dech ? " · " + dech : ""}` : `Emplacement ${s}`}
                  style={{
                    fontFamily:MONO, fontSize: isMobile ? 12 : 9,
                    width:CHIP, height:CHIP, flexShrink:0,
                    borderRadius:4, padding:0, lineHeight:1,
                    cursor: isOcc && !isSel ? "not-allowed" : "pointer",
                    border: isSel ? `2px solid ${accent}` : dc ? `1px solid ${dc.border}` : `1px solid ${isOcc ? (dark?"#4a3a00":"#d97706") : border}`,
                    background: isSel ? accent+"33"
                      : dc  ? dc.bg
                      : isOcc ? (dark ? "#2a1a00" : "#fde68a")
                      : (dark ? "#1a1a1a" : "#f0f0f0"),
                    color: isSel ? accent : dc ? dc.color : isOcc ? "#92400e" : muted,
                    fontWeight: isSel ? 700 : 400,
                    opacity: isOcc && !isSel ? 0.85 : 1,
                  }}>{s}</button>
                </div>
              );
            };
            const hovItemGlobal = hoveredChip
              ? (selectedLine?.items ?? []).find((it) => it.position === hoveredChip && occupiedPositions.has(hoveredChip))
              : undefined;
            return (
              <div style={{ borderRadius:6, border:`1px solid ${dark?"#4ade80":"#16a34a"}`,
                background: dark?"#071a0b":"#f0fdf4", padding:"8px 10px", overflow:"hidden" }}>
                <div style={{ fontSize: isMobile ? 12 : 10, color:muted, marginBottom:5 }}>
                  PAR EMPLACEMENT
                  {qaaPosition && <span style={{color:accent, fontWeight:700, marginLeft:6}}>→ {qaaPosition}</span>}
                </div>
                {/* Scrollable chip grid */}
                <div style={{ overflowX:"auto", paddingBottom:2 }}>
                  <div style={{ display:"inline-block", minWidth:"max-content" }}>
                    <div style={{ display:"flex", gap:GAP, marginLeft:half, marginBottom:GAP }}>
                      {evens.map(chip)}
                    </div>
                    <div style={{ display:"flex", gap:GAP }}>
                      {odds.map(chip)}
                    </div>
                  </div>
                </div>
                {hovItemGlobal && (
                  <div style={{ marginTop:4, borderRadius:5,
                    background:surface, border:`1px solid ${dark?"#4ade80":"#16a34a"}`,
                    padding:"6px 10px", display:"flex", gap:12, flexWrap:"wrap", alignItems:"center" }}>
                    <span style={{ fontSize:9, color:muted, textTransform:"uppercase", letterSpacing:"0.1em", flexShrink:0 }}>
                      {selectedLine?.name}{String(parseInt(hoveredChip!)).padStart(2,"0")}
                    </span>
                    <span style={{ fontFamily:MONO, fontSize: isMobile ? 12 : 10, color:text, fontWeight:700, flex:1, minWidth:0, wordBreak:"break-all" }}>
                      {hovItemGlobal.code.startsWith("∅") ? "∅ vide" : hovItemGlobal.code}
                    </span>
                    <span style={{ fontSize: isMobile ? 11 : 9, color: hovItemGlobal.weight===null?amber:text, flexShrink:0 }}>
                      {hovItemGlobal.weight !== null ? `${fmtW(hovItemGlobal.weight)} kg` : "? kg"}
                    </span>
                    {hovItemGlobal.dechargement && (
                      <span style={{ fontSize: isMobile ? 11 : 9, color:dark?"#86efac":"#166534", fontWeight:700, flexShrink:0 }}>
                        {hovItemGlobal.dechargement}
                      </span>
                    )}
                  </div>
                )}
              </div>
            );
          })()}

          {cameraError && (
            <div style={{ fontSize: isMobile ? 14 : 11, color:"#ef4444", padding:"7px 12px",
              background:dark?"#2d0a0a":"#fee2e2", borderRadius:4, border:"1px solid #ef4444" }}>⚠ {cameraError}</div>
          )}
        </div>

        {/* ── Detail of selected line ──────────────────────────── */}
        {selectedLine && (
          <div style={{ flex:1, minWidth:0, width:"100%" }}>
            <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:10 }}>
              <div style={{ fontSize:12, fontWeight:700, color:accent }}>
                {selectedLine.name}
                <span style={{ marginLeft:8, fontSize:10, color:muted, fontWeight:400 }}>
                  {selectedLine.items.length} art. · {fmtW(selectedLine.items.reduce((s,it)=>s+(it.weight??0),0))} kg
                </span>
              </div>
              <button onClick={() => setClearConfirmId(selectedLine.id)}
                style={{ ...btnBase, background:"transparent", border:`1px solid ${danger}`, color:danger }}>Vider</button>
            </div>
            {selectedLine.items.length === 0 ? (
              <div style={{ fontSize:11, color:muted }}>Aucun article — scannez ou saisissez une réf.</div>
            ) : (
              <>
                {/* collapsible header */}
                <div onClick={() => setItemsOpen(v => !v)}
                  style={{ display:"flex", alignItems:"center", gap:8, cursor:"pointer", userSelect:"none" as const,
                    marginBottom:4, padding:"3px 0" }}>
                  <span style={{ fontSize: isMobile ? 12 : 9, color:muted, textTransform:"uppercase", letterSpacing:"0.1em", flex:1 }}>
                    {selectedLine.items.length} article{selectedLine.items.length>1?"s":""}
                  </span>
                  <span style={{ fontSize:9, color:muted }}>{itemsOpen ? "▲" : "▼"}</span>
                </div>
                {itemsOpen && (
                  <div style={{ borderRadius:6, border:`1px solid ${dark?"#4ade80":"#16a34a"}`,
                    background: dark?"#071a0b":"#f0fdf4", overflow:"hidden" }}>
                  <div style={{ overflowX:"auto" }}>
                  <div style={{ maxHeight:380, overflowY:"auto" }}>
                    <table style={{ borderCollapse:"collapse", minWidth:"100%" }}>
                      {(() => {
                        const showPos = selectedLine.items.some((it) => it.position);
                        const showDechargement = selectedLine.items.some((it) => it.dechargement);
                        const showWagon = wagons.length > 0;
                        return (
                          <>
                          <thead>
                            <tr>
                              <th style={thS}>#</th>
                              {showWagon && <th style={thS}>Wagon</th>}
                              <th style={thS}>Réf.</th>
                              <th style={{...thS, textAlign:"right"}}>Poids (kg)</th>
                              {showPos && <th style={thS}>Pos.</th>}
                              {showDechargement && <th style={thS}>Déchargement</th>}
                              <th style={thS}>Source</th>
                              <th style={thS}></th>
                            </tr>
                          </thead>
                          <tbody>
                            {selectedLine.items.map((it, i) => {
                              const isVide = it.code.startsWith("∅");
                              const wag = showWagon ? wagons.find((w) => w.id === it.wagonId) : undefined;
                              return (
                              <tr key={it.code}>
                                <td style={tdS()}><span style={{color: muted}}>{i+1}</span></td>
                                {showWagon && (
                                  <td style={tdS()}>
                                    {wag
                                      ? <span style={{ fontSize:9, padding:"1px 6px", borderRadius:8,
                                          background: accent+"22", border:`1px solid ${accent}`, color:accent,
                                          fontFamily:MONO, whiteSpace:"nowrap" }}>{wag.serial}</span>
                                      : <span style={{ color:muted }}>—</span>}
                                  </td>
                                )}
                                <td style={{...tdS(), color: text}}>{isVide ? "∅ vide" : it.code}</td>
                                <td style={{...tdS(true), color: it.weight===null?amber:text}}>{fmtW(it.weight)}</td>
                                {showPos && <td style={{...tdS(), fontFamily:MONO, fontWeight:700, color: it.position?accent:muted}}>{it.position ? `${selectedLine.name}${String(parseInt(it.position)).padStart(2,"0")}` : "—"}</td>}
                                {showDechargement && <td style={{...tdS(), color: it.dechargement?text:muted}}>{it.dechargement ?? "—"}</td>}
                                <td style={tdS()}>
                                  <span style={{ fontSize:9, padding:"1px 6px", borderRadius:8,
                                    background: it.fromExcel?(dark?"#0a200f":"#dcfce7"):(dark?"#1a100a":"#fef3c7"),
                                    border:`1px solid ${it.fromExcel?accent:amber}`,
                                    color: it.fromExcel?accent:amber }}>
                                    {isVide ? "∅" : (it.fromExcel?"Excel":"Manuel")}
                                  </span>
                                </td>
                                <td style={tdS()}>
                                  <button onClick={()=>removeItem(selectedLine.id, it.code)}
                                    style={{...btnBase, padding:"1px 6px", background:"transparent", border:"none", color: muted, fontSize:12}}>×</button>
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
                  </div>
                  </div>
                )}
              </>
            )}
          </div>
        )}
      </div>

      {/* ── Wagons summary ──────────────────────────────────────── */}
      {wagons.length > 0 && (
        <div style={{ marginTop:24, maxWidth:"100%" }}>
          <div onClick={() => setWagonSumOpen(v => !v)}
            style={{ display:"flex", alignItems:"center", gap:8, cursor:"pointer", userSelect:"none" as const, marginBottom: wagonSumOpen ? 10 : 0 }}>
            <span style={{ fontSize: isMobile ? 13 : 10, color:muted, textTransform: isMobile ? "none" : "uppercase", letterSpacing: isMobile ? 0 : "0.12em", flex:1 }}>Par wagon</span>
            <span style={{ fontSize:9, color:muted }}>{wagonSumOpen ? "▲" : "▼"}</span>
          </div>
          {wagonSumOpen && (
            <div style={{ borderRadius:6, border:`1px solid ${dark?"#4ade80":"#16a34a"}`, background:dark?"#071a0b":"#f0fdf4", overflow:"hidden" }}>
            <div style={{ overflowX:"auto" }}>
            <table style={{ borderCollapse:"collapse", minWidth:"100%" }}>
            <thead>
              <tr style={{ background:dark?"#1a0e00":"#fff0ee" }}>
                <th style={{...thS, cursor:"pointer", userSelect:"none" as const}}
                  onDoubleClick={() => setWagonSortAsc(v => v === null ? true : v === true ? false : null)}
                  title="Double-clic pour trier">
                  N° série{wagonSortAsc === true ? " ▲" : wagonSortAsc === false ? " ▼" : ""}
                </th>
                <th style={thS}>Port</th>
                <th style={{...thS, textAlign:"right"}}>Attendus</th>
                <th style={{...thS, textAlign:"right"}}>Scannés</th>
                <th style={{...thS, textAlign:"right"}}>Restants</th>
                <th style={{...thS, textAlign:"right"}}>Poids (kg)</th>
              </tr>
            </thead>
            <tbody>
              {wagonSummary.map((w) => (
                <tr key={w.id}
                  onClick={() => setSelectedWagonId(w.id === selectedWagonId ? null : w.id)}
                  style={{ cursor:"pointer", background: w.id===selectedWagonId ? accent+"28" : (w.done ? (dark?"#0a1a0a":"#f0fff0") : undefined) }}>
                  <td style={{...tdS(), fontWeight:700, color: w.id===selectedWagonId ? accent : text}}>{w.serial}</td>
                  <td style={tdS()}>{w.port ?? <span style={{color:muted}}>—</span>}</td>
                  <td style={tdS(true)}>{w.coilCount > 0 ? w.coilCount : "∞"}</td>
                  <td style={{...tdS(true), color: w.done ? (dark?"#86efac":"#166534") : (w.scanned > 0 ? accent : muted), fontWeight: w.done ? 700 : 400}}>
                    {w.scanned}{w.done ? " ✓" : ""}
                  </td>
                  <td style={{...tdS(true), color: w.coilCount > 0 && w.coilCount - w.scanned > 0 ? (dark?"#fca5a5":"#dc2626") : muted}}>
                    {w.coilCount > 0 ? Math.max(0, w.coilCount - w.scanned) : "–"}
                  </td>
                  <td style={{...tdS(true), color: w.hasNull ? amber : text}}>{fmtW(w.weight)}{w.hasNull?" *":""}</td>
                </tr>
              ))}
            </tbody>
          </table>
          </div>
          </div>
          )}
        </div>
      )}

      {/* ── Summary table ───────────────────────────────────────── */}
      {lines.length > 0 && (
        <div style={{ marginTop:24, maxWidth:"100%" }}>
          <div onClick={() => setRecapOpen(v => !v)}
            style={{ display:"flex", alignItems:"center", gap:8, cursor:"pointer", userSelect:"none" as const, marginBottom: recapOpen ? 10 : 0 }}>
            <span style={{ fontSize: isMobile ? 13 : 10, color:muted, textTransform: isMobile ? "none" : "uppercase", letterSpacing: isMobile ? 0 : "0.12em", flex:1 }}>Récapitulatif</span>
            <span style={{ fontSize:9, color:muted }}>{recapOpen ? "▲" : "▼"}</span>
          </div>
          {recapOpen && (
            <div style={{ borderRadius:6, border:`1px solid ${dark?"#4ade80":"#16a34a"}`, background:dark?"#071a0b":"#f0fdf4", overflow:"hidden" }}>
            <div style={{ overflowX:"auto" }}>
            <table style={{ borderCollapse:"collapse", minWidth:"100%" }}>
            <thead>
              <tr style={{ background:dark?"#1a1a1a":"#f0f0f0" }}>
                <th style={{...thS, cursor:"pointer", userSelect:"none" as const}}
                  onDoubleClick={() => setRecapSortAsc(v => v === null ? true : v === true ? false : null)}
                  title="Double-clic pour trier">
                  Ligne{recapSortAsc === true ? " ▲" : recapSortAsc === false ? " ▼" : ""}
                </th>
                <th style={{...thS, textAlign:"right"}}>Qté</th>
                <th style={{...thS, textAlign:"right"}}>Poids total (kg)</th>
              </tr>
            </thead>
            <tbody>
              {lineSummary.map((l, lIdx) => {
              if (l.qty === 0) return null;
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
          </div>
          </div>
          )}
          {recapOpen && anyNull && <div style={{ fontSize: isMobile ? 13 : 10, color:amber, marginTop:6 }}>* Poids inconnu pour certains articles (non saisi)</div>}
        </div>
      )}

      {/* ── Vider confirm modal ──────────────────────────────────── */}
      {clearConfirmId && (() => {
        const lineName = lines.find((l) => l.id === clearConfirmId)?.name ?? "?";
        return (
          <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.55)", display:"flex",
            alignItems:"center", justifyContent:"center", zIndex:1000 }}
            onClick={(e) => { if (e.target === e.currentTarget) setClearConfirmId(null); }}>
            <div style={{ background:surface, border:`2px solid ${danger}`, borderRadius:10,
              padding:"24px 20px", width:"min(340px, calc(100vw - 32px))", fontFamily:MONO,
              boxShadow:"0 8px 40px rgba(0,0,0,0.4)", boxSizing:"border-box" as const }}>
              <div style={{ fontSize: isMobile ? 16 : 13, fontWeight:700, color:danger, marginBottom:10 }}>Vider la ligne {lineName} ?</div>
              <div style={{ fontSize: isMobile ? 13 : 11, color:muted, marginBottom:20 }}>
                Tous les articles de cette ligne seront supprimés définitivement.
              </div>
              <div style={{ display:"flex", gap:8, justifyContent:"flex-end" }}>
                <button onClick={() => setClearConfirmId(null)}
                  style={{ ...btnBase, padding: isMobile ? "10px 16px" : "7px 12px", background:"transparent",
                    border:`1px solid ${border}`, color:muted }}>Annuler</button>
                <button onClick={() => { clearLine(clearConfirmId!); setClearConfirmId(null); }}
                  style={{ ...btnBase, padding: isMobile ? "10px 16px" : "7px 12px", background:danger+"22",
                    border:`1px solid ${danger}`, color:danger, fontWeight:700 }}>Vider</button>
              </div>
            </div>
          </div>
        );
      })()}

      {/* ── Wagon add/edit modal ────────────────────────────────── */}
      {(addingWagon || editingWagonId) && (
        <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.6)", display:"flex",
          alignItems:"center", justifyContent:"center", zIndex:1000 }}
          onClick={(e) => { if (e.target === e.currentTarget) cancelEditWagon(); }}>
          <div style={{ background:surface, border:`1px solid ${border}`, borderRadius:10,
            padding:"24px 20px", width:"min(380px, calc(100vw - 32px))", fontFamily:MONO,
            boxShadow:"0 8px 40px rgba(0,0,0,0.4)", boxSizing:"border-box" as const }}>

            <div style={{ fontSize: isMobile ? 14 : 10, color:muted, textTransform: isMobile ? "none" : "uppercase",
              letterSpacing: isMobile ? 0 : "0.12em", marginBottom:16 }}>
              {editingWagonId ? "Modifier le wagon" : "Nouveau wagon"}
            </div>

            {/* N° série */}
            <div style={{ marginBottom:12 }}>
              <div style={{ fontSize: isMobile ? 13 : 10, color:muted, marginBottom:4 }}>N° série *</div>
              <input value={wagonSerial} onChange={(e) => setWagonSerial(e.target.value)}
                onKeyDown={(e) => { if(e.key==="Escape") cancelEditWagon(); }}
                placeholder="ex : W-0042" autoFocus
                style={{ width:"100%", boxSizing:"border-box" as const,
                  fontFamily:MONO, fontSize: isMobile ? 15 : 12,
                  padding: isMobile ? "10px 12px" : "7px 10px", borderRadius:5,
                  background:dark?"#0f0800":"#fff",
                  border:`1px solid ${wagonSerial.trim() ? accent : border}`, color:text, outline:"none" }} />
            </div>

            {/* Coils */}
            <div style={{ marginBottom:12 }}>
              <div style={{ fontSize: isMobile ? 13 : 10, color:muted, marginBottom:4 }}>Coils attendus <span style={{fontWeight:400}}>(0 = illimité)</span></div>
              <input value={wagonCoils} onChange={(e) => setWagonCoils(e.target.value.replace(/[^0-9]/g,""))}
                onKeyDown={(e) => { if(e.key==="Escape") cancelEditWagon(); }}
                placeholder="ex : 24"
                style={{ width:"100%", boxSizing:"border-box" as const,
                  fontFamily:MONO, fontSize: isMobile ? 15 : 12,
                  padding: isMobile ? "10px 12px" : "7px 10px", borderRadius:5,
                  background:dark?"#0f0800":"#fff", border:`1px solid ${border}`, color:text, outline:"none" }} />
            </div>

            {/* Port */}
            <div style={{ marginBottom:20 }}>
              <div style={{ fontSize: isMobile ? 13 : 10, color:muted, marginBottom:4 }}>Port <span style={{fontWeight:400}}>(optionnel)</span></div>
              <input value={wagonPort} onChange={(e) => setWagonPort(e.target.value)}
                list="wagon-ports-list"
                onKeyDown={(e) => { if(e.key==="Escape") cancelEditWagon(); }}
                placeholder={existingPorts.length > 0 ? "Choisir ou saisir…" : "ex : DUNKERQUE"}
                style={{ width:"100%", boxSizing:"border-box" as const,
                  fontFamily:MONO, fontSize: isMobile ? 15 : 12,
                  padding: isMobile ? "10px 12px" : "7px 10px", borderRadius:5,
                  background:dark?"#0f0800":"#fff", border:`1px solid ${wagonPort.trim() ? accent : border}`,
                  color:text, outline:"none" }} />
              <datalist id="wagon-ports-list">
                {existingPorts.map((p) => <option key={p} value={p} />)}
              </datalist>
              {existingPorts.length > 0 && (
                <div style={{ display:"flex", flexWrap:"wrap", gap:4, marginTop:6 }}>
                  {existingPorts.map((p) => (
                    <button key={p} onClick={() => setWagonPort(p)}
                      style={{ fontFamily:MONO, fontSize: isMobile ? 12 : 9, padding:"2px 8px",
                        borderRadius:10, cursor:"pointer", border:`1px solid ${wagonPort===p?accent:border}`,
                        background: wagonPort===p ? accent+"22" : "transparent",
                        color: wagonPort===p ? accent : muted }}>{p}</button>
                  ))}
                </div>
              )}
            </div>

            {/* Actions */}
            <div style={{ display:"flex", gap:8, justifyContent:"flex-end" }}>
              {editingWagonId && (
                <button onClick={() => { deleteWagon(editingWagonId!); cancelEditWagon(); }}
                  style={{ ...btnBase, padding: isMobile ? "10px 16px" : "7px 12px", background:"transparent",
                    border:`1px solid ${danger}`, color:danger, marginRight:"auto" }}>Supprimer</button>
              )}
              <button onClick={cancelEditWagon}
                style={{ ...btnBase, padding: isMobile ? "10px 16px" : "7px 12px", background:"transparent",
                  border:`1px solid ${border}`, color:muted }}>Annuler</button>
              <button onClick={addWagon}
                disabled={!wagonSerial.trim() || wagonCoils.trim() === ""}
                style={{ ...btnBase, padding: isMobile ? "10px 16px" : "7px 12px", background:accent+"22",
                  border:`1px solid ${accent}`, color:accent, fontWeight:700,
                  opacity:(!wagonSerial.trim() || wagonCoils.trim() === "") ? 0.4 : 1 }}>
                {editingWagonId ? "Enregistrer" : "Créer"}
              </button>
            </div>
          </div>
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
                          : isOcc  ? (dark ? "#2a1a00" : "#fde68a")
                          : (dark ? "#1c1c1c" : "#f0f0f0"),
                        color: isSel ? accent
                          : isOcc  ? "#92400e"
                          : muted,
                        fontWeight: isSel ? 700 : 400,
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
                placeholder="ex : 11.258"
                style={{ width:"100%", fontFamily:MONO, fontSize: isMobile ? 16 : 13,
                  padding: isMobile ? "12px 10px" : "8px 10px", borderRadius:5,
                  background:bg, border:`1px solid ${pending.weight.trim()==="" ? "#ef4444" : border}`, color:text, outline:"none",
                  boxSizing:"border-box" as const }} />
            </div>
            {/* Déchargement */}
            <div style={{ marginBottom:14 }}>
              <div style={{ fontSize: isMobile ? 13 : 10, color:muted, marginBottom:5 }}>Déchargement <span style={{fontWeight:400}}>(optionnel)</span></div>
              <input value={pending.dechargement}
                onChange={(e) => setPending({ ...pending, dechargement: e.target.value })}
                list="pending-dechargement-list"
                onKeyDown={(e) => { if (e.key==="Enter") confirmPending(); if (e.key==="Escape") setPending(null); }}
                placeholder={existingDechargements.length > 0 ? "Choisir ou saisir…" : "ex : PORT-SUD"}
                style={{ width:"100%", fontFamily:MONO, fontSize: isMobile ? 15 : 12,
                  padding: isMobile ? "10px 10px" : "7px 10px", borderRadius:5,
                  background:bg, border:`1px solid ${pending.dechargement.trim() ? accent : border}`, color:text, outline:"none",
                  boxSizing:"border-box" as const }} />
              <datalist id="pending-dechargement-list">
                {existingDechargements.map((d) => <option key={d} value={d} />)}
              </datalist>
              {existingDechargements.length > 0 && (
                <div style={{ display:"flex", flexWrap:"wrap", gap:4, marginTop:6 }}>
                  {existingDechargements.map((d) => (
                    <button key={d} onClick={() => setPending({ ...pending, dechargement: d })}
                      style={{ fontFamily:MONO, fontSize: isMobile ? 12 : 9, padding:"2px 8px",
                        borderRadius:10, cursor:"pointer", border:`1px solid ${pending.dechargement===d?accent:border}`,
                        background: pending.dechargement===d ? accent+"22" : "transparent",
                        color: pending.dechargement===d ? accent : muted }}>{d}</button>
                  ))}
                </div>
              )}
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
    </div>
  );
}
