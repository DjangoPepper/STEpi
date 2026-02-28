import { useEffect, useRef, useState, useCallback } from "react";
import { BrowserMultiFormatReader } from "@zxing/browser";
import type { CellValue } from "./types";
import { useWindowWidth } from "./useWindowWidth";

/* ── BarcodeDetector type declaration (not yet in TS lib) ─────────── */
interface BarcodeDetectorResult { rawValue: string; format: string; }
interface BarcodeDetectorConstructor {
  new (opts?: { formats: string[] }): { detect(src: HTMLVideoElement | ImageBitmap): Promise<BarcodeDetectorResult[]> };
  getSupportedFormats(): Promise<string[]>;
}
declare const BarcodeDetector: BarcodeDetectorConstructor;

/* ── LS helper ────────────────────────────────────────────────────── */
const LS = {
  get<T>(key: string, fallback: T): T {
    try { const v = localStorage.getItem(key); return v ? (JSON.parse(v) as T) : fallback; } catch { return fallback; }
  },
  set<T>(key: string, val: T) {
    try { localStorage.setItem(key, JSON.stringify(val)); } catch {}
  },
};

/* ── Types ────────────────────────────────────────────────────────── */
interface DestConfig { id: string; name: string; color: string; }
interface ScanResult { rawValue: string; matched: boolean; rowIndices: number[]; destName: string; colHeader: string; }
interface ScaneProps { dark: boolean; }

const MONO = "'IBM Plex Mono', 'Fira Mono', monospace";

export default function Scane({ dark }: ScaneProps) {
  const bg      = dark ? "#0d0d0d" : "#f5f5f5";
  const surface = dark ? "#141414" : "#ffffff";
  const text    = dark ? "#e8e8e0" : "#1a1a1a";
  const accent  = dark ? "#6ee7b7" : "#059669";
  const muted   = dark ? "#555"    : "#888";
  const border  = dark ? "#2a2a2a" : "#d0d0d0";

  const vw       = useWindowWidth();
  const isMobile = vw < 640;
  const pad      = isMobile ? "14px 10px" : "24px 28px";

  /* ── Pointage data ───────────────────────────────────────────── */
  const [headers,       setHeaders]       = useState<string[]>([]);
  const [rows,          setRows]          = useState<CellValue[][]>([]);
  const [destinations,  setDestinations]  = useState<DestConfig[]>([]);
  const [selectedDestId,setSelectedDestId]= useState<string | null>(null);
  const [rowDestMap,    setRowDestMap]    = useState<Record<number, string>>({});
  const [searchColIdx,  setSearchColIdx]  = useState<number>(-1);

  /* ── Camera / detector ───────────────────────────────────────── */
  const videoRef      = useRef<HTMLVideoElement>(null);
  const streamRef     = useRef<MediaStream | null>(null);
  const detectorRef   = useRef<{ detect(src: HTMLVideoElement): Promise<BarcodeDetectorResult[]> } | null>(null);
  const rafRef        = useRef<number | null>(null);
  const zxingReader   = useRef<BrowserMultiFormatReader | null>(null);
  const zxingControls = useRef<{ stop(): void } | null>(null);

  const [cameraOn,    setCameraOn]    = useState(false);
  const [cameraError, setCameraError] = useState<string | null>(null);
  const [useNative,   setUseNative]   = useState<boolean>(false);

  /* ── Scan feedback ───────────────────────────────────────────── */
  const [lastScan,   setLastScan]   = useState<ScanResult | null>(null);
  const [flashColor, setFlashColor] = useState<string | null>(null);
  const [scanCount,  setScanCount]  = useState(0);

  /* ── Check support ───────────────────────────────────────────── */
  useEffect(() => { setUseNative(typeof BarcodeDetector !== "undefined"); }, []);

  /* ── Reload pointage data ────────────────────────────────────── */
  const reloadData = useCallback(() => {
    setHeaders      (LS.get<string[]>              ("ptg_headers",       []));
    setRows         (LS.get<CellValue[][]>         ("ptg_rows",          []));
    setDestinations (LS.get<DestConfig[]>          ("ptg_destinations",  []));
    setSelectedDestId(LS.get<string|null>          ("ptg_selectedDestId",null));
    setRowDestMap   (LS.get<Record<number,string>> ("ptg_rowDestMap",    {}));
  }, []);

  useEffect(() => { reloadData(); }, [reloadData]);

  /* ── Start camera ────────────────────────────────────────────── */
  const startCamera = useCallback(async () => {
    setCameraError(null);
    const native = typeof BarcodeDetector !== "undefined";
    setUseNative(native);
    try {
      if (native) {
        /* ── Native BarcodeDetector path ── */
        const stream = await navigator.mediaDevices.getUserMedia({
          video: { facingMode: { ideal: "environment" }, width: { ideal: 1280 }, height: { ideal: 720 } },
        });
        streamRef.current = stream;
        if (videoRef.current) { videoRef.current.srcObject = stream; await videoRef.current.play(); }
        const formats = await BarcodeDetector.getSupportedFormats();
        detectorRef.current = new BarcodeDetector({ formats });
        setCameraOn(true);
      } else {
        /* ── ZXing fallback path ── */
        if (!videoRef.current) return;
        const reader = new BrowserMultiFormatReader();
        zxingReader.current = reader;
        const lastTimes: Record<string, number> = {};
        const DEBOUNCE = 1800;
        const controls = await reader.decodeFromConstraints(
          { video: { facingMode: { ideal: "environment" }, width: { ideal: 1280 }, height: { ideal: 720 } } },
          videoRef.current,
          (result, _err) => {
            if (!result) return;
            const raw = result.getText();
            const now = Date.now();
            if ((now - (lastTimes[raw] ?? 0)) > DEBOUNCE) {
              lastTimes[raw] = now;
              assignMatchRef.current(raw);
            }
          }
        );
        zxingControls.current = controls;
        setCameraOn(true);
      }
    } catch (e) {
      setCameraError(e instanceof Error ? e.message : "Erreur caméra inconnue");
    }
  }, []);

  /* ── Stop camera ─────────────────────────────────────────────── */
  const stopCamera = useCallback(() => {
    /* native */
    if (rafRef.current) cancelAnimationFrame(rafRef.current);
    rafRef.current = null;
    streamRef.current?.getTracks().forEach((t) => t.stop());
    streamRef.current = null;
    if (videoRef.current) videoRef.current.srcObject = null;
    /* zxing */
    zxingControls.current?.stop();
    zxingControls.current = null;
    zxingReader.current = null;
    setCameraOn(false);
  }, []);

  /* ── Match & assign ──────────────────────────────────────────── */  /* Stable ref so ZXing callback always calls latest version without stale closure */
  const assignMatchRef = useRef<(raw: string) => void>(() => {});  const assignMatch = useCallback((raw: string) => {
    const fHeaders  = LS.get<string[]>              ("ptg_headers",       []);
    const fRows     = LS.get<CellValue[][]>         ("ptg_rows",          []);
    const fDests    = LS.get<DestConfig[]>          ("ptg_destinations",  []);
    const fSelId    = LS.get<string|null>           ("ptg_selectedDestId",null);
    const fMap      = LS.get<Record<number,string>> ("ptg_rowDestMap",    {});

    setHeaders(fHeaders); setRows(fRows); setDestinations(fDests);
    setSelectedDestId(fSelId); setRowDestMap(fMap);

    if (!fSelId) {
      setFlashColor("#b45309");
      setTimeout(() => setFlashColor(null), 800);
      setLastScan({ rawValue: raw, matched: false, rowIndices: [], destName: "—", colHeader: "Aucune destination sélectionnée" });
      return;
    }
    const dest = fDests.find((d) => d.id === fSelId);
    if (!dest) return;

    const cols: number[] = searchColIdx === -1 ? fHeaders.map((_, i) => i) : [searchColIdx];
    const matched: number[] = [];
    let matchedCol = "";

    fRows.forEach((row, ri) => {
      for (const ci of cols) {
        if (String(row[ci] ?? "").trim() === raw.trim()) {
          matched.push(ri);
          if (!matchedCol) matchedCol = fHeaders[ci] ?? `Col ${ci+1}`;
          break;
        }
      }
    });

    if (matched.length > 0) {
      const newMap = { ...fMap };
      matched.forEach((ri) => { newMap[ri] = fSelId; });
      LS.set("ptg_rowDestMap", newMap);
      setRowDestMap(newMap);
      setFlashColor(dest.color);
      setScanCount((n) => n + matched.length);
      setLastScan({ rawValue: raw, matched: true, rowIndices: matched, destName: dest.name, colHeader: matchedCol });
    } else {
      setFlashColor("#b45309");
      const colLabel = searchColIdx === -1 ? "toutes colonnes" : (fHeaders[searchColIdx] ?? `Col ${searchColIdx+1}`);
      setLastScan({ rawValue: raw, matched: false, rowIndices: [], destName: dest.name, colHeader: colLabel });
    }
    setTimeout(() => setFlashColor(null), 800);
  }, [searchColIdx]);

  /* Keep ref in sync so ZXing callback is always fresh */
  useEffect(() => { assignMatchRef.current = assignMatch; }, [assignMatch]);

  /* ── Native scan loop (BarcodeDetector only) ─────────────────────── */
  useEffect(() => {
    if (!cameraOn || !useNative || !detectorRef.current) return;
    let alive = true;
    const lastTimes: Record<string, number> = {};
    const DEBOUNCE = 1800;

    const tick = async () => {
      if (!alive) return;
      if (videoRef.current && videoRef.current.readyState >= 2 && detectorRef.current) {
        try {
          const results = await detectorRef.current.detect(videoRef.current);
          for (const r of results) {
            const now = Date.now();
            if ((now - (lastTimes[r.rawValue] ?? 0)) > DEBOUNCE) {
              lastTimes[r.rawValue] = now;
              assignMatch(r.rawValue);
            }
          }
        } catch { /* ignore frame errors */ }
      }
      rafRef.current = requestAnimationFrame(tick);
    };
    rafRef.current = requestAnimationFrame(tick);
    return () => { alive = false; if (rafRef.current) cancelAnimationFrame(rafRef.current); };
  }, [cameraOn, useNative, assignMatch]);

  useEffect(() => () => { stopCamera(); }, [stopCamera]);

  const selectedDest = destinations.find((d) => d.id === selectedDestId) ?? null;

  return (
    <div style={{ padding: pad, fontFamily: MONO, color: text, background: bg, minHeight: "calc(100vh - 44px)", boxSizing: "border-box", overflowX: "hidden" }}>

      {/* Header */}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 20 }}>
        <h2 style={{ fontSize: 13, letterSpacing: "0.18em", textTransform: "uppercase", color: accent, margin: 0 }}>
          ⬛ Scan Code-barres / QR
        </h2>
        <button onClick={reloadData}
          style={{ fontFamily: MONO, fontSize: 10, letterSpacing: "0.1em", padding: "4px 10px",
            background: "transparent", border: `1px solid ${border}`, borderRadius: 3, color: muted, cursor: "pointer" }}>
          ↺ Actualiser données
        </button>
      </div>

      {/* Engine badge */}
      {cameraOn && (
        <div style={{ marginBottom: 10, display: "inline-block", padding: "3px 10px",
          background: useNative ? (dark?"#0a200f":"#dcfce7") : (dark?"#0a0f20":"#e0e7ff"),
          border: `1px solid ${useNative ? accent : (dark?"#818cf8":"#4f46e5")}`,
          borderRadius: 4, fontSize: 10, color: useNative ? accent : (dark?"#818cf8":"#4f46e5") }}>
          {useNative ? "⚡ BarcodeDetector (natif)" : "⚙ ZXing (compatible tous navigateurs)"}
        </div>
      )}

      {/* Status chips */}
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 18 }}>
        <div style={{ padding: "5px 14px", borderRadius: 4, fontSize: 11, fontWeight: 700,
          background: selectedDest ? selectedDest.color + "22" : (dark?"#1a1a1a":"#f0f0f0"),
          border: `1px solid ${selectedDest ? selectedDest.color : border}`,
          color: selectedDest ? selectedDest.color : muted }}>
          {selectedDest ? `▶ ${selectedDest.name}` : "— Aucune destination —"}
        </div>
        <div style={{ padding: "5px 14px", borderRadius: 4, fontSize: 11,
          background: dark?"#1a1a1a":"#f0f0f0", border: `1px solid ${border}`, color: accent }}>
          {scanCount} affectation{scanCount !== 1 ? "s" : ""}
        </div>
        <div style={{ padding: "5px 14px", borderRadius: 4, fontSize: 11,
          background: dark?"#1a1a1a":"#f0f0f0", border: `1px solid ${border}`, color: muted }}>
          {rows.length} ligne{rows.length !== 1 ? "s" : ""} Excel
        </div>
      </div>

      {/* Column selector */}
      <div style={{ marginBottom: 20, display: "flex", alignItems: "center", gap: 10 }}>
        <span style={{ fontSize: 11, color: muted, textTransform: "uppercase", letterSpacing: "0.1em" }}>Colonne :</span>
        <select value={searchColIdx} onChange={(e) => setSearchColIdx(Number(e.target.value))}
          style={{ fontFamily: MONO, fontSize: 11, padding: "4px 8px", borderRadius: 4,
            background: surface, border: `1px solid ${border}`, color: text, cursor: "pointer" }}>
          <option value={-1}>Toutes les colonnes</option>
          {headers.map((h, i) => <option key={i} value={i}>{h || `Col ${i+1}`}</option>)}
        </select>
        {rows.length === 0 && (
          <span style={{ fontSize: 11, color: muted }}>— aucun fichier chargé dans Pointage</span>
        )}
      </div>

      {/* Camera + Controls */}
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap", alignItems: "flex-start", flexDirection: isMobile ? "column" : "row" }}>

        {/* Video */}
        <div style={{ position: "relative", borderRadius: 8, overflow: "hidden",
          border: `2px solid ${flashColor ?? border}`,
          boxShadow: flashColor ? `0 0 20px ${flashColor}55` : "none",
          transition: "border-color 0.15s, box-shadow 0.15s",
          width: isMobile ? "100%" : 480, aspectRatio: isMobile ? "4/3" : undefined,
          minHeight: isMobile ? undefined : 270,
          background: dark ? "#0a0a0a" : "#ddd",
          display: "flex", alignItems: "center", justifyContent: "center" }}>
          <video ref={videoRef} playsInline muted
            style={{ width: "100%", height: "100%", display: cameraOn ? "block" : "none", objectFit: "cover" }} />
          {!cameraOn && (
            <div style={{ textAlign: "center", color: muted, fontSize: 12, padding: 20 }}>
              <div style={{ fontSize: 40, marginBottom: 8 }}>📷</div>
              <div>Caméra arrêtée</div>
            </div>
          )}
          {/* Corner viewfinder */}
          {cameraOn && (
            <div style={{ position: "absolute", inset: 0, pointerEvents: "none" }}>
              <div style={{ position: "absolute", top: "50%", left: "50%",
                transform: "translate(-50%,-50%)", width: 200, height: 200 }}>
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

        {/* Right panel */}
        <div style={{ flex: 1, minWidth: isMobile ? "unset" : 240, width: isMobile ? "100%" : undefined, display: "flex", flexDirection: "column", gap: 12 }}>
          <button onClick={cameraOn ? stopCamera : startCamera}
            style={{ fontFamily: MONO, fontSize: 12, letterSpacing: "0.1em", padding: "10px 18px",
              background: cameraOn ? (dark?"#2d0a0a":"#fee2e2") : (dark?"#0a200f":"#dcfce7"),
              border: `1px solid ${cameraOn ? "#ef4444" : accent}`,
              borderRadius: 5, color: cameraOn ? "#ef4444" : accent,
              cursor: "pointer", fontWeight: 700 }}>
            {cameraOn ? "⏹ Arrêter la caméra" : "▶ Démarrer la caméra"}
          </button>

          {cameraError && (
            <div style={{ fontSize: 11, color: "#ef4444", padding: "8px 12px",
              background: dark?"#2d0a0a":"#fee2e2", borderRadius: 4, border: "1px solid #ef4444" }}>
              ⚠ {cameraError}
            </div>
          )}

          {/* Scan result */}
          {lastScan && (
            <div style={{ padding: "12px 14px", borderRadius: 6,
              background: lastScan.matched ? (dark?"#0a1f0f":"#f0fdf4") : (dark?"#1c1000":"#fffbeb"),
              border: `1px solid ${lastScan.matched ? accent : "#d97706"}` }}>
              <div style={{ fontSize: 10, color: muted, textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: 5 }}>
                Dernier scan
              </div>
              <div style={{ fontSize: 12, fontWeight: 700, color: text, wordBreak: "break-all", marginBottom: 6 }}>
                {lastScan.rawValue}
              </div>
              {lastScan.matched ? (
                <>
                  <div style={{ fontSize: 11, color: accent }}>
                    ✓ {lastScan.rowIndices.length} correspondance{lastScan.rowIndices.length > 1 ? "s" : ""}
                    {" · "}{lastScan.colHeader}
                  </div>
                  <div style={{ fontSize: 11, color: muted }}>
                    Ligne{lastScan.rowIndices.length > 1 ? "s" : ""} : {lastScan.rowIndices.map((i)=>i+1).join(", ")}
                  </div>
                  <div style={{ marginTop: 8, display: "inline-block", padding: "2px 10px", borderRadius: 12,
                    background: (selectedDest?.color ?? accent) + "33",
                    border: `1px solid ${selectedDest?.color ?? accent}`,
                    color: selectedDest?.color ?? accent, fontSize: 11, fontWeight: 700 }}>
                    → {lastScan.destName}
                  </div>
                </>
              ) : (
                <div style={{ fontSize: 11, color: "#d97706" }}>
                  {lastScan.colHeader === "Aucune destination sélectionnée"
                    ? "⚠ Sélectionner une destination dans l'onglet Pointage"
                    : `✗ Aucune correspondance · ${lastScan.colHeader}`}
                </div>
              )}
            </div>
          )}

          {/* How-to */}
          <div style={{ fontSize: 10, color: muted, lineHeight: 1.8, borderTop: `1px solid ${border}`, paddingTop: 10 }}>
            <div>1. Sélectionner une destination dans <strong style={{color:text}}>Pointage</strong></div>
            <div>2. Cliquer <strong style={{color:text}}>Démarrer la caméra</strong></div>
            <div>3. Pointer vers un code-barres ou QR code</div>
            <div>4. L'affectation est enregistrée automatiquement</div>
          </div>
        </div>
      </div>

      {/* Affectations actives */}
      {Object.keys(rowDestMap).length > 0 && (
        <div style={{ marginTop: 28 }}>
          <div style={{ fontSize: 10, color: muted, letterSpacing: "0.12em", textTransform: "uppercase", marginBottom: 8 }}>
            Affectations actives ({Object.keys(rowDestMap).length})
          </div>
          <div style={{ maxHeight: 140, overflowY: "auto", display: "flex", flexWrap: "wrap", gap: 6 }}>
            {Object.entries(rowDestMap).slice(0, 100).map(([ri, destId]) => {
              const d = destinations.find((x) => x.id === destId);
              const row = rows[Number(ri)];
              const label = row ? String(row[0] ?? `L${Number(ri)+1}`) : `L${Number(ri)+1}`;
              return (
                <div key={ri} style={{ padding: "2px 8px", borderRadius: 10, fontSize: 10,
                  background: (d?.color ?? muted) + "22", border: `1px solid ${d?.color ?? border}`,
                  color: d?.color ?? muted }}>
                  {label}
                </div>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}
