import { useState, useRef, useEffect, type CSSProperties } from "react";
import * as XLSX from "xlsx";
import type { PointageData } from "./types";
import type { CellValue } from "./types";

/* ─── Types ─────────────────────────────────────────────── */
type SortDir = "none" | "asc" | "desc";

interface DestConfig {
  id: string;
  name: string;
  color: string;
}

interface HistoryEntry {
  rowIndex: number;
  row: CellValue[];
  destId: string | null;
}

/* ─── Constants ─────────────────────────────────────────── */
const MONO = "'IBM Plex Mono', 'Fira Mono', monospace";

const DEST_PALETTE = [
  "#FF6200", "#0091D5", "#00A550", "#9B2FAD",
  "#FA008C", "#00B4AE", "#E4A600", "#D4006E",
  "#006AB3", "#007857", "#FF3A20", "#6A1FAD",
];

/* ─── localStorage helpers ──────────────────────────────── */
const LS = {
  get<T>(key: string, fallback: T): T {
    try {
      const v = localStorage.getItem(key);
      return v !== null ? (JSON.parse(v) as T) : fallback;
    } catch { return fallback; }
  },
  set<T>(key: string, val: T) {
    try { localStorage.setItem(key, JSON.stringify(val)); } catch { /* ignore */ }
  },
};

/* ─── Utilities ─────────────────────────────────────────── */
function cellStr(v: CellValue): string {
  return v == null ? "" : String(v);
}

function applyFormat(value: string, pattern: string): string {
  const chunks = pattern.trim().split(/\s+/).map(Number).filter((n) => n > 0 && isFinite(n));
  if (!chunks.length) return value;
  let pos = 0;
  const parts: string[] = [];
  for (const size of chunks) {
    if (pos >= value.length) break;
    parts.push(value.slice(pos, pos + size));
    pos += size;
  }
  if (pos < value.length) parts.push(value.slice(pos));
  return parts.join(" ");
}

function detectHeaderRow(raw: CellValue[][]): { headers: string[]; rows: CellValue[][] } {
  if (raw.length === 0) return { headers: [], rows: [] };
  let headerIdx = 0;
  for (let i = 0; i < Math.min(5, raw.length); i++) {
    const row = raw[i];
    const filled = row.filter((c) => c !== null && c !== undefined && c !== "").length;
    const strings = row.filter((c) => typeof c === "string" && c.trim() !== "").length;
    if (filled > 0 && strings / filled > 0.6) { headerIdx = i; break; }
  }
  const headers = raw[headerIdx].map((h, i) =>
    h !== null && h !== undefined && h !== "" ? String(h) : `Col ${i + 1}`
  );
  return { headers, rows: raw.slice(headerIdx + 1) };
}

/* ─── Props ─────────────────────────────────────────────── */
interface PointageProps {
  dark: boolean;
  initialData: PointageData | null;
}

/* ═══════════════════ COMPONENT ════════════════════════════ */
export default function Pointage({ dark, initialData }: PointageProps) {
  /* Theme */
  const bg      = dark ? "#0d0d0d" : "#f5f5f5";
  const surface = dark ? "#141414" : "#fff";
  const text    = dark ? "#e8e8e0" : "#1a1a1a";
  const accent  = dark ? "#6ee7b7" : "#059669";
  const muted   = dark ? "#555"    : "#888";
  const border  = dark ? "#222"    : "#ddd";
  const hdrBg   = dark ? "#0b0b0b" : "#f0f0f0";
  const rowAlt  = dark ? "#111"    : "#f9f9f9";

  /* File state */
  const [headers, setHeaders] = useState<string[]>(() => LS.get<string[]>("ptg_headers", []));
  const [rows, setRows]       = useState<CellValue[][]>(() => LS.get<CellValue[][]>("ptg_rows", []));
  const [fileName, setFileName] = useState<string | null>(() => LS.get<string | null>("ptg_fileName", null));
  const fileRef = useRef<HTMLInputElement>(null);

  /* Sort */
  const [sortCol, setSortCol] = useState<number | null>(() => LS.get<number | null>("ptg_sortCol", null));
  const [sortDir, setSortDir] = useState<SortDir>(() => LS.get<SortDir>("ptg_sortDir", "none"));

  /* Format patterns */
  const [fmtPattern, setFmtPattern] = useState<Record<number, string>>(() => LS.get<Record<number, string>>("ptg_fmtPattern", {}));
  const [fmtOpen, setFmtOpen]       = useState<Record<number, boolean>>({});

  /* Destinations */
  const [destinations, setDestinations]     = useState<DestConfig[]>(() => LS.get<DestConfig[]>("ptg_destinations", []));
  const [selectedDestId, setSelectedDestId] = useState<string | null>(() => LS.get<string | null>("ptg_selectedDestId", null));
  const [rowDestMap, setRowDestMap]         = useState<Record<number, string>>(() => LS.get<Record<number, string>>("ptg_rowDestMap", {}));
  const [forcedDestIds, setForcedDestIds]   = useState<string[]>(() => LS.get<string[]>("ptg_forceddests", []));
  const [newDestName, setNewDestName]       = useState("");
  const [editingDestId, setEditingDestId]   = useState<string | null>(null);
  const [editingDestName, setEditingDestName] = useState("");

  /* Search */
  const [searchQuery, setSearchQuery] = useState("");

  /* History */
  const [history, setHistory] = useState<HistoryEntry[]>(() => LS.get<HistoryEntry[]>("ptg_history", []));

  /* Panel open states — independent */
  const [searchOpen, setSearchOpen] = useState(false);
  const [destOpen,   setDestOpen]   = useState(false);
  const [histOpen,   setHistOpen]   = useState(false);

  /* ─── Persist ───────────────────────────────────────────── */
  useEffect(() => { LS.set("ptg_headers",      headers);       }, [headers]);
  useEffect(() => { LS.set("ptg_rows",         rows);          }, [rows]);
  useEffect(() => { LS.set("ptg_fileName",     fileName);      }, [fileName]);
  useEffect(() => { LS.set("ptg_sortCol",      sortCol);       }, [sortCol]);
  useEffect(() => { LS.set("ptg_sortDir",      sortDir);       }, [sortDir]);
  useEffect(() => { LS.set("ptg_fmtPattern",   fmtPattern);    }, [fmtPattern]);
  useEffect(() => { LS.set("ptg_destinations", destinations);  }, [destinations]);
  useEffect(() => { LS.set("ptg_selectedDestId", selectedDestId); }, [selectedDestId]);
  useEffect(() => { LS.set("ptg_rowDestMap",   rowDestMap);    }, [rowDestMap]);
  useEffect(() => { LS.set("ptg_forceddests",  forcedDestIds); }, [forcedDestIds]);
  useEffect(() => { LS.set("ptg_history",      history);       }, [history]);

  /* ─── Load from initialData (bridge from ExcelCleaner) ─── */
  useEffect(() => {
    if (!initialData) return;
    setHeaders(initialData.headers);
    setRows(initialData.rows);
    setFileName(initialData.fileName);
    setRowDestMap({});
    setHistory([]);
    setSortCol(null);
    setSortDir("none");
  }, [initialData]);

  /* ─── File loading ──────────────────────────────────────── */
  const loadFile = (file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target!.result as ArrayBuffer);
      const wb   = XLSX.read(data, { type: "array" });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const raw  = XLSX.utils.sheet_to_json<CellValue[]>(ws, { header: 1, defval: null }) as CellValue[][];
      const { headers: h, rows: r } = detectHeaderRow(raw);
      setHeaders(h); setRows(r); setFileName(file.name);
      setRowDestMap({}); setHistory([]);
      setSortCol(null); setSortDir("none");
    };
    reader.readAsArrayBuffer(file);
  };

  /* ─── Sorted + filtered indices ─────────────────────────── */
  const sortedIndices = (() => {
    const idx = rows.map((_, i) => i);
    if (sortCol === null || sortDir === "none") return idx;
    const dir = sortDir === "asc" ? 1 : -1;
    if (sortCol === -1) {
      return [...idx].sort((a, b) => {
        const da = destinations.find(d => d.id === rowDestMap[a])?.name ?? "";
        const db = destinations.find(d => d.id === rowDestMap[b])?.name ?? "";
        return da.localeCompare(db) * dir;
      });
    }
    return [...idx].sort((a, b) => {
      const va = rows[a][sortCol], vb = rows[b][sortCol];
      if (va == null && vb == null) return 0;
      if (va == null) return dir; if (vb == null) return -dir;
      if (typeof va === "number" && typeof vb === "number") return (va - vb) * dir;
      return String(va).localeCompare(String(vb)) * dir;
    });
  })();

  const filteredIndices = (() => {
    const q = searchQuery.trim().toLowerCase();
    if (!q) return sortedIndices;
    return sortedIndices.filter((ri) =>
      rows[ri].some((cell) => cellStr(cell).toLowerCase().includes(q))
    );
  })();

  const cycleSortCol = (ci: number) => {
    if (sortCol !== ci) { setSortCol(ci); setSortDir("asc"); return; }
    if (sortDir === "asc")  { setSortDir("desc"); return; }
    setSortCol(null); setSortDir("none");
  };

  const sortIcon = (ci: number) => {
    if (sortCol !== ci || sortDir === "none")
      return <span style={{ opacity: 0.3, fontSize: 11, lineHeight: 1 }}>−</span>;
    return sortDir === "asc"
      ? <span style={{ color: accent, fontSize: 11, lineHeight: 1 }}>↓</span>
      : <span style={{ color: accent, fontSize: 11, lineHeight: 1 }}>↑</span>;
  };

  /* ─── Destinations ──────────────────────────────────────── */
  const addDestination = () => {
    const name = newDestName.trim(); if (!name) return;
    const color = DEST_PALETTE[destinations.length % DEST_PALETTE.length];
    const dest: DestConfig = { id: `d_${Date.now()}`, name, color };
    setDestinations((p) => [...p, dest]);
    setNewDestName("");
    if (!selectedDestId) setSelectedDestId(dest.id);
  };
  const commitRename = (id: string) => {
    const name = editingDestName.trim();
    if (name) setDestinations((p) => p.map((d) => d.id === id ? { ...d, name } : d));
    setEditingDestId(null);
  };
  const removeDest = (id: string) => {
    setDestinations((p) => p.filter((d) => d.id !== id));
    if (selectedDestId === id) setSelectedDestId(null);
    setRowDestMap((p) => { const n = { ...p }; Object.keys(n).forEach((k) => { if (n[+k] === id) delete n[+k]; }); return n; });
    setForcedDestIds((p) => p.filter(i => i !== id));
  };
  const moveDest = (id: string, dir: -1 | 1) => {
    setDestinations(prev => {
      const idx = prev.findIndex(d => d.id === id);
      if (idx < 0) return prev;
      const next = idx + dir;
      if (next < 0 || next >= prev.length) return prev;
      const arr = [...prev];
      [arr[idx], arr[next]] = [arr[next], arr[idx]];
      return arr;
    });
  };
  const toggleForced = (id: string) => {
    setForcedDestIds(p => p.includes(id) ? p.filter(i => i !== id) : [...p, id]);
  };
  const getDestById = (id?: string | null) => destinations.find((d) => d.id === id);

  /* ─── Row click ─────────────────────────────────────────── */
  const handleRowClick = (rowIndex: number) => {
    if (selectedDestId) setRowDestMap((p) => ({ ...p, [rowIndex]: selectedDestId }));
    setHistory((p) => [{ rowIndex, row: rows[rowIndex], destId: selectedDestId }, ...p].slice(0, 10));
  };

  /* ─── Format popover ────────────────────────────────────── */
  const closeAllFmt = (except?: number) =>
    setFmtOpen((p) => { const n: Record<number, boolean> = {}; if (except !== undefined) n[except] = !p[except]; return n; });

  /* ─── Style helpers ─────────────────────────────────────── */
  const inpStyle = (extra?: CSSProperties): CSSProperties => ({
    fontFamily: MONO, fontSize: 11, padding: "6px 10px",
    background: dark ? "#0d0d0d" : "#f8f8f8",
    border: `1px solid ${border}`, borderRadius: 3,
    color: text, outline: "none", ...extra,
  });
  const btnStyle = (extra?: CSSProperties): CSSProperties => ({
    fontFamily: MONO, fontSize: 10, letterSpacing: "0.1em",
    textTransform: "uppercase", cursor: "pointer", border: "none",
    borderRadius: 3, padding: "6px 14px", ...extra,
  });

  /* ════════════════════ RENDER ══════════════════════════════ */
  return (
    <div style={{ background: bg, minHeight: "calc(100vh - 44px)", color: text, fontFamily: MONO }}
      onClick={() => closeAllFmt()}>

      {/* ── Upload zone ─────────────────────────────────────── */}
      {rows.length === 0 && (
        <div style={{ maxWidth: 900, margin: "0 auto", padding: "40px 24px" }}>
          <div
            onClick={(e) => { e.stopPropagation(); fileRef.current?.click(); }}
            onDragOver={(e) => e.preventDefault()}
            onDrop={(e) => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) loadFile(f); }}
            style={{
              border: `1.5px dashed ${border}`, borderRadius: 4,
              padding: "60px 40px", textAlign: "center", cursor: "pointer",
              background: surface, transition: "border-color 0.2s, background 0.2s",
            }}
            onMouseEnter={(e) => {
              (e.currentTarget as HTMLElement).style.borderColor = accent;
              (e.currentTarget as HTMLElement).style.background = dark ? "#0f2a20" : "#f0fdf4";
            }}
            onMouseLeave={(e) => {
              (e.currentTarget as HTMLElement).style.borderColor = border;
              (e.currentTarget as HTMLElement).style.background = surface;
            }}
          >
            <div style={{ fontSize: 32, marginBottom: 16, opacity: 0.2, color: text }}>⊞</div>
            <div style={{ fontSize: 13, color: muted, marginBottom: 8 }}>
              Glissez un fichier Excel ou cliquez pour sélectionner
            </div>
            <div style={{ fontSize: 10, color: dark ? "#3a3a3a" : "#ccc", letterSpacing: "0.1em", textTransform: "uppercase" }}>
              .xlsx · .xls · .csv
            </div>
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }}
              onChange={(e) => { const f = e.target.files?.[0]; if (f) loadFile(f); }} />
          </div>
        </div>
      )}

      {rows.length > 0 && (
        <div style={{ padding: "20px 24px" }}>

          {/* File bar */}
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
            <span style={{ fontSize: 11, color: muted }}>
              {fileName}&nbsp;—&nbsp;{rows.length} lignes&nbsp;·&nbsp;{headers.length} colonnes
            </span>
            <button
              onClick={() => {
                setRows([]); setHeaders([]); setFileName(null);
                setRowDestMap({}); setHistory([]);
                setSortCol(null); setSortDir("none");
              }}
              style={{ ...btnStyle(), background: "transparent", border: `1px solid ${border}`, color: muted }}
            >
              × Fermer
            </button>
          </div>

          {/* ══ PANEL TOGGLE TOOLBAR ══ */}
          <div style={{
            display: "flex", gap: 6, marginBottom: 0,
            borderBottom: `1px solid ${border}`,
            background: surface,
            borderRadius: "4px 4px 0 0",
            padding: "6px 10px",
          }}>
            {/* Recherche button */}
            <button
              onClick={() => setSearchOpen(v => !v)}
              style={{
                fontFamily: MONO, fontSize: 10, letterSpacing: "0.12em",
                textTransform: "uppercase", padding: "5px 14px",
                background: searchOpen ? (dark ? "#0f2a20" : "#e8f5ee") : "transparent",
                border: `1px solid ${searchOpen ? accent : border}`,
                borderRadius: 3, color: searchOpen ? accent : muted,
                cursor: "pointer", display: "flex", alignItems: "center", gap: 6,
              }}
            >
              ⌕ Recherche
              {searchQuery && (
                <span style={{ background: accent, color: "#000", borderRadius: 2,
                  padding: "0 5px", fontSize: 9, fontWeight: 700 }}>
                  {filteredIndices.length}
                </span>
              )}
              <span style={{ fontSize: 9, opacity: 0.6 }}>{searchOpen ? "▲" : "▼"}</span>
            </button>

            {/* Destinations button */}
            <button
              onClick={() => setDestOpen(v => !v)}
              style={{
                fontFamily: MONO, fontSize: 10, letterSpacing: "0.12em",
                textTransform: "uppercase", padding: "5px 14px",
                background: destOpen ? (dark ? "#0f2a20" : "#e8f5ee") : "transparent",
                border: `1px solid ${destOpen ? accent : border}`,
                borderRadius: 3, color: destOpen ? accent : muted,
                cursor: "pointer", display: "flex", alignItems: "center", gap: 6,
              }}
            >
              Destinations
              {destinations.length > 0 && (
                <span style={{ background: dark ? "#222" : "#eee", color: muted,
                  borderRadius: 2, padding: "0 5px", fontSize: 9 }}>
                  {destinations.length}
                </span>
              )}
              {selectedDestId && getDestById(selectedDestId) && (() => {
                const d = getDestById(selectedDestId)!;
                return (
                  <span style={{ background: d.color, color: "#fff",
                    borderRadius: 2, padding: "0 6px", fontSize: 9, fontWeight: 700 }}>
                    {d.name}
                  </span>
                );
              })()}
              <span style={{ fontSize: 9, opacity: 0.6 }}>{destOpen ? "▲" : "▼"}</span>
            </button>

            {/* Historique button */}
            <button
              onClick={() => setHistOpen(v => !v)}
              style={{
                fontFamily: MONO, fontSize: 10, letterSpacing: "0.12em",
                textTransform: "uppercase", padding: "5px 14px",
                background: histOpen ? (dark ? "#0f2a20" : "#e8f5ee") : "transparent",
                border: `1px solid ${histOpen ? accent : border}`,
                borderRadius: 3, color: histOpen ? accent : muted,
                cursor: "pointer", display: "flex", alignItems: "center", gap: 6,
              }}
            >
              Historique
              {history.length > 0 && (
                <span style={{ background: dark ? "#222" : "#eee", color: muted,
                  borderRadius: 2, padding: "0 5px", fontSize: 9 }}>
                  {history.length}
                </span>
              )}
              <span style={{ fontSize: 9, opacity: 0.6 }}>{histOpen ? "▲" : "▼"}</span>
            </button>
          </div>

          {/* ══ SEARCH PANEL ══ */}
          {searchOpen && (
            <div style={{
              background: surface, border: `1px solid ${border}`,
              borderTop: "none", padding: "12px 16px", marginBottom: 0,
            }}>
              <div style={{
                display: "flex", alignItems: "center",
                border: `1px solid ${searchQuery ? accent : border}`, borderRadius: 4,
                background: dark ? "#0d0d0d" : "#f8f8f8",
              }}>
                <span style={{ padding: "6px 10px", color: muted, fontSize: 13, flexShrink: 0 }}>⌕</span>
                <input
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  placeholder="Recherche dans toutes les colonnes…"
                  style={{ ...inpStyle(), flex: 1, border: "none", background: "transparent", padding: "6px 10px 6px 0" }}
                />
                {searchQuery && (
                  <button onClick={() => setSearchQuery("")}
                    style={{ background: "transparent", border: "none", color: muted,
                      cursor: "pointer", padding: "6px 10px", fontSize: 13 }}>✕</button>
                )}
              </div>
              {searchQuery && (
                <div style={{ fontSize: 10, color: accent, marginTop: 8 }}>
                  {filteredIndices.length}&nbsp;/&nbsp;{rows.length} lignes
                </div>
              )}
            </div>
          )}

          {/* ══ DESTINATIONS PANEL ══ */}
          {destOpen && (
            <div style={{
              background: surface, border: `1px solid ${border}`,
              borderTop: "none", padding: 16, marginBottom: 0,
            }}>
              <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
                <input
                  value={newDestName}
                  onChange={(e) => setNewDestName(e.target.value)}
                  onKeyDown={(e) => e.key === "Enter" && addDestination()}
                  placeholder="Nom de la destination…"
                  style={{ ...inpStyle(), flex: 1 }}
                />
                <button onClick={addDestination}
                  style={{ ...btnStyle(), background: accent, color: "#000", fontWeight: 700 }}>
                  + Ajouter
                </button>
              </div>
              {destinations.length === 0 && (
                <div style={{ fontSize: 11, color: muted }}>Aucune destination — ajoutez-en une.</div>
              )}
              <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                {destinations.map((dest) => {
                  const isSel  = selectedDestId === dest.id;
                  const isEdit = editingDestId === dest.id;
                  return (
                    <div key={dest.id}
                      onClick={(e) => { e.stopPropagation(); if (!isEdit) setSelectedDestId(isSel ? null : dest.id); }}
                      style={{
                        display: "flex", alignItems: "center", gap: 6,
                        padding: "5px 10px", borderRadius: 4,
                        background: dest.color, color: "#fff",
                        fontSize: 11, cursor: "pointer", fontFamily: MONO,
                        outline: isSel ? "2px solid #fff" : "none", outlineOffset: 2,
                        opacity: isSel ? 1 : 0.78, transition: "opacity 0.15s",
                      }}
                    >
                      {isEdit ? (
                        <input autoFocus value={editingDestName}
                          onChange={(e) => setEditingDestName(e.target.value)}
                          onKeyDown={(e) => { if (e.key === "Enter" || e.key === "Escape") commitRename(dest.id); e.stopPropagation(); }}
                          onBlur={() => commitRename(dest.id)}
                          onClick={(e) => e.stopPropagation()}
                          style={{ fontFamily: MONO, fontSize: 11, background: "rgba(255,255,255,0.2)",
                            border: "1px solid rgba(255,255,255,0.5)", borderRadius: 2,
                            color: "#fff", width: 110, outline: "none", padding: "1px 5px" }}
                        />
                      ) : (
                        <>
                          <span onDoubleClick={(e) => { e.stopPropagation(); setEditingDestId(dest.id); setEditingDestName(dest.name); }}>
                            {dest.name}
                          </span>
                          {isSel && <span style={{ fontSize: 9, opacity: 0.85 }}>✓</span>}
                          <span title="Monter" onClick={(e) => { e.stopPropagation(); moveDest(dest.id, -1); }}
                            style={{ fontSize: 11, opacity: 0.6, cursor: "pointer", lineHeight: 1 }}>↑</span>
                          <span title="Descendre" onClick={(e) => { e.stopPropagation(); moveDest(dest.id, 1); }}
                            style={{ fontSize: 11, opacity: 0.6, cursor: "pointer", lineHeight: 1 }}>↓</span>
                          <span
                            title={forcedDestIds.includes(dest.id) ? "Retirer du tally forcé" : "Forcer à 0 dans le tally"}
                            onClick={(e) => { e.stopPropagation(); toggleForced(dest.id); }}
                            style={{
                              fontSize: 9, cursor: "pointer", fontWeight: 700,
                              opacity: forcedDestIds.includes(dest.id) ? 1 : 0.45,
                              background: forcedDestIds.includes(dest.id) ? "rgba(255,255,255,0.25)" : "transparent",
                              borderRadius: 2, padding: "0 3px",
                            }}>0</span>
                          <span title="Supprimer" onClick={(e) => { e.stopPropagation(); removeDest(dest.id); }}
                            style={{ fontSize: 10, opacity: 0.55, cursor: "pointer" }}>✕</span>
                        </>
                      )}
                    </div>
                  );
                })}
              </div>
              {selectedDestId && (
                <div style={{ fontSize: 10, color: muted, marginTop: 10 }}>
                  Cliquez une ligne du tableau pour l'assigner à «&nbsp;
                  <span style={{ color: getDestById(selectedDestId)?.color }}>
                    {getDestById(selectedDestId)?.name}
                  </span>&nbsp;». Double-clic sur un nom pour le renommer.
                </div>
              )}
            </div>
          )}

          {/* ══ HISTORY PANEL ══ */}
          {histOpen && (
            <div style={{
              background: surface, border: `1px solid ${border}`,
              borderTop: "none", padding: 16, marginBottom: 0,
            }}>
              {history.length === 0 && <div style={{ fontSize: 11, color: muted }}>Aucune ligne cliquée.</div>}
              <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                {history.map((entry, i) => {
                  const dest = getDestById(entry.destId);
                  return (
                    <div key={i} style={{
                      display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap",
                      padding: "5px 10px", borderRadius: 4,
                      background: dest ? dest.color + "22" : (dark ? "#1a1a1a" : "#f5f5f5"),
                      border: `1px solid ${dest ? dest.color + "55" : border}`,
                      fontSize: 11, color: text,
                    }}>
                      {dest && (
                        <span style={{ background: dest.color, color: "#fff",
                          padding: "1px 7px", borderRadius: 3, fontSize: 9,
                          letterSpacing: "0.1em", flexShrink: 0 }}>{dest.name}</span>
                      )}
                      <span style={{ color: muted, flexShrink: 0 }}>#{entry.rowIndex + 1}</span>
                      {entry.row.slice(0, 5).map((v, ci) => (
                        <span key={ci} style={{ maxWidth: 140, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                          {cellStr(v)}
                        </span>
                      ))}
                      {entry.row.length > 5 && <span style={{ color: muted }}>…</span>}
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* spacer between panels and table */}
          <div style={{ height: 12 }} />

          {/* ══ TABLE ══ */}
          <div style={{ overflowX: "auto", borderRadius: 6, border: `1px solid ${border}` }}
            onClick={(e) => e.stopPropagation()}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11, fontFamily: MONO }}>
              <thead>
                <tr>
                  {[...headers, "DEST"].map((h, ci) => {
                    const isDestCol = ci === headers.length;
                    return (
                      <th key={ci} style={{
                        background: hdrBg, borderBottom: `2px solid ${border}`,
                        padding: "8px 10px", textAlign: "left",
                        fontSize: 10, letterSpacing: "0.12em", textTransform: "uppercase",
                        color: accent, whiteSpace: "nowrap", position: "relative",
                      }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 5 }}>
                          <span>{h}</span>
                          {isDestCol ? (
                            <button onClick={() => cycleSortCol(-1)} title="Trier par destination"
                              style={{ background: "transparent", border: "none", cursor: "pointer",
                                padding: "0 2px", display: "flex", alignItems: "center" }}>
                              {sortIcon(-1)}
                            </button>
                          ) : (
                            <>
                              <button onClick={() => cycleSortCol(ci)} title="Trier"
                                style={{ background: "transparent", border: "none", cursor: "pointer",
                                  padding: "0 2px", display: "flex", alignItems: "center" }}>
                                {sortIcon(ci)}
                              </button>
                              <button
                                onClick={(e) => { e.stopPropagation(); closeAllFmt(ci); }}
                                title="Format visuel"
                                style={{
                                  background: fmtPattern[ci] ? accent + "33" : "transparent",
                                  border: fmtPattern[ci] ? `1px solid ${accent}66` : "none",
                                  cursor: "pointer",
                                  color: fmtPattern[ci] ? accent : muted + "bb",
                                  padding: "1px 4px", borderRadius: 2, fontSize: 12, lineHeight: 1,
                                }}>⋯</button>
                            </>
                          )}
                        </div>

                        {/* Format popover */}
                        {!isDestCol && fmtOpen[ci] && (
                          <div onClick={(e) => e.stopPropagation()} style={{
                            position: "absolute", top: "calc(100% + 4px)", left: 0, zIndex: 99,
                            background: surface, border: `1px solid ${border}`,
                            borderRadius: 5, padding: 12, minWidth: 210,
                            boxShadow: "0 8px 32px rgba(0,0,0,0.45)",
                          }}>
                            <div style={{ fontSize: 9, color: muted, marginBottom: 6, letterSpacing: "0.12em", textTransform: "uppercase" }}>
                              Format visuel
                            </div>
                            <input autoFocus
                              value={fmtPattern[ci] || ""}
                              onChange={(e) => setFmtPattern((p) => ({ ...p, [ci]: e.target.value }))}
                              onKeyDown={(e) => { if (e.key === "Enter" || e.key === "Escape") setFmtOpen((p) => ({ ...p, [ci]: false })); }}
                              placeholder="ex: 4 2 3"
                              style={{ ...inpStyle(), width: "100%", boxSizing: "border-box" }}
                            />
                            <div style={{ fontSize: 9, color: muted, marginTop: 8, lineHeight: 1.65 }}>
                              Groupes de car. séparés par espace.<br />
                              «&nbsp;4 2 3&nbsp;» → XXXX XX XXX reste
                            </div>
                            {fmtPattern[ci] && (
                              <button
                                onClick={() => setFmtPattern((p) => { const n = { ...p }; delete n[ci]; return n; })}
                                style={{ ...btnStyle(), background: "transparent", border: `1px solid ${border}`,
                                  color: muted, marginTop: 8, width: "100%" }}>
                                Effacer le format
                              </button>
                            )}
                          </div>
                        )}
                      </th>
                    );
                  })}
                </tr>
              </thead>
              <tbody>
                {filteredIndices.map((rowIndex, i) => {
                  const row    = rows[rowIndex];
                  const destId = rowDestMap[rowIndex];
                  const dest   = getDestById(destId);
                  const rowBg  = dest ? dest.color + "22" : (i % 2 === 0 ? surface : rowAlt);
                  return (
                    <tr key={rowIndex} onClick={() => handleRowClick(rowIndex)}
                      style={{ background: rowBg, cursor: selectedDestId ? "pointer" : "default", transition: "filter 0.1s" }}
                      onMouseEnter={(e) => { (e.currentTarget as HTMLElement).style.filter = "brightness(1.12)"; }}
                      onMouseLeave={(e) => { (e.currentTarget as HTMLElement).style.filter = "none"; }}>
                      {row.map((cell, ci) => {
                        const raw = cellStr(cell);
                        const fmt = fmtPattern[ci] ? applyFormat(raw, fmtPattern[ci]) : raw;
                        return (
                          <td key={ci} style={{ padding: "5px 10px", borderBottom: `1px solid ${border}`,
                            whiteSpace: "nowrap", color: text }}>
                            {fmt}
                          </td>
                        );
                      })}
                      {/* DEST cell */}
                      <td style={{ padding: "4px 10px", borderBottom: `1px solid ${border}`, whiteSpace: "nowrap" }}>
                        {dest ? (
                          <span
                            style={{
                              display: "inline-flex", alignItems: "center", gap: 5,
                              background: dest.color, color: "#fff",
                              padding: "2px 8px", borderRadius: 3,
                              fontSize: 9, letterSpacing: "0.08em", fontWeight: 600,
                            }}
                          >
                            {dest.name}
                            <span
                              title="Désassigner"
                              onClick={(e) => {
                                e.stopPropagation();
                                setRowDestMap((p) => { const n = { ...p }; delete n[rowIndex]; return n; });
                              }}
                              style={{ opacity: 0.55, cursor: "pointer", fontSize: 10 }}>✕</span>
                          </span>
                        ) : (
                          <span style={{ color: muted, fontSize: 10 }}>—</span>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}
