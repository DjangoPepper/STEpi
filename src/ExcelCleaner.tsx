import { useState, useCallback, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import type { PointageData } from "./types";
import { useWindowWidth } from "./useWindowWidth";

type CellValue = string | number | boolean | null;
type RawData = CellValue[][];

interface ExcelCleanerProps {
  dark: boolean;
  onDarkToggle: () => void;
  onSendToPointage: (data: PointageData) => void;
}

interface ParsedData {
  headers: string[];
  rows: CellValue[][];
  headerRowIndex: number;
}

function detectHeaders(data: RawData): ParsedData {
  if (data.length === 0) return { headers: [], rows: [], headerRowIndex: 0 };
  let headerRowIndex = 0;
  for (let i = 0; i < Math.min(5, data.length); i++) {
    const row = data[i];
    const stringCount = row.filter((c) => typeof c === "string" && c.trim() !== "").length;
    const filledCount = row.filter((c) => c !== null && c !== undefined && c !== "").length;
    if (filledCount > 0 && stringCount / filledCount > 0.6) { headerRowIndex = i; break; }
  }
  const headers = data[headerRowIndex].map((h, i) =>
    h !== null && h !== undefined && h !== "" ? String(h) : `Colonne ${i + 1}`
  );
  return { headers, rows: data.slice(headerRowIndex + 1), headerRowIndex };
}

function computeRepetitiveValues(rows: RawData, colIndex: number, threshold = 0.35): Set<string> {
  const counts: Record<string, number> = {};
  let total = 0;
  for (const row of rows) {
    const val = row[colIndex];
    if (val !== null && val !== undefined && val !== "") {
      const key = String(val); counts[key] = (counts[key] || 0) + 1; total++;
    }
  }
  const result = new Set<string>();
  if (total === 0) return result;
  for (const [val, count] of Object.entries(counts))
    if (count / total >= threshold && count > 1) result.add(val);
  return result;
}

// Modal for adding a new row
function AddRowModal({
  headers, hiddenCols, onClose, onAdd, dark,
}: {
  headers: string[]; hiddenCols: Set<number>;
  onClose: () => void; onAdd: (row: CellValue[]) => void; dark: boolean;
}) {
  const [values, setValues] = useState<string[]>(headers.map(() => ""));

  const handleAdd = () => {
    onAdd(values.map((v) => (v.trim() === "" ? null : v.trim())));
    onClose();
  };

  const visibleCols = headers.map((h, i) => ({ h, i })).filter(({ i }) => !hiddenCols.has(i));

  return (
    <div style={{
      position: "fixed", inset: 0, background: "rgba(0,0,0,0.6)",
      display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100,
    }}
      onClick={onClose}
    >
      <div
        style={{
          background: dark ? "#141414" : "#fff",
          border: `1px solid ${dark ? "#2a2a2a" : "#ddd"}`,
          borderRadius: 6, padding: 28, minWidth: "min(360px, calc(100vw - 32px))", maxWidth: 600,
          maxHeight: "80vh", overflowY: "auto",
          boxShadow: "0 20px 60px rgba(0,0,0,0.5)",
        }}
        onClick={(e) => e.stopPropagation()}
      >
        <div style={{ fontSize: 11, letterSpacing: "0.15em", textTransform: "uppercase", color: dark ? "#6ee7b7" : "#059669", marginBottom: 20, fontFamily: "'IBM Plex Mono', monospace" }}>
          + Ajouter une ligne
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {visibleCols.map(({ h, i }) => (
            <div key={i} style={{ display: "flex", flexDirection: "column", gap: 4 }}>
              <label style={{
                fontSize: 9, letterSpacing: "0.12em", textTransform: "uppercase",
                color: dark ? "#555" : "#999", fontFamily: "'IBM Plex Mono', monospace",
              }}>{h}</label>
              <input
                value={values[i]}
                onChange={(e) => { const next = [...values]; next[i] = e.target.value; setValues(next); }}
                onKeyDown={(e) => { if (e.key === "Enter") handleAdd(); if (e.key === "Escape") onClose(); }}
                style={{
                  fontFamily: "'IBM Plex Mono', monospace", fontSize: 12,
                  padding: "7px 10px",
                  background: dark ? "#0d0d0d" : "#f8f8f8",
                  border: `1px solid ${dark ? "#2a2a2a" : "#e0e0e0"}`,
                  borderRadius: 3,
                  color: dark ? "#e8e8e0" : "#1a1a1a",
                  outline: "none",
                  transition: "border-color 0.15s",
                }}
                onFocus={(e) => { e.target.style.borderColor = dark ? "#6ee7b7" : "#059669"; }}
                onBlur={(e) => { e.target.style.borderColor = dark ? "#2a2a2a" : "#e0e0e0"; }}
                placeholder={`Valeur pour ${h}…`}
              />
            </div>
          ))}
        </div>
        <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 24 }}>
          <button
            onClick={onClose}
            style={{
              fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, letterSpacing: "0.08em",
              textTransform: "uppercase", padding: "8px 16px",
              border: `1px solid ${dark ? "#333" : "#ddd"}`, borderRadius: 3,
              cursor: "pointer", background: "transparent",
              color: dark ? "#666" : "#999",
            }}
          >Annuler</button>
          <button
            onClick={handleAdd}
            style={{
              fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, letterSpacing: "0.08em",
              textTransform: "uppercase", padding: "8px 16px",
              border: `1px solid ${dark ? "#166534" : "#059669"}`, borderRadius: 3,
              cursor: "pointer",
              background: dark ? "#0a1f10" : "#f0fdf4",
              color: dark ? "#86efac" : "#059669",
            }}
          >Ajouter</button>
        </div>
      </div>
    </div>
  );
}

export default function ExcelCleaner({ dark, onDarkToggle: _onDarkToggle, onSendToPointage }: ExcelCleanerProps) {
  const vw       = useWindowWidth();
  const isMobile = vw < 640;
  const [fileName, setFileName] = useState<string | null>(null);
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [activeSheet, setActiveSheet] = useState<string | null>(null);
  const [parsed, setParsed] = useState<ParsedData | null>(null);
  const [hiddenCols, setHiddenCols] = useState<Set<number>>(new Set());
  const [hiddenRows, setHiddenRows] = useState<Set<number>>(new Set());
  const [editingHeader, setEditingHeader] = useState<number | null>(null);
  const [headers, setHeaders] = useState<string[]>([]);
  const [addedRows, setAddedRows] = useState<CellValue[][]>([]);
  const [selectMode, setSelectMode] = useState<"none" | "col" | "row">("none");
  const [selectedItems, setSelectedItems] = useState<Set<number>>(new Set());
  const [dimRepeated, setDimRepeated] = useState(true);
  const [exportFileName, setExportFileName] = useState("données_nettoyées");
  const [showAddRow, setShowAddRow] = useState(false);
  const [hiddenSheets, setHiddenSheets] = useState<Set<string>>(new Set());
  const [sheetSelectMode, setSheetSelectMode] = useState<"none" | "delete" | "keep">("none");
  const [selectedSheets, setSelectedSheets] = useState<Set<string>>(new Set());
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Merged rows: original visible rows + added rows
  const allRows = useMemo(() => {
    if (!parsed) return [];
    return [...addedRows, ...parsed.rows];
  }, [parsed, addedRows]);

  const loadSheet = useCallback((wb: XLSX.WorkBook, sheetName: string) => {
    const sheet = wb.Sheets[sheetName];
    const raw: RawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null }) as RawData;
    const result = detectHeaders(raw);
    setParsed(result);
    setHeaders(result.headers);
    setHiddenCols(new Set()); setHiddenRows(new Set());
    setSelectMode("none"); setSelectedItems(new Set());
    setEditingHeader(null); setAddedRows([]);
  }, []);

  const handleFile = useCallback((file: File) => {
    setFileName(file.name);
    const baseName = file.name.replace(/\.[^/.]+$/, "");
    setExportFileName(`${baseName}_cleaned`);
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer);
      const wb = XLSX.read(data, { type: "array" });
      setWorkbook(wb); setSheetNames(wb.SheetNames);
      setActiveSheet(wb.SheetNames[0]);
      setHiddenSheets(new Set());
      setSheetSelectMode("none");
      setSelectedSheets(new Set());
      loadSheet(wb, wb.SheetNames[0]);
    };
    reader.readAsArrayBuffer(file);
  }, [loadSheet]);

  const handleSheetChange = (name: string) => {
    if (!workbook) return;
    setActiveSheet(name); loadSheet(workbook, name);
  };

  const onDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  }, [handleFile]);

  const toggleSelectItem = (idx: number) => {
    setSelectedItems((prev) => {
      const next = new Set(prev);
      next.has(idx) ? next.delete(idx) : next.add(idx); return next;
    });
  };

  const repetitiveByCol = useMemo(() => {
    if (!parsed || !dimRepeated) return new Map<number, Set<string>>();
    const map = new Map<number, Set<string>>();
    headers.forEach((_, ci) => map.set(ci, computeRepetitiveValues(allRows, ci)));
    return map;
  }, [parsed, headers, dimRepeated, allRows]);

  const applyColAction = () => {
    if (selectMode !== "col") return;
    setHiddenCols((prev) => new Set([...prev, ...selectedItems]));
    setSelectedItems(new Set()); setSelectMode("none");
  };

  const applyRowDeletion = () => {
    if (selectMode !== "row") return;
    setHiddenRows((prev) => new Set([...prev, ...selectedItems]));
    setSelectedItems(new Set()); setSelectMode("none");
  };

  const exportClean = () => {
    if (!workbook) return;
    const wb2 = XLSX.utils.book_new();
    const visibleSheets = sheetNames.filter((n) => !hiddenSheets.has(n));

    visibleSheets.forEach((sheetName) => {
      // For the active sheet, use current edited state; for others, use raw data
      if (sheetName === activeSheet && parsed) {
        const visibleHeaders = headers.filter((_, i) => !hiddenCols.has(i));
        const visibleRows = allRows
          .filter((_, i) => !hiddenRows.has(i))
          .map((row) => {
            const padded = [...row];
            while (padded.length < headers.length) padded.push(null);
            return padded.filter((_, ci) => !hiddenCols.has(ci));
          });
        const ws = XLSX.utils.aoa_to_sheet([visibleHeaders, ...visibleRows]);
        XLSX.utils.book_append_sheet(wb2, ws, sheetName);
      } else {
        // Copy sheet as-is from the original workbook
        const ws = workbook.Sheets[sheetName];
        XLSX.utils.book_append_sheet(wb2, ws, sheetName);
      }
    });

    if (wb2.SheetNames.length === 0) return;
    XLSX.writeFile(wb2, `${exportFileName.trim() || "données_nettoyées"}.xlsx`);
  };

  const visibleColCount = parsed ? headers.filter((_, i) => !hiddenCols.has(i)).length : 0;
  const visibleRowCount = allRows.filter((_, i) => !hiddenRows.has(i)).length;

  const applySheetAction = () => {
    if (sheetSelectMode === "delete") {
      const next = new Set(hiddenSheets);
      selectedSheets.forEach((n) => next.add(n));
      setHiddenSheets(next);
      // If active sheet was hidden, switch to first visible
      if (selectedSheets.has(activeSheet ?? "")) {
        const firstVisible = sheetNames.find((n) => !next.has(n));
        if (firstVisible && workbook) { setActiveSheet(firstVisible); loadSheet(workbook, firstVisible); }
      }
    } else {
      // keep selected, hide the rest
      const next = new Set<string>(sheetNames.filter((n) => !selectedSheets.has(n)));
      setHiddenSheets(next);
      if (next.has(activeSheet ?? "")) {
        const firstVisible = sheetNames.find((n) => !next.has(n));
        if (firstVisible && workbook) { setActiveSheet(firstVisible); loadSheet(workbook, firstVisible); }
      }
    }
    setSelectedSheets(new Set());
    setSheetSelectMode("none");
  };

  const actionLabel = () => {
    const n = selectedItems.size;
    if (selectMode === "col") return `Supprimer ${n} col.`;
    return `Supprimer ${n} ligne(s)`;
  };

  // ─── theme tokens ───────────────────────────────────────────────────────────
  const t = {
    bg:        dark ? "#0d0d0d" : "#f5f5f3",
    bgCard:    dark ? "#111"    : "#fff",
    bgHeader:  dark ? "#131313": "#f0f0ee",
    bgRowNum:  dark ? "#161616": "#f7f7f5",
    border:    dark ? "#1e1e1e": "#e4e4e0",
    border2:   dark ? "#2a2a2a": "#d0d0cc",
    text:      dark ? "#e8e8e0": "#1a1a1a",
    textMuted: dark ? "#666"   : "#888",
    textDim:   dark ? "#3a3a3a": "#ccc",
    textRowNum:dark ? "#888"   : "#aaa",
    textTh:    dark ? "#aaa"   : "#666",
    accent:    dark ? "#6ee7b7": "#059669",
    accentBg:  dark ? "#0f2a20": "#f0fdf4",
    accentBor: dark ? "#166534": "#059669",
    cellRepeat:dark ? "#7c3aed": "#c8c8c4",
    cellRepeatBg: dark ? "#1a1030": "transparent",
    scrollBar: dark ? "#3a3a3a": "#c8c8c4",
    rowHover:  dark ? "#111"   : "#fafaf8",
    selRowBg:  dark ? "#0f2a20": "#ecfdf5",
    selRowTxt: dark ? "#6ee7b7": "#059669",
  };

  const css = `
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@300;400;500;600&family=Space+Grotesk:wght@300;400;500&display=swap');
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { background: ${t.bg}; }
    ::-webkit-scrollbar { width: 6px; height: 6px; }
    ::-webkit-scrollbar-track { background: transparent; }
    ::-webkit-scrollbar-thumb { background: ${t.scrollBar}; border-radius: 3px; }

    .drop-zone {
      border: 1.5px dashed ${t.border2}; border-radius: 4px; padding: 60px 40px;
      text-align: center; cursor: pointer; transition: all 0.2s; background: ${t.bgCard};
    }
    .drop-zone:hover { border-color: ${t.accent}; background: ${t.accentBg}; }

    .btn {
      font-family: 'IBM Plex Mono', monospace; font-size: 11px; font-weight: 500;
      letter-spacing: 0.08em; text-transform: uppercase; padding: 8px 16px;
      border: 1px solid ${t.border2}; border-radius: 3px;
      cursor: pointer; background: ${t.bgCard}; color: ${t.text}; transition: all 0.15s;
    }
    .btn:hover { border-color: ${dark ? "#555" : "#aaa"}; background: ${dark ? "#1a1a1a" : "#f0f0ee"}; }
    .btn.active { background: ${t.accentBg}; border-color: ${t.accent}; color: ${t.accent}; }
    .btn.danger { border-color: ${dark ? "#7f1d1d" : "#fca5a5"}; color: ${dark ? "#fca5a5" : "#dc2626"}; }
    .btn.danger:hover { background: ${dark ? "#1f0a0a" : "#fff5f5"}; border-color: #ef4444; }
    .btn.confirm-keep { border-color: ${t.accentBor}; color: ${dark ? "#86efac" : "#059669"}; background: ${t.accentBg}; }
    .btn.confirm-keep:hover { border-color: ${t.accent}; }
    .btn.confirm-del  { border-color: ${dark ? "#7f1d1d" : "#fca5a5"}; color: ${dark ? "#fca5a5" : "#dc2626"}; background: ${dark ? "#1f0a0a" : "#fff5f5"}; }
    .btn.confirm-del:hover { border-color: #ef4444; }
    .btn.success { border-color: ${t.accentBor}; color: ${dark ? "#86efac" : "#059669"}; }
    .btn.success:hover { background: ${t.accentBg}; border-color: ${t.accent}; }

    .sheet-tabs { display: flex; border-bottom: 1px solid ${t.border}; margin-bottom: 20px; overflow-x: auto; }
    .sheet-tab {
      font-family: 'IBM Plex Mono', monospace; font-size: 11px; padding: 8px 18px;
      cursor: pointer; border: 1px solid transparent; border-bottom: none;
      border-radius: 4px 4px 0 0; color: ${t.textMuted}; background: transparent;
      transition: all 0.15s; white-space: nowrap; letter-spacing: 0.05em;
      position: relative; top: 1px;
    }
    .sheet-tab:hover { color: ${t.text}; background: ${t.bgHeader}; border-color: ${t.border}; border-bottom-color: transparent; }
    .sheet-tab.active { color: ${t.accent}; background: ${t.bg}; border-color: ${t.border2}; border-bottom-color: ${t.bg}; }
    .sheet-tab.tab-sel-delete { color: ${dark ? "#fca5a5" : "#dc2626"} !important; background: ${dark ? "#2a0e0e" : "#fff5f5"} !important; border-color: #ef4444 !important; border-bottom-color: transparent !important; }
    .sheet-tab.tab-sel-keep   { color: ${dark ? "#86efac" : "#059669"} !important; background: ${t.accentBg} !important; border-color: ${t.accent} !important; border-bottom-color: transparent !important; }
    .sheet-tab.tab-hidden { opacity: 0.35; text-decoration: line-through; }
    .sheet-tab.tab-selectable { cursor: pointer; }
    .sheet-tab.tab-selectable:hover { opacity: 1; }

    table { border-collapse: collapse; width: 100%; font-size: 12px; }
    thead tr th {
      background: ${t.bgHeader}; color: ${t.textTh}; font-weight: 500;
      font-size: 10px; letter-spacing: 0.1em; text-transform: uppercase;
      padding: 10px 14px; text-align: left; border-bottom: 1px solid ${t.border};
      white-space: nowrap; position: sticky; top: 0; z-index: 2;
    }
    th.col-selectable { cursor: pointer; }
    th.col-selectable:hover { background: ${dark ? "#1a2e24" : "#ecfdf5"}; color: ${t.accent}; }
    th.col-sel-delete { background: ${dark ? "#2a0e0e" : "#fff5f5"} !important; color: ${dark ? "#fca5a5" : "#dc2626"} !important; border-bottom-color: #ef4444 !important; }

    td {
      padding: 8px 14px; border-bottom: 1px solid ${t.border};
      color: ${dark ? "#c8c8c0" : "#333"}; white-space: nowrap;
      max-width: 220px; overflow: hidden; text-overflow: ellipsis;
    }
    tr:hover td { background: ${t.rowHover}; }
    tr.row-selected td { background: ${t.selRowBg} !important; color: ${t.selRowTxt}; }
    tr.added-row td { background: ${dark ? "#0a1a10" : "#f0fdf4"} !important; }
    tr.added-row:hover td { background: ${dark ? "#0f2218" : "#ecfdf5"} !important; }

    .td-rownum {
      color: ${t.textRowNum} !important; font-size: 11px !important; background: ${t.bgRowNum} !important;
      border-right: 1px solid ${t.border} !important;
      min-width: 46px; width: 46px; text-align: center !important;
      user-select: none; font-variant-numeric: tabular-nums;
    }
    .th-rownum {
      background: ${dark ? "#0f0f0f" : "#ebebea"} !important;
      border-right: 1px solid ${t.border} !important;
      width: 46px; min-width: 46px; text-align: center !important;
      color: ${t.textDim} !important;
    }
    tr:hover .td-rownum { background: ${dark ? "#1a1a1a" : "#efefed"} !important; }
    tr.row-selected .td-rownum { background: ${dark ? "#0a2018" : "#d1fae5"} !important; color: ${t.selRowTxt} !important; }
    .td-rownum.row-selectable { cursor: pointer; }
    .td-rownum.row-selectable:hover { color: ${t.accent} !important; background: ${dark ? "#0f2218" : "#ecfdf5"} !important; }

    .cell-repeated { color: ${t.cellRepeat} !important; font-style: italic; ${dark ? `background: ${t.cellRepeatBg};` : ""} }
    tr:hover .cell-repeated { color: ${dark ? "#a78bfa" : "#b0b0ac"} !important; ${dark ? "background: #1f1540;" : ""} }

    .header-label { display: flex; align-items: center; gap: 5px; cursor: pointer; }
    .header-label:hover .edit-pencil { opacity: 1; }
    .edit-pencil { opacity: 0; font-size: 9px; color: ${t.accent}; transition: opacity 0.15s; flex-shrink: 0; }

    .header-input {
      background: transparent; border: none; border-bottom: 1px solid ${t.accent};
      color: ${t.accent}; font-family: 'IBM Plex Mono', monospace;
      font-size: 10px; letter-spacing: 0.08em; text-transform: uppercase;
      width: 100%; min-width: 50px; outline: none; padding: 2px 0;
    }

    .badge { display: inline-block; font-size: 9px; padding: 2px 6px; border-radius: 2px; letter-spacing: 0.05em; }
    .badge-hidden { background: ${dark ? "#2a1515" : "#fff1f1"}; color: ${dark ? "#f87171" : "#dc2626"}; margin-left: 4px; }

    .info-bar {
      margin-bottom: 12px; padding: 8px 14px;
      background: ${t.accentBg}; border: 1px solid ${dark ? "#1a3a28" : "#a7f3d0"};
      border-radius: 3px; font-size: 11px; color: ${dark ? "#6ee7b7" : "#059669"};
    }
    .info-bar.del { background: ${dark ? "#1f0a0a" : "#fff5f5"}; border-color: ${dark ? "#3a1515" : "#fca5a5"}; color: ${dark ? "#fca5a5" : "#dc2626"}; }

    .dim-chip {
      display: flex; align-items: center; gap: 6px;
      font-size: 10px; letter-spacing: 0.06em; text-transform: uppercase;
      color: ${t.textMuted}; cursor: pointer; padding: 7px 12px;
      border: 1px solid ${t.border2}; border-radius: 3px; background: ${t.bgCard};
      transition: all 0.15s; user-select: none;
    }
    .dim-chip:hover { border-color: ${dark ? "#444" : "#bbb"}; }
    .dim-chip.on { color: #a78bfa; border-color: ${dark ? "#3d2a6a" : "#c4b5fd"}; background: ${dark ? "#160f2a" : "#f5f3ff"}; }
    .dim-dot { width: 7px; height: 7px; border-radius: 50%; background: currentColor; flex-shrink: 0; }

    .theme-toggle {
      display: flex; align-items: center; gap: 6px;
      font-size: 10px; letter-spacing: 0.06em; text-transform: uppercase;
      color: ${t.textMuted}; cursor: pointer; padding: 7px 12px;
      border: 1px solid ${t.border2}; border-radius: 3px; background: ${t.bgCard};
      transition: all 0.15s; user-select: none;
    }
    .theme-toggle:hover { border-color: ${dark ? "#555" : "#aaa"}; color: ${t.text}; }

    .export-name-wrap { display: flex; align-items: center; border: 1px solid ${t.accentBor}; border-radius: 3px; overflow: hidden; }
    .export-name-input {
      font-family: 'IBM Plex Mono', monospace; font-size: 11px; letter-spacing: 0.05em;
      background: ${t.accentBg}; color: ${dark ? "#86efac" : "#059669"};
      border: none; outline: none; padding: 8px 10px; width: 180px;
    }
    .export-name-ext {
      font-family: 'IBM Plex Mono', monospace; font-size: 11px;
      color: ${dark ? "#3a6a4a" : "#86efac"}; background: ${t.accentBg};
      padding: 8px 6px 8px 0; pointer-events: none;
    }
    .export-btn {
      font-family: 'IBM Plex Mono', monospace; font-size: 11px; font-weight: 500;
      letter-spacing: 0.08em; text-transform: uppercase; padding: 8px 14px;
      border: none; border-left: 1px solid ${t.accentBor};
      cursor: pointer; background: ${dark ? "#0f2a18" : "#d1fae5"}; color: ${dark ? "#86efac" : "#059669"};
      transition: all 0.15s;
    }
    .export-btn:hover { background: ${dark ? "#1a3a24" : "#a7f3d0"}; }

    @media (max-width: 639px) {
      .btn {
        font-size: 15px !important; padding: 13px 18px !important;
        letter-spacing: 0 !important; text-transform: none !important;
        min-height: 46px;
      }
      .sheet-tab { font-size: 14px !important; padding: 12px 16px !important; letter-spacing: 0 !important; }
      .export-name-input { font-size: 15px !important; padding: 13px 10px !important; min-height: 46px; width: 140px !important; }
      .export-name-ext { font-size: 15px !important; padding: 13px 4px 13px 0 !important; }
      .export-btn { font-size: 15px !important; padding: 13px 16px !important; letter-spacing: 0 !important; text-transform: none !important; min-height: 46px; }
      .dim-chip { font-size: 13px !important; padding: 11px 14px !important; letter-spacing: 0 !important; text-transform: none !important; min-height: 46px; }
      .theme-toggle { font-size: 13px !important; padding: 11px 14px !important; letter-spacing: 0 !important; text-transform: none !important; }
      .info-bar { font-size: 14px !important; padding: 12px 16px !important; }
      table { font-size: 14px !important; }
      thead tr th { font-size: 12px !important; padding: 12px 14px !important; }
      td { padding: 11px 14px !important; }
    }
  `;

  return (
    <div style={{ fontFamily: "'IBM Plex Mono', monospace", minHeight: "100vh", background: t.bg, color: t.text, transition: "background 0.2s, color 0.2s" }}>
      <style>{css}</style>

      {showAddRow && parsed && (
        <AddRowModal
          headers={headers}
          hiddenCols={hiddenCols}
          dark={dark}
          onClose={() => setShowAddRow(false)}
          onAdd={(row) => {
            const padded = headers.map((_, i) => row[i] ?? null);
            setAddedRows((prev) => [padded, ...prev]);
          }}
        />
      )}

      <div style={{ maxWidth: 900, margin: "0 auto", padding: isMobile ? "16px 10px" : "32px 24px" }}>
        {/* Top bar */}
        <div style={{ marginBottom: isMobile ? 16 : 32, borderBottom: `1px solid ${t.border}`, paddingBottom: isMobile ? 12 : 24, display: "flex", justifyContent: "space-between", alignItems: "flex-end", flexWrap: "wrap", gap: 8 }}>
          <div>
            <div style={{ fontSize: isMobile ? 13 : 10, letterSpacing: isMobile ? 0 : "0.2em", color: t.textDim, textTransform: isMobile ? "none" : "uppercase", marginBottom: 8 }}>Outil de nettoyage</div>
            <h1 style={{ fontSize: 24, fontWeight: 600, letterSpacing: "-0.02em", color: t.text, fontFamily: "Space Grotesk, sans-serif" }}>
              Excel <span style={{ color: t.accent }}>Cleaner</span>
            </h1>
          </div>
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            {parsed && (
              <>
                <span style={{ fontSize: isMobile ? 14 : 11, color: t.textMuted, marginRight: 4 }}>
                  <span style={{ color: t.accent, fontWeight: 600 }}>{visibleColCount}</span> col ·{" "}
                  <span style={{ color: t.accent, fontWeight: 600 }}>{visibleRowCount}</span> lignes
                  {sheetNames.length > 1 && (
                    <span style={{ marginLeft: 8, color: t.textMuted }}>·{" "}
                      <span style={{ color: t.accent, fontWeight: 600 }}>{sheetNames.length - hiddenSheets.size}</span>/{sheetNames.length} onglets
                    </span>
                  )}
                  {(hiddenCols.size > 0 || hiddenRows.size > 0) && (
                    <span style={{ marginLeft: 8, color: "#ef4444" }}>({hiddenCols.size + hiddenRows.size} masqués)</span>
                  )}
                  {addedRows.length > 0 && (
                    <span style={{ marginLeft: 8, color: dark ? "#86efac" : "#059669" }}>+{addedRows.length} ajoutée(s)</span>
                  )}
                </span>
                {/* Send to Pointage */}
                <button
                  className="btn success"
                  title="Envoyer les données nettoyées vers l'onglet Pointage"
                  onClick={() => {
                    if (!parsed) return;
                    const visHdrs = headers.filter((_, i) => !hiddenCols.has(i));
                    const visRows = allRows
                      .filter((_, i) => !hiddenRows.has(i))
                      .map((row) => {
                        const padded = [...row];
                        while (padded.length < headers.length) padded.push(null);
                        return padded.filter((_, ci) => !hiddenCols.has(ci));
                      });
                    onSendToPointage({ headers: visHdrs, rows: visRows, fileName: fileName ?? "fichier" });
                  }}
                >
                  ↗ Pointage
                </button>
                {/* Export */}
                <div className="export-name-wrap">
                  <input
                    className="export-name-input"
                    value={exportFileName}
                    onChange={(e) => setExportFileName(e.target.value)}
                    onKeyDown={(e) => { if (e.key === "Enter") exportClean(); }}
                    title="Nom du fichier exporté"
                    spellCheck={false}
                  />
                  <span className="export-name-ext">.xlsx</span>
                  <button className="export-btn" onClick={exportClean}>↓ Exporter</button>
                </div>
                <button className="btn" onClick={() => { setParsed(null); setFileName(null); setWorkbook(null); setSheetNames([]); setActiveSheet(null); setAddedRows([]); setHiddenSheets(new Set()); setSheetSelectMode("none"); setSelectedSheets(new Set()); }}>
                  ✕ Réinitialiser
                </button>
              </>
            )}
          </div>
        </div>

        {!parsed ? (
          <div className="drop-zone" onDrop={onDrop} onDragOver={(e) => e.preventDefault()} onClick={() => fileInputRef.current?.click()}>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }}
              onChange={(e) => { const f = e.target.files?.[0]; if (f) handleFile(f); }} />
            <div style={{ fontSize: 32, marginBottom: 16, opacity: 0.2 }}>⊞</div>
            <div style={{ fontSize: 13, color: t.textMuted, marginBottom: 8 }}>Glissez un fichier Excel ou cliquez pour sélectionner</div>
            <div style={{ fontSize: 10, color: t.textDim, letterSpacing: "0.1em", textTransform: "uppercase" }}>.xlsx · .xls · .csv</div>
          </div>
        ) : (
          <div>
            {/* Sheet tabs */}
            {sheetNames.length > 1 && (
              <div style={{ marginBottom: 20 }}>
                {/* Tab controls bar */}
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                  <span style={{ fontSize: 10, letterSpacing: "0.1em", textTransform: "uppercase", color: t.textMuted }}>
                    Onglets ({sheetNames.length - hiddenSheets.size}/{sheetNames.length} visibles)
                  </span>
                  <div style={{ flex: 1 }} />
                  {sheetSelectMode === "none" ? (
                    <>
                      <button className="btn" style={{ fontSize: 10, padding: "5px 11px" }}
                        onClick={() => { setSheetSelectMode("delete"); setSelectedSheets(new Set()); }}>
                        ✕ Supprimer onglets
                      </button>
                      <button className="btn" style={{ fontSize: 10, padding: "5px 11px" }}
                        onClick={() => { setSheetSelectMode("keep"); setSelectedSheets(new Set()); }}>
                        ✓ Conserver onglets
                      </button>
                      {hiddenSheets.size > 0 && (
                        <button className="btn" style={{ fontSize: 10, padding: "5px 11px" }}
                          onClick={() => setHiddenSheets(new Set())}>
                          Restaurer onglets
                        </button>
                      )}
                    </>
                  ) : (
                    <>
                      <span style={{ fontSize: 10, color: sheetSelectMode === "delete" ? (dark ? "#fca5a5" : "#dc2626") : (dark ? "#86efac" : "#059669"), letterSpacing: "0.06em" }}>
                        {sheetSelectMode === "delete" ? "→ Cliquez les onglets à supprimer" : "→ Cliquez les onglets à conserver"}
                      </span>
                      {selectedSheets.size > 0 && (
                        <button
                          className={`btn ${sheetSelectMode === "keep" ? "confirm-keep" : "confirm-del"}`}
                          style={{ fontSize: 10, padding: "5px 11px" }}
                          onClick={applySheetAction}
                        >
                          {sheetSelectMode === "delete" ? `Supprimer ${selectedSheets.size}` : `Garder ${selectedSheets.size}`}
                        </button>
                      )}
                      <button className="btn" style={{ fontSize: 10, padding: "5px 11px" }}
                        onClick={() => { setSheetSelectMode("none"); setSelectedSheets(new Set()); }}>
                        Annuler
                      </button>
                    </>
                  )}
                </div>
                {/* Tabs row */}
                <div className="sheet-tabs" style={{ marginBottom: 0 }}>
                  {sheetNames.map((name) => {
                    const isHidden = hiddenSheets.has(name);
                    const isActive = activeSheet === name;
                    const isSel = selectedSheets.has(name);
                    const inSelectMode = sheetSelectMode !== "none";
                    let cls = "sheet-tab";
                    if (isHidden) cls += " tab-hidden";
                    if (isActive && !inSelectMode) cls += " active";
                    if (inSelectMode) cls += " tab-selectable";
                    if (isSel && sheetSelectMode === "delete") cls += " tab-sel-delete";
                    if (isSel && sheetSelectMode === "keep") cls += " tab-sel-keep";
                    return (
                      <button
                        key={name}
                        className={cls}
                        onClick={() => {
                          if (inSelectMode) {
                            setSelectedSheets((prev) => {
                              const next = new Set(prev);
                              next.has(name) ? next.delete(name) : next.add(name);
                              return next;
                            });
                          } else if (!isHidden) {
                            handleSheetChange(name);
                          }
                        }}
                        title={isHidden ? `${name} (masqué — cliquer Restaurer onglets)` : inSelectMode ? `Sélectionner "${name}"` : name}
                      >
                        {name}
                        {isHidden && <span style={{ marginLeft: 4, fontSize: 8, opacity: 0.6 }}>✕</span>}
                      </button>
                    );
                  })}
                </div>
              </div>
            )}

            {/* Toolbar */}
            <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 14, flexWrap: "wrap" }}>
              <span style={{ fontSize: 11, color: t.textDim, marginRight: 4 }}>
                📄 <span style={{ color: t.textMuted }}>{fileName}</span>
                {activeSheet && sheetNames.length > 1 && <span style={{ color: t.textDim, marginLeft: 6 }}>· {activeSheet}</span>}
              </span>
              <div style={{ flex: 1 }} />

              {/* Dim chip */}
              <div className={`dim-chip ${dimRepeated ? "on" : ""}`} onClick={() => setDimRepeated((v) => !v)}>
                <span className="dim-dot" />Griser répétitions
              </div>

              {/* Add row */}
              <button className="btn" onClick={() => setShowAddRow(true)}>+ Ligne</button>


              <button className={`btn ${selectMode === "col" ? "active" : ""}`}
                onClick={() => { setSelectMode(selectMode === "col" ? "none" : "col"); setSelectedItems(new Set()); }}>
                {selectMode === "col" ? "✓ " : ""}Colonnes
              </button>
              <button className={`btn ${selectMode === "row" ? "active" : ""}`}
                onClick={() => { setSelectMode(selectMode === "row" ? "none" : "row"); setSelectedItems(new Set()); }}>
                {selectMode === "row" ? "✓ " : ""}Lignes
              </button>

              {selectedItems.size > 0 && (
                <button className="btn confirm-del"
                  onClick={selectMode === "col" ? applyColAction : applyRowDeletion}>
                  {actionLabel()}
                </button>
              )}
              {(hiddenCols.size > 0 || hiddenRows.size > 0) && (
                <button className="btn" onClick={() => { setHiddenCols(new Set()); setHiddenRows(new Set()); }}>Restaurer tout</button>
              )}
            </div>

            {/* Context hints */}
            {selectMode === "col" && (
              <div className="info-bar del">
                → Cliquez sur les en-têtes de colonnes à supprimer, puis confirmez
              </div>
            )}
            {selectMode === "row" && <div className="info-bar del">→ Cliquez sur les numéros de lignes à supprimer, puis confirmez</div>}
            {selectMode === "none" && (
              <div style={{ marginBottom: 10, fontSize: 10, color: t.textDim, letterSpacing: "0.04em" }}>
                Cliquez sur un en-tête pour le renommer · Activez Colonnes ou Lignes pour sélectionner
              </div>
            )}

            {/* Table */}
            <div style={{ overflow: "auto", border: `1px solid ${t.border}`, borderRadius: 4, maxHeight: "calc(100vh - 340px)" }}>
              <table>
                <thead>
                  <tr>
                    <th className="th-rownum">#</th>
                    {headers.map((h, ci) => {
                      if (hiddenCols.has(ci)) return null;
                      const isSel = selectMode === "col" && selectedItems.has(ci);
                      const isEditing = editingHeader === ci;
                      let thClass = selectMode === "col" ? "col-selectable" : "";
                      if (isSel) thClass += " col-sel-delete";
                      return (
                        <th key={ci} className={thClass}
                          onClick={() => {
                            if (selectMode === "col") { toggleSelectItem(ci); return; }
                            if (selectMode === "none") setEditingHeader(ci);
                          }}
                        >
                          {isEditing ? (
                            <input className="header-input" value={headers[ci]} autoFocus
                              onChange={(e) => { const next = [...headers]; next[ci] = e.target.value; setHeaders(next); }}
                              onBlur={() => setEditingHeader(null)}
                              onKeyDown={(e) => { if (e.key === "Enter" || e.key === "Escape") setEditingHeader(null); }}
                              onClick={(e) => e.stopPropagation()}
                            />
                          ) : (
                            <span className="header-label" title="Cliquer pour renommer">
                              {h}{selectMode === "none" && <span className="edit-pencil">✎</span>}
                            </span>
                          )}
                        </th>
                      );
                    })}
                  </tr>
                </thead>
                <tbody>
                  {allRows.map((row, ri) => {
                    if (hiddenRows.has(ri)) return null;
                    const isRowSel = selectMode === "row" && selectedItems.has(ri);
                    const isAdded = ri < addedRows.length;
                    return (
                      <tr key={ri} className={`${isRowSel ? "row-selected" : ""} ${isAdded ? "added-row" : ""}`}>
                        <td
                          className={`td-rownum ${selectMode === "row" ? "row-selectable" : ""}`}
                          onClick={() => selectMode === "row" && toggleSelectItem(ri)}
                          title={isAdded ? "Ligne ajoutée manuellement" : undefined}
                        >
                          {isAdded ? <span style={{ color: t.accent, fontSize: 9 }}>+{ri + 1}</span> : ri + 1}
                        </td>
                        {headers.map((_, ci) => {
                          if (hiddenCols.has(ci)) return null;
                          const cell = row[ci] ?? null;
                          const strVal = cell !== null ? String(cell) : null;
                          const isRepeated = !isAdded && dimRepeated && strVal !== null && (repetitiveByCol.get(ci)?.has(strVal) ?? false);
                          return (
                            <td key={ci} title={strVal ?? ""} className={isRepeated ? "cell-repeated" : ""}>
                              {strVal ?? <span style={{ color: t.textDim }}>—</span>}
                            </td>
                          );
                        })}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            {/* Footer */}
            <div style={{ marginTop: 10, display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 8 }}>
              {(hiddenCols.size > 0 || hiddenRows.size > 0) && (
                <div style={{ fontSize: 11, color: t.textMuted, display: "flex", gap: 16, flexWrap: "wrap" }}>
                  {hiddenCols.size > 0 && (
                    <span>Colonnes masquées :
                      {[...hiddenCols].map((i) => <span key={i} className="badge badge-hidden">{headers[i] || `col ${i + 1}`}</span>)}
                    </span>
                  )}
                  {hiddenRows.size > 0 && <span><span className="badge badge-hidden">{hiddenRows.size} ligne(s) masquée(s)</span></span>}
                </div>
              )}
              {dimRepeated && (
                <div style={{ fontSize: 10, color: dark ? "#7c3aed" : "#a78bfa", letterSpacing: "0.04em" }}>
                  ◆ Valeurs grisées = ≥35% des lignes de leur colonne
                </div>
              )}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
