import { useState, useEffect } from "react";

/* ─── localStorage helpers ─────────────────────────── */
const LS = {
  get<T>(key: string, fallback: T): T {
    try {
      const v = localStorage.getItem(key);
      return v !== null ? (JSON.parse(v) as T) : fallback;
    } catch { return fallback; }
  },
};

/* ─── Types ──────────────────────────────────────────────── */
type CellValue = string | number | boolean | null;
interface DestConfig { id: string; name: string; color: string; }

/* ─── Constants ──────────────────────────────────────────── */
const MONO = "'IBM Plex Mono', 'Fira Mono', monospace";
const WEIGHT_HEADERS = ["poids", "pds", "tons", "ton", "weight", "kg", "tonne", "tonnes"];

function numVal(v: CellValue): number {
  if (v == null) return 0;
  const n = typeof v === "number" ? v : parseFloat(String(v).replace(",", "."));
  return isNaN(n) ? 0 : n;
}

/* Strip thousand-separators (spaces, nbsp, thin-nbsp) then parse */
function parseRaw(s: string): number {
  return parseFloat(s.replace(/[\s\u00a0\u202f]/g, "").replace(",", "."));
}
/* Format a raw input string as a localized number on blur */
function fmtInput(s: string): string {
  const n = parseRaw(s);
  if (isNaN(n) || s.trim() === "") return s;
  return n % 1 === 0
    ? n.toLocaleString("fr-FR")
    : n.toLocaleString("fr-FR", { minimumFractionDigits: 2, maximumFractionDigits: 6 });
}

/* ─── Tally row ─────────────────────────────────────────── */
interface TallyRow {
  dest: DestConfig | null; // null = unassigned
  count: number;
  weight: number;
}

/* ─── Component ─────────────────────────────────────────── */
interface TallyProps { dark: boolean; }

export default function Tally({ dark }: TallyProps) {
  /* Theme */
  const bg     = dark ? "#0d0d0d" : "#f5f5f5";
  const surface = dark ? "#141414" : "#fff";
  const text   = dark ? "#e8e8e0" : "#1a1a1a";
  const accent = dark ? "#6ee7b7" : "#059669";
  const muted  = dark ? "#555"    : "#888";
  const border = dark ? "#222"    : "#ddd";
  const hdrBg  = dark ? "#0b0b0b" : "#f0f0f0";

  /* State — refreshed on every mount */
  const [headers, setHeaders]         = useState<string[]>([]);
  const [rows, setRows]               = useState<CellValue[][]>([]);
  const [destinations, setDestinations] = useState<DestConfig[]>([]);
  const [rowDestMap, setRowDestMap]   = useState<Record<number, string>>({});
  /* Maxi per destination id — editable, persisted */
  const [maxi, setMaxi] = useState<Record<string, string>>({});

  /* Forced-to-zero destinations (ids) */
  const [forcedDestIds, setForcedDestIds] = useState<string[]>([]);

  /* Weight column: -1 = none; index = chosen column. Persisted in localStorage. */
  const [weightCol, setWeightCol] = useState<number>(-2); // -2 = not yet initialised

  useEffect(() => {
    const hdrs = LS.get<string[]>("ptg_headers", []);
    const rws  = LS.get<CellValue[][]>("ptg_rows", []);
    const dsts = LS.get<DestConfig[]>("ptg_destinations", []);
    const rdm  = LS.get<Record<number, string>>("ptg_rowDestMap", {});
    setHeaders(hdrs);
    setRows(rws);
    setDestinations(dsts);
    setRowDestMap(rdm);
    setMaxi(LS.get<Record<string, string>>("tly_maxi", {}));
    setForcedDestIds(LS.get<string[]>("ptg_forceddests", []));
    // Restore or auto-detect weight column
    const saved = LS.get<number | null>("tly_weightCol", null);
    if (saved !== null) {
      setWeightCol(saved);
    } else {
      const auto = hdrs.findIndex(h => WEIGHT_HEADERS.includes(h.toLowerCase().trim()));
      setWeightCol(auto !== -1 ? auto : -1);
    }
  }, []);

  const hasWeight = weightCol >= 0 && rows.length > 0;

  /* Build tally */
  const destMap = new Map<string, DestConfig>(destinations.map(d => [d.id, d]));

  const tallyMap = new Map<string | null, TallyRow>();
  // Pre-populate with known destinations in order
  for (const d of destinations) {
    tallyMap.set(d.id, { dest: d, count: 0, weight: 0 });
  }
  // null = unassigned
  tallyMap.set(null, { dest: null, count: 0, weight: 0 });

  for (let ri = 0; ri < rows.length; ri++) {
    const destId: string | null = rowDestMap[ri] ?? null;
    const key = destId && destMap.has(destId) ? destId : null;
    const entry = tallyMap.get(key)!;
    entry.count += 1;
    if (hasWeight) entry.weight += numVal(rows[ri][weightCol]);
  }

  const tallyRows = [...tallyMap.values()].filter(r => r.count > 0 || r.dest !== null);
  const assignedRows   = tallyRows.filter(r => r.dest !== null && (r.count > 0 || forcedDestIds.includes(r.dest.id)));
  const unassigned     = tallyMap.get(null)!;
  const assignedCount  = assignedRows.reduce((s, r) => s + r.count, 0);
  const assignedWeight = assignedRows.reduce((s, r) => s + r.weight, 0);
  const totalCount     = [...tallyMap.values()].reduce((s, r) => s + r.count, 0);
  const totalWeight    = [...tallyMap.values()].reduce((s, r) => s + r.weight, 0);

  const setMaxiVal = (id: string, val: string) => {
    const next = { ...maxi, [id]: val };
    setMaxi(next);
    try { localStorage.setItem("tly_maxi", JSON.stringify(next)); } catch { /* ignore */ }
  };

  /* Q_MAXI per destination — déchargement */
  const [qMaxi, setQMaxi] = useState<Record<string, string>>({});
  useEffect(() => { setQMaxi(LS.get<Record<string, string>>("tly_qmaxi", {})); }, []);
  const setQMaxiVal = (id: string, val: string) => {
    const next = { ...qMaxi, [id]: val };
    setQMaxi(next);
    try { localStorage.setItem("tly_qmaxi", JSON.stringify(next)); } catch { /* ignore */ }
  };

  /* P_MAXI (poids) per destination — déchargement */
  const [dPMaxi, setDPMaxi] = useState<Record<string, string>>({});
  useEffect(() => { setDPMaxi(LS.get<Record<string, string>>("tly_dpmaxi", {})); }, []);
  const setDPMaxiVal = (id: string, val: string) => {
    const next = { ...dPMaxi, [id]: val };
    setDPMaxi(next);
    try { localStorage.setItem("tly_dpmaxi", JSON.stringify(next)); } catch { /* ignore */ }
  };

  /* Déchargement aggregates */
  const qStMaxiSum   = assignedRows.reduce((s, r) => { const v = parseRaw(qMaxi[r.dest!.id] ?? ""); return s + (isNaN(v) ? 0 : v); }, 0);
  const qStHasMaxi   = assignedRows.some(r => !isNaN(parseRaw(qMaxi[r.dest!.id] ?? "")));
  const qStRestant   = qStHasMaxi ? qStMaxiSum - assignedCount : NaN;
  const dpStMaxiSum  = assignedRows.reduce((s, r) => { const v = parseRaw(dPMaxi[r.dest!.id] ?? ""); return s + (isNaN(v) ? 0 : v); }, 0);
  const dpStHasMaxi  = assignedRows.some(r => !isNaN(parseRaw(dPMaxi[r.dest!.id] ?? "")));
  const dpStRestant  = dpStHasMaxi ? dpStMaxiSum - assignedWeight : NaN;

  /* Estimation table aggregates */
  const estRows = assignedRows.map(r => {
    const maxiRaw = maxi[r.dest!.id] ?? "";
    const maxiNum = parseRaw(maxiRaw);
    const hasMaxi = !isNaN(maxiNum) && maxiRaw.trim() !== "";
    const pMoy    = r.count > 0 ? r.weight / r.count : 0;
    const pRest   = hasMaxi ? maxiNum - r.weight : NaN;
    const estim   = pMoy > 0 && !isNaN(pRest) ? pRest / pMoy : NaN;
    return { maxiNum: hasMaxi ? maxiNum : NaN, pRest, pMoy, estim };
  });
  const stMaxiSum   = estRows.reduce((s, r) => s + (isNaN(r.maxiNum) ? 0 : r.maxiNum), 0);
  const stHasMaxi   = estRows.some(r => !isNaN(r.maxiNum));
  const stPRestant  = stHasMaxi ? stMaxiSum - assignedWeight : NaN;
  const stPMoyne    = assignedCount > 0 ? assignedWeight / assignedCount : 0;
  const stEstim     = stPMoyne > 0 && !isNaN(stPRestant) ? stPRestant / stPMoyne : NaN;

  const noData = headers.length === 0;

  const [chargOpen,    setChargOpen]    = useState(false);
  const [dechargeOpen, setDechargeOpen] = useState(false);
  const [cumulOpen,    setCumulOpen]    = useState(true);

  /* Cumul table — PREVIOUS_Q / PREVIOUS_P per destination + extra free rows */
  interface ExtraRow { id: string; afterDestId: string; label: string; q: string; p: string; }
  const [prevQ,     setPrevQ]     = useState<Record<string,string>>({});
  const [prevP,     setPrevP]     = useState<Record<string,string>>({});
  const [extraRows, setExtraRows] = useState<ExtraRow[]>([]);
  useEffect(() => {
    setPrevQ(LS.get<Record<string,string>>("tly_prevq", {}));
    setPrevP(LS.get<Record<string,string>>("tly_prevp", {}));
    setExtraRows(LS.get<ExtraRow[]>("tly_extrarows", []));
  }, []);
  const setPrevQVal = (id: string, v: string) => { const n = {...prevQ,[id]:v}; setPrevQ(n); try{localStorage.setItem("tly_prevq",JSON.stringify(n));}catch{} };
  const setPrevPVal = (id: string, v: string) => { const n = {...prevP,[id]:v}; setPrevP(n); try{localStorage.setItem("tly_prevp",JSON.stringify(n));}catch{} };
  const blurPrevQ   = (id: string, v: string) => { const f=fmtInput(v); if(f!==v) setPrevQVal(id,f); };
  const blurPrevP   = (id: string, v: string) => { const f=fmtInput(v); if(f!==v) setPrevPVal(id,f); };
  const updateExtra = (id: string, field: keyof ExtraRow, val: string) => {
    const n = extraRows.map(r => r.id===id ? {...r,[field]:val} : r);
    setExtraRows(n); try{localStorage.setItem("tly_extrarows",JSON.stringify(n));}catch{};
  };
  const blurExtra = (id: string, field: "q"|"p", val: string) => { const f=fmtInput(val); if(f!==val) updateExtra(id,field,f); };
  const removeExtra = (id: string) => {
    const n = extraRows.filter(r => r.id!==id);
    setExtraRows(n); try{localStorage.setItem("tly_extrarows",JSON.stringify(n));}catch{};
  };
  const inpCumul = (extra?: React.CSSProperties): React.CSSProperties => ({
    width: 90, textAlign: "right" as const, fontFamily: MONO, fontSize: 11,
    background: dark ? "#0d0d0d" : "#fff", color: dark ? "#93c5fd" : "#1d4ed8",
    border: `1px solid ${dark?"#333":"#ccc"}`, borderRadius: 3,
    padding: "3px 7px", outline: "none", ...extra,
  });
  const inpLabel = (): React.CSSProperties => ({
    width: 130, fontFamily: MONO, fontSize: 11,
    background: dark ? "#0d0d0d" : "#fff", color: text,
    border: `1px solid ${dark?"#333":"#ccc"}`, borderRadius: 3,
    padding: "3px 7px", outline: "none",
  });

  /* Formatting */
  const fmtWeight = (w: number) =>
    w === 0 ? "—" : w % 1 === 0 ? w.toLocaleString("fr-FR") : w.toLocaleString("fr-FR", { minimumFractionDigits: 2, maximumFractionDigits: 3 });

  const cellStyle = (extra?: React.CSSProperties): React.CSSProperties => ({
    padding: "9px 14px",
    borderBottom: `1px solid ${border}`,
    fontSize: 12,
    fontFamily: MONO,
    color: text,
    whiteSpace: "nowrap",
    ...extra,
  });

  const thStyle = (extra?: React.CSSProperties): React.CSSProperties => ({
    padding: "8px 14px",
    background: hdrBg,
    borderBottom: `2px solid ${border}`,
    fontSize: 10,
    fontFamily: MONO,
    letterSpacing: "0.12em",
    textTransform: "uppercase" as const,
    color: accent,
    textAlign: "left" as const,
    whiteSpace: "nowrap",
    ...extra,
  });

  return (
    <div style={{ padding: "24px 28px", fontFamily: MONO, color: text, background: bg, minHeight: "calc(100vh - 44px)" }}>
      {/* Header */}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 20, flexWrap: "wrap", gap: 12 }}>
        <span style={{ fontSize: 11, letterSpacing: "0.2em", textTransform: "uppercase", color: accent, fontWeight: 600 }}>
          Tally
        </span>
        {!noData && (
          <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
            <span style={{ fontSize: 10, color: muted }}>
              {rows.length} ligne{rows.length !== 1 ? "s" : ""} · {destinations.length} destination{destinations.length !== 1 ? "s" : ""}
            </span>
            {/* Column selector */}
            <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 10, color: muted, fontFamily: MONO }}>
              <span style={{ letterSpacing: "0.08em", textTransform: "uppercase" }}>Col. poids</span>
              <select
                value={weightCol}
                onChange={(e) => {
                  const v = Number(e.target.value);
                  setWeightCol(v);
                  try { localStorage.setItem("tly_weightCol", JSON.stringify(v)); } catch { /* ignore */ }
                }}
                style={{
                  fontFamily: MONO, fontSize: 10, background: surface,
                  color: weightCol >= 0 ? accent : muted,
                  border: `1px solid ${border}`, borderRadius: 3,
                  padding: "3px 7px", cursor: "pointer",
                  outline: "none",
                }}
              >
                <option value={-1}>— aucune —</option>
                {headers.map((h, i) => (
                  <option key={i} value={i}>{h || `Col ${i + 1}`}</option>
                ))}
              </select>
            </label>
          </div>
        )}
      </div>

      {noData ? (
        <div style={{ color: muted, fontSize: 12, marginTop: 40, textAlign: "center" }}>
          Aucune donnée — chargez un fichier dans l'onglet Pointage.
        </div>
      ) : (
        <div style={{ overflowX: "auto", borderRadius: 6, border: `1px solid ${border}`, background: surface }}>
          <div style={{ padding: "8px 14px", borderBottom: `1px solid ${border}`, fontSize: 10, letterSpacing: "0.18em", textTransform: "uppercase", color: accent, fontWeight: 700, fontFamily: MONO }}>
            Today
          </div>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                <th style={thStyle()}>Destination</th>
                <th style={thStyle({ textAlign: "right" })}>Quantité</th>
                {hasWeight && <th style={thStyle({ textAlign: "right" })}>Poids</th>}
              </tr>
            </thead>
            <tbody>
              {/* Assigned destinations */}
              {assignedRows.map(r => (
                <tr key={r.dest!.id} style={{ transition: "background 0.1s" }}
                  onMouseEnter={e => (e.currentTarget.style.background = dark ? "#1a1a1a" : "#fafafa")}
                  onMouseLeave={e => (e.currentTarget.style.background = "")}>
                  <td style={cellStyle()}>
                    <span style={{
                      display: "inline-block",
                      background: r.dest!.color,
                      color: "#fff",
                      padding: "2px 10px",
                      borderRadius: 3,
                      fontSize: 10,
                      letterSpacing: "0.1em",
                      fontWeight: 600,
                    }}>{r.dest!.name}</span>
                  </td>
                  <td style={cellStyle({ textAlign: "right", fontWeight: 600 })}>
                    {r.count.toLocaleString("fr-FR")}
                  </td>
                  {hasWeight && (
                    <td style={cellStyle({ textAlign: "right" })}>{fmtWeight(r.weight)}</td>
                  )}
                </tr>
              ))}
              {/* Sous-total assigné */}
              {assignedRows.length > 0 && (
                <tr style={{ background: dark ? "#0a1a12" : "#f0faf4", borderTop: `2px solid ${border}` }}>
                  <td style={cellStyle({ color: accent, fontSize: 10, letterSpacing: "0.12em", textTransform: "uppercase", fontStyle: "italic" })}>
                    Sous-total assigné
                  </td>
                  <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: accent })}>
                    {assignedCount.toLocaleString("fr-FR")}
                  </td>
                  {hasWeight && (
                    <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: accent })}>{fmtWeight(assignedWeight)}</td>
                  )}
                </tr>
              )}
              {/* Unassigned */}
              {unassigned.count > 0 && (
                <tr style={{ background: dark ? "#111" : "#fdf8f0" }}>
                  <td style={cellStyle({ color: muted, fontStyle: "italic" })}>
                    — non assigné
                  </td>
                  <td style={cellStyle({ textAlign: "right", color: muted })}>{unassigned.count.toLocaleString("fr-FR")}</td>
                  {hasWeight && (
                    <td style={cellStyle({ textAlign: "right", color: muted })}>{fmtWeight(unassigned.weight)}</td>
                  )}
                </tr>
              )}
              {/* Total */}
              <tr style={{ background: dark ? "#0e2016" : "#e8f5ee" }}>
                <td style={cellStyle({ color: accent, letterSpacing: "0.15em", fontWeight: 700, textTransform: "uppercase", fontSize: 10 })}>
                  Total
                </td>
                <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: accent })}>{totalCount.toLocaleString("fr-FR")}</td>
                {hasWeight && (
                  <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: accent })}>{fmtWeight(totalWeight)}</td>
                )}
              </tr>
            </tbody>
          </table>
        </div>
      )}

      {/* ══ CUMUL TABLE ══ */}
      {!noData && assignedRows.length > 0 && (
        <div style={{ marginTop: 20, overflowX: "auto", borderRadius: 6, border: `1px solid ${border}`, background: surface }}>
          <div onClick={() => setCumulOpen(v => !v)} style={{ padding: "8px 14px", borderBottom: cumulOpen ? `1px solid ${border}` : "none", fontSize: 10, letterSpacing: "0.18em", textTransform: "uppercase", color: accent, fontWeight: 700, fontFamily: MONO, cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center", userSelect: "none" }}>
            <span>Cumul</span><span style={{ fontSize: 12 }}>{cumulOpen ? "▲" : "▼"}</span>
          </div>
          {cumulOpen && (
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                <th style={thStyle()}>Destination</th>
                <th style={thStyle({ textAlign: "right" })}>Quantité</th>
                {hasWeight && <th style={thStyle({ textAlign: "right" })}>Poids</th>}
                <th style={thStyle({ textAlign: "right", color: dark ? "#93c5fd" : "#1d4ed8" })}>Previous_Q</th>
                {hasWeight && <th style={thStyle({ textAlign: "right", color: dark ? "#93c5fd" : "#1d4ed8" })}>Previous_P</th>}
                <th style={thStyle({ textAlign: "right", color: dark ? "#6ee7b7" : "#059669" })}>TTL_Q</th>
                {hasWeight && <th style={thStyle({ textAlign: "right", color: dark ? "#6ee7b7" : "#059669" })}>TTL_P</th>}
                <th style={thStyle({ width: 32 })}></th>
              </tr>
            </thead>
            <tbody>
              {assignedRows.map((r) => {
                const pqRaw  = prevQ[r.dest!.id] ?? "";
                const ppRaw  = prevP[r.dest!.id] ?? "";
                const pqNum  = parseRaw(pqRaw);
                const ppNum  = parseRaw(ppRaw);
                const ttlQ   = r.count + (isNaN(pqNum) ? 0 : pqNum);
                const ttlP   = r.weight + (isNaN(ppNum) ? 0 : ppNum);
                const extras = extraRows.filter(e => e.afterDestId === r.dest!.id);
                return (
                  <>
                    <tr key={r.dest!.id}
                      onMouseEnter={e => (e.currentTarget.style.background = dark ? "#1a1a1a" : "#fafafa")}
                      onMouseLeave={e => (e.currentTarget.style.background = "")}>
                      <td style={cellStyle()}>
                        <span style={{ display: "inline-block", background: r.dest!.color, color: "#fff",
                          padding: "2px 10px", borderRadius: 3, fontSize: 10, letterSpacing: "0.1em", fontWeight: 600 }}>
                          {r.dest!.name}
                        </span>
                      </td>
                      <td style={cellStyle({ textAlign: "right", fontWeight: 600 })}>{r.count.toLocaleString("fr-FR")}</td>
                      {hasWeight && <td style={cellStyle({ textAlign: "right" })}>{fmtWeight(r.weight)}</td>}
                      <td style={cellStyle({ padding: "4px 8px", textAlign: "right" })}>
                        <input type="text" inputMode="decimal" value={pqRaw}
                          onChange={e => setPrevQVal(r.dest!.id, e.target.value)}
                          onBlur={e => blurPrevQ(r.dest!.id, e.target.value)}
                          placeholder="—" style={inpCumul()} />
                      </td>
                      {hasWeight && <td style={cellStyle({ padding: "4px 8px", textAlign: "right" })}>
                        <input type="text" inputMode="decimal" value={ppRaw}
                          onChange={e => setPrevPVal(r.dest!.id, e.target.value)}
                          onBlur={e => blurPrevP(r.dest!.id, e.target.value)}
                          placeholder="—" style={inpCumul()} />
                      </td>}
                      <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: dark ? "#6ee7b7" : "#059669" })}>{ttlQ.toLocaleString("fr-FR")}</td>
                      {hasWeight && <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: dark ? "#6ee7b7" : "#059669" })}>{fmtWeight(ttlP)}</td>}
                      <td style={cellStyle({ textAlign: "center", padding: "4px 8px" })}></td>
                    </tr>
                    {extras.map(er => (
                      <tr key={er.id} style={{ background: dark ? "#111" : "#fafdf8" }}>
                        <td style={cellStyle({ padding: "4px 8px" })} colSpan={1}>
                          <input type="text" value={er.label} onChange={e => updateExtra(er.id,"label",e.target.value)}
                            placeholder="libellé…" style={{...inpLabel(), width: 150}} />
                        </td>
                        <td style={cellStyle({ padding: "4px 8px", textAlign: "right" })}>
                          <input type="text" inputMode="decimal" value={er.q} onChange={e => updateExtra(er.id,"q",e.target.value)}
                            onBlur={e => blurExtra(er.id,"q",e.target.value)} placeholder="—" style={inpCumul()} />
                        </td>
                        {hasWeight && <td style={cellStyle({ padding: "4px 8px", textAlign: "right" })}>
                          <input type="text" inputMode="decimal" value={er.p} onChange={e => updateExtra(er.id,"p",e.target.value)}
                            onBlur={e => blurExtra(er.id,"p",e.target.value)} placeholder="—" style={inpCumul()} />
                        </td>}
                        <td style={cellStyle({ padding: "4px 8px" })} colSpan={hasWeight ? 4 : 3}></td>
                        <td style={cellStyle({ padding: "4px 8px", textAlign: "center" })}>
                          <button onClick={() => removeExtra(er.id)} title="Supprimer"
                            style={{ background: "transparent", border: "none", color: muted, cursor: "pointer", fontSize: 13, padding: "0 4px" }}>✕</button>
                        </td>
                      </tr>
                    ))}

                  </>
                );
              })}
              {/* Total */}
              {(() => {
                const totPrevQ = assignedRows.reduce((s,r) => { const pq = parseRaw(prevQ[r.dest!.id]??""); return s + (isNaN(pq)?0:pq); }, 0);
                const totPrevP = assignedRows.reduce((s,r) => { const pp = parseRaw(prevP[r.dest!.id]??""); return s + (isNaN(pp)?0:pp); }, 0);
                const totTTLQ = assignedCount + totPrevQ;
                const totTTLP = assignedWeight + totPrevP;
                return (
                  <tr style={{ background: dark ? "#0e2016" : "#e8f5ee" }}>
                    <td style={cellStyle({ color: accent, fontSize: 10, letterSpacing: "0.15em", textTransform: "uppercase", fontWeight: 700 })}>Total</td>
                    <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: accent })}>{assignedCount.toLocaleString("fr-FR")}</td>
                    {hasWeight && <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: accent })}>{fmtWeight(assignedWeight)}</td>}
                    <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: accent })}>{totPrevQ.toLocaleString("fr-FR")}</td>
                    {hasWeight && <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: accent })}>{fmtWeight(totPrevP)}</td>}
                    <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: dark ? "#6ee7b7" : "#059669" })}>{totTTLQ.toLocaleString("fr-FR")}</td>
                    {hasWeight && <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: dark ? "#6ee7b7" : "#059669" })}>{fmtWeight(totTTLP)}</td>}
                    <td style={cellStyle({})}></td>
                  </tr>
                );
              })()}
            </tbody>
          </table>
          )}
        </div>
      )}

      {/* ══ CHARGEMENT TABLE ══ */}
      {!noData && assignedRows.length > 0 && (
        <div style={{ marginTop: 20, overflowX: "auto", borderRadius: 6, border: `1px solid ${border}`, background: surface }}>
          <div onClick={() => setChargOpen(v => !v)} style={{ padding: "8px 14px", borderBottom: chargOpen ? `1px solid ${border}` : "none", fontSize: 10, letterSpacing: "0.18em", textTransform: "uppercase", color: accent, fontWeight: 700, fontFamily: MONO, cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center", userSelect: "none" }}>
            <span>Chargement</span><span style={{ fontSize: 12 }}>{chargOpen ? "▲" : "▼"}</span>
          </div>
          {chargOpen && <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                <th style={thStyle()}>Destination</th>
                <th style={thStyle({ textAlign: "right" })}>Quantité</th>
                {hasWeight && <th style={thStyle({ textAlign: "right" })}>Poids</th>}
                <th style={thStyle({ textAlign: "right", color: dark ? "#fbbf24" : "#b45309" })}>P_Maxi</th>
                {hasWeight && <th style={thStyle({ textAlign: "right", color: dark ? "#f87171" : "#dc2626" })}>P_Restant</th>}
                {hasWeight && <th style={thStyle({ textAlign: "right" })}>P_Moyne</th>}
                {hasWeight && <th style={thStyle({ textAlign: "right", color: dark ? "#a78bfa" : "#7c3aed" })}>Q_Estimation</th>}
              </tr>
            </thead>
            <tbody>
              {assignedRows.map(r => {
                const maxiRaw  = maxi[r.dest!.id] ?? "";
                const maxiNum  = parseRaw(maxiRaw);
                const hasMaxi  = !isNaN(maxiNum) && maxiRaw.trim() !== "";
                const pMoy     = r.count > 0 ? r.weight / r.count : 0;
                const pRestant = hasMaxi ? maxiNum - r.weight : NaN;
                const estim    = pMoy > 0 && !isNaN(pRestant) ? pRestant / pMoy : NaN;
                const restColor = isNaN(pRestant) ? muted : pRestant >= 0 ? (dark ? "#6ee7b7" : "#059669") : (dark ? "#f87171" : "#dc2626");
                return (
                  <tr key={r.dest!.id}
                    onMouseEnter={e => (e.currentTarget.style.background = dark ? "#1a1a1a" : "#fafafa")}
                    onMouseLeave={e => (e.currentTarget.style.background = "")}>
                    <td style={cellStyle()}>
                      <span style={{ display: "inline-block", background: r.dest!.color, color: "#fff",
                        padding: "2px 10px", borderRadius: 3, fontSize: 10, letterSpacing: "0.1em", fontWeight: 600 }}>
                        {r.dest!.name}
                      </span>
                    </td>
                    <td style={cellStyle({ textAlign: "right", fontWeight: 600 })}>{r.count.toLocaleString("fr-FR")}</td>
                    {hasWeight && <td style={cellStyle({ textAlign: "right" })}>{fmtWeight(r.weight)}</td>}
                    <td style={cellStyle({ textAlign: "right", padding: "4px 8px" })}>
                      <input
                        type="text" inputMode="decimal"
                        value={maxiRaw}
                        onChange={e => setMaxiVal(r.dest!.id, e.target.value)}
                        onBlur={e => { const f = fmtInput(e.target.value); if (f !== e.target.value) setMaxiVal(r.dest!.id, f); }}
                        placeholder="—"
                        style={{
                          width: 90, textAlign: "right", fontFamily: MONO, fontSize: 12,
                          background: dark ? "#0d0d0d" : "#fff",
                          color: dark ? "#fbbf24" : "#b45309",
                          border: `1px solid ${dark ? "#333" : "#ccc"}`, borderRadius: 3,
                          padding: "4px 8px", outline: "none",
                        }}
                      />
                    </td>
                    {hasWeight && (
                      <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: restColor })}>
                        {isNaN(pRestant) ? "—" : fmtWeight(pRestant)}
                      </td>
                    )}
                    {hasWeight && (
                      <td style={cellStyle({ textAlign: "right", color: muted })}>
                        {pMoy > 0 ? fmtWeight(pMoy) : "—"}
                      </td>
                    )}
                    {hasWeight && (
                      <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: dark ? "#a78bfa" : "#7c3aed" })}>
                        {isNaN(estim) ? "—" : Math.round(estim).toLocaleString("fr-FR")}
                      </td>
                    )}
                  </tr>
                );
              })}
              {/* Sous-total estimation */}
              <tr style={{ background: dark ? "#0a1a12" : "#f0faf4", borderTop: `2px solid ${border}` }}>
                <td style={cellStyle({ color: accent, fontSize: 10, letterSpacing: "0.12em", textTransform: "uppercase", fontStyle: "italic" })}>Sous-total assigné</td>
                <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: accent })}>{assignedCount.toLocaleString("fr-FR")}</td>
                {hasWeight && <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: accent })}>{fmtWeight(assignedWeight)}</td>}
                <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: stHasMaxi ? (dark ? "#fbbf24" : "#b45309") : muted })}>{stHasMaxi ? fmtWeight(stMaxiSum) : "—"}</td>
                {hasWeight && <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: isNaN(stPRestant) ? muted : stPRestant >= 0 ? accent : (dark ? "#f87171" : "#dc2626") })}>{isNaN(stPRestant) ? "—" : fmtWeight(stPRestant)}</td>}
                {hasWeight && <td style={cellStyle({ textAlign: "right", color: muted })}>{stPMoyne > 0 ? fmtWeight(stPMoyne) : "—"}</td>}
                {hasWeight && <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: dark ? "#a78bfa" : "#7c3aed" })}>{isNaN(stEstim) ? "—" : Math.round(stEstim).toLocaleString("fr-FR")}</td>}
              </tr>
              {/* Sous-total non assigné */}
              {unassigned.count > 0 && (
                <tr style={{ background: dark ? "#111" : "#fdf8f0" }}>
                  <td style={cellStyle({ color: muted, fontSize: 10, letterSpacing: "0.12em", textTransform: "uppercase", fontStyle: "italic" })}>Sous-total non assigné</td>
                  <td style={cellStyle({ textAlign: "right", color: muted })}>{unassigned.count.toLocaleString("fr-FR")}</td>
                  {hasWeight && <td style={cellStyle({ textAlign: "right", color: muted })}>{fmtWeight(unassigned.weight)}</td>}
                  <td style={cellStyle({ textAlign: "right", color: muted })}>—</td>
                  {hasWeight && <td style={cellStyle({ textAlign: "right", color: muted })}>—</td>}
                  {hasWeight && <td style={cellStyle({ textAlign: "right", color: muted })}>—</td>}
                  {hasWeight && <td style={cellStyle({ textAlign: "right", color: muted })}>—</td>}
                </tr>
              )}
              {/* Total avec non assigné */}
              {(() => {
                return (
                  <tr style={{ background: dark ? "#0e2016" : "#e8f5ee" }}>
                    <td style={cellStyle({ color: accent, fontSize: 10, letterSpacing: "0.15em", textTransform: "uppercase", fontWeight: 700 })}>Total</td>
                    <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: accent })}>{totalCount.toLocaleString("fr-FR")}</td>
                    {hasWeight && <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: accent })}>{fmtWeight(totalWeight)}</td>}
                    <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: stHasMaxi ? (dark ? "#fbbf24" : "#b45309") : muted })}>{stHasMaxi ? fmtWeight(stMaxiSum) : "—"}</td>
                    {hasWeight && <td style={cellStyle({ textAlign: "right", color: muted })}>—</td>}
                    {hasWeight && <td style={cellStyle({ textAlign: "right", color: muted })}>—</td>}
                    {hasWeight && <td style={cellStyle({ textAlign: "right", color: muted })}>—</td>}
                  </tr>
                );
              })()}
            </tbody>
          </table>}
        </div>
      )}

      {/* ══ DÉCHARGEMENT TABLE ══ */}
      {!noData && assignedRows.length > 0 && (
        <div style={{ marginTop: 20, overflowX: "auto", borderRadius: 6, border: `1px solid ${border}`, background: surface }}>
          <div onClick={() => setDechargeOpen(v => !v)} style={{ padding: "8px 14px", borderBottom: dechargeOpen ? `1px solid ${border}` : "none", fontSize: 10, letterSpacing: "0.18em", textTransform: "uppercase", color: accent, fontWeight: 700, fontFamily: MONO, cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center", userSelect: "none" }}>
            <span>Déchargement</span><span style={{ fontSize: 12 }}>{dechargeOpen ? "▲" : "▼"}</span>
          </div>
          {dechargeOpen && <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                <th style={thStyle()}>Destination</th>
                <th style={thStyle({ textAlign: "right" })}>Quantité</th>
                {hasWeight && <th style={thStyle({ textAlign: "right" })}>Poids</th>}
                <th style={thStyle({ textAlign: "right", color: dark ? "#fbbf24" : "#b45309" })}>Q_Maxi</th>
                <th style={thStyle({ textAlign: "right", color: dark ? "#f87171" : "#dc2626" })}>Q_Restant</th>
                <th style={thStyle({ textAlign: "right", color: dark ? "#fbbf24" : "#b45309" })}>P_Maxi</th>
                {hasWeight && <th style={thStyle({ textAlign: "right", color: dark ? "#f87171" : "#dc2626" })}>P_Restant</th>}
              </tr>
            </thead>
            <tbody>
              {assignedRows.map(r => {
                const qRaw   = qMaxi[r.dest!.id] ?? "";
                const qNum   = parseRaw(qRaw);
                const hasQ   = !isNaN(qNum) && qRaw.trim() !== "";
                const qRest  = hasQ ? qNum - r.count : NaN;
                const qColor = isNaN(qRest) ? muted : qRest >= 0 ? (dark ? "#6ee7b7" : "#059669") : (dark ? "#f87171" : "#dc2626");
                const dpRaw  = dPMaxi[r.dest!.id] ?? "";
                const dpNum  = parseRaw(dpRaw);
                const hasDp  = !isNaN(dpNum) && dpRaw.trim() !== "";
                const dpRest = hasDp ? dpNum - r.weight : NaN;
                const dpColor = isNaN(dpRest) ? muted : dpRest >= 0 ? (dark ? "#6ee7b7" : "#059669") : (dark ? "#f87171" : "#dc2626");
                return (
                  <tr key={r.dest!.id}
                    onMouseEnter={e => (e.currentTarget.style.background = dark ? "#1a1a1a" : "#fafafa")}
                    onMouseLeave={e => (e.currentTarget.style.background = "")}>
                    <td style={cellStyle()}>
                      <span style={{ display: "inline-block", background: r.dest!.color, color: "#fff",
                        padding: "2px 10px", borderRadius: 3, fontSize: 10, letterSpacing: "0.1em", fontWeight: 600 }}>
                        {r.dest!.name}
                      </span>
                    </td>
                    <td style={cellStyle({ textAlign: "right", fontWeight: 600 })}>{r.count.toLocaleString("fr-FR")}</td>
                    {hasWeight && <td style={cellStyle({ textAlign: "right" })}>{fmtWeight(r.weight)}</td>}
                    <td style={cellStyle({ textAlign: "right", padding: "4px 8px" })}>
                      <input type="text" inputMode="decimal"
                        value={qRaw}
                        onChange={e => setQMaxiVal(r.dest!.id, e.target.value)}
                        onBlur={e => { const f = fmtInput(e.target.value); if (f !== e.target.value) setQMaxiVal(r.dest!.id, f); }}
                        placeholder="—"
                        style={{
                          width: 90, textAlign: "right", fontFamily: MONO, fontSize: 12,
                          background: dark ? "#0d0d0d" : "#fff",
                          color: dark ? "#fbbf24" : "#b45309",
                          border: `1px solid ${dark ? "#333" : "#ccc"}`, borderRadius: 3,
                          padding: "4px 8px", outline: "none",
                        }}
                      />
                    </td>
                    <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: qColor })}>
                      {isNaN(qRest) ? "—" : qRest.toLocaleString("fr-FR")}
                    </td>
                    <td style={cellStyle({ textAlign: "right", padding: "4px 8px" })}>
                      <input type="text" inputMode="decimal"
                        value={dpRaw}
                        onChange={e => setDPMaxiVal(r.dest!.id, e.target.value)}
                        onBlur={e => { const f = fmtInput(e.target.value); if (f !== e.target.value) setDPMaxiVal(r.dest!.id, f); }}
                        placeholder="—"
                        style={{
                          width: 90, textAlign: "right", fontFamily: MONO, fontSize: 12,
                          background: dark ? "#0d0d0d" : "#fff",
                          color: dark ? "#fbbf24" : "#b45309",
                          border: `1px solid ${dark ? "#333" : "#ccc"}`, borderRadius: 3,
                          padding: "4px 8px", outline: "none",
                        }}
                      />
                    </td>
                    {hasWeight && (
                      <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: dpColor })}>
                        {isNaN(dpRest) ? "—" : fmtWeight(dpRest)}
                      </td>
                    )}
                  </tr>
                );
              })}
              {/* Sous-total */}
              <tr style={{ background: dark ? "#0a1a12" : "#f0faf4", borderTop: `2px solid ${border}` }}>
                <td style={cellStyle({ color: accent, fontSize: 10, letterSpacing: "0.12em", textTransform: "uppercase", fontStyle: "italic" })}>Sous-total assigné</td>
                <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: accent })}>{assignedCount.toLocaleString("fr-FR")}</td>
                {hasWeight && <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: accent })}>{fmtWeight(assignedWeight)}</td>}
                <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: qStHasMaxi ? (dark ? "#fbbf24" : "#b45309") : muted })}>{qStHasMaxi ? qStMaxiSum.toLocaleString("fr-FR") : "—"}</td>
                <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: isNaN(qStRestant) ? muted : qStRestant >= 0 ? accent : (dark ? "#f87171" : "#dc2626") })}>{isNaN(qStRestant) ? "—" : qStRestant.toLocaleString("fr-FR")}</td>
                <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: dpStHasMaxi ? (dark ? "#fbbf24" : "#b45309") : muted })}>{dpStHasMaxi ? dpStMaxiSum.toLocaleString("fr-FR", { minimumFractionDigits: 2, maximumFractionDigits: 3 }) : "—"}</td>
                {hasWeight && <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: isNaN(dpStRestant) ? muted : dpStRestant >= 0 ? accent : (dark ? "#f87171" : "#dc2626") })}>{isNaN(dpStRestant) ? "—" : dpStRestant.toLocaleString("fr-FR", { minimumFractionDigits: 2, maximumFractionDigits: 3 })}</td>}
              </tr>
              {/* Total */}
              {(() => {
                const qTotRest  = qStHasMaxi ? qStMaxiSum - totalCount : NaN;
                const dpTotRest = dpStHasMaxi ? dpStMaxiSum - totalWeight : NaN;
                return (
                  <tr style={{ background: dark ? "#0e2016" : "#e8f5ee" }}>
                    <td style={cellStyle({ color: accent, fontSize: 10, letterSpacing: "0.15em", textTransform: "uppercase", fontWeight: 700 })}>Total</td>
                    <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: accent })}>{totalCount.toLocaleString("fr-FR")}</td>
                    {hasWeight && <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: accent })}>{fmtWeight(totalWeight)}</td>}
                    <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: qStHasMaxi ? (dark ? "#fbbf24" : "#b45309") : muted })}>{qStHasMaxi ? qStMaxiSum.toLocaleString("fr-FR") : "—"}</td>
                    <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: isNaN(qTotRest) ? muted : qTotRest >= 0 ? accent : (dark ? "#f87171" : "#dc2626") })}>{isNaN(qTotRest) ? "—" : qTotRest.toLocaleString("fr-FR")}</td>
                    <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: dpStHasMaxi ? (dark ? "#fbbf24" : "#b45309") : muted })}>{dpStHasMaxi ? dpStMaxiSum.toLocaleString("fr-FR", { minimumFractionDigits: 2, maximumFractionDigits: 3 }) : "—"}</td>
                    {hasWeight && <td style={cellStyle({ textAlign: "right", fontWeight: 700, color: isNaN(dpTotRest) ? muted : dpTotRest >= 0 ? accent : (dark ? "#f87171" : "#dc2626") })}>{isNaN(dpTotRest) ? "—" : dpTotRest.toLocaleString("fr-FR", { minimumFractionDigits: 2, maximumFractionDigits: 3 })}</td>}
                  </tr>
                );
              })()}
            </tbody>
          </table>}
        </div>
      )}
    </div>
  );
}
