/* ─── Plan ───────────────────────────────────────────────────────────────── */
import { useState, useEffect } from "react";
import { dechargementColor } from "./dechargementColor";
import { useWindowWidth } from "./useWindowWidth";

const MONO = "'IBM Plex Mono', 'Fira Mono', monospace";

interface HangarItem { code: string; weight: number | null; position?: string; dechargement?: string; }
interface HangarLine { id: string; name: string; items: HangarItem[]; }

const LS = {
  get<T>(key: string, def: T): T {
    try { const v = localStorage.getItem(key); return v ? (JSON.parse(v) as T) : def; } catch { return def; }
  },
};

const LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");
const ODDS  = Array.from({ length: 26 }, (_, i) => String(i * 2 + 1)); /* 1,3,5…51 */
const EVENS = Array.from({ length: 25 }, (_, i) => String(i * 2 + 2)); /* 2,4,6…50 */

const LINE_BG_DARK  = ["#1c1305","#051c09","#06091c","#1a0512","#051a1a","#14190a","#190a05"];
const LINE_BG_LIGHT = ["#fff7ed","#f0fdf4","#eff6ff","#fdf4ff","#f0fdfa","#fefce8","#fff1f2"];

const fmtW = (w: number | null) =>
  w === null ? "—" : w.toLocaleString("fr-FR", { minimumFractionDigits: 2, maximumFractionDigits: 3 });

interface PlanProps { dark: boolean; }

export default function Plan({ dark }: PlanProps) {
  const vw       = useWindowWidth();
  const isMobile = vw < 640;

  const [f12Mode, setF12Mode] = useState<boolean>(() => {
    try { const v = localStorage.getItem("hgr_f12mode"); return v ? JSON.parse(v) : false; } catch { return false; }
  });

  /* persist + keep in sync with HANGAR tab */
  useEffect(() => {
    try { localStorage.setItem("hgr_f12mode", JSON.stringify(f12Mode)); } catch {}
  }, [f12Mode]);

  /* accent: orange for QAA, red for F12 — mirrors hangar.tsx line 153 */
  const accentQAA = "#f97316";
  const accentF12 = dark ? "#fca5a5" : "#ef4444";
  const accent    = f12Mode ? accentF12 : accentQAA;

  const bg     = dark ? "#0f0800" : "#fff5f2";
  const text   = dark ? "#fed7aa" : "#9f1239";
  const muted  = dark ? "#a16207" : "#be185d";
  const border = dark ? "#3d1a00" : "#fecdd3";

  const pad = isMobile ? "14px 10px" : "24px 28px";

  const [lines, setLines] = useState<HangarLine[]>([]);

  useEffect(() => {
    const load = () => {
      const f12 = (() => { try { return JSON.parse(localStorage.getItem("hgr_f12mode") ?? "false") as boolean; } catch { return false; } })();
      setF12Mode(f12);
      setLines(LS.get<HangarLine[]>(`hgr_lines2_${f12 ? "F12" : "QAA"}`, []));
    };
    load();
    const id = setInterval(load, 2000);
    return () => clearInterval(id);
  }, []);

  /* ── Recherche multi-produits ─────────────────────────────── */  const [searchOpen, setSearchOpen] = useState(false);  const [searchRaw, setSearchRaw] = useState("");

  /* parse : séparateurs = saut de ligne, virgule, point-virgule, espace */
  const searchCodes = searchRaw.trim()
    ? searchRaw.split(/[\n,;\s]+/).map((s) => s.trim()).filter(Boolean)
    : [];
  const searchSet = new Set(searchCodes.map((c) => c.toUpperCase()));

  /* résultats : pour chaque code cherché, trouver toutes les occurrences */
  type SearchHit = { code: string; lineName: string; position: string | undefined; dechargement: string | undefined };
  const searchResults: Map<string, SearchHit[]> = new Map();
  if (searchSet.size > 0) {
    for (const line of lines) {
      for (const item of line.items) {
        const up = item.code.toUpperCase();
        if (searchSet.has(up)) {
          const hits = searchResults.get(up) ?? [];
          hits.push({ code: item.code, lineName: line.name, position: item.position, dechargement: item.dechargement });
          searchResults.set(up, hits);
        }
      }
    }
  }

  /* map letter → items */
  const lineMap = new Map<string, HangarItem[]>();
  for (const l of lines) lineMap.set(l.name.toUpperCase(), l.items);

  const CHIP = isMobile ? 28 : 20;
  const GAP  = 2;
  const half = (CHIP + GAP) / 2;

  return (
    <div style={{ fontFamily: MONO, color: text, background: bg, minHeight: "calc(100vh - 44px)" }}>
      {/* Animation clignotante pour les chips trouvés */}
      <style>{`@keyframes plan-blink { 0%,49%{opacity:1} 50%,100%{opacity:0.15} }`}</style>
      <div style={{ maxWidth: 900, margin: "0 auto", padding: pad, boxSizing: "border-box", overflowX: "hidden" }}>

        {/* Header */}
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 20, flexWrap: "wrap" }}>
          <h2 style={{ fontSize: isMobile ? 17 : 13, letterSpacing: isMobile ? 0 : "0.18em",
            textTransform: isMobile ? "none" : "uppercase", color: accent, margin: 0 }}>
            📋 Plan hangar
          </h2>

          {/* QAA / F12 pill toggle */}
          <button
            onClick={() => setF12Mode((v) => !v)}
            title={f12Mode ? "Site F12" : "Site QAA"}
            style={{
              display: "flex", alignItems: "center", gap: 0,
              padding: 0, border: `1px solid ${accent}`,
              borderRadius: 20, overflow: "hidden", cursor: "pointer",
              background: "transparent", flexShrink: 0,
              fontSize: isMobile ? 12 : 10, fontFamily: MONO,
            }}
          >
            <span style={{
              padding: isMobile ? "7px 11px" : "4px 9px",
              background: !f12Mode ? accentQAA+"33" : "transparent",
              color: !f12Mode ? accentQAA : muted,
              fontWeight: !f12Mode ? 700 : 400,
              transition: "background 0.2s, color 0.2s",
            }}>QAA</span>
            <span style={{
              padding: isMobile ? "7px 11px" : "4px 9px",
              background: f12Mode ? accentF12+"33" : "transparent",
              color: f12Mode ? accentF12 : muted,
              fontWeight: f12Mode ? 700 : 400,
              transition: "background 0.2s, color 0.2s",
            }}>F12</span>
          </button>

          <span style={{ fontSize: isMobile ? 12 : 10, color: muted }}>
            — {lines.length} ligne{lines.length !== 1 ? "s" : ""} · {lines.reduce((s, l) => s + l.items.length, 0)} bobines
          </span>

          {/* Bouton toggle recherche */}
          <button
            onClick={() => setSearchOpen((v) => !v)}
            title={searchOpen ? "Fermer la recherche" : "Rechercher des produits"}
            style={{
              marginLeft: "auto", fontFamily: MONO, cursor: "pointer",
              fontSize: isMobile ? 12 : 10, padding: isMobile ? "6px 12px" : "3px 10px",
              borderRadius: 20, border: `1px solid ${searchOpen || searchCodes.length > 0 ? accent : border}`,
              background: searchOpen ? accent+"22" : "transparent",
              color: searchOpen || searchCodes.length > 0 ? accent : muted,
              fontWeight: searchOpen ? 700 : 400,
              transition: "background 0.2s, color 0.2s",
            }}>
            {searchCodes.length > 0 ? `🔍 ${searchCodes.length} code${searchCodes.length > 1 ? "s" : ""}` : "🔍 Recherche"}
          </button>
        </div>

        {/* ── Panneau de recherche (rétractable) ───────────────── */}
        {searchOpen && (
        <div style={{ marginBottom: 16 }}>
          <textarea
            value={searchRaw}
            onChange={(e) => setSearchRaw(e.target.value)}
            placeholder={"Rechercher des produits\nUn code par ligne, ou séparés par virgule / espace"}
            rows={isMobile ? 3 : 2}
            style={{
              width: "100%", boxSizing: "border-box",
              fontFamily: MONO, fontSize: isMobile ? 13 : 11,
              padding: isMobile ? "10px 12px" : "7px 10px",
              borderRadius: 6, border: `1px solid ${searchCodes.length > 0 ? accent : border}`,
              background: dark ? "#1a0e00" : "#fff",
              color: text, outline: "none", resize: "vertical",
              transition: "border-color 0.2s",
            }}
          />

          {/* Résultats */}
          {searchCodes.length > 0 && (
            <div style={{ marginTop: 8, display: "flex", flexDirection: "column", gap: 4 }}>
              {searchCodes.map((code) => {
                const up   = code.toUpperCase();
                const hits = searchResults.get(up) ?? [];
                return (
                  <div key={code} style={{
                    display: "flex", alignItems: "flex-start", gap: 8, flexWrap: "wrap",
                    padding: isMobile ? "6px 10px" : "4px 8px",
                    borderRadius: 5,
                    border: `1px solid ${hits.length > 0 ? accent : (dark ? "#3a3a3a" : "#e5e5e5")}`,
                    background: hits.length > 0 ? accent+"11" : (dark ? "#111" : "#fafafa"),
                  }}>
                    <span style={{ fontFamily: MONO, fontSize: isMobile ? 12 : 10, fontWeight: 700,
                      color: hits.length > 0 ? accent : muted, flexShrink: 0, minWidth: 80 }}>
                      {code}
                    </span>
                    {hits.length === 0 ? (
                      <span style={{ fontSize: isMobile ? 11 : 9, color: muted, fontStyle: "italic" }}>introuvable</span>
                    ) : hits.map((h, i) => (
                      <span key={i} style={{
                        fontFamily: MONO, fontSize: isMobile ? 11 : 9,
                        padding: "1px 6px", borderRadius: 10,
                        background: accent+"33", color: accent, fontWeight: 700,
                      }}>
                        {h.lineName}{h.position ? String(parseInt(h.position)).padStart(2, "0") : "?"}
                        {h.dechargement ? ` · ${h.dechargement}` : ""}
                      </span>
                    ))}
                  </div>
                );
              })}
            </div>
          )}
        </div>
        )}

        {/* Lignes par lettre */}
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {LETTERS.map((letter, li) => {
            const items = lineMap.get(letter) ?? [];
            const occupied = new Set(items.map((it) => it.position).filter(Boolean) as string[]);
            const vide     = new Set(items.filter((it) => it.code.startsWith("∅")).map((it) => it.position).filter(Boolean) as string[]);
            const isEmpty  = items.length === 0;
            const lineBg   = dark ? LINE_BG_DARK[li % LINE_BG_DARK.length] : LINE_BG_LIGHT[li % LINE_BG_LIGHT.length];
            const hasMatch = searchSet.size > 0 && items.some((it) => searchSet.has(it.code.toUpperCase()));

            const chip = (s: string) => {
              const isOcc   = occupied.has(s);
              const isVide  = vide.has(s);
              const item    = items.find((it) => it.position === s);
              const dech    = item?.dechargement ?? "";
              const dc      = isOcc && !isVide && dech ? dechargementColor(dech, dark) : null;
              const isFound = isOcc && searchSet.size > 0 && !!item && searchSet.has(item.code.toUpperCase());
              const label   = isOcc
                ? (isVide ? "∅" : (item?.code?.slice(0, 4) ?? s))
                : s;
              return (
                <div key={s}
                  title={isOcc ? `${letter}${s.padStart(2,"0")} : ${item?.code ?? "?"} ${item?.weight != null ? fmtW(item.weight)+" kg" : ""}${dech ? " · "+dech : ""}` : `${letter}${s.padStart(2,"0")} libre`}
                  style={{
                    fontFamily: MONO,
                    fontSize: isMobile ? 9 : 7,
                    width: CHIP, height: CHIP, flexShrink: 0,
                    borderRadius: 3, lineHeight: `${CHIP}px`, textAlign: "center",
                    overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
                    outline: isFound ? `2px solid ${accent}` : "none",
                    outlineOffset: 1,
                    animation: isFound ? "plan-blink 0.8s step-start infinite" : "none",
                    border: isOcc
                      ? (dc ? `1px solid ${dc.border}` : (isVide ? `1px solid #6b7280` : `1px solid ${accent}`))
                      : `1px solid ${dark ? "#2a2a2a" : "#e5e5e5"}`,
                    background: isFound
                      ? accent+"55"
                      : (isOcc
                        ? (dc ? dc.bg : (isVide ? (dark ? "#1a1a1a" : "#f3f4f6") : (dark ? "#2a1a00" : "#fde68a")))
                        : (dark ? "#111" : "#f9f9f9")),
                    color: isFound ? "#fff" : (isOcc ? (dc ? dc.color : (isVide ? muted : "#92400e")) : (dark ? "#333" : "#ccc")),
                    fontWeight: isOcc ? 700 : 400,
                    cursor: "default",
                    boxShadow: isFound ? `0 0 6px ${accent}88` : "none",
                  }}>
                  {label}
                </div>
              );
            };

            return (
              <div key={letter} style={{
                borderRadius: 6,
                border: `1px solid ${hasMatch ? accent : (isEmpty ? border : (dark ? "#3d1a00" : "#fecdd3"))}`,
                background: isEmpty ? (dark ? "#111" : "#fafafa") : lineBg,
                padding: "6px 8px",
                opacity: isEmpty && searchSet.size > 0 ? 0.25 : (isEmpty ? 0.45 : 1),
                boxShadow: hasMatch ? `0 0 0 1px ${accent}44` : "none",
              }}>
                <div style={{ display: "flex", alignItems: "flex-start", gap: 8 }}>
                  {/* Letter badge */}
                  <div style={{
                    width: isMobile ? 28 : 22, height: isMobile ? 28 : 22, flexShrink: 0,
                    borderRadius: 4, background: isEmpty ? (dark?"#1a1a1a":"#e5e5e5") : accent,
                    color: isEmpty ? muted : "#fff",
                    display: "flex", alignItems: "center", justifyContent: "center",
                    fontSize: isMobile ? 13 : 10, fontWeight: 700,
                  }}>{letter}</div>

                  {/* Chip grid */}
                  <div style={{ flex: 1, minWidth: 0, overflowX: "auto" }}>
                    <div style={{ display: "inline-block", minWidth: "max-content" }}>
                      <div style={{ display: "flex", gap: GAP, marginLeft: half, marginBottom: GAP }}>
                        {EVENS.map(chip)}
                      </div>
                      <div style={{ display: "flex", gap: GAP }}>
                        {ODDS.map(chip)}
                      </div>
                    </div>
                  </div>

                  {/* Count */}
                  <div style={{ flexShrink: 0, fontSize: isMobile ? 11 : 9, color: isEmpty ? muted : text,
                    minWidth: 28, textAlign: "right", paddingTop: 2 }}>
                    {isEmpty ? "—" : `${items.length}`}
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

