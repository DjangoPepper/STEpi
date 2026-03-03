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

  const bg     = dark ? "#0f0800" : "#fff5f2";
  const text   = dark ? "#fed7aa" : "#9f1239";
  const muted  = dark ? "#a16207" : "#be185d";
  const border = dark ? "#3d1a00" : "#fecdd3";
  const accent = dark ? "#f97316" : "#f97316";

  const pad = isMobile ? "14px 10px" : "24px 28px";

  const [lines, setLines] = useState<HangarLine[]>([]);

  useEffect(() => {
    const load = () => setLines(LS.get<HangarLine[]>("hgr_lines2", []));
    load();
    const id = setInterval(load, 2000);
    return () => clearInterval(id);
  }, []);

  /* map letter → items */
  const lineMap = new Map<string, HangarItem[]>();
  for (const l of lines) lineMap.set(l.name.toUpperCase(), l.items);

  const CHIP = isMobile ? 28 : 20;
  const GAP  = 2;
  const half = (CHIP + GAP) / 2;

  return (
    <div style={{ fontFamily: MONO, color: text, background: bg, minHeight: "calc(100vh - 44px)" }}>
      <div style={{ maxWidth: 900, margin: "0 auto", padding: pad, boxSizing: "border-box", overflowX: "hidden" }}>

        {/* Header */}
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 20 }}>
          <h2 style={{ fontSize: isMobile ? 17 : 13, letterSpacing: isMobile ? 0 : "0.18em",
            textTransform: isMobile ? "none" : "uppercase", color: accent, margin: 0 }}>
            📋 Plan hangar
          </h2>
          <span style={{ fontSize: isMobile ? 12 : 10, color: muted }}>
            — {lines.length} ligne{lines.length !== 1 ? "s" : ""} · {lines.reduce((s, l) => s + l.items.length, 0)} bobines
          </span>
        </div>

        {/* One row per letter */}
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {LETTERS.map((letter, li) => {
            const items = lineMap.get(letter) ?? [];
            const occupied = new Set(items.map((it) => it.position).filter(Boolean) as string[]);
            const vide     = new Set(items.filter((it) => it.code.startsWith("∅")).map((it) => it.position).filter(Boolean) as string[]);
            const isEmpty  = items.length === 0;
            const lineBg   = dark ? LINE_BG_DARK[li % LINE_BG_DARK.length] : LINE_BG_LIGHT[li % LINE_BG_LIGHT.length];

            const chip = (s: string) => {
              const isOcc  = occupied.has(s);
              const isVide = vide.has(s);
              const item   = items.find((it) => it.position === s);
              const dech   = item?.dechargement ?? "";
              const dc     = isOcc && !isVide && dech ? dechargementColor(dech, dark) : null;
              const label  = isOcc
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
                    border: isOcc
                      ? (dc ? `1px solid ${dc.border}` : (isVide ? `1px solid #6b7280` : `1px solid ${accent}`))
                      : `1px solid ${dark ? "#2a2a2a" : "#e5e5e5"}`,
                    background: isOcc
                      ? (dc ? dc.bg : (isVide ? (dark ? "#1a1a1a" : "#f3f4f6") : (dark ? "#2a1a00" : "#fde68a")))
                      : (dark ? "#111" : "#f9f9f9"),
                    color: isOcc ? (dc ? dc.color : (isVide ? muted : "#92400e")) : (dark ? "#333" : "#ccc"),
                    fontWeight: isOcc ? 700 : 400,
                    cursor: "default",
                  }}>
                  {label}
                </div>
              );
            };

            return (
              <div key={letter} style={{
                borderRadius: 6,
                border: `1px solid ${isEmpty ? border : (dark ? "#3d1a00" : "#fecdd3")}`,
                background: isEmpty ? (dark ? "#111" : "#fafafa") : lineBg,
                padding: "6px 8px",
                opacity: isEmpty ? 0.45 : 1,
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

