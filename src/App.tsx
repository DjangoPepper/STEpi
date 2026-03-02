import { useState } from "react";
import ExcelCleaner from "./ExcelCleaner";
import Pointage from "./pointage";
import Scane from "./scane";
import Tally from "./tally";
import Hangar from "./hangar";
import type { PointageData } from "./types";
import { useWindowWidth } from "./useWindowWidth";

type TabId = "cleaner" | "pointage" | "scan" | "tally" | "hangar";

const TABS: { id: TabId; label: string }[] = [
  { id: "cleaner", label: "Cleaner" },
  { id: "pointage",      label: "Pointage" },
  { id: "scan",          label: "Scan" },
  { id: "tally",         label: "Tally" },
  { id: "hangar",        label: "Hangar" },
];

const MONO = "'IBM Plex Mono', 'Fira Mono', monospace";

export default function App() {
  const [active, setActive]           = useState<TabId>("cleaner");
  const [dark, setDark]               = useState(true);
  const [pointageData, setPointageData] = useState<PointageData | null>(null);
  const vw      = useWindowWidth();
  const isMobile = vw < 640;

  const SURFACE = dark ? "#141414" : "#f0f0f0";
  const BORDER  = dark ? "#2a2a2a" : "#d0d0d0";
  const ACCENT  = dark ? "#6ee7b7" : "#059669";
  const MUTED   = dark ? "#555"    : "#888";
  const TEXT    = dark ? "#e8e8e0" : "#1a1a1a";
  const BG      = dark ? "#0d0d0d" : "#f5f5f5";

  const handleSendToPointage = (data: PointageData) => {
    setPointageData(data);
    setActive("pointage");
  };

  return (
    <div style={{ minHeight: "100vh", background: BG, color: TEXT, fontFamily: MONO, display: "flex", flexDirection: "column" }}>
      {/* ── Tab bar ─────────────────────────────────────────────── */}
      <nav style={{
        display: "flex", alignItems: "center", gap: 0,
        borderBottom: `1px solid ${BORDER}`,
        background: SURFACE, paddingLeft: isMobile ? 4 : 16, paddingRight: 4,
        overflowX: "auto", WebkitOverflowScrolling: "touch" as "touch",
        scrollbarWidth: "none" as "none",
      }}>
        <div style={{ display: "flex", flex: 1, alignItems: "flex-end", flexShrink: 0 }}>
          {TABS.map((tab) => {
            const isActive = tab.id === active;
            return (
              <button
                key={tab.id}
                onClick={() => setActive(tab.id)}
                style={{
                  fontFamily: isMobile ? "-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif" : MONO,
                  fontSize: isMobile ? 15 : 11, letterSpacing: isMobile ? 0 : "0.1em",
                  textTransform: isMobile ? "none" : "uppercase", padding: isMobile ? "16px 14px 14px" : "12px 22px 10px",
                  border: "none", whiteSpace: "nowrap", flexShrink: 0,
                  borderBottom: isActive ? `2px solid ${ACCENT}` : "2px solid transparent",
                  background: "transparent", color: isActive ? ACCENT : MUTED,
                  cursor: "pointer", transition: "color 0.15s, border-color 0.15s",
                  outline: "none", marginBottom: -1,
                }}
                onMouseEnter={(e) => { if (!isActive) (e.currentTarget as HTMLButtonElement).style.color = TEXT; }}
                onMouseLeave={(e) => { if (!isActive) (e.currentTarget as HTMLButtonElement).style.color = MUTED; }}
              >
                {tab.label}
              </button>
            );
          })}
        </div>

        {/* Dark/light toggle — always visible */}
        <button
          onClick={() => setDark((d) => !d)}
          style={{
            fontFamily: isMobile ? "-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif" : MONO,
            fontSize: isMobile ? 15 : 11, letterSpacing: isMobile ? 0 : "0.12em",
            textTransform: isMobile ? "none" : "uppercase", padding: isMobile ? "12px 16px" : "8px 14px",
            background: "transparent", border: `1px solid ${BORDER}`,
            borderRadius: 3, color: MUTED, cursor: "pointer", marginLeft: 4, flexShrink: 0,
            transition: "color 0.15s, border-color 0.15s",
          }}
        >
          {dark ? (isMobile ? "☀" : "☀ Clair") : (isMobile ? "☾" : "☾ Sombre")}
        </button>
      </nav>

      {/* ── Page content ────────────────────────────────────────── */}
      <main style={{ flex: 1, overflow: "auto" }}>
        {active === "cleaner" && (
          <ExcelCleaner
            dark={dark}
            onDarkToggle={() => setDark((d) => !d)}
            onSendToPointage={handleSendToPointage}
          />
        )}
        {active === "pointage" && <Pointage dark={dark} initialData={pointageData} />}
        {active === "scan"     && <Scane dark={dark} />}
        {active === "tally"    && <Tally dark={dark} />}
        {active === "hangar"   && <Hangar dark={dark} />}
      </main>
    </div>
  );
}
