/* ─── Plan ───────────────────────────────────────────────────────────────── */

const MONO = "'IBM Plex Mono', 'Fira Mono', monospace";

interface PlanProps { dark: boolean; }

export default function Plan({ dark }: PlanProps) {
  const bg      = dark ? "#111111" : "#ffffff";
  const surface = dark ? "#1c1c1c" : "#f8fafc";
  const text    = dark ? "#e2e8f0" : "#1e293b";
  const accent  = dark ? "#e2e8f0" : "#6b7280";
  const muted   = dark ? "#64748b" : "#9ca3af";
  const border  = dark ? "#2e2e2e" : "#e2e8f0";

  return (
    <div style={{ minHeight: "100vh", background: bg, color: text, fontFamily: MONO, padding: "32px 24px" }}>
      <div style={{
        maxWidth: 900, margin: "0 auto",
        background: surface, border: `1px solid ${border}`,
        borderRadius: 8, padding: "40px 32px", textAlign: "center",
      }}>
        <div style={{ fontSize: 48, marginBottom: 16 }}>📋</div>
        <h1 style={{ color: accent, fontSize: 22, letterSpacing: "0.12em", textTransform: "uppercase", margin: "0 0 12px" }}>
          Plan
        </h1>
        <p style={{ color: muted, fontSize: 13, letterSpacing: "0.08em" }}>
          Module en cours de développement
        </p>
      </div>
    </div>
  );
}
