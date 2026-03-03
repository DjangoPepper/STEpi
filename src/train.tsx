/* ─── Train ──────────────────────────────────────────────────────────────── */

const MONO = "'IBM Plex Mono', 'Fira Mono', monospace";

interface TrainProps { dark: boolean; }

export default function Train({ dark }: TrainProps) {
  const bg      = dark ? "#0f0c00" : "#fffbeb";
  const surface = dark ? "#1a1500" : "#fff";
  const text    = dark ? "#fef08a" : "#78350f";
  const accent  = dark ? "#fbbf24" : "#d97706";
  const muted   = dark ? "#a16207" : "#a16207";
  const border  = dark ? "#3d2e00" : "#fde68a";

  return (
    <div style={{ minHeight: "100vh", background: bg, color: text, fontFamily: MONO, padding: "32px 24px" }}>
      <div style={{
        maxWidth: 900, margin: "0 auto",
        background: surface, border: `1px solid ${border}`,
        borderRadius: 8, padding: "40px 32px", textAlign: "center",
      }}>
        <div style={{ fontSize: 48, marginBottom: 16 }}>🚂</div>
        <h1 style={{ color: accent, fontSize: 22, letterSpacing: "0.12em", textTransform: "uppercase", margin: "0 0 12px" }}>
          Train
        </h1>
        <p style={{ color: muted, fontSize: 13, letterSpacing: "0.08em" }}>
          Module en cours de développement
        </p>
      </div>
    </div>
  );
}
