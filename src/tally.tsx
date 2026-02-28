interface TallyProps { dark: boolean; }

export default function Tally({ dark }: TallyProps) {
  const bg     = dark ? "#0d0d0d" : "#f5f5f5";
  const text   = dark ? "#e8e8e0" : "#1a1a1a";
  const accent = dark ? "#6ee7b7" : "#059669";
  const muted  = dark ? "#555"    : "#888";
  return (
    <div style={{ padding: 40, fontFamily: "'IBM Plex Mono', monospace", color: text, background: bg, minHeight: "calc(100vh - 44px)" }}>
      <h2 style={{ fontSize: 14, letterSpacing: "0.15em", textTransform: "uppercase", color: accent }}>
        Tally
      </h2>
      <p style={{ color: muted, fontSize: 12 }}>Page tally — à implémenter.</p>
    </div>
  );
}
