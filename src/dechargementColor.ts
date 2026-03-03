/* ── Couleur par déchargement ─────────────────────────────────────────────
   Retourne un jeu de couleurs cohérent (fond, bordure, texte) basé sur la
   valeur de déchargement. Même valeur → même couleur, partout dans l'appli.
   ──────────────────────────────────────────────────────────────────────── */

// [darkBg, darkBorder, darkText, lightBg, lightBorder, lightText]
const PALETTE: [string, string, string, string, string, string][] = [
  ["#1a0a00", "#f97316", "#fdba74", "#fff7ed", "#f97316", "#c2410c"], // orange
  ["#001a2e", "#0ea5e9", "#7dd3fc", "#f0f9ff", "#0ea5e9", "#0369a1"], // sky
  ["#001a06", "#22c55e", "#86efac", "#f0fdf4", "#22c55e", "#15803d"], // green
  ["#160a28", "#a78bfa", "#c4b5fd", "#f5f3ff", "#8b5cf6", "#6d28d9"], // violet
  ["#200010", "#ec4899", "#f9a8d4", "#fdf2f8", "#ec4899", "#be185d"], // pink
  ["#001a18", "#2dd4bf", "#99f6e4", "#f0fdfa", "#14b8a6", "#0f766e"], // teal
  ["#1a1400", "#eab308", "#fde047", "#fefce8", "#ca8a04", "#a16207"], // yellow
  ["#1a0000", "#f87171", "#fca5a5", "#fff1f2", "#f87171", "#b91c1c"], // red
  ["#06071a", "#818cf8", "#c7d2fe", "#eef2ff", "#6366f1", "#4338ca"], // indigo
  ["#0a1a00", "#84cc16", "#bef264", "#f7fee7", "#65a30d", "#3f6212"], // lime
];

function hash(s: string): number {
  let h = 0;
  for (let i = 0; i < s.length; i++) h = (Math.imul(31, h) + s.charCodeAt(i)) >>> 0;
  return h;
}

export interface DechargementColors {
  bg: string;
  border: string;
  color: string;
}

export function dechargementColor(dechargement: string, dark: boolean): DechargementColors {
  const idx = hash(dechargement.trim().toLowerCase()) % PALETTE.length;
  const p = PALETTE[idx];
  return dark
    ? { bg: p[0], border: p[1], color: p[2] }
    : { bg: p[3], border: p[4], color: p[5] };
}
