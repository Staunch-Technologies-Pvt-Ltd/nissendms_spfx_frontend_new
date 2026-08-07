/* eslint-disable no-bitwise */
/**
 * vesselImagePool.ts
 * --------------------------------------------------------------------------
 * Generates highly-detailed, realistic side-profile vessel SVG illustrations.
 * Each vessel type gets a unique ship silhouette matching the screenshot style:
 * - Realistic hull shape with bow, stern, waterline
 * - Type-specific deck equipment (containers, cranes, pipes, A-frame)
 * - Multiple distinct color palettes assigned by vessel ID hash
 * - Ocean water with waves, sky with clouds and sun glow
 */

/* eslint-disable max-len */

interface VesselPalette {
  id: number;
  skyTop: string;
  skyBot: string;
  waterTop: string;
  waterBot: string;
  hull: string;
  hullShade: string;
  keel: string;
  superstructure: string;
  funnel: string;
  funnelAccent: string;
  deckColor: string;
}

const PALETTES: VesselPalette[] = [
  // 0: Deep Royal Blue & Crimson (Ocean Star)
  {
    id: 0,
    skyTop: '#b3d4f0', skyBot: '#dceefa',
    waterTop: '#1565a8', waterBot: '#0a3d72',
    hull: '#1a365d', hullShade: '#0f2942', keel: '#c53030',
    superstructure: '#f8fafc', funnel: '#dd6b20', funnelAccent: '#1a365d',
    deckColor: '#243650',
  },
  // 1: Emerald Green & Charcoal (Sea Breeze)
  {
    id: 1,
    skyTop: '#a7f3d0', skyBot: '#ecfdf5',
    waterTop: '#059669', waterBot: '#047857',
    hull: '#064e3b', hullShade: '#022c22', keel: '#9f1239',
    superstructure: '#f8fafc', funnel: '#eab308', funnelAccent: '#064e3b',
    deckColor: '#16181b',
  },
  // 2: Midnight Cobalt & Wine Red (Blue Horizon)
  {
    id: 2,
    skyTop: '#cbe5ff', skyBot: '#f0f9ff',
    waterTop: '#1d4ed8', waterBot: '#1e3a8a',
    hull: '#1e40af', hullShade: '#172554', keel: '#881337',
    superstructure: '#ffffff', funnel: '#06b6d4', funnelAccent: '#1e40af',
    deckColor: '#0d1640',
  },
  // 3: Deep Burgundy & Bronze (Pacific Dawn)
  {
    id: 3,
    skyTop: '#fed7aa', skyBot: '#fff7ed',
    waterTop: '#0284c7', waterBot: '#0f766e',
    hull: '#701a75', hullShade: '#4a044e', keel: '#92400e',
    superstructure: '#fdf4ff', funnel: '#f59e0b', funnelAccent: '#701a75',
    deckColor: '#4a044e',
  },
  // 4: Vibrant Coral Red & Navy (Atlantic Wave)
  {
    id: 4,
    skyTop: '#dbeafe', skyBot: '#eff6ff',
    waterTop: '#2563eb', waterBot: '#1e40af',
    hull: '#dc2626', hullShade: '#991b1b', keel: '#450a0a',
    superstructure: '#ffffff', funnel: '#1d4ed8', funnelAccent: '#1e40af',
    deckColor: '#1d2b3a',
  },
  // 5: Sunset Amber & Charcoal (Golden Pearl)
  {
    id: 5,
    skyTop: '#fef08a', skyBot: '#fefce8',
    waterTop: '#0369a1', waterBot: '#0c4a6e',
    hull: '#78350f', hullShade: '#451a03', keel: '#b91c1c',
    superstructure: '#fffbeb', funnel: '#fbbf24', funnelAccent: '#78350f',
    deckColor: '#0f1c28',
  },
  // 6: Steel Slate & Flame Orange (Titan Guard)
  {
    id: 6,
    skyTop: '#cbd5e1', skyBot: '#f1f5f9',
    waterTop: '#0f766e', waterBot: '#134e4a',
    hull: '#334155', hullShade: '#1e293b', keel: '#be123c',
    superstructure: '#f8fafc', funnel: '#f97316', funnelAccent: '#334155',
    deckColor: '#0f2414',
  },
  // 7: Royal Violet & Gold
  {
    id: 7,
    skyTop: '#e9d5ff', skyBot: '#faf5ff',
    waterTop: '#4338ca', waterBot: '#312e81',
    hull: '#4c1d95', hullShade: '#2e1065', keel: '#d97706',
    superstructure: '#ffffff', funnel: '#8b5cf6', funnelAccent: '#4c1d95',
    deckColor: '#2e1065',
  },
  // 8: Deep Aqua Teal & Coral
  {
    id: 8,
    skyTop: '#ccfbf1', skyBot: '#f0fdf4',
    waterTop: '#0d9488', waterBot: '#115e59',
    hull: '#115e59', hullShade: '#042f2e', keel: '#e11d48',
    superstructure: '#f0fdf4', funnel: '#14b8a6', funnelAccent: '#115e59',
    deckColor: '#042f2e',
  },
  // 9: Terracotta Copper & Flame
  {
    id: 9,
    skyTop: '#ffedd5', skyBot: '#fff7ed',
    waterTop: '#0284c7', waterBot: '#075985',
    hull: '#c2410c', hullShade: '#7c2d12', keel: '#292524',
    superstructure: '#fff7ed', funnel: '#fb923c', funnelAccent: '#c2410c',
    deckColor: '#1c1917',
  },
];

function getHash(str: string): number {
  let h = 5381;
  for (let i = 0; i < str.length; i++) {
    h = (((h << 5) + h) ^ str.charCodeAt(i)) >>> 0;
  }
  return h;
}

/**
 * Generates a detailed side-profile ship SVG.
 * ViewBox: 400 x 200  (wider canvas for more detail)
 */
export function generateVesselSVG(name: string, type: string, id: string): string {
  const hash = getHash(id + name + type);

  // Determine palette index
  let palIdx = hash % PALETTES.length;
  // Exact overrides for screenshot vessels
  const exact: Record<string, number> = {
    'Ocean Star': 0, 'Sea Breeze': 1, 'Blue Horizon': 2,
    'Pacific Dawn': 3, 'Atlantic Wave': 4, 'Golden Pearl': 5,
  };
  if (exact[name] !== undefined) palIdx = exact[name];

  const p = PALETTES[palIdx];
  const t = (type || 'Container Ship').toLowerCase();

  /* ── Deck equipment based on vessel type ──────────────────────────── */

  let deckEquipment = '';

  if (t.indexOf('container') !== -1) {
    // Stacked multicolored containers
    const colors = ['#e53935','#1e88e5','#43a047','#fb8c00','#8e24aa','#039be5','#c0ca33','#6d4c41'];
    let stacks = '';
    for (let col = 0; col < 14; col++) {
      const cx = 42 + col * 18;
      const layers = 1 + ((col + hash) % 3); // 1–3 containers tall
      for (let layer = 0; layer < layers; layer++) {
        const cy = 88 - layer * 13;
        const cc = colors[(col * 2 + layer + hash) % colors.length];
        stacks += `<rect x="${cx}" y="${cy}" width="16" height="12" fill="${cc}" rx="0.8" stroke="#111" stroke-width="0.4"/>`;
        // Rib line on container
        stacks += `<line x1="${cx+2}" y1="${cy+3}" x2="${cx+14}" y2="${cy+3}" stroke="rgba(0,0,0,0.15)" stroke-width="0.4"/>`;
        stacks += `<line x1="${cx+2}" y1="${cy+7}" x2="${cx+14}" y2="${cy+7}" stroke="rgba(0,0,0,0.15)" stroke-width="0.4"/>`;
      }
    }
    deckEquipment = stacks;

    // Aft superstructure (right side)
    deckEquipment += `
      <rect x="298" y="52" width="38" height="48" fill="${p.superstructure}" rx="2" stroke="#c0c8d0" stroke-width="0.5"/>
      <rect x="300" y="56" width="34" height="5" fill="#64b5f6"/>
      <rect x="302" y="64" width="7" height="7" fill="#90a4ae" rx="1"/>
      <rect x="313" y="64" width="7" height="7" fill="#90a4ae" rx="1"/>
      <rect x="324" y="64" width="7" height="7" fill="#90a4ae" rx="1"/>
      <rect x="302" y="74" width="7" height="6" fill="#90a4ae" rx="1"/>
      <rect x="313" y="74" width="7" height="6" fill="#90a4ae" rx="1"/>
      <rect x="324" y="74" width="7" height="6" fill="#90a4ae" rx="1"/>
      <rect x="310" y="28" width="14" height="24" fill="${p.funnel}" rx="2"/>
      <rect x="313" y="22" width="8" height="6" fill="${p.funnelAccent}" rx="1"/>
      <rect x="308" y="26" width="18" height="2" fill="${p.funnelAccent}"/>
    `;
  } else if (t.indexOf('bulk') !== -1) {
    // 4 bulk cargo holds with deck cranes
    for (let h2 = 0; h2 < 4; h2++) {
      const hx = 50 + h2 * 55;
      deckEquipment += `
        <rect x="${hx}" y="82" width="40" height="10" fill="#334155" rx="2"/>
        <rect x="${hx+2}" y="78" width="36" height="4" fill="#475569" rx="1"/>
        <rect x="${hx+2}" y="76" width="36" height="2" fill="#64748b"/>
        <!-- crane mast -->
        <rect x="${hx+18}" y="44" width="4" height="34" fill="#e2e8f0"/>
        <!-- crane boom -->
        <line x1="${hx+20}" y1="44" x2="${hx+42}" y2="74" stroke="#e2e8f0" stroke-width="2"/>
        <line x1="${hx+20}" y1="44" x2="${hx-2}" y2="74" stroke="#e2e8f0" stroke-width="2"/>
        <!-- crane wire -->
        <line x1="${hx+42}" y1="74" x2="${hx+42}" y2="80" stroke="#bbb" stroke-width="1"/>
      `;
    }
    deckEquipment += `
      <rect x="302" y="50" width="40" height="48" fill="${p.superstructure}" rx="2" stroke="#c0c8d0" stroke-width="0.5"/>
      <rect x="304" y="54" width="36" height="6" fill="#64b5f6"/>
      <rect x="306" y="64" width="7" height="7" fill="#90a4ae" rx="1"/>
      <rect x="317" y="64" width="7" height="7" fill="#90a4ae" rx="1"/>
      <rect x="328" y="64" width="7" height="7" fill="#90a4ae" rx="1"/>
      <rect x="306" y="75" width="7" height="5" fill="#90a4ae" rx="1"/>
      <rect x="317" y="75" width="7" height="5" fill="#90a4ae" rx="1"/>
      <rect x="314" y="24" width="14" height="26" fill="${p.funnel}" rx="2"/>
      <rect x="317" y="18" width="8" height="6" fill="${p.funnelAccent}" rx="1"/>
    `;
  } else if (t.indexOf('tanker') !== -1 || t.indexOf('chemical') !== -1) {
    // Long flat pipe manifold deck
    deckEquipment += `
      <rect x="44" y="84" width="240" height="4" fill="#b0bec5" rx="1"/>
      <rect x="44" y="82" width="240" height="2" fill="#ffd54f"/>
    `;
    for (let m = 0; m < 8; m++) {
      const mx = 54 + m * 30;
      deckEquipment += `
        <rect x="${mx}" y="72" width="10" height="14" fill="#607d8b" rx="1"/>
        <rect x="${mx+1}" y="66" width="8" height="6" fill="#90a4ae" rx="0.5"/>
        <circle cx="${mx+5}" cy="64" r="3" fill="${m % 2 === 0 ? '#ef5350' : '#42a5f5'}"/>
        <line x1="${mx+5}" y1="72" x2="${mx+5}" y2="80" stroke="#90a4ae" stroke-width="1.5"/>
      `;
    }
    deckEquipment += `
      <rect x="304" y="46" width="42" height="52" fill="${p.superstructure}" rx="2" stroke="#c0c8d0" stroke-width="0.5"/>
      <rect x="306" y="50" width="38" height="6" fill="#64b5f6"/>
      <rect x="308" y="60" width="8" height="8" fill="#90a4ae" rx="1"/>
      <rect x="320" y="60" width="8" height="8" fill="#90a4ae" rx="1"/>
      <rect x="332" y="60" width="8" height="8" fill="#90a4ae" rx="1"/>
      <rect x="308" y="72" width="8" height="6" fill="#90a4ae" rx="1"/>
      <rect x="320" y="72" width="8" height="6" fill="#90a4ae" rx="1"/>
      <rect x="332" y="72" width="8" height="6" fill="#90a4ae" rx="1"/>
      <rect x="316" y="20" width="16" height="26" fill="${p.funnel}" rx="2"/>
      <rect x="319" y="14" width="10" height="6" fill="${p.funnelAccent}" rx="1"/>
      <rect x="314" y="44" width="20" height="2" fill="${p.funnelAccent}"/>
    `;
  } else if (t.indexOf('offshore') !== -1 || t.indexOf('support') !== -1) {
    // Forward superstructure + open aft deck
    deckEquipment += `
      <!-- Forward superstructure -->
      <rect x="42" y="38" width="58" height="60" fill="${p.superstructure}" rx="2" stroke="#c0c8d0" stroke-width="0.5"/>
      <rect x="44" y="44" width="54" height="8" fill="#64b5f6"/>
      <rect x="46" y="56" width="9" height="9" fill="#90a4ae" rx="1"/>
      <rect x="59" y="56" width="9" height="9" fill="#90a4ae" rx="1"/>
      <rect x="72" y="56" width="9" height="9" fill="#90a4ae" rx="1"/>
      <rect x="85" y="56" width="9" height="9" fill="#90a4ae" rx="1"/>
      <rect x="46" y="70" width="9" height="8" fill="#90a4ae" rx="1"/>
      <rect x="59" y="70" width="9" height="8" fill="#90a4ae" rx="1"/>
      <rect x="72" y="70" width="9" height="8" fill="#90a4ae" rx="1"/>
      <!-- Helideck -->
      <rect x="28" y="32" width="60" height="6" fill="#546e7a"/>
      <circle cx="58" cy="35" r="3" fill="rgba(255,255,255,0.25)" stroke="#fff" stroke-width="0.8"/>
      <!-- Exhaust stack -->
      <rect x="104" y="36" width="14" height="28" fill="${p.funnel}" rx="2"/>
      <rect x="107" y="30" width="8" height="6" fill="${p.funnelAccent}" rx="1"/>
      <!-- Aft crane / A-frame -->
      <rect x="155" y="60" width="4" height="38" fill="#e2e8f0"/>
      <line x1="157" y1="60" x2="185" y2="90" stroke="#e2e8f0" stroke-width="2.5"/>
      <line x1="157" y1="60" x2="129" y2="90" stroke="#e2e8f0" stroke-width="2.5"/>
      <!-- Aft deck equipment -->
      <rect x="175" y="72" width="28" height="16" fill="#d97706" rx="2"/>
      <rect x="210" y="68" width="36" height="20" fill="#1565c0" rx="2"/>
      <rect x="255" y="74" width="22" height="14" fill="#cc2424" rx="2"/>
      <rect x="125" y="82" width="150" height="8" fill="#455a64"/>
    `;
  } else if (t.indexOf('gas') !== -1) {
    // LNG spherical tanks on deck
    for (let g = 0; g < 3; g++) {
      const gx = 80 + g * 75;
      deckEquipment += `
        <circle cx="${gx}" cy="72" r="26" fill="#e3f2fd" stroke="#90caf9" stroke-width="1.5"/>
        <ellipse cx="${gx}" cy="72" rx="24" ry="22" fill="#bbdefb"/>
        <ellipse cx="${gx-6}" cy="66" rx="8" ry="6" fill="rgba(255,255,255,0.35)"/>
        <line x1="${gx}" y1="46" x2="${gx}" y2="96" stroke="#42a5f5" stroke-width="1" opacity="0.5"/>
        <line x1="${gx-24}" y1="72" x2="${gx+24}" y2="72" stroke="#42a5f5" stroke-width="1" opacity="0.5"/>
      `;
    }
    deckEquipment += `
      <rect x="310" y="48" width="40" height="50" fill="${p.superstructure}" rx="2" stroke="#c0c8d0" stroke-width="0.5"/>
      <rect x="312" y="52" width="36" height="6" fill="#64b5f6"/>
      <rect x="314" y="62" width="8" height="8" fill="#90a4ae" rx="1"/>
      <rect x="326" y="62" width="8" height="8" fill="#90a4ae" rx="1"/>
      <rect x="338" y="62" width="8" height="8" fill="#90a4ae" rx="1"/>
      <rect x="320" y="22" width="14" height="26" fill="${p.funnel}" rx="2"/>
      <rect x="323" y="16" width="8" height="6" fill="${p.funnelAccent}" rx="1"/>
    `;
  } else {
    // General Cargo — midship cranes + aft house
    deckEquipment += `
      <rect x="48" y="78" width="50" height="14" fill="#334155" rx="2"/>
      <rect x="50" y="76" width="46" height="2" fill="#475569"/>
      <rect x="120" y="78" width="50" height="14" fill="#334155" rx="2"/>
      <rect x="122" y="76" width="46" height="2" fill="#475569"/>
      <!-- Fore crane -->
      <rect x="90" y="38" width="5" height="38" fill="#e2e8f0"/>
      <line x1="92.5" y1="38" x2="120" y2="74" stroke="#e2e8f0" stroke-width="2"/>
      <line x1="92.5" y1="38" x2="54" y2="74" stroke="#e2e8f0" stroke-width="2"/>
      <line x1="120" y1="74" x2="120" y2="82" stroke="#aaa" stroke-width="1.5"/>
      <!-- Mid crane -->
      <rect x="170" y="38" width="5" height="38" fill="#e2e8f0"/>
      <line x1="172.5" y1="38" x2="200" y2="74" stroke="#e2e8f0" stroke-width="2"/>
      <line x1="172.5" y1="38" x2="138" y2="74" stroke="#e2e8f0" stroke-width="2"/>
      <line x1="200" y1="74" x2="200" y2="82" stroke="#aaa" stroke-width="1.5"/>
    `;
    deckEquipment += `
      <rect x="300" y="50" width="40" height="48" fill="${p.superstructure}" rx="2" stroke="#c0c8d0" stroke-width="0.5"/>
      <rect x="302" y="54" width="36" height="6" fill="#64b5f6"/>
      <rect x="304" y="64" width="7" height="7" fill="#90a4ae" rx="1"/>
      <rect x="315" y="64" width="7" height="7" fill="#90a4ae" rx="1"/>
      <rect x="326" y="64" width="7" height="7" fill="#90a4ae" rx="1"/>
      <rect x="304" y="75" width="7" height="6" fill="#90a4ae" rx="1"/>
      <rect x="315" y="75" width="7" height="6" fill="#90a4ae" rx="1"/>
      <rect x="314" y="24" width="14" height="26" fill="${p.funnel}" rx="2"/>
      <rect x="317" y="18" width="8" height="6" fill="${p.funnelAccent}" rx="1"/>
    `;
  }

  /* ── Complete SVG ─────────────────────────────────────────────────── */
  const svg = `
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 200" width="100%" height="100%" style="display:block">
  <defs>
    <linearGradient id="sky_${palIdx}" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0%" stop-color="${p.skyTop}"/>
      <stop offset="100%" stop-color="${p.skyBot}"/>
    </linearGradient>
    <linearGradient id="water_${palIdx}" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0%" stop-color="${p.waterTop}"/>
      <stop offset="100%" stop-color="${p.waterBot}"/>
    </linearGradient>
    <linearGradient id="hull_${palIdx}" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0%" stop-color="${p.hull}"/>
      <stop offset="100%" stop-color="${p.hullShade}"/>
    </linearGradient>
    <linearGradient id="superHull_${palIdx}" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0%" stop-color="${p.hull}" stop-opacity="0.7"/>
      <stop offset="100%" stop-color="${p.hull}"/>
    </linearGradient>
  </defs>

  <!-- ── Sky ── -->
  <rect x="0" y="0" width="400" height="126" fill="url(#sky_${palIdx})"/>

  <!-- Sun glow -->
  <circle cx="360" cy="34" r="22" fill="#fff9e0" opacity="0.55"/>
  <circle cx="360" cy="34" r="14" fill="#ffe082" opacity="0.7"/>

  <!-- Clouds -->
  <ellipse cx="60" cy="28" rx="32" ry="10" fill="#fff" opacity="0.65"/>
  <ellipse cx="80" cy="24" rx="20" ry="8" fill="#fff" opacity="0.55"/>
  <ellipse cx="195" cy="22" rx="38" ry="11" fill="#fff" opacity="0.5"/>
  <ellipse cx="218" cy="18" rx="24" ry="9" fill="#fff" opacity="0.45"/>
  <ellipse cx="310" cy="30" rx="22" ry="8" fill="#fff" opacity="0.4"/>

  <!-- Horizon haze -->
  <rect x="0" y="118" width="400" height="8" fill="${p.waterTop}" opacity="0.35"/>

  <!-- ── Ocean ── -->
  <rect x="0" y="124" width="400" height="76" fill="url(#water_${palIdx})"/>

  <!-- Wave ripples -->
  <path d="M0,130 Q25,127 50,130 Q75,133 100,130 Q125,127 150,130 Q175,133 200,130 Q225,127 250,130 Q275,133 300,130 Q325,127 350,130 Q375,133 400,130" stroke="#fff" stroke-width="0.9" fill="none" opacity="0.4"/>
  <path d="M0,138 Q30,135 60,138 Q90,141 120,138 Q150,135 180,138 Q210,141 240,138 Q270,135 300,138 Q330,141 360,138 Q380,141 400,138" stroke="#fff" stroke-width="0.6" fill="none" opacity="0.28"/>
  <path d="M0,148 Q40,145 80,148 Q120,151 160,148 Q200,145 240,148 Q280,151 320,148 Q360,145 400,148" stroke="#fff" stroke-width="0.5" fill="none" opacity="0.18"/>
  <path d="M0,158 Q50,155 100,158 Q150,161 200,158 Q250,155 300,158 Q350,161 400,158" stroke="#fff" stroke-width="0.4" fill="none" opacity="0.12"/>

  <!-- ── Deck equipment (drawn BEHIND the hull shape) ── -->
  ${deckEquipment}

  <!-- ── Main Hull shape ── -->
  <!-- Hull body -->
  <path d="
    M 24,98
    L 360,98
    Q 378,98 385,108
    L 390,124
    L 14,124
    Q 10,124 16,114
    Z
  " fill="url(#hull_${palIdx})"/>

  <!-- Hull highlight strip (deck rail top) -->
  <rect x="24" y="96" width="338" height="4" fill="${p.deckColor}" rx="1"/>

  <!-- Sharp bow -->
  <path d="M 360,98 L 392,110 L 388,124 L 360,124 Z" fill="${p.hull}"/>
  <path d="M 382,100 L 397,116 L 395,124 L 388,124 L 392,110 Z" fill="${p.hullShade}"/>

  <!-- Stern curve -->
  <path d="M 24,98 L 12,108 L 14,124 L 24,124 Z" fill="${p.hullShade}"/>

  <!-- Waterline / keel stripe -->
  <rect x="14" y="118" width="378" height="6" fill="${p.keel}" rx="0"/>

  <!-- Portholes row -->
  <circle cx="44" cy="110" r="2.2" fill="#fff" opacity="0.7"/>
  <circle cx="60" cy="110" r="2.2" fill="#fff" opacity="0.7"/>
  <circle cx="76" cy="110" r="2.2" fill="#fff" opacity="0.7"/>
  <circle cx="92" cy="110" r="2.2" fill="#fff" opacity="0.7"/>
  <circle cx="108" cy="110" r="2.2" fill="#fff" opacity="0.7"/>
  <circle cx="124" cy="110" r="2.2" fill="#fff" opacity="0.7"/>
  <circle cx="140" cy="110" r="2.2" fill="#fff" opacity="0.7"/>
  <circle cx="156" cy="110" r="2.2" fill="#fff" opacity="0.7"/>
  <circle cx="172" cy="110" r="2.2" fill="#fff" opacity="0.7"/>
  <circle cx="188" cy="110" r="2.2" fill="#fff" opacity="0.7"/>
  <circle cx="204" cy="110" r="2.2" fill="#fff" opacity="0.7"/>
  <circle cx="220" cy="110" r="2.2" fill="#fff" opacity="0.7"/>
  <circle cx="236" cy="110" r="2.2" fill="#fff" opacity="0.7"/>
  <circle cx="252" cy="110" r="2.2" fill="#fff" opacity="0.7"/>
  <circle cx="268" cy="110" r="2.2" fill="#fff" opacity="0.7"/>
  <circle cx="284" cy="110" r="2.2" fill="#fff" opacity="0.7"/>

  <!-- Hull shadow band -->
  <rect x="14" y="114" width="378" height="4" fill="rgba(0,0,0,0.12)"/>

  <!-- Water reflection / wake -->
  <path d="M 390,120 Q 397,126 393,134 Q 388,140 380,136" stroke="#fff" stroke-width="1.4" fill="none" opacity="0.55"/>
  <path d="M 14,120 Q 7,126 11,134 Q 16,140 24,136" stroke="#fff" stroke-width="1.4" fill="none" opacity="0.55"/>
  <path d="M 0,140 Q 40,136 80,140 Q 120,144 160,140" stroke="#fff" stroke-width="0.6" fill="none" opacity="0.2"/>
  <path d="M 240,140 Q 280,136 320,140 Q 360,144 400,140" stroke="#fff" stroke-width="0.6" fill="none" opacity="0.2"/>
</svg>`.trim();

  return 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
}

export function resolveImgUrl(val: any, name?: string, type?: string, id?: string): string {
  if (typeof val === 'string' && val.length > 10 && val.indexOf('data:image/svg+xml') === 0) return val;
  return generateVesselSVG(name || 'Vessel', type || 'Container Ship', id || 'v0');
}

export function getVesselImageForId(id: string, name?: string, type?: string): string {
  return generateVesselSVG(name || 'Vessel', type || 'Container Ship', id || 'v0');
}

export function pickRandomVesselImage(vesselType?: string): string {
  const randId = 'v_' + Math.floor(Math.random() * 100000);
  return generateVesselSVG('New Vessel', vesselType || 'Container Ship', randId);
}

export function getVesselImagePool(): string[] {
  return [
    generateVesselSVG('Ocean Star', 'Container Ship', 'v1'),
    generateVesselSVG('Sea Breeze', 'Bulk Carrier', 'v2'),
    generateVesselSVG('Blue Horizon', 'Oil Tanker', 'v3'),
    generateVesselSVG('Pacific Dawn', 'General Cargo', 'v4'),
    generateVesselSVG('Atlantic Wave', 'Offshore Support', 'v5'),
    generateVesselSVG('Golden Pearl', 'Container Ship', 'v6'),
  ];
}

export const VESSEL_IMAGE_POOL: string[] = getVesselImagePool();

export async function loadVesselImagePool(): Promise<string[]> {
  return getVesselImagePool();
}
