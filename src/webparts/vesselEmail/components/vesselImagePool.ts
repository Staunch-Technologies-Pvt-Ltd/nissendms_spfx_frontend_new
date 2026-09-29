/* eslint-disable no-bitwise */
/**
 * vesselImagePool.ts
 * --------------------------------------------------------------------------
 * Generates the vessel card banner artwork shown at the top of each card in
 * the Vessels grid.
 *
 * Design language: "Night Harbor" — a detailed, realistic side-profile
 * vessel silhouette (real hull, bow, stern, waterline, type-specific deck
 * equipment — the same level of ship detail as before), but rendered as a
 * moody, professional night scene instead of a flat cartoon-daylight
 * illustration: a deep tone-on-tone gradient, the hull picked out with a
 * thin accent rim-light, lit bridge/porthole windows that actually glow
 * (a blurred duplicate layer behind the crisp shape), an ambient floodlight
 * glow behind the superstructure, and a soft reflection of the lit hull on
 * the water below. No rainbow container colours, no sun/clouds — this is
 * meant to read like a premium fleet-management product, client-presentable
 * out of the box.
 *
 * Each vessel still gets a distinct look: the palette (dark base + one
 * signature glow colour) is chosen deterministically from the vessel's
 * id/name/type hash, and the deck equipment is chosen by vessel type
 * (container / bulk / tanker / gas / offshore / general cargo), so the
 * fleet grid reads as varied but consistently professional.
 */

/* eslint-disable max-len */

interface VesselPalette {
  id: number;
  name: string;
  skyTop: string;     // background gradient, top
  skyBot: string;     // background gradient, bottom (also the "night water" base)
  hull: string;       // hull top tone
  hullShade: string;  // hull bottom / shadow tone
  superstructure: string; // deckhouse / funnel base tone (slightly lighter neutral)
  glow: string;       // the one signature accent — rim light, lit windows, glow
}

// Eight dark, tone-on-tone "night harbor" palettes — a deep neutral base
// with exactly one signature glow colour each, so the fleet grid stays
// varied without ever looking like a rainbow cartoon.
const PALETTES: VesselPalette[] = [
  { id: 0, name: 'Navy Steel',    skyTop: '#0a1626', skyBot: '#0e2038', hull: '#26374d', hullShade: '#101b29', superstructure: '#33455c', glow: '#38bdf8' }, // cyan
  { id: 1, name: 'Charcoal Ember',skyTop: '#181310', skyBot: '#211a15', hull: '#332a23', hullShade: '#17120e', superstructure: '#3e3229', glow: '#f2a65a' }, // amber
  { id: 2, name: 'Deep Teal',     skyTop: '#06201e', skyBot: '#0a2e2a', hull: '#1d3934', hullShade: '#0a1917', superstructure: '#24443e', glow: '#2dd4bf' }, // teal
  { id: 3, name: 'Onyx Frost',    skyTop: '#12151f', skyBot: '#1a2030', hull: '#2b3145', hullShade: '#12151f', superstructure: '#333a52', glow: '#93c5fd' }, // ice blue
  { id: 4, name: 'Midnight Copper', skyTop: '#190f0a', skyBot: '#241609', hull: '#3a2717', hullShade: '#170e08', superstructure: '#452e1b', glow: '#e08a4b' }, // copper
  { id: 5, name: 'Graphite Gold', skyTop: '#141416', skyBot: '#1c1c1f', hull: '#2f2f34', hullShade: '#131315', superstructure: '#39393f', glow: '#e3c565' }, // gold
  { id: 6, name: 'Forest Night',  skyTop: '#0a1811', skyBot: '#0e2519', hull: '#1c3626', hullShade: '#09150e', superstructure: '#24422f', glow: '#7be3a8' }, // mint
  { id: 7, name: 'Storm Indigo',  skyTop: '#0e1224', skyBot: '#151a34', hull: '#262c4d', hullShade: '#0f1226', superstructure: '#2e355c', glow: '#a5b4fc' }, // periwinkle
];

function getHash(str: string): number {
  let h = 5381;
  for (let i = 0; i < str.length; i++) {
    h = (((h << 5) + h) ^ str.charCodeAt(i)) >>> 0;
  }
  return h;
}

/**
 * Generates a detailed, realistic side-profile ship silhouette rendered as a
 * moody, glowing "night harbor" scene.
 * ViewBox: 400 x 200.
 */
export function generateVesselSVG(name: string, type: string, id: string): string {
  const hash = getHash(id + name + type);
  const p = PALETTES[hash % PALETTES.length];
  const t = (type || 'Container Ship').toLowerCase();
  const uid = `${p.id}_${hash % 99991}`;

  const glowFilter = `glow_${uid}`;
  const softGlowFilter = `softGlow_${uid}`;
  const skyGrad = `sky_${uid}`;
  const waterGrad = `water_${uid}`;
  const hullGrad = `hull_${uid}`;
  const fadeMask = `fade_${uid}`;

  /* ── A pair of "lit window" rects: a crisp opaque tile + a blurred glow
     copy behind it, so bridge/cabin windows actually read as lit at night. */
  const litWindow = (x: number, y: number, w: number, h: number, rx = 1): string => `
    <rect x="${x - 1.5}" y="${y - 1.5}" width="${w + 3}" height="${h + 3}" rx="${rx + 1}" fill="${p.glow}" opacity="0.55" filter="url(#${softGlowFilter})"/>
    <rect x="${x}" y="${y}" width="${w}" height="${h}" rx="${rx}" fill="${p.glow}" opacity="0.92"/>
  `;

  const litPort = (cx: number, cy: number, r = 2.3): string => `
    <circle cx="${cx}" cy="${cy}" r="${r + 1.6}" fill="${p.glow}" opacity="0.45" filter="url(#${softGlowFilter})"/>
    <circle cx="${cx}" cy="${cy}" r="${r}" fill="${p.glow}" opacity="0.9"/>
  `;

  /* ── Deck equipment, by vessel type. Dark neutral masses (recognisable
     silhouettes) with a thin accent rim on load-bearing edges — no rainbow
     fills. Superstructure + funnel + bridge windows are shared below. ── */
  let deckEquipment = '';

  if (t.indexOf('container') !== -1) {
    for (let col = 0; col < 14; col++) {
      const cx = 42 + col * 18;
      const layers = 1 + ((col + hash) % 3);
      for (let layer = 0; layer < layers; layer++) {
        const cy = 88 - layer * 13;
        deckEquipment += `<rect x="${cx}" y="${cy}" width="16" height="12" rx="0.8" fill="${p.hullShade}"/>`;
        deckEquipment += `<rect x="${cx}" y="${cy}" width="16" height="2" fill="${p.glow}" opacity="0.14"/>`;
      }
    }
  } else if (t.indexOf('bulk') !== -1) {
    for (let h2 = 0; h2 < 4; h2++) {
      const hx = 50 + h2 * 55;
      deckEquipment += `
        <rect x="${hx}" y="82" width="40" height="10" rx="2" fill="${p.hullShade}"/>
        <rect x="${hx + 2}" y="78" width="36" height="3" fill="${p.superstructure}"/>
        <rect x="${hx + 18}" y="44" width="3" height="34" fill="${p.superstructure}"/>
        <line x1="${hx + 19.5}" y1="44" x2="${hx + 42}" y2="74" stroke="${p.glow}" stroke-width="1" opacity="0.4"/>
        <line x1="${hx + 19.5}" y1="44" x2="${hx - 2}" y2="74" stroke="${p.glow}" stroke-width="1" opacity="0.4"/>
      `;
    }
  } else if (t.indexOf('tanker') !== -1 || t.indexOf('chemical') !== -1) {
    deckEquipment += `<rect x="44" y="84" width="240" height="4" rx="1" fill="${p.hullShade}"/>`;
    for (let m = 0; m < 8; m++) {
      const mx = 54 + m * 30;
      deckEquipment += `
        <rect x="${mx}" y="72" width="10" height="14" rx="1" fill="${p.hullShade}"/>
        <rect x="${mx + 1}" y="66" width="8" height="6" rx="0.5" fill="${p.superstructure}"/>
        ${m % 3 === 0 ? litPort(mx + 5, 64, 2.4) : `<circle cx="${mx + 5}" cy="64" r="2.4" fill="${p.superstructure}"/>`}
      `;
    }
  } else if (t.indexOf('offshore') !== -1 || t.indexOf('support') !== -1) {
    deckEquipment += `
      <rect x="42" y="38" width="58" height="60" rx="2" fill="${p.superstructure}"/>
      ${litWindow(46, 56, 9, 9)} ${litWindow(59, 56, 9, 9)} ${litWindow(72, 56, 9, 9)} ${litWindow(85, 56, 9, 9)}
      <rect x="46" y="70" width="9" height="8" rx="1" fill="${p.hullShade}"/>
      <rect x="59" y="70" width="9" height="8" rx="1" fill="${p.hullShade}"/>
      <rect x="72" y="70" width="9" height="8" rx="1" fill="${p.hullShade}"/>
      <rect x="28" y="32" width="60" height="6" fill="${p.hullShade}"/>
      <circle cx="58" cy="35" r="3" fill="none" stroke="${p.glow}" stroke-width="1" opacity="0.7"/>
      <rect x="104" y="36" width="14" height="28" rx="2" fill="${p.hullShade}"/>
      <rect x="106" y="34" width="10" height="3" fill="${p.glow}" opacity="0.45"/>
      <rect x="155" y="60" width="3" height="38" fill="${p.superstructure}"/>
      <line x1="156.5" y1="60" x2="185" y2="90" stroke="${p.glow}" stroke-width="1" opacity="0.4"/>
      <line x1="156.5" y1="60" x2="129" y2="90" stroke="${p.glow}" stroke-width="1" opacity="0.4"/>
      <rect x="175" y="72" width="28" height="16" rx="2" fill="${p.hullShade}"/>
      <rect x="210" y="68" width="36" height="20" rx="2" fill="${p.hullShade}"/>
      <rect x="255" y="74" width="22" height="14" rx="2" fill="${p.hullShade}"/>
      <rect x="125" y="82" width="150" height="6" fill="${p.hullShade}"/>
    `;
  } else if (t.indexOf('gas') !== -1) {
    for (let g = 0; g < 3; g++) {
      const gx = 80 + g * 75;
      deckEquipment += `
        <circle cx="${gx}" cy="72" r="26" fill="${p.hullShade}"/>
        <circle cx="${gx}" cy="72" r="26" fill="none" stroke="${p.glow}" stroke-width="1" opacity="0.5"/>
        <ellipse cx="${gx - 7}" cy="64" rx="9" ry="6" fill="${p.glow}" opacity="0.12"/>
      `;
    }
  } else {
    deckEquipment += `
      <rect x="48" y="78" width="50" height="14" rx="2" fill="${p.hullShade}"/>
      <rect x="120" y="78" width="50" height="14" rx="2" fill="${p.hullShade}"/>
      <rect x="90" y="38" width="4" height="38" fill="${p.superstructure}"/>
      <line x1="92" y1="38" x2="120" y2="74" stroke="${p.glow}" stroke-width="1" opacity="0.4"/>
      <line x1="92" y1="38" x2="54" y2="74" stroke="${p.glow}" stroke-width="1" opacity="0.4"/>
      <rect x="170" y="38" width="4" height="38" fill="${p.superstructure}"/>
      <line x1="172" y1="38" x2="200" y2="74" stroke="${p.glow}" stroke-width="1" opacity="0.4"/>
      <line x1="172" y1="38" x2="138" y2="74" stroke="${p.glow}" stroke-width="1" opacity="0.4"/>
    `;
  }

  // Shared aft superstructure — bridge deck, lit windows, funnel with a
  // glowing cap ring. Position matches the original detailed layout.
  const superstructure = `
    <ellipse cx="322" cy="52" rx="46" ry="36" fill="${p.glow}" opacity="0.10" filter="url(#${softGlowFilter})"/>
    <rect x="300" y="50" width="40" height="48" rx="2" fill="${p.superstructure}"/>
    ${litWindow(302, 54, 36, 6, 1.5)}
    ${litPort(305.5, 68, 2.2)} ${litPort(316.5, 68, 2.2)} ${litPort(327.5, 68, 2.2)}
    <rect x="302" y="76" width="7" height="6" rx="1" fill="${p.hullShade}"/>
    <rect x="313" y="76" width="7" height="6" rx="1" fill="${p.hullShade}"/>
    <rect x="324" y="76" width="7" height="6" rx="1" fill="${p.hullShade}"/>
    <rect x="310" y="26" width="14" height="26" rx="2" fill="${p.hullShade}"/>
    <rect x="309" y="24" width="16" height="3" rx="1.5" fill="${p.glow}" opacity="0.7"/>
    <circle cx="317" cy="24" r="5" fill="${p.glow}" opacity="0.18" filter="url(#${softGlowFilter})"/>
  `;

  /* ── Hull, portholes, waterline (shared, realistic profile) ── */
  const hullTopY = 96;
  const hull = `
    <path d="M 24,98 L 360,98 Q 378,98 385,108 L 390,124 L 14,124 Q 10,124 16,114 Z" fill="url(#${hullGrad})"/>
    <rect x="24" y="${hullTopY}" width="338" height="2.2" fill="${p.glow}" opacity="0.55"/>
    <path d="M 360,98 L 392,110 L 388,124 L 360,124 Z" fill="${p.hull}"/>
    <path d="M 382,100 L 397,116 L 395,124 L 388,124 L 392,110 Z" fill="${p.hullShade}"/>
    <path d="M 24,98 L 12,108 L 14,124 L 24,124 Z" fill="${p.hullShade}"/>
    <rect x="14" y="119" width="378" height="5" fill="${p.glow}" opacity="0.16"/>
    ${[44, 76, 108, 140, 172, 204, 236, 268].map(cx => litPort(cx, 110, 1.7)).join('')}
    <rect x="14" y="114" width="378" height="4" fill="rgba(0,0,0,0.25)"/>
  `;

  /* ── Reflection: the lit hull + windows mirrored below the waterline,
     faded out with a gradient mask so it reads as a soft glow on water. ── */
  const reflection = `
    <g transform="translate(0,248) scale(1,-1)" mask="url(#${fadeMask})" opacity="0.55">
      ${hull}
      ${superstructure}
    </g>
  `;

  const svg = `
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 200" width="100%" height="100%" style="display:block">
  <defs>
    <linearGradient id="${skyGrad}" x1="0" y1="0" x2="0.3" y2="1">
      <stop offset="0%" stop-color="${p.skyTop}"/>
      <stop offset="100%" stop-color="${p.skyBot}"/>
    </linearGradient>
    <linearGradient id="${waterGrad}" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0%" stop-color="${p.skyBot}"/>
      <stop offset="100%" stop-color="${p.hullShade}"/>
    </linearGradient>
    <linearGradient id="${hullGrad}" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0%" stop-color="${p.hull}"/>
      <stop offset="100%" stop-color="${p.hullShade}"/>
    </linearGradient>
    <linearGradient id="fadeGrad_${uid}" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0%" stop-color="#ffffff" stop-opacity="0.9"/>
      <stop offset="100%" stop-color="#ffffff" stop-opacity="0"/>
    </linearGradient>
    <mask id="${fadeMask}">
      <rect x="0" y="0" width="400" height="200" fill="url(#fadeGrad_${uid})"/>
    </mask>
    <filter id="${softGlowFilter}" x="-60%" y="-60%" width="220%" height="220%">
      <feGaussianBlur stdDeviation="2.4"/>
    </filter>
    <filter id="${glowFilter}" x="-80%" y="-80%" width="260%" height="260%">
      <feGaussianBlur stdDeviation="10"/>
    </filter>
  </defs>

  <!-- Night sky / harbor backdrop -->
  <rect x="0" y="0" width="400" height="126" fill="url(#${skyGrad})"/>
  <!-- Ambient floodlight glow behind the deckhouse -->
  <ellipse cx="322" cy="46" rx="90" ry="60" fill="${p.glow}" opacity="0.14" filter="url(#${glowFilter})"/>
  <!-- Horizon -->
  <rect x="0" y="118" width="400" height="8" fill="${p.skyBot}" opacity="0.55"/>

  <!-- Water -->
  <rect x="0" y="124" width="400" height="76" fill="url(#${waterGrad})"/>
  <!-- Reflection of the lit hull -->
  ${reflection}
  <!-- Faint moonlit ripples -->
  <path d="M0,150 Q40,146 80,150 Q120,154 160,150 Q200,146 240,150 Q280,154 320,150 Q360,146 400,150" stroke="${p.glow}" stroke-width="0.6" fill="none" opacity="0.18"/>
  <path d="M0,164 Q50,160 100,164 Q150,168 200,164 Q250,160 300,164 Q350,168 400,164" stroke="${p.glow}" stroke-width="0.5" fill="none" opacity="0.12"/>

  <!-- Deck equipment (behind hull silhouette) -->
  ${deckEquipment}
  ${superstructure}

  <!-- Hull -->
  ${hull}
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
    generateVesselSVG('Golden Pearl', 'Gas Carrier', 'v6'),
  ];
}

export const VESSEL_IMAGE_POOL: string[] = getVesselImagePool();

export async function loadVesselImagePool(): Promise<string[]> {
  return getVesselImagePool();
}
