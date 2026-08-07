"use strict";
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
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
exports.__esModule = true;
exports.loadVesselImagePool = exports.VESSEL_IMAGE_POOL = exports.getVesselImagePool = exports.pickRandomVesselImage = exports.getVesselImageForId = exports.resolveImgUrl = exports.generateVesselSVG = void 0;
var PALETTES = [
    // 0: Deep Royal Blue & Crimson (Ocean Star)
    {
        id: 0,
        skyTop: '#b3d4f0', skyBot: '#dceefa',
        waterTop: '#1565a8', waterBot: '#0a3d72',
        hull: '#1a365d', hullShade: '#0f2942', keel: '#c53030',
        superstructure: '#f8fafc', funnel: '#dd6b20', funnelAccent: '#1a365d',
        deckColor: '#243650'
    },
    // 1: Emerald Green & Charcoal (Sea Breeze)
    {
        id: 1,
        skyTop: '#a7f3d0', skyBot: '#ecfdf5',
        waterTop: '#059669', waterBot: '#047857',
        hull: '#064e3b', hullShade: '#022c22', keel: '#9f1239',
        superstructure: '#f8fafc', funnel: '#eab308', funnelAccent: '#064e3b',
        deckColor: '#16181b'
    },
    // 2: Midnight Cobalt & Wine Red (Blue Horizon)
    {
        id: 2,
        skyTop: '#cbe5ff', skyBot: '#f0f9ff',
        waterTop: '#1d4ed8', waterBot: '#1e3a8a',
        hull: '#1e40af', hullShade: '#172554', keel: '#881337',
        superstructure: '#ffffff', funnel: '#06b6d4', funnelAccent: '#1e40af',
        deckColor: '#0d1640'
    },
    // 3: Deep Burgundy & Bronze (Pacific Dawn)
    {
        id: 3,
        skyTop: '#fed7aa', skyBot: '#fff7ed',
        waterTop: '#0284c7', waterBot: '#0f766e',
        hull: '#701a75', hullShade: '#4a044e', keel: '#92400e',
        superstructure: '#fdf4ff', funnel: '#f59e0b', funnelAccent: '#701a75',
        deckColor: '#4a044e'
    },
    // 4: Vibrant Coral Red & Navy (Atlantic Wave)
    {
        id: 4,
        skyTop: '#dbeafe', skyBot: '#eff6ff',
        waterTop: '#2563eb', waterBot: '#1e40af',
        hull: '#dc2626', hullShade: '#991b1b', keel: '#450a0a',
        superstructure: '#ffffff', funnel: '#1d4ed8', funnelAccent: '#1e40af',
        deckColor: '#1d2b3a'
    },
    // 5: Sunset Amber & Charcoal (Golden Pearl)
    {
        id: 5,
        skyTop: '#fef08a', skyBot: '#fefce8',
        waterTop: '#0369a1', waterBot: '#0c4a6e',
        hull: '#78350f', hullShade: '#451a03', keel: '#b91c1c',
        superstructure: '#fffbeb', funnel: '#fbbf24', funnelAccent: '#78350f',
        deckColor: '#0f1c28'
    },
    // 6: Steel Slate & Flame Orange (Titan Guard)
    {
        id: 6,
        skyTop: '#cbd5e1', skyBot: '#f1f5f9',
        waterTop: '#0f766e', waterBot: '#134e4a',
        hull: '#334155', hullShade: '#1e293b', keel: '#be123c',
        superstructure: '#f8fafc', funnel: '#f97316', funnelAccent: '#334155',
        deckColor: '#0f2414'
    },
    // 7: Royal Violet & Gold
    {
        id: 7,
        skyTop: '#e9d5ff', skyBot: '#faf5ff',
        waterTop: '#4338ca', waterBot: '#312e81',
        hull: '#4c1d95', hullShade: '#2e1065', keel: '#d97706',
        superstructure: '#ffffff', funnel: '#8b5cf6', funnelAccent: '#4c1d95',
        deckColor: '#2e1065'
    },
    // 8: Deep Aqua Teal & Coral
    {
        id: 8,
        skyTop: '#ccfbf1', skyBot: '#f0fdf4',
        waterTop: '#0d9488', waterBot: '#115e59',
        hull: '#115e59', hullShade: '#042f2e', keel: '#e11d48',
        superstructure: '#f0fdf4', funnel: '#14b8a6', funnelAccent: '#115e59',
        deckColor: '#042f2e'
    },
    // 9: Terracotta Copper & Flame
    {
        id: 9,
        skyTop: '#ffedd5', skyBot: '#fff7ed',
        waterTop: '#0284c7', waterBot: '#075985',
        hull: '#c2410c', hullShade: '#7c2d12', keel: '#292524',
        superstructure: '#fff7ed', funnel: '#fb923c', funnelAccent: '#c2410c',
        deckColor: '#1c1917'
    },
];
function getHash(str) {
    var h = 5381;
    for (var i = 0; i < str.length; i++) {
        h = (((h << 5) + h) ^ str.charCodeAt(i)) >>> 0;
    }
    return h;
}
/**
 * Generates a detailed side-profile ship SVG.
 * ViewBox: 400 x 200  (wider canvas for more detail)
 */
function generateVesselSVG(name, type, id) {
    var hash = getHash(id + name + type);
    // Determine palette index
    var palIdx = hash % PALETTES.length;
    // Exact overrides for screenshot vessels
    var exact = {
        'Ocean Star': 0, 'Sea Breeze': 1, 'Blue Horizon': 2,
        'Pacific Dawn': 3, 'Atlantic Wave': 4, 'Golden Pearl': 5
    };
    if (exact[name] !== undefined)
        palIdx = exact[name];
    var p = PALETTES[palIdx];
    var t = (type || 'Container Ship').toLowerCase();
    /* ── Deck equipment based on vessel type ──────────────────────────── */
    var deckEquipment = '';
    if (t.indexOf('container') !== -1) {
        // Stacked multicolored containers
        var colors = ['#e53935', '#1e88e5', '#43a047', '#fb8c00', '#8e24aa', '#039be5', '#c0ca33', '#6d4c41'];
        var stacks = '';
        for (var col = 0; col < 14; col++) {
            var cx = 42 + col * 18;
            var layers = 1 + ((col + hash) % 3); // 1–3 containers tall
            for (var layer = 0; layer < layers; layer++) {
                var cy = 88 - layer * 13;
                var cc = colors[(col * 2 + layer + hash) % colors.length];
                stacks += "<rect x=\"".concat(cx, "\" y=\"").concat(cy, "\" width=\"16\" height=\"12\" fill=\"").concat(cc, "\" rx=\"0.8\" stroke=\"#111\" stroke-width=\"0.4\"/>");
                // Rib line on container
                stacks += "<line x1=\"".concat(cx + 2, "\" y1=\"").concat(cy + 3, "\" x2=\"").concat(cx + 14, "\" y2=\"").concat(cy + 3, "\" stroke=\"rgba(0,0,0,0.15)\" stroke-width=\"0.4\"/>");
                stacks += "<line x1=\"".concat(cx + 2, "\" y1=\"").concat(cy + 7, "\" x2=\"").concat(cx + 14, "\" y2=\"").concat(cy + 7, "\" stroke=\"rgba(0,0,0,0.15)\" stroke-width=\"0.4\"/>");
            }
        }
        deckEquipment = stacks;
        // Aft superstructure (right side)
        deckEquipment += "\n      <rect x=\"298\" y=\"52\" width=\"38\" height=\"48\" fill=\"".concat(p.superstructure, "\" rx=\"2\" stroke=\"#c0c8d0\" stroke-width=\"0.5\"/>\n      <rect x=\"300\" y=\"56\" width=\"34\" height=\"5\" fill=\"#64b5f6\"/>\n      <rect x=\"302\" y=\"64\" width=\"7\" height=\"7\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"313\" y=\"64\" width=\"7\" height=\"7\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"324\" y=\"64\" width=\"7\" height=\"7\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"302\" y=\"74\" width=\"7\" height=\"6\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"313\" y=\"74\" width=\"7\" height=\"6\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"324\" y=\"74\" width=\"7\" height=\"6\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"310\" y=\"28\" width=\"14\" height=\"24\" fill=\"").concat(p.funnel, "\" rx=\"2\"/>\n      <rect x=\"313\" y=\"22\" width=\"8\" height=\"6\" fill=\"").concat(p.funnelAccent, "\" rx=\"1\"/>\n      <rect x=\"308\" y=\"26\" width=\"18\" height=\"2\" fill=\"").concat(p.funnelAccent, "\"/>\n    ");
    }
    else if (t.indexOf('bulk') !== -1) {
        // 4 bulk cargo holds with deck cranes
        for (var h2 = 0; h2 < 4; h2++) {
            var hx = 50 + h2 * 55;
            deckEquipment += "\n        <rect x=\"".concat(hx, "\" y=\"82\" width=\"40\" height=\"10\" fill=\"#334155\" rx=\"2\"/>\n        <rect x=\"").concat(hx + 2, "\" y=\"78\" width=\"36\" height=\"4\" fill=\"#475569\" rx=\"1\"/>\n        <rect x=\"").concat(hx + 2, "\" y=\"76\" width=\"36\" height=\"2\" fill=\"#64748b\"/>\n        <!-- crane mast -->\n        <rect x=\"").concat(hx + 18, "\" y=\"44\" width=\"4\" height=\"34\" fill=\"#e2e8f0\"/>\n        <!-- crane boom -->\n        <line x1=\"").concat(hx + 20, "\" y1=\"44\" x2=\"").concat(hx + 42, "\" y2=\"74\" stroke=\"#e2e8f0\" stroke-width=\"2\"/>\n        <line x1=\"").concat(hx + 20, "\" y1=\"44\" x2=\"").concat(hx - 2, "\" y2=\"74\" stroke=\"#e2e8f0\" stroke-width=\"2\"/>\n        <!-- crane wire -->\n        <line x1=\"").concat(hx + 42, "\" y1=\"74\" x2=\"").concat(hx + 42, "\" y2=\"80\" stroke=\"#bbb\" stroke-width=\"1\"/>\n      ");
        }
        deckEquipment += "\n      <rect x=\"302\" y=\"50\" width=\"40\" height=\"48\" fill=\"".concat(p.superstructure, "\" rx=\"2\" stroke=\"#c0c8d0\" stroke-width=\"0.5\"/>\n      <rect x=\"304\" y=\"54\" width=\"36\" height=\"6\" fill=\"#64b5f6\"/>\n      <rect x=\"306\" y=\"64\" width=\"7\" height=\"7\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"317\" y=\"64\" width=\"7\" height=\"7\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"328\" y=\"64\" width=\"7\" height=\"7\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"306\" y=\"75\" width=\"7\" height=\"5\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"317\" y=\"75\" width=\"7\" height=\"5\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"314\" y=\"24\" width=\"14\" height=\"26\" fill=\"").concat(p.funnel, "\" rx=\"2\"/>\n      <rect x=\"317\" y=\"18\" width=\"8\" height=\"6\" fill=\"").concat(p.funnelAccent, "\" rx=\"1\"/>\n    ");
    }
    else if (t.indexOf('tanker') !== -1 || t.indexOf('chemical') !== -1) {
        // Long flat pipe manifold deck
        deckEquipment += "\n      <rect x=\"44\" y=\"84\" width=\"240\" height=\"4\" fill=\"#b0bec5\" rx=\"1\"/>\n      <rect x=\"44\" y=\"82\" width=\"240\" height=\"2\" fill=\"#ffd54f\"/>\n    ";
        for (var m = 0; m < 8; m++) {
            var mx = 54 + m * 30;
            deckEquipment += "\n        <rect x=\"".concat(mx, "\" y=\"72\" width=\"10\" height=\"14\" fill=\"#607d8b\" rx=\"1\"/>\n        <rect x=\"").concat(mx + 1, "\" y=\"66\" width=\"8\" height=\"6\" fill=\"#90a4ae\" rx=\"0.5\"/>\n        <circle cx=\"").concat(mx + 5, "\" cy=\"64\" r=\"3\" fill=\"").concat(m % 2 === 0 ? '#ef5350' : '#42a5f5', "\"/>\n        <line x1=\"").concat(mx + 5, "\" y1=\"72\" x2=\"").concat(mx + 5, "\" y2=\"80\" stroke=\"#90a4ae\" stroke-width=\"1.5\"/>\n      ");
        }
        deckEquipment += "\n      <rect x=\"304\" y=\"46\" width=\"42\" height=\"52\" fill=\"".concat(p.superstructure, "\" rx=\"2\" stroke=\"#c0c8d0\" stroke-width=\"0.5\"/>\n      <rect x=\"306\" y=\"50\" width=\"38\" height=\"6\" fill=\"#64b5f6\"/>\n      <rect x=\"308\" y=\"60\" width=\"8\" height=\"8\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"320\" y=\"60\" width=\"8\" height=\"8\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"332\" y=\"60\" width=\"8\" height=\"8\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"308\" y=\"72\" width=\"8\" height=\"6\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"320\" y=\"72\" width=\"8\" height=\"6\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"332\" y=\"72\" width=\"8\" height=\"6\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"316\" y=\"20\" width=\"16\" height=\"26\" fill=\"").concat(p.funnel, "\" rx=\"2\"/>\n      <rect x=\"319\" y=\"14\" width=\"10\" height=\"6\" fill=\"").concat(p.funnelAccent, "\" rx=\"1\"/>\n      <rect x=\"314\" y=\"44\" width=\"20\" height=\"2\" fill=\"").concat(p.funnelAccent, "\"/>\n    ");
    }
    else if (t.indexOf('offshore') !== -1 || t.indexOf('support') !== -1) {
        // Forward superstructure + open aft deck
        deckEquipment += "\n      <!-- Forward superstructure -->\n      <rect x=\"42\" y=\"38\" width=\"58\" height=\"60\" fill=\"".concat(p.superstructure, "\" rx=\"2\" stroke=\"#c0c8d0\" stroke-width=\"0.5\"/>\n      <rect x=\"44\" y=\"44\" width=\"54\" height=\"8\" fill=\"#64b5f6\"/>\n      <rect x=\"46\" y=\"56\" width=\"9\" height=\"9\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"59\" y=\"56\" width=\"9\" height=\"9\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"72\" y=\"56\" width=\"9\" height=\"9\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"85\" y=\"56\" width=\"9\" height=\"9\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"46\" y=\"70\" width=\"9\" height=\"8\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"59\" y=\"70\" width=\"9\" height=\"8\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"72\" y=\"70\" width=\"9\" height=\"8\" fill=\"#90a4ae\" rx=\"1\"/>\n      <!-- Helideck -->\n      <rect x=\"28\" y=\"32\" width=\"60\" height=\"6\" fill=\"#546e7a\"/>\n      <circle cx=\"58\" cy=\"35\" r=\"3\" fill=\"rgba(255,255,255,0.25)\" stroke=\"#fff\" stroke-width=\"0.8\"/>\n      <!-- Exhaust stack -->\n      <rect x=\"104\" y=\"36\" width=\"14\" height=\"28\" fill=\"").concat(p.funnel, "\" rx=\"2\"/>\n      <rect x=\"107\" y=\"30\" width=\"8\" height=\"6\" fill=\"").concat(p.funnelAccent, "\" rx=\"1\"/>\n      <!-- Aft crane / A-frame -->\n      <rect x=\"155\" y=\"60\" width=\"4\" height=\"38\" fill=\"#e2e8f0\"/>\n      <line x1=\"157\" y1=\"60\" x2=\"185\" y2=\"90\" stroke=\"#e2e8f0\" stroke-width=\"2.5\"/>\n      <line x1=\"157\" y1=\"60\" x2=\"129\" y2=\"90\" stroke=\"#e2e8f0\" stroke-width=\"2.5\"/>\n      <!-- Aft deck equipment -->\n      <rect x=\"175\" y=\"72\" width=\"28\" height=\"16\" fill=\"#d97706\" rx=\"2\"/>\n      <rect x=\"210\" y=\"68\" width=\"36\" height=\"20\" fill=\"#1565c0\" rx=\"2\"/>\n      <rect x=\"255\" y=\"74\" width=\"22\" height=\"14\" fill=\"#cc2424\" rx=\"2\"/>\n      <rect x=\"125\" y=\"82\" width=\"150\" height=\"8\" fill=\"#455a64\"/>\n    ");
    }
    else if (t.indexOf('gas') !== -1) {
        // LNG spherical tanks on deck
        for (var g = 0; g < 3; g++) {
            var gx = 80 + g * 75;
            deckEquipment += "\n        <circle cx=\"".concat(gx, "\" cy=\"72\" r=\"26\" fill=\"#e3f2fd\" stroke=\"#90caf9\" stroke-width=\"1.5\"/>\n        <ellipse cx=\"").concat(gx, "\" cy=\"72\" rx=\"24\" ry=\"22\" fill=\"#bbdefb\"/>\n        <ellipse cx=\"").concat(gx - 6, "\" cy=\"66\" rx=\"8\" ry=\"6\" fill=\"rgba(255,255,255,0.35)\"/>\n        <line x1=\"").concat(gx, "\" y1=\"46\" x2=\"").concat(gx, "\" y2=\"96\" stroke=\"#42a5f5\" stroke-width=\"1\" opacity=\"0.5\"/>\n        <line x1=\"").concat(gx - 24, "\" y1=\"72\" x2=\"").concat(gx + 24, "\" y2=\"72\" stroke=\"#42a5f5\" stroke-width=\"1\" opacity=\"0.5\"/>\n      ");
        }
        deckEquipment += "\n      <rect x=\"310\" y=\"48\" width=\"40\" height=\"50\" fill=\"".concat(p.superstructure, "\" rx=\"2\" stroke=\"#c0c8d0\" stroke-width=\"0.5\"/>\n      <rect x=\"312\" y=\"52\" width=\"36\" height=\"6\" fill=\"#64b5f6\"/>\n      <rect x=\"314\" y=\"62\" width=\"8\" height=\"8\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"326\" y=\"62\" width=\"8\" height=\"8\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"338\" y=\"62\" width=\"8\" height=\"8\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"320\" y=\"22\" width=\"14\" height=\"26\" fill=\"").concat(p.funnel, "\" rx=\"2\"/>\n      <rect x=\"323\" y=\"16\" width=\"8\" height=\"6\" fill=\"").concat(p.funnelAccent, "\" rx=\"1\"/>\n    ");
    }
    else {
        // General Cargo — midship cranes + aft house
        deckEquipment += "\n      <rect x=\"48\" y=\"78\" width=\"50\" height=\"14\" fill=\"#334155\" rx=\"2\"/>\n      <rect x=\"50\" y=\"76\" width=\"46\" height=\"2\" fill=\"#475569\"/>\n      <rect x=\"120\" y=\"78\" width=\"50\" height=\"14\" fill=\"#334155\" rx=\"2\"/>\n      <rect x=\"122\" y=\"76\" width=\"46\" height=\"2\" fill=\"#475569\"/>\n      <!-- Fore crane -->\n      <rect x=\"90\" y=\"38\" width=\"5\" height=\"38\" fill=\"#e2e8f0\"/>\n      <line x1=\"92.5\" y1=\"38\" x2=\"120\" y2=\"74\" stroke=\"#e2e8f0\" stroke-width=\"2\"/>\n      <line x1=\"92.5\" y1=\"38\" x2=\"54\" y2=\"74\" stroke=\"#e2e8f0\" stroke-width=\"2\"/>\n      <line x1=\"120\" y1=\"74\" x2=\"120\" y2=\"82\" stroke=\"#aaa\" stroke-width=\"1.5\"/>\n      <!-- Mid crane -->\n      <rect x=\"170\" y=\"38\" width=\"5\" height=\"38\" fill=\"#e2e8f0\"/>\n      <line x1=\"172.5\" y1=\"38\" x2=\"200\" y2=\"74\" stroke=\"#e2e8f0\" stroke-width=\"2\"/>\n      <line x1=\"172.5\" y1=\"38\" x2=\"138\" y2=\"74\" stroke=\"#e2e8f0\" stroke-width=\"2\"/>\n      <line x1=\"200\" y1=\"74\" x2=\"200\" y2=\"82\" stroke=\"#aaa\" stroke-width=\"1.5\"/>\n    ";
        deckEquipment += "\n      <rect x=\"300\" y=\"50\" width=\"40\" height=\"48\" fill=\"".concat(p.superstructure, "\" rx=\"2\" stroke=\"#c0c8d0\" stroke-width=\"0.5\"/>\n      <rect x=\"302\" y=\"54\" width=\"36\" height=\"6\" fill=\"#64b5f6\"/>\n      <rect x=\"304\" y=\"64\" width=\"7\" height=\"7\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"315\" y=\"64\" width=\"7\" height=\"7\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"326\" y=\"64\" width=\"7\" height=\"7\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"304\" y=\"75\" width=\"7\" height=\"6\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"315\" y=\"75\" width=\"7\" height=\"6\" fill=\"#90a4ae\" rx=\"1\"/>\n      <rect x=\"314\" y=\"24\" width=\"14\" height=\"26\" fill=\"").concat(p.funnel, "\" rx=\"2\"/>\n      <rect x=\"317\" y=\"18\" width=\"8\" height=\"6\" fill=\"").concat(p.funnelAccent, "\" rx=\"1\"/>\n    ");
    }
    /* ── Complete SVG ─────────────────────────────────────────────────── */
    var svg = "\n<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 400 200\" width=\"100%\" height=\"100%\" style=\"display:block\">\n  <defs>\n    <linearGradient id=\"sky_".concat(palIdx, "\" x1=\"0\" y1=\"0\" x2=\"0\" y2=\"1\">\n      <stop offset=\"0%\" stop-color=\"").concat(p.skyTop, "\"/>\n      <stop offset=\"100%\" stop-color=\"").concat(p.skyBot, "\"/>\n    </linearGradient>\n    <linearGradient id=\"water_").concat(palIdx, "\" x1=\"0\" y1=\"0\" x2=\"0\" y2=\"1\">\n      <stop offset=\"0%\" stop-color=\"").concat(p.waterTop, "\"/>\n      <stop offset=\"100%\" stop-color=\"").concat(p.waterBot, "\"/>\n    </linearGradient>\n    <linearGradient id=\"hull_").concat(palIdx, "\" x1=\"0\" y1=\"0\" x2=\"0\" y2=\"1\">\n      <stop offset=\"0%\" stop-color=\"").concat(p.hull, "\"/>\n      <stop offset=\"100%\" stop-color=\"").concat(p.hullShade, "\"/>\n    </linearGradient>\n    <linearGradient id=\"superHull_").concat(palIdx, "\" x1=\"0\" y1=\"0\" x2=\"0\" y2=\"1\">\n      <stop offset=\"0%\" stop-color=\"").concat(p.hull, "\" stop-opacity=\"0.7\"/>\n      <stop offset=\"100%\" stop-color=\"").concat(p.hull, "\"/>\n    </linearGradient>\n  </defs>\n\n  <!-- \u2500\u2500 Sky \u2500\u2500 -->\n  <rect x=\"0\" y=\"0\" width=\"400\" height=\"126\" fill=\"url(#sky_").concat(palIdx, ")\"/>\n\n  <!-- Sun glow -->\n  <circle cx=\"360\" cy=\"34\" r=\"22\" fill=\"#fff9e0\" opacity=\"0.55\"/>\n  <circle cx=\"360\" cy=\"34\" r=\"14\" fill=\"#ffe082\" opacity=\"0.7\"/>\n\n  <!-- Clouds -->\n  <ellipse cx=\"60\" cy=\"28\" rx=\"32\" ry=\"10\" fill=\"#fff\" opacity=\"0.65\"/>\n  <ellipse cx=\"80\" cy=\"24\" rx=\"20\" ry=\"8\" fill=\"#fff\" opacity=\"0.55\"/>\n  <ellipse cx=\"195\" cy=\"22\" rx=\"38\" ry=\"11\" fill=\"#fff\" opacity=\"0.5\"/>\n  <ellipse cx=\"218\" cy=\"18\" rx=\"24\" ry=\"9\" fill=\"#fff\" opacity=\"0.45\"/>\n  <ellipse cx=\"310\" cy=\"30\" rx=\"22\" ry=\"8\" fill=\"#fff\" opacity=\"0.4\"/>\n\n  <!-- Horizon haze -->\n  <rect x=\"0\" y=\"118\" width=\"400\" height=\"8\" fill=\"").concat(p.waterTop, "\" opacity=\"0.35\"/>\n\n  <!-- \u2500\u2500 Ocean \u2500\u2500 -->\n  <rect x=\"0\" y=\"124\" width=\"400\" height=\"76\" fill=\"url(#water_").concat(palIdx, ")\"/>\n\n  <!-- Wave ripples -->\n  <path d=\"M0,130 Q25,127 50,130 Q75,133 100,130 Q125,127 150,130 Q175,133 200,130 Q225,127 250,130 Q275,133 300,130 Q325,127 350,130 Q375,133 400,130\" stroke=\"#fff\" stroke-width=\"0.9\" fill=\"none\" opacity=\"0.4\"/>\n  <path d=\"M0,138 Q30,135 60,138 Q90,141 120,138 Q150,135 180,138 Q210,141 240,138 Q270,135 300,138 Q330,141 360,138 Q380,141 400,138\" stroke=\"#fff\" stroke-width=\"0.6\" fill=\"none\" opacity=\"0.28\"/>\n  <path d=\"M0,148 Q40,145 80,148 Q120,151 160,148 Q200,145 240,148 Q280,151 320,148 Q360,145 400,148\" stroke=\"#fff\" stroke-width=\"0.5\" fill=\"none\" opacity=\"0.18\"/>\n  <path d=\"M0,158 Q50,155 100,158 Q150,161 200,158 Q250,155 300,158 Q350,161 400,158\" stroke=\"#fff\" stroke-width=\"0.4\" fill=\"none\" opacity=\"0.12\"/>\n\n  <!-- \u2500\u2500 Deck equipment (drawn BEHIND the hull shape) \u2500\u2500 -->\n  ").concat(deckEquipment, "\n\n  <!-- \u2500\u2500 Main Hull shape \u2500\u2500 -->\n  <!-- Hull body -->\n  <path d=\"\n    M 24,98\n    L 360,98\n    Q 378,98 385,108\n    L 390,124\n    L 14,124\n    Q 10,124 16,114\n    Z\n  \" fill=\"url(#hull_").concat(palIdx, ")\"/>\n\n  <!-- Hull highlight strip (deck rail top) -->\n  <rect x=\"24\" y=\"96\" width=\"338\" height=\"4\" fill=\"").concat(p.deckColor, "\" rx=\"1\"/>\n\n  <!-- Sharp bow -->\n  <path d=\"M 360,98 L 392,110 L 388,124 L 360,124 Z\" fill=\"").concat(p.hull, "\"/>\n  <path d=\"M 382,100 L 397,116 L 395,124 L 388,124 L 392,110 Z\" fill=\"").concat(p.hullShade, "\"/>\n\n  <!-- Stern curve -->\n  <path d=\"M 24,98 L 12,108 L 14,124 L 24,124 Z\" fill=\"").concat(p.hullShade, "\"/>\n\n  <!-- Waterline / keel stripe -->\n  <rect x=\"14\" y=\"118\" width=\"378\" height=\"6\" fill=\"").concat(p.keel, "\" rx=\"0\"/>\n\n  <!-- Portholes row -->\n  <circle cx=\"44\" cy=\"110\" r=\"2.2\" fill=\"#fff\" opacity=\"0.7\"/>\n  <circle cx=\"60\" cy=\"110\" r=\"2.2\" fill=\"#fff\" opacity=\"0.7\"/>\n  <circle cx=\"76\" cy=\"110\" r=\"2.2\" fill=\"#fff\" opacity=\"0.7\"/>\n  <circle cx=\"92\" cy=\"110\" r=\"2.2\" fill=\"#fff\" opacity=\"0.7\"/>\n  <circle cx=\"108\" cy=\"110\" r=\"2.2\" fill=\"#fff\" opacity=\"0.7\"/>\n  <circle cx=\"124\" cy=\"110\" r=\"2.2\" fill=\"#fff\" opacity=\"0.7\"/>\n  <circle cx=\"140\" cy=\"110\" r=\"2.2\" fill=\"#fff\" opacity=\"0.7\"/>\n  <circle cx=\"156\" cy=\"110\" r=\"2.2\" fill=\"#fff\" opacity=\"0.7\"/>\n  <circle cx=\"172\" cy=\"110\" r=\"2.2\" fill=\"#fff\" opacity=\"0.7\"/>\n  <circle cx=\"188\" cy=\"110\" r=\"2.2\" fill=\"#fff\" opacity=\"0.7\"/>\n  <circle cx=\"204\" cy=\"110\" r=\"2.2\" fill=\"#fff\" opacity=\"0.7\"/>\n  <circle cx=\"220\" cy=\"110\" r=\"2.2\" fill=\"#fff\" opacity=\"0.7\"/>\n  <circle cx=\"236\" cy=\"110\" r=\"2.2\" fill=\"#fff\" opacity=\"0.7\"/>\n  <circle cx=\"252\" cy=\"110\" r=\"2.2\" fill=\"#fff\" opacity=\"0.7\"/>\n  <circle cx=\"268\" cy=\"110\" r=\"2.2\" fill=\"#fff\" opacity=\"0.7\"/>\n  <circle cx=\"284\" cy=\"110\" r=\"2.2\" fill=\"#fff\" opacity=\"0.7\"/>\n\n  <!-- Hull shadow band -->\n  <rect x=\"14\" y=\"114\" width=\"378\" height=\"4\" fill=\"rgba(0,0,0,0.12)\"/>\n\n  <!-- Water reflection / wake -->\n  <path d=\"M 390,120 Q 397,126 393,134 Q 388,140 380,136\" stroke=\"#fff\" stroke-width=\"1.4\" fill=\"none\" opacity=\"0.55\"/>\n  <path d=\"M 14,120 Q 7,126 11,134 Q 16,140 24,136\" stroke=\"#fff\" stroke-width=\"1.4\" fill=\"none\" opacity=\"0.55\"/>\n  <path d=\"M 0,140 Q 40,136 80,140 Q 120,144 160,140\" stroke=\"#fff\" stroke-width=\"0.6\" fill=\"none\" opacity=\"0.2\"/>\n  <path d=\"M 240,140 Q 280,136 320,140 Q 360,144 400,140\" stroke=\"#fff\" stroke-width=\"0.6\" fill=\"none\" opacity=\"0.2\"/>\n</svg>").trim();
    return 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
}
exports.generateVesselSVG = generateVesselSVG;
function resolveImgUrl(val, name, type, id) {
    if (typeof val === 'string' && val.length > 10 && val.indexOf('data:image/svg+xml') === 0)
        return val;
    return generateVesselSVG(name || 'Vessel', type || 'Container Ship', id || 'v0');
}
exports.resolveImgUrl = resolveImgUrl;
function getVesselImageForId(id, name, type) {
    return generateVesselSVG(name || 'Vessel', type || 'Container Ship', id || 'v0');
}
exports.getVesselImageForId = getVesselImageForId;
function pickRandomVesselImage(vesselType) {
    var randId = 'v_' + Math.floor(Math.random() * 100000);
    return generateVesselSVG('New Vessel', vesselType || 'Container Ship', randId);
}
exports.pickRandomVesselImage = pickRandomVesselImage;
function getVesselImagePool() {
    return [
        generateVesselSVG('Ocean Star', 'Container Ship', 'v1'),
        generateVesselSVG('Sea Breeze', 'Bulk Carrier', 'v2'),
        generateVesselSVG('Blue Horizon', 'Oil Tanker', 'v3'),
        generateVesselSVG('Pacific Dawn', 'General Cargo', 'v4'),
        generateVesselSVG('Atlantic Wave', 'Offshore Support', 'v5'),
        generateVesselSVG('Golden Pearl', 'Container Ship', 'v6'),
    ];
}
exports.getVesselImagePool = getVesselImagePool;
exports.VESSEL_IMAGE_POOL = getVesselImagePool();
function loadVesselImagePool() {
    return __awaiter(this, void 0, void 0, function () {
        return __generator(this, function (_a) {
            return [2 /*return*/, getVesselImagePool()];
        });
    });
}
exports.loadVesselImagePool = loadVesselImagePool;
