"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
exports.renderReportsPage = renderReportsPage;
var React = __importStar(require("react"));
function renderReportsPage(host) {
    return (React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 20 } },
        React.createElement("div", { style: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 12 } },
            React.createElement("div", null,
                React.createElement("h2", { style: { margin: 0, fontSize: 20, fontWeight: 700, color: '#0f172a' } }, "Reports"),
                React.createElement("p", { style: { margin: '4px 0 0', fontSize: 13, color: '#64748b' } }, "Generate and view analytics reports.")),
            React.createElement("button", { onClick: function () { return alert('Downloading report PDF...'); }, style: {
                    background: '#0078d4', color: '#fff', border: 'none', borderRadius: 6,
                    padding: '8px 18px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
                    display: 'inline-flex', alignItems: 'center', gap: 6,
                } }, "\u2B07 Download Report")),
        React.createElement("div", { style: { background: '#fff', borderRadius: 8, padding: 12, border: '1px solid #e2e8f0', display: 'flex', gap: 12, alignItems: 'center' } },
            React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 8, fontSize: 13, color: '#334155' } },
                React.createElement("span", null, "\uD83D\uDCC5"),
                React.createElement("input", { type: "text", defaultValue: "May 1, 2024 - May 31, 2024", style: { border: '1px solid #cbd5e1', borderRadius: 6, padding: '6px 12px', fontSize: 13, outline: 'none' } })),
            React.createElement("select", { style: { border: '1px solid #cbd5e1', borderRadius: 6, padding: '6px 12px', fontSize: 13, outline: 'none', background: '#fff' } },
                React.createElement("option", null, "All Vessels"),
                React.createElement("option", null, "Ocean Star"),
                React.createElement("option", null, "Sea Breeze"))),
        React.createElement("div", { style: { display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: 16 } },
            React.createElement("div", { style: { background: '#fff', borderRadius: 8, padding: 18, border: '1px solid #e2e8f0' } },
                React.createElement("div", { style: { fontSize: 12, color: '#64748b', fontWeight: 600, textTransform: 'uppercase' } }, "Total Documents"),
                React.createElement("div", { style: { fontSize: 26, fontWeight: 800, color: '#0078d4', marginTop: 4 } }, "156")),
            React.createElement("div", { style: { background: '#fff', borderRadius: 8, padding: 18, border: '1px solid #e2e8f0' } },
                React.createElement("div", { style: { fontSize: 12, color: '#64748b', fontWeight: 600, textTransform: 'uppercase' } }, "Expired Documents"),
                React.createElement("div", { style: { fontSize: 26, fontWeight: 800, color: '#ef4444', marginTop: 4 } }, "8")),
            React.createElement("div", { style: { background: '#fff', borderRadius: 8, padding: 18, border: '1px solid #e2e8f0' } },
                React.createElement("div", { style: { fontSize: 12, color: '#64748b', fontWeight: 600, textTransform: 'uppercase' } }, "Expiring in 30 days"),
                React.createElement("div", { style: { fontSize: 26, fontWeight: 800, color: '#f59e0b', marginTop: 4 } }, "23")),
            React.createElement("div", { style: { background: '#fff', borderRadius: 8, padding: 18, border: '1px solid #e2e8f0' } },
                React.createElement("div", { style: { fontSize: 12, color: '#64748b', fontWeight: 600, textTransform: 'uppercase' } }, "Valid Documents"),
                React.createElement("div", { style: { fontSize: 26, fontWeight: 800, color: '#10b981', marginTop: 4 } }, "125"))),
        React.createElement("div", { style: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 20 } },
            React.createElement("div", { style: { background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', padding: 20 } },
                React.createElement("h3", { style: { margin: '0 0 16px', fontSize: 15, fontWeight: 700, color: '#0f172a' } }, "Documents by Type"),
                React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 24 } },
                    React.createElement("div", { style: { position: 'relative', width: 140, height: 140 } },
                        React.createElement("svg", { width: "140", height: "140", viewBox: "0 0 42 42" },
                            React.createElement("circle", { cx: "21", cy: "21", r: "15.91549430918954", fill: "transparent", stroke: "#e2e8f0", strokeWidth: "5" }),
                            React.createElement("circle", { cx: "21", cy: "21", r: "15.91549430918954", fill: "transparent", stroke: "#2563eb", strokeWidth: "5", strokeDasharray: "29 71", strokeDashoffset: "25" }),
                            React.createElement("circle", { cx: "21", cy: "21", r: "15.91549430918954", fill: "transparent", stroke: "#f59e0b", strokeWidth: "5", strokeDasharray: "26 74", strokeDashoffset: "96" }),
                            React.createElement("circle", { cx: "21", cy: "21", r: "15.91549430918954", fill: "transparent", stroke: "#10b981", strokeWidth: "5", strokeDasharray: "22 78", strokeDashoffset: "70" }),
                            React.createElement("circle", { cx: "21", cy: "21", r: "15.91549430918954", fill: "transparent", stroke: "#8b5cf6", strokeWidth: "5", strokeDasharray: "13 87", strokeDashoffset: "48" }),
                            React.createElement("circle", { cx: "21", cy: "21", r: "15.91549430918954", fill: "transparent", stroke: "#06b6d4", strokeWidth: "5", strokeDasharray: "10 90", strokeDashoffset: "35" }))),
                    React.createElement("div", { style: { flex: 1, display: 'flex', flexDirection: 'column', gap: 6, fontSize: 12 } },
                        React.createElement("div", { style: { display: 'flex', justifyContent: 'space-between' } },
                            React.createElement("span", { style: { color: '#475569' } }, "\u25CF Insurance"),
                            React.createElement("span", { style: { fontWeight: 600 } }, "45 (29%)")),
                        React.createElement("div", { style: { display: 'flex', justifyContent: 'space-between' } },
                            React.createElement("span", { style: { color: '#475569' } }, "\u25CF Crew"),
                            React.createElement("span", { style: { fontWeight: 600 } }, "40 (26%)")),
                        React.createElement("div", { style: { display: 'flex', justifyContent: 'space-between' } },
                            React.createElement("span", { style: { color: '#475569' } }, "\u25CF Maintenance"),
                            React.createElement("span", { style: { fontWeight: 600 } }, "35 (22%)")),
                        React.createElement("div", { style: { display: 'flex', justifyContent: 'space-between' } },
                            React.createElement("span", { style: { color: '#475569' } }, "\u25CF Certificate"),
                            React.createElement("span", { style: { fontWeight: 600 } }, "20 (13%)")),
                        React.createElement("div", { style: { display: 'flex', justifyContent: 'space-between' } },
                            React.createElement("span", { style: { color: '#475569' } }, "\u25CF Survey"),
                            React.createElement("span", { style: { fontWeight: 600 } }, "16 (10%)")),
                        React.createElement("div", { style: { borderTop: '1px solid #e2e8f0', paddingTop: 6, marginTop: 4, display: 'flex', justifyContent: 'space-between', fontWeight: 700 } },
                            React.createElement("span", null, "Total"),
                            React.createElement("span", null, "156"))))),
            React.createElement("div", { style: { background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', padding: 20 } },
                React.createElement("h3", { style: { margin: '0 0 16px', fontSize: 15, fontWeight: 700, color: '#0f172a' } }, "Documents by Vessel"),
                React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 12 } },
                    [
                        { name: 'Ocean Star', count: 45, max: 50 },
                        { name: 'Sea Breeze', count: 32, max: 50 },
                        { name: 'Blue Horizon', count: 28, max: 50 },
                        { name: 'Pacific Dawn', count: 26, max: 50 },
                        { name: 'Atlantic Wave', count: 25, max: 50 },
                    ].map(function (v) { return (React.createElement("div", { key: v.name, style: { display: 'flex', alignItems: 'center', gap: 12, fontSize: 12 } },
                        React.createElement("span", { style: { width: 90, color: '#475569', fontWeight: 500 } }, v.name),
                        React.createElement("div", { style: { flex: 1, background: '#f1f5f9', borderRadius: 4, height: 16, overflow: 'hidden' } },
                            React.createElement("div", { style: { width: "".concat((v.count / v.max) * 100, "%"), background: '#2563eb', height: '100%', borderRadius: 4 } })),
                        React.createElement("span", { style: { width: 24, fontWeight: 600, textAlign: 'right' } }, v.count))); }),
                    React.createElement("div", { style: { fontSize: 10, color: '#94a3b8', textAlign: 'center', marginTop: 8 } }, "No. of Documents"))))));
}
