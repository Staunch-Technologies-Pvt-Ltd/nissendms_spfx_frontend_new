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
exports.renderDashboard = renderDashboard;
var React = __importStar(require("react"));
var constants_1 = require("../constants");
var vesselImagePool_1 = require("../vesselImagePool");
function renderDashboard(host) {
    var _a = host.state, rows = _a.rows, vessels = _a.vessels, documentsList = _a.documentsList;
    var userDisplayName = host.props.userDisplayName || 'Priya';
    var totalDocs = rows.length > 0 ? rows.length : 156;
    var expiringSoon = 23;
    var pendingApprovals = 7;
    var totalVessels = vessels.length;
    return (React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 20 } },
        React.createElement("div", { style: {
                background: 'linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%)',
                borderRadius: 12, padding: '24px 32px', display: 'flex', justifyContent: 'space-between',
                alignItems: 'center', border: '1px solid #bfdbfe', position: 'relative', overflow: 'hidden',
            } },
            React.createElement("div", null,
                React.createElement("h1", { style: { margin: 0, fontSize: 22, fontWeight: 700, color: '#1e3a8a' } },
                    "Welcome back, ",
                    userDisplayName,
                    "! \uD83D\uDC4B"),
                React.createElement("p", { style: { margin: '6px 0 0', fontSize: 13, color: '#3b82f6' } }, "Here's what's happening with your vessels and documents.")),
            React.createElement("div", { style: { height: 85, width: 200, borderRadius: 8, overflow: 'hidden', flexShrink: 0, boxShadow: '0 4px 12px rgba(37,99,235,0.15)' } },
                React.createElement("img", { src: (0, vesselImagePool_1.resolveImgUrl)(null), alt: "Vessel Illustration", style: { width: '100%', height: '100%', objectFit: 'cover' } }))),
        React.createElement("div", { style: { display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: 16 } },
            React.createElement("div", { style: { background: '#fff', borderRadius: 10, padding: 20, border: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between', alignItems: 'center' } },
                React.createElement("div", null,
                    React.createElement("div", { style: { fontSize: 28, fontWeight: 800, color: '#0f172a' } }, totalDocs),
                    React.createElement("div", { style: { fontSize: 12, color: '#64748b', fontWeight: 500, marginTop: 2 } }, "Total Documents"),
                    React.createElement("button", { onClick: function () { return host._goToView('list'); }, style: { background: 'none', border: 'none', color: '#2563eb', fontSize: 11, fontWeight: 600, padding: 0, marginTop: 8, cursor: 'pointer' } }, "View all \u2192")),
                React.createElement("div", { style: { width: 44, height: 44, borderRadius: '50%', background: '#f0f9ff', color: '#0284c7', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20 } }, "\uD83D\uDCC4")),
            React.createElement("div", { style: { background: '#fff', borderRadius: 10, padding: 20, border: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between', alignItems: 'center' } },
                React.createElement("div", null,
                    React.createElement("div", { style: { fontSize: 28, fontWeight: 800, color: '#0f172a' } }, expiringSoon),
                    React.createElement("div", { style: { fontSize: 12, color: '#64748b', fontWeight: 500, marginTop: 2 } }, "Expiring Soon"),
                    React.createElement("button", { onClick: function () { return host._goToView('list'); }, style: { background: 'none', border: 'none', color: '#2563eb', fontSize: 11, fontWeight: 600, padding: 0, marginTop: 8, cursor: 'pointer' } }, "View all \u2192")),
                React.createElement("div", { style: { width: 44, height: 44, borderRadius: '50%', background: '#ecfdf5', color: '#10b981', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20 } }, "\u2705")),
            React.createElement("div", { style: { background: '#fff', borderRadius: 10, padding: 20, border: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between', alignItems: 'center' } },
                React.createElement("div", null,
                    React.createElement("div", { style: { fontSize: 28, fontWeight: 800, color: '#0f172a' } },
                        "0",
                        pendingApprovals),
                    React.createElement("div", { style: { fontSize: 12, color: '#64748b', fontWeight: 500, marginTop: 2 } }, "Pending Approvals"),
                    React.createElement("button", { onClick: function () { return host._goToView('approvals'); }, style: { background: 'none', border: 'none', color: '#2563eb', fontSize: 11, fontWeight: 600, padding: 0, marginTop: 8, cursor: 'pointer' } }, "View all \u2192")),
                React.createElement("div", { style: { width: 44, height: 44, borderRadius: '50%', background: '#fffbeb', color: '#f59e0b', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20 } }, "\u23F0")),
            React.createElement("div", { style: { background: '#fff', borderRadius: 10, padding: 20, border: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between', alignItems: 'center' } },
                React.createElement("div", null,
                    React.createElement("div", { style: { fontSize: 28, fontWeight: 800, color: '#0f172a' } }, totalVessels),
                    React.createElement("div", { style: { fontSize: 12, color: '#64748b', fontWeight: 500, marginTop: 2 } }, "Total Vessels"),
                    React.createElement("button", { onClick: function () { return host._goToView('vessels'); }, style: { background: 'none', border: 'none', color: '#2563eb', fontSize: 11, fontWeight: 600, padding: 0, marginTop: 8, cursor: 'pointer' } }, "View all \u2192")),
                React.createElement("div", { style: { width: 44, height: 44, borderRadius: '50%', background: '#eff6ff', color: '#2563eb', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20 } }, "\uD83D\uDEA2"))),
        React.createElement("div", { style: { display: 'grid', gridTemplateColumns: '2fr 1fr', gap: 20 } },
            React.createElement("div", { style: { background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', padding: 20 } },
                React.createElement("div", { style: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 } },
                    React.createElement("h3", { style: { margin: 0, fontSize: 15, fontWeight: 700, color: '#0f172a' } }, "Recent Documents"),
                    React.createElement("button", { onClick: function () { return host._goToView('list'); }, style: { background: 'none', border: 'none', color: '#2563eb', fontSize: 12, fontWeight: 600, cursor: 'pointer' } }, "View all")),
                React.createElement("table", { style: { width: '100%', borderCollapse: 'collapse', fontSize: 13 } },
                    React.createElement("thead", null,
                        React.createElement("tr", { style: { borderBottom: '1px solid #e2e8f0', textTransform: 'uppercase', fontSize: 11, color: '#64748b', textAlign: 'left' } },
                            React.createElement("th", { style: { padding: '8px 12px' } }, "Name"),
                            React.createElement("th", { style: { padding: '8px 12px' } }, "Vessel"),
                            React.createElement("th", { style: { padding: '8px 12px' } }, "Type"),
                            React.createElement("th", { style: { padding: '8px 12px' } }, "Modified"))),
                    React.createElement("tbody", null, documentsList.slice(3, 7).map(function (doc, idx) { return (React.createElement("tr", { key: doc.id, style: { borderBottom: '1px solid #f1f5f9' } },
                        React.createElement("td", { style: { padding: '10px 12px', fontWeight: 600, color: '#1e293b', display: 'flex', alignItems: 'center', gap: 8 } },
                            React.createElement("span", { style: { color: doc.name.endsWith('.pdf') ? '#ef4444' : doc.name.endsWith('.docx') ? '#2563eb' : '#10b981' } }, "\uD83D\uDCC4"),
                            doc.name),
                        React.createElement("td", { style: { padding: '10px 12px', color: '#475569' } }, doc.vessel),
                        React.createElement("td", { style: { padding: '10px 12px' } }, (0, constants_1.badge)(doc.type === 'Insurance' ? 'purple' : doc.type === 'Crew' ? 'blue' : doc.type === 'Maintenance' ? 'green' : 'orange', doc.type)),
                        React.createElement("td", { style: { padding: '10px 12px', color: '#64748b', fontSize: 12 } }, doc.modified))); })))),
            React.createElement("div", { style: { background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', padding: 20 } },
                React.createElement("div", { style: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 } },
                    React.createElement("h3", { style: { margin: 0, fontSize: 15, fontWeight: 700, color: '#0f172a' } }, "Document Expiry Overview"),
                    React.createElement("button", { onClick: function () { return host._goToView('reports'); }, style: { background: 'none', border: 'none', color: '#2563eb', fontSize: 12, fontWeight: 600, cursor: 'pointer' } }, "View all")),
                React.createElement("div", { style: { display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 16 } },
                    React.createElement("div", { style: { position: 'relative', width: 140, height: 140 } },
                        React.createElement("svg", { width: "140", height: "140", viewBox: "0 0 42 42" },
                            React.createElement("circle", { cx: "21", cy: "21", r: "15.91549430918954", fill: "transparent", stroke: "#e2e8f0", strokeWidth: "5" }),
                            React.createElement("circle", { cx: "21", cy: "21", r: "15.91549430918954", fill: "transparent", stroke: "#10b981", strokeWidth: "5", strokeDasharray: "67 33", strokeDashoffset: "25" }),
                            React.createElement("circle", { cx: "21", cy: "21", r: "15.91549430918954", fill: "transparent", stroke: "#f59e0b", strokeWidth: "5", strokeDasharray: "24 76", strokeDashoffset: "58" }),
                            React.createElement("circle", { cx: "21", cy: "21", r: "15.91549430918954", fill: "transparent", stroke: "#ef4444", strokeWidth: "5", strokeDasharray: "9 91", strokeDashoffset: "34" })),
                        React.createElement("div", { style: { position: 'absolute', inset: 0, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center' } },
                            React.createElement("span", { style: { fontSize: 20, fontWeight: 800, color: '#0f172a' } }, "96"),
                            React.createElement("span", { style: { fontSize: 10, color: '#64748b', textTransform: 'uppercase' } }, "Total"))),
                    React.createElement("div", { style: { width: '100%', display: 'flex', flexDirection: 'column', gap: 8, fontSize: 12 } },
                        React.createElement("div", { style: { display: 'flex', justifyContent: 'space-between', alignItems: 'center' } },
                            React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 6 } },
                                React.createElement("span", { style: { width: 10, height: 10, borderRadius: '50%', background: '#ef4444' } }),
                                React.createElement("span", { style: { color: '#475569' } }, "Expired")),
                            React.createElement("span", { style: { fontWeight: 600, color: '#0f172a' } }, "8 (9%)")),
                        React.createElement("div", { style: { display: 'flex', justifyContent: 'space-between', alignItems: 'center' } },
                            React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 6 } },
                                React.createElement("span", { style: { width: 10, height: 10, borderRadius: '50%', background: '#f59e0b' } }),
                                React.createElement("span", { style: { color: '#475569' } }, "Expiring in 30 days")),
                            React.createElement("span", { style: { fontWeight: 600, color: '#0f172a' } }, "23 (24%)")),
                        React.createElement("div", { style: { display: 'flex', justifyContent: 'space-between', alignItems: 'center' } },
                            React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 6 } },
                                React.createElement("span", { style: { width: 10, height: 10, borderRadius: '50%', background: '#10b981' } }),
                                React.createElement("span", { style: { color: '#475569' } }, "Valid")),
                            React.createElement("span", { style: { fontWeight: 600, color: '#0f172a' } }, "65 (67%)"))))))));
}
