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
exports.renderArchivePage = renderArchivePage;
var React = __importStar(require("react"));
function renderArchivePage(host) {
    var archiveList = host.state.archiveList;
    return (React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 16 } },
        React.createElement("div", { style: { background: '#ede8f5', border: '1px solid #b4a0d4', borderRadius: 8, padding: '16px 20px' } },
            React.createElement("h2", { style: { margin: 0, fontSize: 18, fontWeight: 700, color: '#5c2d91' } }, "\uD83D\uDCE6 Archive"),
            React.createElement("p", { style: { margin: '4px 0 0', fontSize: 12, color: '#5c2d91', opacity: 0.9 } }, "Archived items stored for historical record-keeping.")),
        React.createElement("div", { style: { background: '#fff', borderRadius: 8, border: '1px solid #e2e8f0', overflow: 'hidden' } },
            React.createElement("table", { style: { width: '100%', borderCollapse: 'collapse', fontSize: 13 } },
                React.createElement("thead", null,
                    React.createElement("tr", { style: { background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 11, textTransform: 'uppercase', textAlign: 'left' } },
                        React.createElement("th", { style: { padding: '10px 16px' } }, "Name"),
                        React.createElement("th", { style: { padding: '10px 16px' } }, "Type"),
                        React.createElement("th", { style: { padding: '10px 16px' } }, "Archived At"))),
                React.createElement("tbody", null, archiveList.length === 0 ? (React.createElement("tr", null,
                    React.createElement("td", { colSpan: 3, style: { padding: 20, textAlign: 'center', color: '#64748b' } }, "No items archived."))) : (archiveList.map(function (item) { return (React.createElement("tr", { key: item.id, style: { borderBottom: '1px solid #f1f5f9' } },
                    React.createElement("td", { style: { padding: '10px 16px', fontWeight: 600 } }, item.name),
                    React.createElement("td", { style: { padding: '10px 16px' } }, item.kind),
                    React.createElement("td", { style: { padding: '10px 16px' } }, item.archived_at || '—'))); })))))));
}
