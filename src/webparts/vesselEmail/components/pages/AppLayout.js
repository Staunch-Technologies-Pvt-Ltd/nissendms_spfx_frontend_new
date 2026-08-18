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
exports.renderLayout = renderLayout;
var React = __importStar(require("react"));
function renderLayout(host, content) {
    var userDisplayName = host.props.userDisplayName || 'Priya';
    return (React.createElement("div", { style: { display: 'flex', alignItems: 'stretch', minHeight: '100vh', background: '#f8fafc', width: '100%', fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif" } },
        host._renderSidebar(),
        React.createElement("div", { style: { flex: 1, display: 'flex', flexDirection: 'column', minWidth: 0 } },
            React.createElement("div", { style: {
                    height: 52, background: '#ffffff', borderBottom: '1px solid #e2e8f0',
                    display: 'flex', alignItems: 'center', justifyContent: 'space-between',
                    padding: '0 24px', flexShrink: 0,
                } },
                React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 12 } },
                    React.createElement("span", { style: { color: '#64748b', fontSize: 14, fontWeight: 500 } }, "Vessel DMS"),
                    React.createElement("span", { style: { color: '#cbd5e1' } }, "/"),
                    React.createElement("span", { style: { color: '#0f172a', fontSize: 14, fontWeight: 600, textTransform: 'capitalize' } }, host.state.view.replace('_', ' '))),
                React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 16 } },
                    React.createElement("div", { style: { position: 'relative', width: 220 } },
                        React.createElement("input", { type: "text", placeholder: "Search host site...", style: {
                                width: '100%', padding: '6px 12px 6px 30px', borderRadius: 16,
                                border: '1px solid #cbd5e1', fontSize: 12, background: '#f8fafc',
                                outline: 'none', boxSizing: 'border-box',
                            } }),
                        React.createElement("span", { style: { position: 'absolute', left: 10, top: '50%', transform: 'translateY(-50%)', color: '#94a3b8', fontSize: 12 } }, "\uD83D\uDD0D")),
                    React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 8 } },
                        React.createElement("div", { style: {
                                width: 28, height: 28, borderRadius: '50%', background: '#0078d4',
                                color: '#fff', fontSize: 12, fontWeight: 600, display: 'flex',
                                alignItems: 'center', justifyContent: 'center',
                            } }, userDisplayName.charAt(0)),
                        React.createElement("span", { style: { fontSize: 13, fontWeight: 600, color: '#334155' } }, userDisplayName)))),
            React.createElement("div", { style: { flex: 1, padding: 24, overflowY: 'auto' } }, content),
            host._renderDocPreviewDrawer())));
}
