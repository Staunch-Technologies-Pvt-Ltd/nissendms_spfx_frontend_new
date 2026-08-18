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
exports.renderDocPreviewDrawer = renderDocPreviewDrawer;
var React = __importStar(require("react"));
function renderDocPreviewDrawer(host) {
    var selectedDocPreview = host.state.selectedDocPreview;
    if (!selectedDocPreview)
        return null;
    var base = host._base();
    var downloadUrl = selectedDocPreview.fileId ? "".concat(base, "/api/files/").concat(selectedDocPreview.fileId, "/content") : null;
    return (React.createElement("div", { style: {
            position: 'fixed',
            top: 0,
            right: 0,
            bottom: 0,
            width: 380,
            maxWidth: '90vw',
            background: '#ffffff',
            boxShadow: '-6px 0 24px rgba(0, 0, 0, 0.18)',
            zIndex: 10000,
            display: 'flex',
            flexDirection: 'column',
            borderLeft: '1px solid #cbd5e1',
            animation: 'slideIn 0.2s ease-out',
        } },
        React.createElement("div", { style: {
                padding: '16px 20px', background: '#0f172a', color: '#fff',
                display: 'flex', alignItems: 'center', justifyContent: 'space-between',
            } },
            React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 8, fontWeight: 700, fontSize: 15 } },
                React.createElement("span", null, "\uD83D\uDCC4"),
                " Document Overview"),
            React.createElement("button", { onClick: function () { return host.setState({ selectedDocPreview: null }); }, style: { border: 'none', background: 'transparent', color: '#cbd5e1', fontSize: 18, cursor: 'pointer', fontWeight: 700 }, title: "Close Drawer" }, "\u2715")),
        React.createElement("div", { style: { flex: 1, padding: 20, overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: 16 } },
            React.createElement("div", { style: {
                    background: '#f0f9ff', borderRadius: 12, padding: 20, border: '1px solid #bae6fd',
                    display: 'flex', flexDirection: 'column', alignItems: 'center', textAlign: 'center', gap: 8,
                } },
                React.createElement("div", { style: { fontSize: 42 } }, "\uD83D\uDCC4"),
                React.createElement("a", { href: downloadUrl || '#', target: "_blank", rel: "noopener noreferrer", onClick: function (e) {
                        if (!downloadUrl) {
                            e.preventDefault();
                            alert("File \"".concat(selectedDocPreview.fileName, "\" stored locally."));
                        }
                    }, style: {
                        fontSize: 15, fontWeight: 700, color: '#0284c7', textDecoration: 'underline',
                        wordBreak: 'break-all', cursor: 'pointer',
                    } }, selectedDocPreview.fileName),
                React.createElement("span", { style: { fontSize: 11, background: '#dbeafe', color: '#1e40af', padding: '2px 8px', borderRadius: 10, fontWeight: 600 } }, "Uploaded Document")),
            React.createElement("div", { style: { background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', padding: 14, display: 'flex', flexDirection: 'column', gap: 10 } },
                React.createElement("div", { style: { fontSize: 12, fontWeight: 700, color: '#475569', textTransform: 'uppercase', letterSpacing: '0.05em', marginBottom: 2 } }, "Metadata Details"),
                React.createElement("div", { style: { display: 'flex', justifyContent: 'space-between', fontSize: 12, borderBottom: '1px solid #f1f5f9', paddingBottom: 6 } },
                    React.createElement("span", { style: { color: '#64748b' } }, "\uD83D\uDEA2 Vessel Name:"),
                    React.createElement("strong", { style: { color: '#0f172a' } }, selectedDocPreview.vesselName)),
                React.createElement("div", { style: { display: 'flex', justifyContent: 'space-between', fontSize: 12, borderBottom: '1px solid #f1f5f9', paddingBottom: 6 } },
                    React.createElement("span", { style: { color: '#64748b' } }, "\uD83D\uDCC1 Group:"),
                    React.createElement("span", { style: { color: '#2563eb', fontWeight: 600 } }, selectedDocPreview.group)),
                React.createElement("div", { style: { display: 'flex', justifyContent: 'space-between', fontSize: 12, borderBottom: '1px solid #f1f5f9', paddingBottom: 6 } },
                    React.createElement("span", { style: { color: '#64748b' } }, "\uD83C\uDFF7\uFE0F Category:"),
                    React.createElement("strong", { style: { color: '#1e293b' } }, selectedDocPreview.category)),
                React.createElement("div", { style: { display: 'flex', justifyContent: 'space-between', fontSize: 12, borderBottom: '1px solid #f1f5f9', paddingBottom: 6 } },
                    React.createElement("span", { style: { color: '#64748b' } }, "\uD83D\uDCBE Size:"),
                    React.createElement("span", { style: { color: '#475569' } }, selectedDocPreview.size || '142 KB')),
                React.createElement("div", { style: { display: 'flex', justifyContent: 'space-between', fontSize: 12, borderBottom: '1px solid #f1f5f9', paddingBottom: 6 } },
                    React.createElement("span", { style: { color: '#64748b' } }, "\uD83D\uDCC5 Date:"),
                    React.createElement("span", { style: { color: '#475569' } }, selectedDocPreview.date || 'Today')),
                React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 2, fontSize: 12 } },
                    React.createElement("span", { style: { color: '#64748b' } }, "\uD83D\uDDFA\uFE0F Folder Path:"),
                    React.createElement("span", { style: { color: '#334155', fontFamily: 'monospace', fontSize: 11, background: '#f8fafc', padding: '4px 6px', borderRadius: 4, wordBreak: 'break-all' } }, selectedDocPreview.folderPath))),
            React.createElement("div", { style: { background: '#f8fafc', borderRadius: 10, border: '1px dashed #cbd5e1', padding: 14, textAlign: 'center' } },
                React.createElement("div", { style: { fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 } },
                    "Document Status: ",
                    React.createElement("span", { style: { color: '#16a34a', fontWeight: 700 } }, "Active / Ready")),
                React.createElement("p", { style: { margin: 0, fontSize: 11, color: '#64748b' } }, "Click open below to view or download the uploaded document directly."))),
        React.createElement("div", { style: { padding: 16, background: '#f8fafc', borderTop: '1px solid #e2e8f0', display: 'flex', flexDirection: 'column', gap: 8 } },
            React.createElement("a", { href: downloadUrl || '#', target: "_blank", rel: "noopener noreferrer", onClick: function (e) {
                    if (!downloadUrl) {
                        e.preventDefault();
                        alert("File \"".concat(selectedDocPreview.fileName, "\" stored locally."));
                    }
                }, style: {
                    display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6,
                    background: '#0284c7', color: '#fff', padding: '9px 16px', borderRadius: 8,
                    fontSize: 13, fontWeight: 600, textDecoration: 'none', textAlign: 'center',
                } }, "\uD83C\uDF10 Open / View Document"),
            React.createElement("button", { onClick: function () {
                    host.setState({
                        view: 'bento_email',
                        bentoUploadVessel: selectedDocPreview.vesselName,
                        selectedDocPreview: null,
                    });
                }, style: {
                    display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6,
                    background: '#0f172a', color: '#fff', border: 'none', padding: '9px 16px', borderRadius: 8,
                    fontSize: 13, fontWeight: 600, cursor: 'pointer',
                } }, "\u2709 Send via Bento Email"),
            React.createElement("button", { onClick: function () { return host.setState({ selectedDocPreview: null }); }, style: {
                    background: '#fff', border: '1px solid #cbd5e1', color: '#475569', padding: '8px 16px',
                    borderRadius: 8, fontSize: 12, fontWeight: 600, cursor: 'pointer',
                } }, "Close Drawer"))));
}
