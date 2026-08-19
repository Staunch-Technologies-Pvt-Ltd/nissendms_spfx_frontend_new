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
exports.renderDeleteModal = renderDeleteModal;
var React = __importStar(require("react"));
/** A single dialog supports both one-vessel and bulk temporary deletion. */
function renderDeleteModal(host) {
    var _a = host.state, selectedVessel = _a.selectedVessel, modalBusy = _a.modalBusy, modalMsg = _a.modalMsg, modalError = _a.modalError;
    var isSuccess = !!modalMsg && !modalBusy && !modalError;
    React.useEffect(function () {
        if (!isSuccess)
            return undefined;
        var timer = setTimeout(function () { return host._closeModal(); }, 15000);
        return function () { return clearTimeout(timer); };
    }, [isSuccess, host]);
    var content = isSuccess
        ? React.createElement("div", { style: { textAlign: 'center', padding: '16px 0 8px' } },
            React.createElement("div", { style: { fontSize: 50, marginBottom: 10 } }, "✅"),
            React.createElement("div", { style: { fontSize: 20, fontWeight: 700, color: '#059669', marginBottom: 8 } }, "Vessel Deleted Successfully"),
            React.createElement("p", { style: { fontSize: 13, color: '#475569', margin: '0 0 20px', lineHeight: 1.5 } }, modalMsg),
            React.createElement("button", { onClick: host._closeModal, style: { background: 'linear-gradient(135deg, #10b981, #059669)', color: '#fff', border: 'none', borderRadius: 8, padding: '10px 28px', fontSize: 13, fontWeight: 700, cursor: 'pointer' } }, "Done"))
        : React.createElement(React.Fragment, null,
            React.createElement("div", { style: { fontSize: 16, fontWeight: 700, marginBottom: 8 } }, "Delete vessel"),
            React.createElement("p", { style: { fontSize: 13, color: '#475569', margin: '0 0 14px' } }, "This is temporary: the vessel folder and its contents will move to the Recycle Bin and can be restored."),
            selectedVessel ? React.createElement("div", { style: { background: '#fff1f2', border: '1px solid #fecdd3', borderRadius: 8, padding: '14px 16px', marginBottom: 14 } },
                React.createElement("div", { style: { fontSize: 15, fontWeight: 700, color: '#881337' } }, selectedVessel.name),
                React.createElement("div", { style: { marginTop: 4, fontSize: 11, color: '#9f1239' } }, selectedVessel.vessel_type || 'Vessel', selectedVessel.imo ? " | IMO ".concat(selectedVessel.imo) : '')) : React.createElement("div", { style: { padding: 16, color: '#64748b', fontSize: 13 } }, "No vessel selected."),
            modalMsg && React.createElement("div", { style: { background: '#dff6dd', color: '#107c10', padding: 8, borderRadius: 6, fontSize: 12, marginBottom: 12 } }, modalMsg),
            modalError && React.createElement("div", { style: { background: '#fde7e9', color: '#a4262c', padding: 8, borderRadius: 6, fontSize: 12, marginBottom: 12 } }, modalError),
            React.createElement("div", { style: { display: 'flex', justifyContent: 'flex-end', gap: 8 } },
                React.createElement("button", { style: { background: '#f1f5f9', border: '1px solid #cbd5e1', borderRadius: 6, padding: '8px 16px', fontSize: 13, cursor: 'pointer' }, onClick: host._closeModal, disabled: modalBusy }, "Cancel"),
                React.createElement("button", { style: { background: '#ef4444', color: '#fff', border: 'none', borderRadius: 6, padding: '8px 16px', fontSize: 13, fontWeight: 600, cursor: selectedVessel ? 'pointer' : 'not-allowed', opacity: selectedVessel ? 1 : 0.55 }, onClick: host._submitDelete, disabled: modalBusy || !selectedVessel }, modalBusy ? 'Moving to Recycle Bin...' : 'Delete and Move to Recycle Bin')));
    return React.createElement("div", { style: { position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999 }, onClick: host._closeModal },
        React.createElement("div", { style: { background: '#fff', borderRadius: 10, padding: '24px 28px', width: 520, maxWidth: '90vw', boxShadow: '0 8px 32px rgba(0,0,0,0.2)' }, onClick: function (e) { return e.stopPropagation(); } }, content));
}
