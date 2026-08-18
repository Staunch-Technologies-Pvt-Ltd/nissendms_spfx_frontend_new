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
    var _a = host.state, vessels = _a.vessels, deleteVesselIds = _a.deleteVesselIds, modalBusy = _a.modalBusy, modalMsg = _a.modalMsg, modalError = _a.modalError;
    var selectedCount = deleteVesselIds.size;
    var allSelected = vessels.length > 0 && vessels.every(function (v) { return deleteVesselIds.has(v.id); });
    var toggleVessel = function (id) {
        host.setState(function (prev) {
            var next = new Set(prev.deleteVesselIds);
            if (next.has(id))
                next.delete(id);
            else
                next.add(id);
            return { deleteVesselIds: next, modalError: null };
        });
    };
    var toggleAll = function () {
        host.setState({
            deleteVesselIds: allSelected ? new Set() : new Set(vessels.map(function (v) { return v.id; })),
            modalError: null,
        });
    };
    return (React.createElement("div", { style: { position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999 }, onClick: host._closeModal },
        React.createElement("div", { style: { background: '#fff', borderRadius: 10, padding: '24px 28px', width: 640, maxWidth: '90vw', boxShadow: '0 8px 32px rgba(0,0,0,0.2)' }, onClick: function (e) { return e.stopPropagation(); } },
            React.createElement("div", { style: { fontSize: 16, fontWeight: 700, marginBottom: 8 } }, "Delete vessels"),
            React.createElement("p", { style: { fontSize: 13, color: '#475569', margin: '0 0 14px' } }, "Select one or more vessels. This is temporary: the vessel folder and its contents move to the Recycle Bin and can be restored."),
            React.createElement("div", { style: { border: '1px solid #e2e8f0', borderRadius: 8, overflow: 'hidden', marginBottom: 14 } },
                React.createElement("label", { style: { display: 'flex', gap: 10, alignItems: 'center', padding: '10px 12px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', fontSize: 13, fontWeight: 700, cursor: modalBusy ? 'default' : 'pointer' } },
                    React.createElement("input", { type: "checkbox", checked: allSelected, onChange: toggleAll, disabled: modalBusy }),
                    "Select all created vessels (",
                    vessels.length,
                    ")"),
                React.createElement("div", { style: { maxHeight: 260, overflowY: 'auto' } }, vessels.length === 0 ? (React.createElement("div", { style: { padding: 16, color: '#64748b', fontSize: 13 } }, "No vessels are available to delete.")) : vessels.map(function (vessel) { return (React.createElement("label", { key: vessel.id, style: { display: 'flex', gap: 10, alignItems: 'center', padding: '10px 12px', borderBottom: '1px solid #f1f5f9', fontSize: 13, cursor: modalBusy ? 'default' : 'pointer', background: deleteVesselIds.has(vessel.id) ? '#fff1f2' : '#fff' } },
                    React.createElement("input", { type: "checkbox", checked: deleteVesselIds.has(vessel.id), onChange: function () { return toggleVessel(vessel.id); }, disabled: modalBusy }),
                    React.createElement("span", { style: { flex: 1, fontWeight: 600, color: '#0f172a' } }, vessel.name),
                    React.createElement("span", { style: { fontSize: 11, color: '#64748b' } },
                        vessel.vessel_type || 'Vessel',
                        vessel.imo ? " | IMO ".concat(vessel.imo) : ''))); }))),
            modalMsg && React.createElement("div", { style: { background: '#dff6dd', color: '#107c10', padding: 8, borderRadius: 6, fontSize: 12, marginBottom: 12 } }, modalMsg),
            modalError && React.createElement("div", { style: { background: '#fde7e9', color: '#a4262c', padding: 8, borderRadius: 6, fontSize: 12, marginBottom: 12 } }, modalError),
            React.createElement("div", { style: { display: 'flex', justifyContent: 'flex-end', gap: 8 } },
                React.createElement("button", { style: { background: '#f1f5f9', border: '1px solid #cbd5e1', borderRadius: 6, padding: '8px 16px', fontSize: 13, cursor: 'pointer' }, onClick: host._closeModal, disabled: modalBusy }, "Cancel"),
                React.createElement("button", { style: { background: '#ef4444', color: '#fff', border: 'none', borderRadius: 6, padding: '8px 16px', fontSize: 13, fontWeight: 600, cursor: selectedCount ? 'pointer' : 'not-allowed', opacity: selectedCount ? 1 : 0.55 }, onClick: host._submitDelete, disabled: modalBusy || selectedCount === 0 }, modalBusy ? 'Moving to Recycle Bin...' : "Delete selected (".concat(selectedCount, ")"))))));
}
