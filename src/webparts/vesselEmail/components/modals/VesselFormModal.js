"use strict";
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
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
exports.renderVesselForm = renderVesselForm;
var React = __importStar(require("react"));
var constants_1 = require("../constants");
function renderVesselForm(host, mode) {
    return React.createElement(VesselFormContent, { host: host, mode: mode });
}
function VesselFormContent(_a) {
    var host = _a.host, mode = _a.mode;
    var _b = host.state, form = _b.form, modalBusy = _b.modalBusy, modalMsg = _b.modalMsg, modalError = _b.modalError;
    var isCreate = mode === 'create';
    var _c = React.useState(0), elapsed = _c[0], setElapsed = _c[1];
    React.useEffect(function () {
        if (!modalBusy) {
            setElapsed(0);
            return;
        }
        var timer = setInterval(function () { return setElapsed(function (e) { return e + 1; }); }, 1000);
        return function () { return clearInterval(timer); };
    }, [modalBusy]);
    var formatTimer = function (sec) {
        var m = Math.floor(sec / 60);
        var s = sec % 60;
        return "".concat(m < 10 ? '0' + m : m, ":").concat(s < 10 ? '0' + s : s);
    };
    var set = function (k) { return function (e) {
        var _a;
        return host.setState({ form: __assign(__assign({}, form), (_a = {}, _a[k] = e.target.value, _a)), modalError: null });
    }; };
    var isSuccess = modalMsg && modalMsg.startsWith('🎉');
    var handleClose = function () {
        host.setState({ modal: 'none', modalMsg: null, modalError: null });
        if (isSuccess) {
            host._loadData();
        }
    };
    return (React.createElement("div", { style: { position: 'fixed', inset: 0, background: 'rgba(15,23,42,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999 }, onClick: function (e) { if (e.target === e.currentTarget && !modalBusy)
            handleClose(); } },
        React.createElement("div", { style: { background: '#fff', borderRadius: 20, padding: '32px 36px', width: 480, maxWidth: '92vw', boxShadow: '0 24px 64px rgba(0,0,0,0.28)', position: 'relative' } },
            !modalBusy && (React.createElement("button", { onClick: handleClose, style: { position: 'absolute', top: 14, right: 14, background: 'none', border: 'none', fontSize: 18, color: '#94a3b8', cursor: 'pointer' }, title: "Close" }, "\u2715")),
            modalBusy ? (React.createElement("div", { style: { textAlign: 'center', padding: '12px 0' } },
                React.createElement("div", { style: { fontSize: 44, marginBottom: 12 } }, "\u23F3"),
                React.createElement("h3", { style: { margin: '0 0 6px', fontSize: 18, fontWeight: 700, color: '#0284c7' } }, "Creating & Provisioning Vessel\u2026"),
                React.createElement("p", { style: { margin: '0 0 16px', fontSize: 13, color: '#475569', lineHeight: 1.5 } },
                    "Registering ",
                    React.createElement("strong", null,
                        "\"",
                        form.name || 'Vessel',
                        "\""),
                    " in DMS database & provisioning SharePoint folder structure. Please wait."),
                React.createElement("div", { style: {
                        display: 'inline-flex', alignItems: 'center', gap: 8,
                        background: '#e0f2fe', color: '#0369a1', border: '1px solid #bae6fd',
                        borderRadius: 20, padding: '8px 18px', fontSize: 14, fontWeight: 700, marginBottom: 20,
                    } },
                    React.createElement("span", { style: { fontSize: 16 } }, "\u23F1\uFE0F"),
                    React.createElement("span", null,
                        "Elapsed Time: ",
                        formatTimer(elapsed))),
                React.createElement("div", { style: { background: '#e2e8f0', borderRadius: 10, height: 8, overflow: 'hidden', marginBottom: 20 } },
                    React.createElement("div", { style: {
                            background: 'linear-gradient(90deg, #0ea5e9, #0284c7, #38bdf8)',
                            height: '100%', width: "".concat(Math.min(96, 20 + elapsed * 15), "%"),
                            transition: 'width 0.8s ease-out', borderRadius: 10,
                        } })),
                React.createElement("div", { style: { background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 12, padding: '14px 16px', textAlign: 'left', fontSize: 12 } },
                    React.createElement("div", { style: { color: '#16a34a', fontWeight: 600, marginBottom: 8, display: 'flex', alignItems: 'center', gap: 8 } },
                        React.createElement("span", { style: { fontSize: 14 } }, "\u2713"),
                        React.createElement("span", null, "Vessel record registered in database")),
                    React.createElement("div", { style: { color: '#0284c7', fontWeight: 600, marginBottom: 8, display: 'flex', alignItems: 'center', gap: 8 } },
                        React.createElement("span", { style: { fontSize: 14 } }, "\u23F3"),
                        React.createElement("span", null, "Creating SharePoint DMS folder tree (Technical & Crewing, Commercial, Insurance)\u2026")),
                    React.createElement("div", { style: { color: '#94a3b8', display: 'flex', alignItems: 'center', gap: 8 } },
                        React.createElement("span", { style: { fontSize: 14 } }, "\u25CB"),
                        React.createElement("span", null, "Linking category subfolders & permissions"))))) : isSuccess ? (
            /* ── STATE 2: SUCCESS SCREEN ── */
            React.createElement("div", { style: { textAlign: 'center', padding: '12px 0' } },
                React.createElement("div", { style: { fontSize: 52, marginBottom: 12 } }, "\uD83C\uDF89"),
                React.createElement("h3", { style: { margin: '0 0 8px', fontSize: 20, fontWeight: 700, color: '#059669' } }, "Vessel Successfully Created & Provisioned!"),
                React.createElement("p", { style: { margin: '0 0 20px', fontSize: 13, color: '#475569', lineHeight: 1.5 } },
                    React.createElement("strong", null,
                        "\"",
                        form.name,
                        "\""),
                    " has been registered in the DMS database and its full SharePoint folder structure has been created."),
                React.createElement("div", { style: { background: '#ecfdf5', border: '1px solid #a7f3d0', borderRadius: 10, padding: '14px 16px', marginBottom: 24, fontSize: 12, color: '#065f46', textAlign: 'left' } },
                    React.createElement("div", { style: { fontWeight: 700, marginBottom: 4 } }, "\u2705 Vessel Ready"),
                    React.createElement("div", null, "\u2022 Registered vessel card added to main grid"),
                    React.createElement("div", null,
                        "\u2022 IMO Number: ",
                        form.imo || '—'),
                    React.createElement("div", null, "\u2022 SharePoint DMS department subfolders created")),
                React.createElement("button", { onClick: handleClose, style: {
                        background: 'linear-gradient(135deg, #10b981, #059669)',
                        color: '#fff', border: 'none', borderRadius: 10,
                        padding: '12px 32px', fontSize: 14, fontWeight: 700, cursor: 'pointer',
                        boxShadow: '0 4px 12px rgba(16,185,129,0.3)',
                    } }, "Done / View Vessels"))) : (
            /* ── STATE 3: FORM ENTRY ── */
            React.createElement(React.Fragment, null,
                React.createElement("div", { style: { display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 24 } },
                    React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 14 } },
                        React.createElement("div", { style: { width: 44, height: 44, borderRadius: 10, background: '#e0f2fe', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 22 } }, "\uD83D\uDEA2"),
                        React.createElement("div", null,
                            React.createElement("div", { style: { fontSize: 18, fontWeight: 700, color: '#0f172a' } }, isCreate ? 'New Vessel' : 'Update Vessel'),
                            React.createElement("div", { style: { fontSize: 12, color: '#64748b', marginTop: 2, maxWidth: 280 } }, isCreate ? 'Provisions the full folder structure across all 3 main folders.' : 'Update the vessel details below.')))),
                modalMsg && React.createElement("div", { style: { background: '#dff6dd', color: '#107c10', padding: '8px 12px', borderRadius: 6, fontSize: 12, marginBottom: 14 } }, modalMsg),
                modalError && React.createElement("div", { style: { background: '#fde7e9', color: '#a4262c', padding: '8px 12px', borderRadius: 6, fontSize: 12, marginBottom: 14 } }, modalError),
                React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 16 } },
                    React.createElement("div", null,
                        React.createElement("label", { style: { display: 'block', fontSize: 13, fontWeight: 600, color: '#334155', marginBottom: 6 } },
                            "Vessel name ",
                            React.createElement("span", { style: { color: '#ef4444' } }, "*")),
                        React.createElement("input", { style: { width: '100%', padding: '10px 14px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 14, outline: 'none', boxSizing: 'border-box', color: '#0f172a' }, value: form.name, onChange: set('name'), placeholder: "e.g. MV Pacific Trader" })),
                    React.createElement("div", null,
                        React.createElement("label", { style: { display: 'block', fontSize: 13, fontWeight: 600, color: '#334155', marginBottom: 6 } },
                            "IMO number ",
                            React.createElement("span", { style: { color: '#ef4444' } }, "*")),
                        React.createElement("input", { style: { width: '100%', padding: '10px 14px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 14, outline: 'none', boxSizing: 'border-box', color: '#0f172a' }, value: form.imo, onChange: function (e) {
                                var val = e.target.value.replace(/\D/g, '').slice(0, 7);
                                host.setState({ form: __assign(__assign({}, form), { imo: val }), modalError: null });
                            }, placeholder: "7 digits, e.g. 9074729", maxLength: 7, inputMode: "numeric" })),
                    React.createElement("div", null,
                        React.createElement("label", { style: { display: 'block', fontSize: 13, fontWeight: 600, color: '#334155', marginBottom: 6 } }, "Ship yard name"),
                        React.createElement("input", { style: { width: '100%', padding: '10px 14px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 14, outline: 'none', boxSizing: 'border-box', color: '#0f172a' }, value: form.shipyard, onChange: set('shipyard'), placeholder: "e.g. Hyundai Heavy Industries" })),
                    React.createElement("div", null,
                        React.createElement("label", { style: { display: 'block', fontSize: 13, fontWeight: 600, color: '#334155', marginBottom: 6 } }, "Hull number"),
                        React.createElement("input", { style: { width: '100%', padding: '10px 14px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 14, outline: 'none', boxSizing: 'border-box', color: '#0f172a' }, value: form.hull_number, onChange: set('hull_number'), placeholder: "e.g. H2456" })),
                    React.createElement("div", null,
                        React.createElement("label", { style: { display: 'block', fontSize: 13, fontWeight: 600, color: '#334155', marginBottom: 6 } }, "Vessel type"),
                        React.createElement("select", { style: { width: '100%', padding: '10px 14px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 14, outline: 'none', background: '#fff', color: form.vessel_type ? '#0f172a' : '#94a3b8', appearance: 'auto' }, value: form.vessel_type, onChange: set('vessel_type') },
                            React.createElement("option", { value: "" }, "Select a type..."),
                            constants_1.VESSEL_TYPES.map(function (t) { return React.createElement("option", { key: t, value: t }, t); })))),
                React.createElement("div", { style: { display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 28 } },
                    React.createElement("button", { style: { background: 'transparent', border: 'none', borderRadius: 8, padding: '10px 20px', fontSize: 14, fontWeight: 500, color: '#475569', cursor: 'pointer' }, onClick: handleClose }, "Cancel"),
                    React.createElement("button", { style: { background: 'linear-gradient(135deg, #0d9488, #0f766e)', color: '#fff', border: 'none', borderRadius: 8, padding: '10px 24px', fontSize: 14, fontWeight: 600, cursor: 'pointer' }, onClick: isCreate ? host._submitCreate : host._submitEdit }, isCreate ? 'Create Vessel' : 'Update Vessel')))))));
}
