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
exports.renderSettingsPage = renderSettingsPage;
var React = __importStar(require("react"));
function renderSettingsPage(host) {
    var _a = host.state, settingsTab = _a.settingsTab, settingsForm = _a.settingsForm, settingsSavedMsg = _a.settingsSavedMsg;
    return (React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 16 } },
        React.createElement("div", null,
            React.createElement("h2", { style: { margin: 0, fontSize: 20, fontWeight: 700, color: '#0f172a' } }, "Settings"),
            React.createElement("p", { style: { margin: '4px 0 0', fontSize: 13, color: '#64748b' } }, "Configure application settings and preferences.")),
        React.createElement("div", { style: { display: 'grid', gridTemplateColumns: '200px 1fr', gap: 24, background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', padding: 24 } },
            React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 4, borderRight: '1px solid #f1f5f9', paddingRight: 16 } }, [
                'General', 'Document Settings', 'Notification Settings',
                'Permission Settings', 'Integration', 'Audit Logs'
            ].map(function (tab) { return (React.createElement("button", { key: tab, onClick: function () { return host.setState({ settingsTab: tab }); }, style: {
                    border: 'none', background: settingsTab === tab ? '#eff6ff' : 'transparent',
                    color: settingsTab === tab ? '#0078d4' : '#475569', fontWeight: settingsTab === tab ? 700 : 500,
                    fontSize: 13, padding: '8px 12px', borderRadius: 6, textAlign: 'left', cursor: 'pointer',
                } }, tab)); })),
            React.createElement("div", { style: { maxWidth: 500 } },
                React.createElement("h3", { style: { margin: '0 0 16px', fontSize: 15, fontWeight: 700, color: '#0f172a' } }, settingsTab),
                settingsSavedMsg && (React.createElement("div", { style: { background: '#dff6dd', color: '#107c10', padding: '8px 12px', borderRadius: 6, fontSize: 13, marginBottom: 16 } }, "\u2713 Settings saved successfully.")),
                settingsTab === 'General' && (React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 14 } },
                    React.createElement("div", null,
                        React.createElement("label", { style: { display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 } }, "Site Title"),
                        React.createElement("input", { type: "text", value: settingsForm.siteTitle, onChange: function (e) { return host.setState({ settingsForm: __assign(__assign({}, settingsForm), { siteTitle: e.target.value }) }); }, style: { width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, outline: 'none', boxSizing: 'border-box' } })),
                    React.createElement("div", null,
                        React.createElement("label", { style: { display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 } }, "Site Description"),
                        React.createElement("textarea", { value: settingsForm.siteDescription, onChange: function (e) { return host.setState({ settingsForm: __assign(__assign({}, settingsForm), { siteDescription: e.target.value }) }); }, style: { width: '100%', height: 60, padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, outline: 'none', boxSizing: 'border-box', resize: 'vertical' } })),
                    React.createElement("div", null,
                        React.createElement("label", { style: { display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 } }, "Date Format"),
                        React.createElement("select", { value: settingsForm.dateFormat, onChange: function (e) { return host.setState({ settingsForm: __assign(__assign({}, settingsForm), { dateFormat: e.target.value }) }); }, style: { width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, outline: 'none', background: '#fff' } },
                            React.createElement("option", { value: "MM/DD/YYYY" }, "MM/DD/YYYY"),
                            React.createElement("option", { value: "DD/MM/YYYY" }, "DD/MM/YYYY"),
                            React.createElement("option", { value: "YYYY-MM-DD" }, "YYYY-MM-DD"))),
                    React.createElement("div", null,
                        React.createElement("label", { style: { display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 } }, "Time Zone"),
                        React.createElement("select", { value: settingsForm.timeZone, onChange: function (e) { return host.setState({ settingsForm: __assign(__assign({}, settingsForm), { timeZone: e.target.value }) }); }, style: { width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, outline: 'none', background: '#fff' } },
                            React.createElement("option", { value: "(UTC+05:30) Chennai, Kolkata, Mumbai, New Delhi" }, "(UTC+05:30) Chennai, Kolkata, Mumbai, New Delhi"),
                            React.createElement("option", { value: "(UTC+00:00) UTC" }, "(UTC+00:00) UTC"))),
                    React.createElement("div", { style: { marginTop: 12 } },
                        React.createElement("button", { onClick: function () {
                                host.setState({ settingsSavedMsg: true });
                                setTimeout(function () { return host.setState({ settingsSavedMsg: false }); }, 2000);
                            }, style: {
                                background: '#0078d4', color: '#fff', border: 'none', borderRadius: 6,
                                padding: '8px 20px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
                            } }, "Save Changes")))),
                settingsTab !== 'General' && (React.createElement("div", { style: { color: '#64748b', fontSize: 13 } },
                    "Configuration settings for ",
                    settingsTab,
                    " are active with system default policies."))))));
}
