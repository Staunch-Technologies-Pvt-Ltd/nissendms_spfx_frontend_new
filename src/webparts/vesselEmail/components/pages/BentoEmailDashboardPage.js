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
var __spreadArray = (this && this.__spreadArray) || function (to, from, pack) {
    if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
        if (ar || !(i in from)) {
            if (!ar) ar = Array.prototype.slice.call(from, 0, i);
            ar[i] = from[i];
        }
    }
    return to.concat(ar || Array.prototype.slice.call(from));
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.renderBentoEmailDashboardPage = renderBentoEmailDashboardPage;
var React = __importStar(require("react"));
var constants_1 = require("../constants");
function renderBentoEmailDashboardPage(host) {
    var _a = host.state, bentoLogs = _a.bentoLogs, panelLoading = _a.panelLoading, bentoStatusFilter = _a.bentoStatusFilter, bentoSearch = _a.bentoSearch, bentoComposeOpen = _a.bentoComposeOpen, bentoComposeForm = _a.bentoComposeForm, bentoComposeBusy = _a.bentoComposeBusy, bentoComposeMsg = _a.bentoComposeMsg, bentoComposeErr = _a.bentoComposeErr, bentoDetailLog = _a.bentoDetailLog, vessels = _a.vessels, bentoUploadFile = _a.bentoUploadFile, bentoUploadVessel = _a.bentoUploadVessel, bentoUploadTag = _a.bentoUploadTag, bentoUploadBusy = _a.bentoUploadBusy, bentoUploadMsg = _a.bentoUploadMsg, bentoUploadErr = _a.bentoUploadErr;
    var totalCount = bentoLogs.length;
    var completedCount = bentoLogs.filter(function (l) { return l.status === 'completed' || l.status === 'success' || l.status === 'sent'; }).length;
    var pendingCount = bentoLogs.filter(function (l) { return l.status === 'pending'; }).length;
    var failedCount = bentoLogs.filter(function (l) { return l.status === 'failed' || l.status === 'error'; }).length;
    var q = bentoSearch.toLowerCase();
    var filtered = bentoLogs.filter(function (log) {
        var _a, _b, _c, _d;
        var st = log.status === 'success' || log.status === 'sent' ? 'completed' : log.status;
        if (bentoStatusFilter !== 'all' && st !== bentoStatusFilter)
            return false;
        if (q) {
            return __spreadArray([
                (_a = log.vessel_name) !== null && _a !== void 0 ? _a : '',
                (_b = log.subject) !== null && _b !== void 0 ? _b : '',
                (_c = log.datasource_tag_used) !== null && _c !== void 0 ? _c : '',
                (_d = log.recipient) !== null && _d !== void 0 ? _d : ''
            ], (log.attachment_names || []), true).some(function (s) { return s.toLowerCase().indexOf(q) !== -1; });
        }
        return true;
    });
    return (React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 16 } },
        React.createElement("div", { style: { background: 'linear-gradient(135deg, #0f172a 0%, #1e293b 100%)', borderRadius: 12, padding: '24px 28px', color: '#fff', boxShadow: '0 4px 16px rgba(0,0,0,0.15)' } },
            React.createElement("div", { style: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 16 } },
                React.createElement("div", null,
                    React.createElement("div", { style: { display: 'inline-flex', alignItems: 'center', gap: 8, background: 'rgba(255,255,255,0.1)', padding: '4px 12px', borderRadius: 20, fontSize: 12, fontWeight: 600, color: '#38bdf8', marginBottom: 8 } }, "\uD83E\uDD16 AI BENTO AUTOMATION ENGINE"),
                    React.createElement("h1", { style: { margin: 0, fontSize: 22, fontWeight: 800 } }, "AI Bento Email Dashboard"),
                    React.createElement("p", { style: { margin: '6px 0 0', fontSize: 13, color: '#94a3b8', maxWidth: 650 } }, "Automated document tagging, status tracking, and Graph email dispatching.")),
                React.createElement("button", { style: { background: 'linear-gradient(135deg, #38bdf8, #0284c7)', color: '#fff', border: 'none', borderRadius: 8, padding: '10px 22px', fontSize: 13, fontWeight: 700, cursor: 'pointer' }, onClick: function () { return host.setState({
                        bentoComposeOpen: true,
                        bentoComposeMsg: null,
                        bentoComposeErr: null,
                        bentoComposeForm: {
                            vessel_name: '',
                            datasource_tag: 'mail',
                            subject_text: host._buildAutoSubject('', 'mail', ''),
                            body: '',
                            file: null,
                            existing_attachment: '',
                            recipient: '',
                        },
                    }); } }, "\u2709 Compose & Dispatch Email")),
            React.createElement("div", { style: { display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: 14, marginTop: 22 } },
                React.createElement("div", { style: { background: 'rgba(255,255,255,0.06)', borderRadius: 10, padding: '14px 18px', border: '1px solid rgba(255,255,255,0.1)' } },
                    React.createElement("div", { style: { fontSize: 11, fontWeight: 700, textTransform: 'uppercase', color: '#94a3b8' } }, "Total Processed"),
                    React.createElement("div", { style: { fontSize: 26, fontWeight: 800, color: '#fff', marginTop: 4 } }, totalCount)),
                React.createElement("div", { style: { background: 'rgba(16, 185, 129, 0.12)', borderRadius: 10, padding: '14px 18px', border: '1px solid rgba(16, 185, 129, 0.25)' } },
                    React.createElement("div", { style: { fontSize: 11, fontWeight: 700, textTransform: 'uppercase', color: '#6ee7b7' } }, "Completed"),
                    React.createElement("div", { style: { fontSize: 26, fontWeight: 800, color: '#34d399', marginTop: 4 } }, completedCount)),
                React.createElement("div", { style: { background: 'rgba(245, 158, 11, 0.12)', borderRadius: 10, padding: '14px 18px', border: '1px solid rgba(245, 158, 11, 0.25)' } },
                    React.createElement("div", { style: { fontSize: 11, fontWeight: 700, textTransform: 'uppercase', color: '#fcd34d' } }, "Pending"),
                    React.createElement("div", { style: { fontSize: 26, fontWeight: 800, color: '#fbbf24', marginTop: 4 } }, pendingCount)),
                React.createElement("div", { style: { background: 'rgba(239, 68, 68, 0.12)', borderRadius: 10, padding: '14px 18px', border: '1px solid rgba(239, 68, 68, 0.25)' } },
                    React.createElement("div", { style: { fontSize: 11, fontWeight: 700, textTransform: 'uppercase', color: '#fca5a5' } }, "Failed"),
                    React.createElement("div", { style: { fontSize: 26, fontWeight: 800, color: '#f87171', marginTop: 4 } }, failedCount)))),
        React.createElement("div", { style: { background: '#fff', borderRadius: 8, border: '1px solid #e2e8f0', overflowX: 'auto' } },
            React.createElement("table", { style: { width: '100%', minWidth: 900, borderCollapse: 'collapse', fontSize: 13 } },
                React.createElement("thead", null,
                    React.createElement("tr", { style: { background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 11, textTransform: 'uppercase', textAlign: 'left', whiteSpace: 'nowrap' } },
                        React.createElement("th", { style: { padding: '10px 16px' } }, "#"),
                        React.createElement("th", { style: { padding: '10px 16px' } }, "Vessel"),
                        React.createElement("th", { style: { padding: '10px 16px' } }, "Tag"),
                        React.createElement("th", { style: { padding: '10px 16px' } }, "Subject"),
                        React.createElement("th", { style: { padding: '10px 16px' } }, "Recipient"),
                        React.createElement("th", { style: { padding: '10px 16px' } }, "Status"),
                        React.createElement("th", { style: { padding: '10px 16px' } }, "Attached File"))),
                React.createElement("tbody", null, filtered.map(function (log) { return (React.createElement("tr", { key: log.id, style: { borderBottom: '1px solid #f1f5f9' } },
                    React.createElement("td", { style: { padding: '10px 16px', color: '#64748b', whiteSpace: 'nowrap' } },
                        "#",
                        log.id),
                    React.createElement("td", { style: { padding: '10px 16px', fontWeight: 600, whiteSpace: 'nowrap' } }, log.vessel_name || '—'),
                    React.createElement("td", { style: { padding: '10px 16px', whiteSpace: 'nowrap' } }, (0, constants_1.badge)('blue', log.datasource_tag_used)),
                    React.createElement("td", { style: { padding: '10px 16px', color: '#1e293b', minWidth: 200 } }, log.subject),
                    React.createElement("td", { style: { padding: '10px 16px', color: '#475569', whiteSpace: 'nowrap' } }, log.recipient),
                    React.createElement("td", { style: { padding: '10px 16px', whiteSpace: 'nowrap' } }, (0, constants_1.badge)(log.status === 'completed' || log.status === 'success' ? 'green' : log.status === 'pending' ? 'orange' : 'red', log.status)),
                    React.createElement("td", { style: { padding: '10px 16px', whiteSpace: 'nowrap' } }, log.attachment_names && log.attachment_names.length > 0
                        ? log.attachment_names.map(function (name, i) { return (React.createElement("span", { key: i, style: { display: 'inline-flex', alignItems: 'center', gap: 4, background: '#f0f9ff', border: '1px solid #bae6fd', borderRadius: 6, padding: '2px 8px', fontSize: 11, color: '#0284c7', fontWeight: 600, marginRight: 4 } },
                            "\uD83D\uDCCE ",
                            name)); })
                        : React.createElement("span", { style: { color: '#94a3b8', fontSize: 11 } }, "\u2014")))); }))))));
}
