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
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g = Object.create((typeof Iterator === "function" ? Iterator : Object).prototype);
    return g.next = verb(0), g["throw"] = verb(1), g["return"] = verb(2), typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.renderApprovalsPage = renderApprovalsPage;
var React = __importStar(require("react"));
var constants_1 = require("../constants");
function renderApprovalsPage(host) {
    var _this = this;
    var _a = host.state, approvalsList = _a.approvalsList, approvalTab = _a.approvalTab, panelLoading = _a.panelLoading;
    var filtered = approvalsList.filter(function (a) { return a.status === approvalTab; });
    return (React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 16 } },
        React.createElement("div", null,
            React.createElement("h2", { style: { margin: 0, fontSize: 20, fontWeight: 700, color: '#0f172a' } }, "Approvals"),
            React.createElement("p", { style: { margin: '4px 0 0', fontSize: 13, color: '#64748b' } }, "Review and take action on pending approvals.")),
        React.createElement("div", { style: { display: 'flex', gap: 8, borderBottom: '1px solid #e2e8f0', paddingBottom: 8 } }, ['Pending', 'Approved', 'Rejected'].map(function (tab) { return (React.createElement("button", { key: tab, onClick: function () { return host.setState({ approvalTab: tab }); }, style: {
                border: 'none', background: approvalTab === tab ? '#eff6ff' : 'transparent',
                color: approvalTab === tab ? '#0078d4' : '#64748b', fontWeight: approvalTab === tab ? 700 : 500,
                fontSize: 13, padding: '6px 14px', borderRadius: 6, cursor: 'pointer',
            } },
            tab,
            " (",
            approvalsList.filter(function (a) { return a.status === tab; }).length,
            ")")); })),
        panelLoading ? (React.createElement("div", { style: { padding: 32, textAlign: 'center', color: '#64748b', fontSize: 13 } }, "\u23F3 Loading approvals...")) : (React.createElement("div", { style: { background: '#fff', borderRadius: 8, border: '1px solid #e2e8f0', overflowX: 'auto' } },
            React.createElement("table", { style: { width: '100%', minWidth: 700, borderCollapse: 'collapse', fontSize: 13 } },
                React.createElement("thead", null,
                    React.createElement("tr", { style: { background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 11, textTransform: 'uppercase', textAlign: 'left' } },
                        React.createElement("th", { style: { padding: '10px 16px', whiteSpace: 'nowrap' } }, "Document Name"),
                        React.createElement("th", { style: { padding: '10px 16px', whiteSpace: 'nowrap' } }, "Vessel"),
                        React.createElement("th", { style: { padding: '10px 16px', whiteSpace: 'nowrap' } }, "Requested By"),
                        React.createElement("th", { style: { padding: '10px 16px', whiteSpace: 'nowrap' } }, "Requested On"),
                        React.createElement("th", { style: { padding: '10px 16px', textAlign: 'right', whiteSpace: 'nowrap' } }, "Action"))),
                React.createElement("tbody", null, filtered.length === 0 ? (React.createElement("tr", null,
                    React.createElement("td", { colSpan: 5, style: { padding: 24, textAlign: 'center', color: '#94a3b8', fontSize: 13 } },
                        "No ",
                        approvalTab.toLowerCase(),
                        " approvals."))) : filtered.map(function (app) { return (React.createElement("tr", { key: app.id, style: { borderBottom: '1px solid #f1f5f9' } },
                    React.createElement("td", { style: { padding: '12px 16px', fontWeight: 600, color: '#1e293b', maxWidth: 200, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }, title: app.documentName }, app.documentName),
                    React.createElement("td", { style: { padding: '12px 16px', color: '#475569', whiteSpace: 'nowrap' } }, app.vessel),
                    React.createElement("td", { style: { padding: '12px 16px', color: '#475569', maxWidth: 180, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }, title: app.requestedBy }, app.requestedBy),
                    React.createElement("td", { style: { padding: '12px 16px', color: '#64748b', whiteSpace: 'nowrap' } }, app.requestedOn),
                    React.createElement("td", { style: { padding: '12px 16px', textAlign: 'right' } }, app.status === 'Pending' ? (React.createElement("div", { style: { display: 'flex', gap: 6, justifyContent: 'flex-end' } },
                        React.createElement("button", { onClick: function () { return __awaiter(_this, void 0, void 0, function () {
                                var userEmail, adminParam, _a, updated, vesselName_1;
                                return __generator(this, function (_b) {
                                    switch (_b.label) {
                                        case 0:
                                            userEmail = host.props.userEmail || '';
                                            adminParam = userEmail ? "?admin=".concat(encodeURIComponent(userEmail)) : '';
                                            _b.label = 1;
                                        case 1:
                                            _b.trys.push([1, 3, , 4]);
                                            return [4 /*yield*/, fetch("".concat(host._base(), "/api/approvals/").concat(app.id, "/approve").concat(adminParam), { method: 'POST', headers: host._headers() })];
                                        case 2:
                                            _b.sent();
                                            return [3 /*break*/, 4];
                                        case 3:
                                            _a = _b.sent();
                                            return [3 /*break*/, 4];
                                        case 4:
                                            updated = approvalsList.map(function (item) { return item.id === app.id ? __assign(__assign({}, item), { status: 'Approved' }) : item; });
                                            host.setState({ approvalsList: updated });
                                            host._filesLoadedForFolders.clear();
                                            host._filesLoadedForVessels.clear();
                                            host._loadData();
                                            // Re-fetch files for the approved item's vessel so list view updates
                                            if (app.vessel && app.vessel !== '—') {
                                                vesselName_1 = app.vessel;
                                                setTimeout(function () {
                                                    host._filesLoadedForVessels.add(vesselName_1);
                                                    host._loadFilesForVessel(vesselName_1).catch(function () { return undefined; });
                                                }, 2000);
                                            }
                                            return [2 /*return*/];
                                    }
                                });
                            }); }, style: { border: 'none', background: '#dff6dd', color: '#107c10', width: 28, height: 28, borderRadius: 4, cursor: 'pointer', fontWeight: 700 }, title: "Approve" }, "\u2713"),
                        React.createElement("button", { onClick: function () { return __awaiter(_this, void 0, void 0, function () {
                                var userEmail, adminParam, _a, updated;
                                return __generator(this, function (_b) {
                                    switch (_b.label) {
                                        case 0:
                                            userEmail = host.props.userEmail || '';
                                            adminParam = userEmail ? "?admin=".concat(encodeURIComponent(userEmail)) : '';
                                            _b.label = 1;
                                        case 1:
                                            _b.trys.push([1, 3, , 4]);
                                            return [4 /*yield*/, fetch("".concat(host._base(), "/api/approvals/").concat(app.id, "/reject").concat(adminParam), {
                                                    method: 'POST',
                                                    headers: __assign(__assign({}, host._headers()), { 'Content-Type': 'application/json' }),
                                                    body: JSON.stringify({ reason: null }),
                                                })];
                                        case 2:
                                            _b.sent();
                                            return [3 /*break*/, 4];
                                        case 3:
                                            _a = _b.sent();
                                            return [3 /*break*/, 4];
                                        case 4:
                                            updated = approvalsList.map(function (item) { return item.id === app.id ? __assign(__assign({}, item), { status: 'Rejected' }) : item; });
                                            host.setState({ approvalsList: updated });
                                            return [2 /*return*/];
                                    }
                                });
                            }); }, style: { border: 'none', background: '#fde7e9', color: '#a4262c', width: 28, height: 28, borderRadius: 4, cursor: 'pointer', fontWeight: 700 }, title: "Reject" }, "\u2715"))) : ((0, constants_1.badge)(app.status === 'Approved' ? 'green' : 'red', app.status))))); })))))));
}
