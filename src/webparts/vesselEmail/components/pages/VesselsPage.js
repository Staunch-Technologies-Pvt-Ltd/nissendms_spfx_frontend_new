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
exports.renderDismissConfirmDialog = renderDismissConfirmDialog;
exports.renderClassifyDialog = renderClassifyDialog;
exports.renderVesselsPage = renderVesselsPage;
var React = __importStar(require("react"));
var constants_1 = require("../constants");
var vesselImagePool_1 = require("../vesselImagePool");
function getSpoVesselFolderUrl(siteUrlProp, vesselName, department) {
    if (department === void 0) { department = 'Technical & Crewing'; }
    var fallbackSite = 'https://nissenkaiunsingapore.sharepoint.com';
    var site = (siteUrlProp && siteUrlProp !== '#') ? siteUrlProp : fallbackSite;
    try {
        var urlObj = new URL(site);
        var basePath = urlObj.pathname.replace(/\/$/, '');
        var folderPath = vesselName
            ? "".concat(basePath, "/Shared Documents/").concat(department, "/Vessels/").concat(vesselName)
            : "".concat(basePath, "/Shared Documents/").concat(department);
        return "".concat(urlObj.origin).concat(basePath, "/Shared Documents/Forms/AllItems.aspx?id=").concat(encodeURIComponent(folderPath));
    }
    catch (_a) {
        return '#';
    }
}
// ── Dismiss Confirm Dialog ────────────────────────────────────────────────────
// Shown when user clicks "Dismiss" on any anomaly item.
// Gives two options: "Keep it here" (close dialog) or "Move to Recycle Bin".
function renderDismissConfirmDialog(host) {
    var anomaly = host.state.spoAnomalyDismissConfirm;
    if (!anomaly)
        return null;
    return (React.createElement("div", { role: "dialog", "aria-modal": "true", "aria-label": "Dismiss Confirmation", style: {
            position: 'fixed', inset: 0, zIndex: 10000,
            background: 'rgba(15,23,42,0.6)',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            padding: 16,
        }, onClick: function (e) { if (e.target === e.currentTarget)
            host.setState({ spoAnomalyDismissConfirm: null }); } },
        React.createElement("div", { style: {
                background: '#fff', borderRadius: 16, boxShadow: '0 24px 64px rgba(0,0,0,0.28)',
                padding: 32, width: '100%', maxWidth: 440, position: 'relative',
            } },
            React.createElement("button", { onClick: function () { return host.setState({ spoAnomalyDismissConfirm: null }); }, style: { position: 'absolute', top: 14, right: 14, background: 'none', border: 'none', fontSize: 18, cursor: 'pointer', color: '#94a3b8' }, title: "Cancel" }, "\u2715"),
            React.createElement("div", { style: { textAlign: 'center', marginBottom: 20 } },
                React.createElement("div", { style: { fontSize: 44, marginBottom: 10 } }, "\uD83D\uDDD1\uFE0F"),
                React.createElement("h3", { style: { margin: '0 0 8px', fontSize: 17, fontWeight: 700, color: '#0f172a' } }, "What would you like to do?"),
                React.createElement("p", { style: { margin: 0, fontSize: 13, color: '#64748b', lineHeight: 1.5 } },
                    "The ",
                    anomaly.item_type,
                    " ",
                    React.createElement("strong", null,
                        "\"",
                        anomaly.name,
                        "\""),
                    " was found outside the expected SharePoint structure.")),
            React.createElement("div", { style: {
                    background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 8,
                    padding: '8px 12px', marginBottom: 22, fontSize: 11, color: '#475569',
                    fontFamily: 'monospace', wordBreak: 'break-all',
                } },
                "\uD83D\uDCC2 ",
                anomaly.spo_path),
            React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 10 } },
                React.createElement("button", { onClick: function () { return host.setState({ spoAnomalyDismissConfirm: null }); }, style: {
                        background: '#f8fafc', color: '#334155',
                        border: '2px solid #e2e8f0', borderRadius: 10, padding: '14px 18px',
                        cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 12, textAlign: 'left',
                        transition: 'border-color 0.15s',
                    } },
                    React.createElement("span", { style: { fontSize: 26 } }, "\uD83D\uDCCC"),
                    React.createElement("div", null,
                        React.createElement("div", { style: { fontWeight: 700, fontSize: 13, marginBottom: 2, color: '#0f172a' } }, "Keep it here"),
                        React.createElement("div", { style: { fontSize: 12, color: '#64748b' } }, "Leave this item listed in the warning section for now. You can classify or dismiss it later."))),
                React.createElement("button", { onClick: function () { return host._moveAnomalyToRecycleBin(anomaly); }, style: {
                        background: 'linear-gradient(135deg, #ef4444, #dc2626)',
                        color: '#fff', border: 'none', borderRadius: 10, padding: '14px 18px',
                        cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 12, textAlign: 'left',
                        transition: 'opacity 0.15s',
                    } },
                    React.createElement("span", { style: { fontSize: 26 } }, "\uD83D\uDDD1\uFE0F"),
                    React.createElement("div", null,
                        React.createElement("div", { style: { fontWeight: 700, fontSize: 13, marginBottom: 2 } }, "Move to Recycle Bin"),
                        React.createElement("div", { style: { fontSize: 12, opacity: 0.9 } }, "Dismiss this warning and move the item record to the Recycle Bin. You can restore it from there.")))))));
}
// ── Classify Dialog ──────────────────────────────────────────────────────────
function renderClassifyDialog(host) {
    var dlg = host.state.spoClassifyDialog;
    if (!dlg)
        return null;
    return React.createElement(ClassifyModalContent, { host: host, dlg: dlg });
}
function ClassifyModalContent(_a) {
    var _this = this;
    var host = _a.host, dlg = _a.dlg;
    var anomaly = dlg.anomaly, provisioning = dlg.provisioning, done = dlg.done, doneNormal = dlg.doneNormal, alreadyExisted = dlg.alreadyExisted, error = dlg.error;
    var _b = React.useState(0), elapsed = _b[0], setElapsed = _b[1];
    React.useEffect(function () {
        if (!provisioning) {
            setElapsed(0);
            return;
        }
        var timer = setInterval(function () { return setElapsed(function (e) { return e + 1; }); }, 1000);
        return function () { return clearInterval(timer); };
    }, [provisioning]);
    var formatTimer = function (sec) {
        var m = Math.floor(sec / 60);
        var s = sec % 60;
        return "".concat(m < 10 ? '0' + m : m, ":").concat(s < 10 ? '0' + s : s);
    };
    var handleVessel = function () { return __awaiter(_this, void 0, void 0, function () {
        var base, newVessel, newRecord_1, e_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    host.setState({ spoClassifyDialog: __assign(__assign({}, dlg), { provisioning: true, error: null }) });
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 4, , 5]);
                    base = host._base();
                    return [4 /*yield*/, host._fetchJson("".concat(base, "/api/vessels"), {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({ name: anomaly.name, imo: null, shipyard: 'Auto-Discovered', vessel_type: 'Bulk Carrier' }),
                        })];
                case 2:
                    newVessel = _a.sent();
                    newRecord_1 = {
                        id: (newVessel === null || newVessel === void 0 ? void 0 : newVessel.id) || "v_".concat(Date.now()),
                        name: (0, constants_1.cleanName)(anomaly.name),
                        imo: (newVessel === null || newVessel === void 0 ? void 0 : newVessel.imo) || '—',
                        status: 'Active',
                        image_url: (0, vesselImagePool_1.pickRandomVesselImage)('Bulk Carrier'),
                    };
                    // 2. Add vessel directly to top of state grid so it appears at the top of the vessel list immediately
                    host.setState(function (prev) { return ({
                        vessels: __spreadArray([newRecord_1], prev.vessels.filter(function (v) { return v.name.toLowerCase() !== newRecord_1.name.toLowerCase(); }), true),
                    }); });
                    // 3. Provision SPO DMS folder structure
                    return [4 /*yield*/, host._provisionVesselFolders(anomaly.name, newRecord_1.id)];
                case 3:
                    // 3. Provision SPO DMS folder structure
                    _a.sent();
                    // 4. Dismiss the anomaly
                    host._dismissAnomaly(anomaly.id);
                    // 5. Transition to success screen!
                    host.setState({ spoClassifyDialog: __assign(__assign({}, dlg), { provisioning: false, done: true, error: null }) });
                    return [3 /*break*/, 5];
                case 4:
                    e_1 = _a.sent();
                    host.setState({ spoClassifyDialog: __assign(__assign({}, dlg), { provisioning: false, done: false, error: (e_1 === null || e_1 === void 0 ? void 0 : e_1.message) || 'Failed to provision vessel.' }) });
                    return [3 /*break*/, 5];
                case 5: return [2 /*return*/];
            }
        });
    }); };
    var handleNormal = function () { return __awaiter(_this, void 0, void 0, function () {
        var result, e_2;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    host.setState({ spoClassifyDialog: __assign(__assign({}, dlg), { provisioning: true, error: null }) });
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, , 4]);
                    return [4 /*yield*/, host._saveNormalFolder(anomaly)];
                case 2:
                    result = _a.sent();
                    // Dismiss anomaly warning
                    host._dismissAnomaly(anomaly.id);
                    // Refresh normal folder list in state
                    host._loadNormalFolders();
                    // Show success screen
                    host.setState({ spoClassifyDialog: __assign(__assign({}, dlg), { provisioning: false, done: false, doneNormal: true, alreadyExisted: (result === null || result === void 0 ? void 0 : result.already_existed) === true, error: null }) });
                    return [3 /*break*/, 4];
                case 3:
                    e_2 = _a.sent();
                    host.setState({ spoClassifyDialog: __assign(__assign({}, dlg), { provisioning: false, done: false, doneNormal: false, error: (e_2 === null || e_2 === void 0 ? void 0 : e_2.message) || 'Failed to save folder.' }) });
                    return [3 /*break*/, 4];
                case 4: return [2 /*return*/];
            }
        });
    }); };
    var handleClose = function () {
        if (done || doneNormal) {
            host._loadData();
        }
        host.setState({ spoClassifyDialog: null });
    };
    var renderBody = function () {
        if (doneNormal) {
            return (React.createElement("div", { style: { textAlign: 'center', padding: '12px 0' } },
                React.createElement("div", { style: { fontSize: 52, marginBottom: 12 } }, "\uD83D\uDCC1"),
                React.createElement("h3", { style: { margin: '0 0 8px', fontSize: 20, fontWeight: 700, color: '#0f172a' } }, alreadyExisted ? 'Folder Already Registered' : 'Folder Confirmed as Normal Folder!'),
                React.createElement("p", { style: { margin: '0 0 20px', fontSize: 13, color: '#475569', lineHeight: 1.5 } },
                    React.createElement("strong", null,
                        "\"",
                        anomaly.name,
                        "\""),
                    " ",
                    alreadyExisted
                        ? 'was already listed as a Normal Folder in the system.'
                        : 'has been saved and will appear in the Normal Folders section below the vessel list.'),
                React.createElement("div", { style: { background: '#f0f9ff', border: '1px solid #bae6fd', borderRadius: 10, padding: '12px 16px', marginBottom: 24, fontSize: 12, color: '#0369a1', textAlign: 'left' } },
                    React.createElement("div", { style: { fontWeight: 700, marginBottom: 4 } }, "\uD83D\uDCCB Saved to Normal Folders"),
                    React.createElement("div", null, "\u2022 Anomaly warning dismissed"),
                    React.createElement("div", null, "\u2022 Folder listed in \"Normal Folders\" section"),
                    React.createElement("div", null, "\u2022 Record stored in database")),
                React.createElement("button", { onClick: handleClose, style: {
                        background: 'linear-gradient(135deg, #475569, #334155)',
                        color: '#fff', border: 'none', borderRadius: 10,
                        padding: '12px 32px', fontSize: 14, fontWeight: 700, cursor: 'pointer',
                        boxShadow: '0 4px 12px rgba(71,85,105,0.3)',
                    } }, "Done / View Vessels")));
        }
        if (provisioning) {
            return (React.createElement("div", { style: { textAlign: 'center', padding: '8px 0' } },
                React.createElement("div", { style: { fontSize: 44, marginBottom: 12 } }, "\u23F3"),
                React.createElement("h3", { style: { margin: '0 0 6px', fontSize: 18, fontWeight: 700, color: '#0284c7' } }, "Provisioning DMS Folder Structure\u2026"),
                React.createElement("p", { style: { margin: '0 0 16px', fontSize: 13, color: '#475569', lineHeight: 1.5 } },
                    "Creating standard DMS folder hierarchy for ",
                    React.createElement("strong", null,
                        "\"",
                        anomaly.name,
                        "\""),
                    " in SharePoint Online. Please wait."),
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
                            height: '100%', width: "".concat(Math.min(96, 15 + elapsed * 12), "%"),
                            transition: 'width 0.8s ease-out', borderRadius: 10,
                        } })),
                React.createElement("div", { style: { background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 12, padding: '14px 16px', textAlign: 'left', fontSize: 12 } },
                    React.createElement("div", { style: { color: '#16a34a', fontWeight: 600, marginBottom: 8, display: 'flex', alignItems: 'center', gap: 8 } },
                        React.createElement("span", { style: { fontSize: 14 } }, "\u2713"),
                        React.createElement("span", null, "Vessel record created in database")),
                    React.createElement("div", { style: { color: '#0284c7', fontWeight: 600, marginBottom: 8, display: 'flex', alignItems: 'center', gap: 8 } },
                        React.createElement("span", { style: { fontSize: 14 } }, "\u23F3"),
                        React.createElement("span", null, "Creating SharePoint DMS folder tree (Technical & Crewing, Month End, Certificates)\u2026")),
                    React.createElement("div", { style: { color: '#94a3b8', display: 'flex', alignItems: 'center', gap: 8 } },
                        React.createElement("span", { style: { fontSize: 14 } }, "\u25CB"),
                        React.createElement("span", null, "Linking category subfolders & permissions")))));
        }
        if (done) {
            return (React.createElement("div", { style: { textAlign: 'center', padding: '12px 0' } },
                React.createElement("div", { style: { fontSize: 52, marginBottom: 12 } }, "\uD83C\uDF89"),
                React.createElement("h3", { style: { margin: '0 0 8px', fontSize: 20, fontWeight: 700, color: '#059669' } }, "Vessel Successfully Provisioned & Classified!"),
                React.createElement("p", { style: { margin: '0 0 20px', fontSize: 13, color: '#475569', lineHeight: 1.5 } },
                    React.createElement("strong", null,
                        "\"",
                        anomaly.name,
                        "\""),
                    " has been registered in the DMS database and its full SharePoint DMS folder tree has been created."),
                React.createElement("div", { style: { background: '#ecfdf5', border: '1px solid #a7f3d0', borderRadius: 10, padding: '12px 16px', marginBottom: 24, fontSize: 12, color: '#065f46', textAlign: 'left' } },
                    React.createElement("div", { style: { fontWeight: 700, marginBottom: 4 } }, "\u2705 Provisioning Complete"),
                    React.createElement("div", null, "\u2022 Registered vessel card added to main grid"),
                    React.createElement("div", null, "\u2022 SharePoint DMS department subfolders created"),
                    React.createElement("div", null, "\u2022 Unrecognised warning dismissed")),
                React.createElement("button", { onClick: handleClose, style: {
                        background: 'linear-gradient(135deg, #10b981, #059669)',
                        color: '#fff', border: 'none', borderRadius: 10,
                        padding: '12px 32px', fontSize: 14, fontWeight: 700, cursor: 'pointer',
                        boxShadow: '0 4px 12px rgba(16,185,129,0.3)',
                    } }, "Done / View Vessels")));
        }
        return (React.createElement(React.Fragment, null,
            React.createElement("div", { style: { fontSize: 42, textAlign: 'center', marginBottom: 12 } }, anomaly.item_type === 'folder' ? '📁' : '📄'),
            React.createElement("h3", { style: { margin: '0 0 6px', fontSize: 18, fontWeight: 700, color: '#0f172a', textAlign: 'center' } },
                "Unrecognised SharePoint ",
                anomaly.item_type === 'folder' ? 'Folder' : 'File',
                " Detected"),
            React.createElement("p", { style: { margin: '0 0 20px', fontSize: 13, color: '#64748b', textAlign: 'center' } },
                "A ",
                anomaly.item_type,
                " named ",
                React.createElement("strong", null,
                    "\"",
                    anomaly.name,
                    "\""),
                " was found directly inside the vessel management area in SharePoint Online but is not registered in DMS."),
            React.createElement("div", { style: { background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 8, padding: '10px 14px', marginBottom: 20, fontSize: 11, color: '#475569', fontFamily: 'monospace', wordBreak: 'break-all' } },
                "\uD83D\uDCC2 ",
                anomaly.spo_path),
            error && (React.createElement("div", { style: { background: '#fef2f2', border: '1px solid #fca5a5', borderRadius: 8, padding: '10px 14px', marginBottom: 16, fontSize: 12, color: '#dc2626', display: 'flex', gap: 8, alignItems: 'center' } },
                React.createElement("span", null, "\u26A0\uFE0F"),
                React.createElement("span", null, error))),
            anomaly.item_type === 'folder' ? (React.createElement(React.Fragment, null,
                React.createElement("p", { style: { margin: '0 0 16px', fontSize: 13, fontWeight: 600, color: '#0f172a', textAlign: 'center' } }, "What is this folder?"),
                React.createElement("div", { style: { display: 'flex', gap: 12, flexDirection: 'column' } },
                    React.createElement("button", { onClick: function () { return handleVessel(); }, style: {
                            background: 'linear-gradient(135deg, #0ea5e9, #0284c7)',
                            color: '#fff', border: 'none', borderRadius: 12, padding: '16px 20px',
                            cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 14, textAlign: 'left',
                            transition: 'opacity 0.2s',
                        } },
                        React.createElement("span", { style: { fontSize: 28 } }, "\uD83D\uDEA2"),
                        React.createElement("div", null,
                            React.createElement("div", { style: { fontWeight: 700, fontSize: 14, marginBottom: 2 } }, "This is a Vessel"),
                            React.createElement("div", { style: { fontSize: 12, opacity: 0.9 } }, "Register it as a vessel and create its full DMS folder structure in SharePoint"))),
                    React.createElement("button", { onClick: handleNormal, style: {
                            background: '#f8fafc', color: '#334155',
                            border: '2px solid #e2e8f0', borderRadius: 12, padding: '16px 20px',
                            cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 14, textAlign: 'left',
                            transition: 'opacity 0.2s',
                        } },
                        React.createElement("span", { style: { fontSize: 28 } }, "\uD83D\uDCC1"),
                        React.createElement("div", null,
                            React.createElement("div", { style: { fontWeight: 700, fontSize: 14, marginBottom: 2, color: '#0f172a' } }, "This is a Normal Folder"),
                            React.createElement("div", { style: { fontSize: 12, color: '#64748b' } }, "Dismiss the warning and keep it listed as an unstructured folder")))))) : (React.createElement("div", { style: { display: 'flex', justifyContent: 'center' } },
                React.createElement("button", { onClick: handleNormal, style: { background: '#f1f5f9', color: '#475569', border: '1px solid #cbd5e1', borderRadius: 8, padding: '8px 20px', fontSize: 13, fontWeight: 600, cursor: 'pointer' } }, "Dismiss Warning")))));
    };
    return (React.createElement("div", { role: "dialog", "aria-modal": "true", "aria-label": "Classify SharePoint Item", style: {
            position: 'fixed', inset: 0, zIndex: 9999,
            background: 'rgba(15,23,42,0.6)',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            padding: 16,
        }, onClick: function (e) { if (e.target === e.currentTarget && !provisioning)
            handleClose(); } },
        React.createElement("div", { style: {
                background: '#fff', borderRadius: 20, boxShadow: '0 24px 64px rgba(0,0,0,0.28)',
                padding: 32, width: '100%', maxWidth: 480, position: 'relative',
            } },
            !provisioning && (React.createElement("button", { onClick: handleClose, style: { position: 'absolute', top: 14, right: 14, background: 'none', border: 'none', fontSize: 18, cursor: 'pointer', color: '#94a3b8' }, title: "Close" }, "\u2715")),
            renderBody())));
}
// ── Unrecognised Folders Section ─────────────────────────────────────────────
function renderUnrecognisedFolders(host, items) {
    return (React.createElement("div", { style: { marginTop: 32, background: 'linear-gradient(135deg, #fffbeb 0%, #fff9e6 100%)', border: '2px solid #f59e0b', borderRadius: 16, padding: 24 } },
        React.createElement("div", { style: { display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', marginBottom: 16, flexWrap: 'wrap', gap: 10 } },
            React.createElement("div", null,
                React.createElement("h3", { style: { margin: '0 0 4px', fontSize: 16, fontWeight: 700, color: '#92400e', display: 'flex', alignItems: 'center', gap: 8 } },
                    "\u26A0\uFE0F Folders Created Outside DMS Records",
                    React.createElement("span", { style: { background: '#f59e0b', color: '#fff', borderRadius: 20, padding: '1px 10px', fontSize: 12, fontWeight: 700 } }, items.length)),
                React.createElement("p", { style: { margin: 0, fontSize: 12, color: '#a16207' } }, "These folders were created directly in SharePoint Online at the vessel management level but are not registered in DMS. Please classify each one as a Vessel or a Normal Folder."))),
        React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 10 } }, items.map(function (item) { return (React.createElement("div", { key: item.id, style: {
                background: '#fff', borderRadius: 12, border: '1px solid #fde68a',
                padding: '14px 18px', display: 'flex', alignItems: 'center',
                justifyContent: 'space-between', gap: 12, flexWrap: 'wrap',
                boxShadow: '0 1px 4px rgba(245,158,11,0.08)',
            } },
            React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 12, flex: 1, minWidth: 200 } },
                React.createElement("span", { style: { fontSize: 28 } }, "\uD83D\uDCC1"),
                React.createElement("div", null,
                    React.createElement("div", { style: { fontWeight: 700, fontSize: 14, color: '#1f1f1f', marginBottom: 2 } }, item.name),
                    React.createElement("div", { style: { fontSize: 11, color: '#78716c', fontFamily: 'monospace' } }, item.spo_path),
                    item.detected_at && (React.createElement("div", { style: { fontSize: 10, color: '#a8a29e', marginTop: 2 } },
                        "Detected: ",
                        new Date(item.detected_at).toLocaleString())))),
            React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' } },
                React.createElement("button", { onClick: function () { return host.setState({ spoClassifyDialog: { anomaly: item, provisioning: false, done: false, error: null } }); }, style: {
                        background: 'linear-gradient(135deg, #f59e0b, #d97706)',
                        color: '#fff', border: 'none', borderRadius: 8, padding: '7px 14px',
                        fontSize: 12, fontWeight: 700, cursor: 'pointer',
                        display: 'inline-flex', alignItems: 'center', gap: 5,
                    } }, "\uD83D\uDD0D Classify"),
                React.createElement("a", { href: getSpoVesselFolderUrl(host.props.siteUrl, item.name), target: "_blank", rel: "noopener noreferrer", style: {
                        background: '#e0f2fe', color: '#0369a1', border: '1px solid #bae6fd',
                        borderRadius: 8, padding: '7px 12px', fontSize: 12, fontWeight: 600,
                        textDecoration: 'none', display: 'inline-flex', alignItems: 'center', gap: 4,
                    } }, "\u2197 Open"),
                React.createElement("button", { onClick: function () { return host.setState({ spoAnomalyDismissConfirm: item }); }, style: {
                        background: '#f1f5f9', color: '#64748b', border: '1px solid #cbd5e1',
                        borderRadius: 8, padding: '7px 12px', fontSize: 12, cursor: 'pointer',
                    }, title: "Dismiss warning" }, "\u2715 Dismiss")))); }))));
}
// ── Vessel-Level Uploaded Files Section ──────────────────────────────────────
function renderVesselLevelFiles(host, items) {
    var formatSize = function (bytes) {
        if (!bytes)
            return '';
        if (bytes < 1024)
            return "".concat(bytes, " B");
        if (bytes < 1048576)
            return "".concat((bytes / 1024).toFixed(1), " KB");
        return "".concat((bytes / 1048576).toFixed(1), " MB");
    };
    return (React.createElement("div", { style: { marginTop: 24, background: 'linear-gradient(135deg, #faf5ff 0%, #f3e8ff 100%)', border: '2px solid #a855f7', borderRadius: 16, padding: 24 } },
        React.createElement("div", { style: { marginBottom: 16 } },
            React.createElement("h3", { style: { margin: '0 0 4px', fontSize: 16, fontWeight: 700, color: '#6b21a8', display: 'flex', alignItems: 'center', gap: 8 } },
                "\uD83D\uDCC4 Files Uploaded Outside DMS Structure",
                React.createElement("span", { style: { background: '#a855f7', color: '#fff', borderRadius: 20, padding: '1px 10px', fontSize: 12, fontWeight: 700 } }, items.length)),
            React.createElement("p", { style: { margin: 0, fontSize: 12, color: '#7c3aed' } }, "These files were uploaded directly to the vessel management area in SharePoint Online \u2014 not inside any vessel's DMS folder structure.")),
        React.createElement("div", { style: { background: '#fff', borderRadius: 10, border: '1px solid #e9d5ff', overflow: 'hidden' } },
            React.createElement("table", { style: { width: '100%', borderCollapse: 'collapse', fontSize: 12 } },
                React.createElement("thead", null,
                    React.createElement("tr", { style: { background: '#f3e8ff', borderBottom: '1px solid #e9d5ff' } },
                        React.createElement("th", { style: { padding: '10px 14px', textAlign: 'left', fontWeight: 700, color: '#6b21a8', width: 36 } }),
                        React.createElement("th", { style: { padding: '10px 14px', textAlign: 'left', fontWeight: 700, color: '#6b21a8' } }, "File Name"),
                        React.createElement("th", { style: { padding: '10px 14px', textAlign: 'left', fontWeight: 700, color: '#6b21a8' } }, "Location"),
                        React.createElement("th", { style: { padding: '10px 14px', textAlign: 'left', fontWeight: 700, color: '#6b21a8' } }, "Detected"),
                        React.createElement("th", { style: { padding: '10px 14px', textAlign: 'center', fontWeight: 700, color: '#6b21a8' } }, "Actions"))),
                React.createElement("tbody", null, items.map(function (item, idx) {
                    var _a;
                    var ext = ((_a = item.name.split('.').pop()) === null || _a === void 0 ? void 0 : _a.toUpperCase()) || '';
                    var extColor = { PDF: '#ef4444', DOCX: '#3b82f6', XLSX: '#10b981', PPTX: '#f59e0b', JPG: '#ec4899', PNG: '#ec4899' };
                    var color = extColor[ext] || '#64748b';
                    return (React.createElement("tr", { key: item.id, style: { borderBottom: idx < items.length - 1 ? '1px solid #f3e8ff' : 'none', background: idx % 2 === 0 ? '#fff' : '#faf5ff' } },
                        React.createElement("td", { style: { padding: '10px 14px' } },
                            React.createElement("span", { style: { background: "".concat(color, "18"), color: color, borderRadius: 4, padding: '2px 5px', fontSize: 10, fontWeight: 700 } }, ext || 'FILE')),
                        React.createElement("td", { style: { padding: '10px 14px', fontWeight: 600, color: '#1e293b' } }, item.name),
                        React.createElement("td", { style: { padding: '10px 14px', color: '#64748b', fontFamily: 'monospace', fontSize: 11, maxWidth: 220, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' } }, item.spo_path),
                        React.createElement("td", { style: { padding: '10px 14px', color: '#94a3b8', whiteSpace: 'nowrap' } }, item.detected_at ? new Date(item.detected_at).toLocaleDateString() : '—'),
                        React.createElement("td", { style: { padding: '10px 14px', textAlign: 'center', whiteSpace: 'nowrap' } },
                            React.createElement("div", { style: { display: 'inline-flex', gap: 6 } },
                                React.createElement("a", { href: getSpoVesselFolderUrl(host.props.siteUrl, item.vessel_name || undefined), target: "_blank", rel: "noopener noreferrer", style: { background: '#e0f2fe', color: '#0369a1', border: '1px solid #bae6fd', borderRadius: 6, padding: '4px 10px', fontSize: 11, fontWeight: 600, textDecoration: 'none' } }, "\u2197 Open"),
                                React.createElement("button", { onClick: function () { return host.setState({ spoAnomalyDismissConfirm: item }); }, style: { background: '#f1f5f9', color: '#64748b', border: '1px solid #cbd5e1', borderRadius: 6, padding: '4px 8px', fontSize: 11, cursor: 'pointer' }, title: "Dismiss" }, "\u2715 Dismiss")))));
                }))))));
}
// ── Normal Folders Section ────────────────────────────────────────────────────
function renderNormalFoldersSection(host) {
    var folders = host.state.normalFolders || [];
    if (folders.length === 0)
        return null;
    return (React.createElement("div", { style: { marginTop: 36 } },
        React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 12, marginBottom: 16 } },
            React.createElement("div", { style: { flex: 1, height: 1, background: 'linear-gradient(to right, #94a3b8, transparent)' } }),
            React.createElement("span", { style: { fontSize: 12, fontWeight: 700, color: '#475569', background: '#f1f5f9', padding: '4px 14px', borderRadius: 20, border: '1px solid #cbd5e1', whiteSpace: 'nowrap' } },
                "\uD83D\uDCC1 Normal Folders (",
                folders.length,
                ")"),
            React.createElement("div", { style: { flex: 1, height: 1, background: 'linear-gradient(to left, #94a3b8, transparent)' } })),
        React.createElement("div", { style: { background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 14, padding: '16px 20px' } },
            React.createElement("p", { style: { margin: '0 0 14px', fontSize: 12, color: '#64748b', lineHeight: 1.5 } },
                "The following folders were found in SharePoint Online but confirmed as ",
                React.createElement("strong", null, "normal (non-vessel) folders"),
                " by a user. They are listed here for reference only."),
            React.createElement("div", { style: { display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: 10 } }, folders.map(function (f, idx) {
                var _a;
                return (React.createElement("div", { key: (_a = f.id) !== null && _a !== void 0 ? _a : "nf-".concat(idx), style: {
                        background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: '12px 14px',
                        display: 'flex', alignItems: 'flex-start', gap: 10,
                    } },
                    React.createElement("span", { style: { fontSize: 24, flexShrink: 0, marginTop: 2 } }, f.item_type === 'folder' ? '📁' : '📄'),
                    React.createElement("div", { style: { minWidth: 0, flex: 1 } },
                        React.createElement("div", { style: { fontWeight: 700, fontSize: 13, color: '#0f172a', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' } }, f.name),
                        React.createElement("div", { style: { fontSize: 10, color: '#94a3b8', fontFamily: 'monospace', marginTop: 2, wordBreak: 'break-all' } }, f.spo_path),
                        f.detected_at && (React.createElement("div", { style: { fontSize: 10, color: '#94a3b8', marginTop: 4 } },
                            "Saved: ",
                            new Date(f.detected_at).toLocaleDateString())))));
            })))));
}
// ── Vessel Card Provision Dialog ──────────────────────────────────────────────
// Shown when user clicks the "📁 Provision" button on a vessel card.
function renderProvisionDialog(host) {
    if (!host.state.spoProvisionDialog)
        return null;
    return React.createElement(ProvisionModalContent, { host: host, dlg: host.state.spoProvisionDialog });
}
function ProvisionModalContent(_a) {
    var _this = this;
    var _b;
    var host = _a.host, dlg = _a.dlg;
    var vessel = dlg.vessel, provisioning = dlg.provisioning, done = dlg.done, error = dlg.error;
    var _c = React.useState(0), elapsed = _c[0], setElapsed = _c[1];
    var isProvisioned = (((_b = host.state.provisionedVesselIds) === null || _b === void 0 ? void 0 : _b.has(vessel.id)) ||
        (host.state.rows && host.state.rows.some(function (r) { return r.vesselName && r.vesselName.toLowerCase() === vessel.name.toLowerCase(); })) ||
        vessel.is_provisioned === true);
    React.useEffect(function () {
        if (!provisioning) {
            setElapsed(0);
            return;
        }
        var timer = setInterval(function () { return setElapsed(function (e) { return e + 1; }); }, 1000);
        return function () { return clearInterval(timer); };
    }, [provisioning]);
    var formatTimer = function (sec) {
        var m = Math.floor(sec / 60);
        var s = sec % 60;
        return "".concat(m < 10 ? '0' + m : m, ":").concat(s < 10 ? '0' + s : s);
    };
    var handleProvision = function () { return __awaiter(_this, void 0, void 0, function () {
        var e_3;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    host.setState({ spoProvisionDialog: __assign(__assign({}, dlg), { provisioning: true, error: null }) });
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, , 4]);
                    return [4 /*yield*/, host._provisionVesselFolders(vessel.name, vessel.id)];
                case 2:
                    _a.sent();
                    host.setState({ spoProvisionDialog: __assign(__assign({}, dlg), { provisioning: false, done: true, error: null }) });
                    return [3 /*break*/, 4];
                case 3:
                    e_3 = _a.sent();
                    host.setState({ spoProvisionDialog: __assign(__assign({}, dlg), { provisioning: false, done: false, error: (e_3 === null || e_3 === void 0 ? void 0 : e_3.message) || 'Failed to provision vessel folders.' }) });
                    return [3 /*break*/, 4];
                case 4: return [2 /*return*/];
            }
        });
    }); };
    var handleClose = function () {
        host.setState({ spoProvisionDialog: null, folderCreationResults: null });
    };
    var renderBody = function () {
        if (provisioning) {
            return (React.createElement("div", { style: { textAlign: 'center', padding: '8px 0' } },
                React.createElement("div", { style: { fontSize: 44, marginBottom: 12 } }, "\u23F3"),
                React.createElement("h3", { style: { margin: '0 0 6px', fontSize: 18, fontWeight: 700, color: '#0284c7' } }, "Provisioning DMS Folder Structure\u2026"),
                React.createElement("p", { style: { margin: '0 0 16px', fontSize: 13, color: '#475569', lineHeight: 1.5 } },
                    "Creating standard DMS folder hierarchy for ",
                    React.createElement("strong", null,
                        "\"",
                        vessel.name,
                        "\""),
                    " in SharePoint Online. Please wait."),
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
                            height: '100%', width: "".concat(Math.min(96, 15 + elapsed * 12), "%"),
                            transition: 'width 0.8s ease-out', borderRadius: 10,
                        } })),
                React.createElement("div", { style: { background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 12, padding: '14px 16px', textAlign: 'left', fontSize: 12 } },
                    React.createElement("div", { style: { color: '#16a34a', fontWeight: 600, marginBottom: 8, display: 'flex', alignItems: 'center', gap: 8 } },
                        React.createElement("span", { style: { fontSize: 14 } }, "\u2713"),
                        React.createElement("span", null, "Vessel record confirmed in database")),
                    React.createElement("div", { style: { color: '#0284c7', fontWeight: 600, marginBottom: 8, display: 'flex', alignItems: 'center', gap: 8 } },
                        React.createElement("span", { style: { fontSize: 14 } }, "\u23F3"),
                        React.createElement("span", null, "Creating SharePoint DMS folder tree (Technical & Crewing, Month End, Certificates)\u2026")),
                    React.createElement("div", { style: { color: '#94a3b8', display: 'flex', alignItems: 'center', gap: 8 } },
                        React.createElement("span", { style: { fontSize: 14 } }, "\u25CB"),
                        React.createElement("span", null, "Linking category subfolders & permissions")))));
        }
        if (done) {
            return (React.createElement("div", { style: { textAlign: 'center', padding: '12px 0' } },
                React.createElement("div", { style: { fontSize: 52, marginBottom: 12 } }, "\uD83C\uDF89"),
                React.createElement("h3", { style: { margin: '0 0 8px', fontSize: 20, fontWeight: 700, color: '#059669' } }, "Vessel Folders Successfully Provisioned!"),
                React.createElement("p", { style: { margin: '0 0 20px', fontSize: 13, color: '#475569', lineHeight: 1.5 } },
                    React.createElement("strong", null,
                        "\"",
                        vessel.name,
                        "\""),
                    " has been fully provisioned \u2014 all SharePoint DMS folder structures have been created."),
                React.createElement("div", { style: { background: '#ecfdf5', border: '1px solid #a7f3d0', borderRadius: 10, padding: '12px 16px', marginBottom: 24, fontSize: 12, color: '#065f46', textAlign: 'left' } },
                    React.createElement("div", { style: { fontWeight: 700, marginBottom: 4 } }, "\u2705 Provisioning Complete"),
                    React.createElement("div", null, "\u2022 SharePoint DMS department subfolders created"),
                    React.createElement("div", null,
                        "\u2022 IMO: ",
                        vessel.imo || '—'),
                    React.createElement("div", null, "\u2022 Vessel is ready for document uploads")),
                React.createElement("button", { onClick: handleClose, style: {
                        background: 'linear-gradient(135deg, #10b981, #059669)',
                        color: '#fff', border: 'none', borderRadius: 10,
                        padding: '12px 32px', fontSize: 14, fontWeight: 700, cursor: 'pointer',
                        boxShadow: '0 4px 12px rgba(16,185,129,0.3)',
                    } }, "Done / View Vessels")));
        }
        if (!provisioning && !done && isProvisioned) {
            return (React.createElement("div", { style: { textAlign: 'center', padding: '8px 0' } },
                React.createElement("div", { style: { fontSize: 48, marginBottom: 10 } }, "\u2705"),
                React.createElement("h3", { style: { margin: '0 0 6px', fontSize: 19, fontWeight: 700, color: '#15803d' } }, "Folders Already Provisioned"),
                React.createElement("p", { style: { margin: '0 0 16px', fontSize: 13, color: '#475569', lineHeight: 1.5 } },
                    "The standard SharePoint DMS folder tree for ",
                    React.createElement("strong", null,
                        "\"",
                        vessel.name,
                        "\""),
                    " is already provisioned and registered."),
                React.createElement("div", { style: { background: '#f0fdf4', border: '1px solid #bbf7d0', borderRadius: 10, padding: '12px 16px', marginBottom: 20, fontSize: 12, color: '#166534', textAlign: 'left' } },
                    React.createElement("div", { style: { fontWeight: 700, marginBottom: 4 } }, "\uD83D\uDCCB Provisioned Folder Structure:"),
                    React.createElement("div", null, "\u2022 Technical & Crewing \u2014 Active"),
                    React.createElement("div", null, "\u2022 Commercial & Chartering \u2014 Active"),
                    React.createElement("div", null, "\u2022 Insurance \u2014 Active")),
                error && (React.createElement("div", { style: { background: '#fef2f2', border: '1px solid #fca5a5', borderRadius: 8, padding: '10px 14px', marginBottom: 16, fontSize: 12, color: '#dc2626', display: 'flex', gap: 8, alignItems: 'center' } },
                    React.createElement("span", null, "\u26A0\uFE0F"),
                    React.createElement("span", null, error))),
                React.createElement("div", { style: { display: 'flex', gap: 10, justifyContent: 'center', flexWrap: 'wrap' } },
                    React.createElement("button", { onClick: handleClose, style: { background: '#f1f5f9', color: '#475569', border: '1px solid #cbd5e1', borderRadius: 8, padding: '10px 20px', fontSize: 13, fontWeight: 600, cursor: 'pointer' } }, "Close"),
                    host.props.siteUrl && (React.createElement("a", { href: getSpoVesselFolderUrl(host.props.siteUrl, vessel.name), target: "_blank", rel: "noopener noreferrer", style: {
                            background: 'linear-gradient(135deg, #0ea5e9, #0284c7)', color: '#fff',
                            border: 'none', borderRadius: 8, padding: '10px 20px',
                            fontSize: 13, fontWeight: 700, textDecoration: 'none',
                            display: 'inline-flex', alignItems: 'center', gap: 6,
                        } }, "Open in SharePoint \u2197")),
                    React.createElement("button", { onClick: handleProvision, style: {
                            background: 'transparent', color: '#0284c7', border: '1px solid #0284c7',
                            borderRadius: 8, padding: '10px 20px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
                        } }, "\uD83D\uDD04 Re-Provision Folders"))));
        }
        return (React.createElement(React.Fragment, null,
            React.createElement("div", { style: { textAlign: 'center', marginBottom: 16 } },
                React.createElement("div", { style: { fontSize: 44, marginBottom: 10 } }, "\uD83D\uDCC1"),
                React.createElement("h3", { style: { margin: '0 0 6px', fontSize: 18, fontWeight: 700, color: '#0f172a' } }, "Provision SharePoint Folders"),
                React.createElement("p", { style: { margin: '0 0 12px', fontSize: 13, color: '#64748b', lineHeight: 1.5 } },
                    "This will create the full DMS folder hierarchy for ",
                    React.createElement("strong", null,
                        "\"",
                        vessel.name,
                        "\""),
                    " in SharePoint Online across all departments.")),
            React.createElement("div", { style: { background: '#f0f9ff', border: '1px solid #bae6fd', borderRadius: 10, padding: '12px 16px', marginBottom: 20, fontSize: 12, color: '#0369a1', textAlign: 'left' } },
                React.createElement("div", { style: { fontWeight: 700, marginBottom: 4 } }, "\uD83D\uDCCB What will be created:"),
                React.createElement("div", null, "\u2022 Technical & Crewing \u2014 monthly sub-folders + categories"),
                React.createElement("div", null, "\u2022 Commercial \u2014 contract and invoice folders"),
                React.createElement("div", null, "\u2022 Insurance \u2014 certificate and policy folders")),
            error && (React.createElement("div", { style: { background: '#fef2f2', border: '1px solid #fca5a5', borderRadius: 8, padding: '10px 14px', marginBottom: 16, fontSize: 12, color: '#dc2626', display: 'flex', gap: 8, alignItems: 'center' } },
                React.createElement("span", null, "\u26A0\uFE0F"),
                React.createElement("span", null, error))),
            React.createElement("div", { style: { display: 'flex', gap: 10, justifyContent: 'center' } },
                React.createElement("button", { onClick: handleClose, style: { background: '#f1f5f9', color: '#475569', border: '1px solid #cbd5e1', borderRadius: 8, padding: '10px 22px', fontSize: 13, fontWeight: 600, cursor: 'pointer' } }, "Cancel"),
                React.createElement("button", { onClick: handleProvision, style: {
                        background: 'linear-gradient(135deg, #0ea5e9, #0284c7)',
                        color: '#fff', border: 'none', borderRadius: 8,
                        padding: '10px 26px', fontSize: 13, fontWeight: 700, cursor: 'pointer',
                        boxShadow: '0 4px 12px rgba(2,132,199,0.3)',
                    } }, "\uD83D\uDCC1 Start Provisioning"))));
    };
    return (React.createElement("div", { role: "dialog", "aria-modal": "true", "aria-label": "Provision Vessel Folders", style: {
            position: 'fixed', inset: 0, zIndex: 9999,
            background: 'rgba(15,23,42,0.6)',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            backdropFilter: 'blur(4px)',
        }, onClick: function (e) { if (e.target === e.currentTarget && !provisioning)
            handleClose(); } },
        React.createElement("div", { style: {
                background: '#fff', borderRadius: 20, padding: '32px 36px',
                width: 480, maxWidth: '92vw',
                boxShadow: '0 24px 64px rgba(0,0,0,0.28)', position: 'relative',
            } },
            !provisioning && (React.createElement("button", { onClick: handleClose, style: { position: 'absolute', top: 14, right: 14, background: 'none', border: 'none', fontSize: 18, cursor: 'pointer', color: '#94a3b8' }, title: "Close" }, "\u2715")),
            renderBody())));
}
// ── Main Page Render ─────────────────────────────────────────────────────────
function renderVesselsPage(host) {
    var _a = host.state, vessels = _a.vessels, vesselsSearch = _a.vesselsSearch, vesselStatusFilter = _a.vesselStatusFilter, vesselTypeFilter = _a.vesselTypeFilter, modal = _a.modal, selectedVessel = _a.selectedVessel, folderProvisioningVesselId = _a.folderProvisioningVesselId, folderCreationError = _a.folderCreationError, folderCreationResults = _a.folderCreationResults, panelLoading = _a.panelLoading, loading = _a.loading;
    var filtered = vessels.filter(function (v) {
        if (vesselStatusFilter !== 'all' && (v.status || 'Active') !== vesselStatusFilter)
            return false;
        if (vesselTypeFilter && vesselTypeFilter !== 'all' && (v.vessel_type || '') !== vesselTypeFilter)
            return false;
        if (vesselsSearch) {
            var q = vesselsSearch.toLowerCase();
            return v.name.toLowerCase().includes(q) ||
                (v.imo || '').includes(vesselsSearch) ||
                (v.vessel_type || '').toLowerCase().includes(q) ||
                (v.shipyard || '').toLowerCase().includes(q);
        }
        return true;
    });
    // All unique vessel types for filter dropdown
    var allTypes = Array.from(new Set(vessels.map(function (v) { return v.vessel_type; }).filter(Boolean)));
    var isLoading = panelLoading || (loading && vessels.length === 0);
    // Segregated anomalies
    var allAnomalies = host.state.folderAnomalies || [];
    var vesselLevelFolders = allAnomalies.filter(function (a) { return a.anomaly_type === 'vessel_level_unmatched' && a.item_type === 'folder'; });
    var vesselLevelFiles = allAnomalies.filter(function (a) { return a.anomaly_type === 'vessel_level_unmatched' && a.item_type === 'file'; });
    return (React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 0 } },
        React.createElement("div", { style: { display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 20, flexWrap: 'wrap', gap: 12 } },
            React.createElement("div", null,
                React.createElement("h2", { style: { margin: 0, fontSize: 22, fontWeight: 700, color: '#0f172a', display: 'flex', alignItems: 'center', gap: 10 } },
                    "\uD83D\uDEA2 Vessels",
                    vessels.length > 0 && (React.createElement("span", { style: { background: '#e0f2fe', color: '#0284c7', borderRadius: 20, padding: '2px 10px', fontSize: 13, fontWeight: 700 } }, vessels.length)),
                    (vesselLevelFolders.length + vesselLevelFiles.length) > 0 && (React.createElement("span", { style: { background: '#fef3c7', color: '#92400e', borderRadius: 20, padding: '2px 10px', fontSize: 12, fontWeight: 700, border: '1px solid #f59e0b', display: 'inline-flex', alignItems: 'center', gap: 4 } },
                        "\u26A0\uFE0F ",
                        vesselLevelFolders.length + vesselLevelFiles.length,
                        " SPO item",
                        vesselLevelFolders.length + vesselLevelFiles.length !== 1 ? 's' : '',
                        " need review"))),
                React.createElement("p", { style: { margin: '4px 0 0', fontSize: 13, color: '#64748b' } }, "Manage fleet vessels, provision SharePoint folders, and view documents.")),
            React.createElement("div", { style: { display: 'flex', gap: 8, flexWrap: 'wrap' } },
                React.createElement("button", { onClick: function () { return host._goToView('vessels').catch(function () { return undefined; }); }, title: "Reload vessel list from database", style: { background: '#f1f5f9', color: '#475569', border: '1px solid #cbd5e1', borderRadius: 8, padding: '9px 14px', fontSize: 13, fontWeight: 600, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6 } }, "\uD83D\uDD04 Refresh"),
                React.createElement("button", { onClick: function () { return host._openDeleteVessel(); }, style: { background: 'linear-gradient(135deg, #f43f5e, #e11d48)', color: '#fff', border: 'none', borderRadius: 8, padding: '9px 18px', fontSize: 13, fontWeight: 600, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6 } }, "\uD83D\uDDD1 Delete"),
                React.createElement("button", { onClick: function () { return selectedVessel ? host._openEditVessel(selectedVessel) : alert('Please select a vessel first.'); }, style: { background: 'linear-gradient(135deg, #38bdf8, #0284c7)', color: '#fff', border: 'none', borderRadius: 8, padding: '9px 18px', fontSize: 13, fontWeight: 600, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6 } }, "\u270F\uFE0F Edit"),
                React.createElement("button", { onClick: host._openCreate, style: { background: 'linear-gradient(135deg, #10b981, #059669)', color: '#fff', border: 'none', borderRadius: 8, padding: '9px 18px', fontSize: 13, fontWeight: 600, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6 } }, "\uFF0B New Vessel"))),
        React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 10, marginBottom: 20, flexWrap: 'wrap' } },
            React.createElement("div", { style: { position: 'relative', flex: '1 1 220px', minWidth: 180 } },
                React.createElement("span", { style: { position: 'absolute', left: 12, top: '50%', transform: 'translateY(-50%)', color: '#94a3b8', fontSize: 14 } }, "\uD83D\uDD0D"),
                React.createElement("input", { type: "text", placeholder: "Search by name, IMO, type, shipyard\u2026", value: vesselsSearch, onChange: function (e) { return host.setState({ vesselsSearch: e.target.value }); }, style: { width: '100%', padding: '10px 14px 10px 38px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 13, outline: 'none', boxSizing: 'border-box', background: '#fff' } })),
            React.createElement("select", { value: vesselStatusFilter, onChange: function (e) { return host.setState({ vesselStatusFilter: e.target.value }); }, style: { padding: '10px 14px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 13, background: '#fff', outline: 'none', minWidth: 130 } },
                React.createElement("option", { value: "all" }, "All Status"),
                React.createElement("option", { value: "Active" }, "Active"),
                React.createElement("option", { value: "In Maintenance" }, "In Maintenance"),
                React.createElement("option", { value: "Inactive" }, "Inactive")),
            React.createElement("select", { value: vesselTypeFilter || 'all', onChange: function (e) { return host.setState({ vesselTypeFilter: e.target.value }); }, style: { padding: '10px 14px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 13, background: '#fff', outline: 'none', minWidth: 140 } },
                React.createElement("option", { value: "all" }, "All Types"),
                allTypes.map(function (t) { return React.createElement("option", { key: t, value: t }, t); })),
            (vesselsSearch || vesselStatusFilter !== 'all' || (vesselTypeFilter && vesselTypeFilter !== 'all')) && (React.createElement("button", { onClick: function () { return host.setState({ vesselsSearch: '', vesselStatusFilter: 'all', vesselTypeFilter: 'all' }); }, style: { background: 'transparent', color: '#64748b', border: '1px solid #cbd5e1', borderRadius: 8, padding: '9px 14px', fontSize: 13, cursor: 'pointer' } }, "\u2715 Clear")),
            React.createElement("span", { style: { marginLeft: 'auto', fontSize: 12, color: '#94a3b8', fontWeight: 500 } },
                filtered.length,
                " vessel",
                filtered.length !== 1 ? 's' : '')),
        folderCreationError && (React.createElement("div", { style: { marginBottom: 12, background: '#fde7e9', border: '1px solid #fca5a5', borderRadius: 8, padding: '10px 14px', fontSize: 12, color: '#a4262c', display: 'flex', justifyContent: 'space-between', alignItems: 'center' } },
            React.createElement("span", null,
                "\u26A0 ",
                folderCreationError),
            React.createElement("button", { onClick: function () { return host.setState({ folderCreationError: null }); }, style: { border: 'none', background: 'none', cursor: 'pointer', color: '#a4262c', fontWeight: 700 } }, "\u2715"))),
        folderCreationResults && !folderCreationError && (React.createElement("div", { style: { marginBottom: 12, background: '#dff6dd', border: '1px solid #86efac', borderRadius: 8, padding: '10px 14px', fontSize: 12, color: '#107c10', display: 'flex', justifyContent: 'space-between', alignItems: 'center' } },
            React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' } },
                React.createElement("span", null,
                    "\u2705 SharePoint folders provisioned \u2014 ",
                    folderCreationResults.filter(function (r) { return r.status === 'created'; }).length,
                    " created, ",
                    folderCreationResults.filter(function (r) { return r.status === 'existed'; }).length,
                    " already existed."),
                host.props.siteUrl && (React.createElement("a", { href: getSpoVesselFolderUrl(host.props.siteUrl), target: "_blank", rel: "noopener noreferrer", style: { color: '#0078d4', fontWeight: 600, textDecoration: 'underline', display: 'inline-flex', alignItems: 'center', gap: 3 } }, "Open SharePoint Folder \u2197"))),
            React.createElement("button", { onClick: function () { return host.setState({ folderCreationResults: null }); }, style: { border: 'none', background: 'none', cursor: 'pointer', color: '#107c10', fontWeight: 700 } }, "\u2715"))),
        isLoading ? (React.createElement("div", { style: { display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', minHeight: 280, gap: 16, background: '#fff', borderRadius: 12, border: '1px solid #e2e8f0' } },
            React.createElement("div", { style: {
                    width: 40, height: 40, border: '3px solid #e2e8f0',
                    borderTopColor: '#0078d4', borderRadius: '50%',
                    animation: 'spin 0.8s linear infinite',
                } }),
            React.createElement("p", { style: { margin: 0, fontSize: 14, color: '#64748b', fontWeight: 500 } }, "Loading vessels from database\u2026"),
            React.createElement("p", { style: { margin: 0, fontSize: 12, color: '#94a3b8' } }, "Fetching vessel records and folder structure"))) : filtered.length === 0 ? (
        /* ── Empty State ── */
        React.createElement("div", { style: { display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', minHeight: 280, gap: 16, background: '#fff', borderRadius: 12, border: '2px dashed #e2e8f0' } },
            React.createElement("span", { style: { fontSize: 48 } }, "\uD83D\uDEA2"),
            React.createElement("div", { style: { textAlign: 'center' } },
                React.createElement("p", { style: { margin: 0, fontSize: 16, fontWeight: 700, color: '#0f172a' } }, vessels.length === 0 ? 'No vessels yet' : 'No vessels match your filter'),
                React.createElement("p", { style: { margin: '6px 0 0', fontSize: 13, color: '#64748b' } }, vessels.length === 0
                    ? 'Create your first vessel to provision its SharePoint folder structure.'
                    : 'Try clearing the search or filters.')),
            vessels.length === 0 && (React.createElement("button", { onClick: host._openCreate, style: { background: '#10b981', color: '#fff', border: 'none', borderRadius: 8, padding: '10px 24px', fontSize: 14, fontWeight: 600, cursor: 'pointer' } }, "\uFF0B Create First Vessel")))) : (
        /* ── Vessel Cards Grid ── */
        React.createElement("div", { style: { display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(320px, 1fr))', gap: 16 } }, filtered.map(function (vessel) {
            var isSelected = (selectedVessel === null || selectedVessel === void 0 ? void 0 : selectedVessel.id) === vessel.id;
            var status = vessel.status || 'Active';
            var statusColor = status === 'Active' ? '#10b981' : status === 'In Maintenance' ? '#f59e0b' : '#ef4444';
            var statusBg = status === 'Active' ? '#f0fdf4' : status === 'In Maintenance' ? '#fffbeb' : '#fef2f2';
            var isProvisioning = folderProvisioningVesselId === vessel.id;
            var imgSrc = (0, vesselImagePool_1.getVesselImageForId)(vessel.id);
            return (React.createElement("div", { key: vessel.id, onClick: function () { return host.setState({ selectedVessel: isSelected ? null : vessel }); }, style: {
                    background: '#fff',
                    borderRadius: 14,
                    border: isSelected ? '2px solid #0078d4' : '1px solid #e2e8f0',
                    boxShadow: isSelected ? '0 0 0 3px rgba(0,120,212,0.15)' : '0 1px 4px rgba(0,0,0,0.06)',
                    overflow: 'hidden',
                    cursor: 'pointer',
                    transition: 'all 0.18s ease',
                    display: 'flex',
                    flexDirection: 'column',
                }, onMouseEnter: function (e) { if (!isSelected)
                    e.currentTarget.style.boxShadow = '0 4px 16px rgba(0,0,0,0.10)'; }, onMouseLeave: function (e) { if (!isSelected)
                    e.currentTarget.style.boxShadow = '0 1px 4px rgba(0,0,0,0.06)'; } },
                React.createElement("div", { style: { position: 'relative', height: 120, overflow: 'hidden', background: '#1e3a5f' } },
                    React.createElement("img", { src: (0, vesselImagePool_1.resolveImgUrl)(imgSrc), alt: vessel.name, style: { width: '100%', height: '100%', objectFit: 'cover', opacity: 0.85 } }),
                    React.createElement("div", { style: {
                            position: 'absolute', inset: 0,
                            background: 'linear-gradient(to bottom, transparent 30%, rgba(15,23,42,0.75) 100%)',
                        } }),
                    React.createElement("span", { style: {
                            position: 'absolute', top: 10, right: 10,
                            background: statusBg, color: statusColor,
                            borderRadius: 20, padding: '2px 10px', fontSize: 11, fontWeight: 700,
                            border: "1px solid ".concat(statusColor, "40"),
                        } }, status),
                    isSelected && (React.createElement("span", { style: {
                            position: 'absolute', top: 10, left: 10,
                            background: '#0078d4', color: '#fff',
                            borderRadius: '50%', width: 22, height: 22,
                            display: 'flex', alignItems: 'center', justifyContent: 'center',
                            fontSize: 13, fontWeight: 700, boxShadow: '0 2px 6px rgba(0,0,0,0.25)',
                        } }, "\u2713")),
                    React.createElement("div", { style: { position: 'absolute', bottom: 10, left: 14, right: 14 } },
                        React.createElement("p", { style: { margin: 0, fontSize: 15, fontWeight: 700, color: '#fff', textShadow: '0 1px 3px rgba(0,0,0,0.5)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' } },
                            "\uD83D\uDEA2 ",
                            vessel.name))),
                React.createElement("div", { style: { padding: '14px 16px', flex: 1, display: 'flex', flexDirection: 'column', gap: 8 } },
                    React.createElement("div", { style: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '6px 12px' } },
                        React.createElement("div", null,
                            React.createElement("span", { style: { fontSize: 10, color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase' } }, "IMO"),
                            React.createElement("p", { style: { margin: 0, fontSize: 13, color: '#1e293b', fontWeight: 600 } }, vessel.imo || '—')),
                        React.createElement("div", null,
                            React.createElement("span", { style: { fontSize: 10, color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase' } }, "Type"),
                            React.createElement("p", { style: { margin: 0, fontSize: 13, color: '#1e293b', fontWeight: 600, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' } }, vessel.vessel_type || '—')),
                        React.createElement("div", null,
                            React.createElement("span", { style: { fontSize: 10, color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase' } }, "Shipyard"),
                            React.createElement("p", { style: { margin: 0, fontSize: 12, color: '#475569', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' } }, vessel.shipyard || '—')),
                        React.createElement("div", null,
                            React.createElement("span", { style: { fontSize: 10, color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase' } }, "Hull No."),
                            React.createElement("p", { style: { margin: 0, fontSize: 12, color: '#475569' } }, vessel.hull_number || '—'))),
                    React.createElement("div", { style: { display: 'flex', gap: 8, marginTop: 4 } },
                        React.createElement("button", { onClick: function (e) {
                                e.stopPropagation();
                                host.setState({ vesselFilter: vessel.name, docMainFolder: null, folderPathStack: [] });
                                void host._goToView('list');
                            }, title: "View documents for this vessel", style: {
                                flex: 1, background: '#eff6ff', color: '#1d4ed8',
                                border: '1px solid #bfdbfe', borderRadius: 7, padding: '7px 10px',
                                fontSize: 11, fontWeight: 700, cursor: 'pointer',
                                display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 4,
                            } }, "\uD83D\uDCC4 View Documents"),
                        (function () {
                            var _a;
                            var provisioned = (((_a = host.state.provisionedVesselIds) === null || _a === void 0 ? void 0 : _a.has(vessel.id)) ||
                                (host.state.rows && host.state.rows.some(function (r) { return r.vesselName && r.vesselName.toLowerCase() === vessel.name.toLowerCase(); })) ||
                                vessel.is_provisioned === true);
                            return (React.createElement("button", { disabled: !!folderProvisioningVesselId, onClick: function (e) {
                                    e.stopPropagation();
                                    host.setState({ spoProvisionDialog: { vessel: vessel, provisioning: false, done: false, error: null } });
                                }, title: provisioned ? "SharePoint DMS folders are already provisioned" : "Create SharePoint folder structure for this vessel", style: {
                                    flex: 1,
                                    border: provisioned ? '1px solid #bbf7d0' : '1px solid #cbd5e1',
                                    borderRadius: 7, padding: '7px 10px',
                                    fontSize: 11, fontWeight: 700, cursor: folderProvisioningVesselId ? 'not-allowed' : 'pointer',
                                    background: isProvisioning ? '#f0f9ff' : provisioned ? '#f0fdf4' : '#f8fafc',
                                    color: isProvisioning ? '#0284c7' : provisioned ? '#15803d' : '#334155',
                                    display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 4,
                                } }, isProvisioning ? (React.createElement(React.Fragment, null,
                                React.createElement("span", { style: { display: 'inline-block', width: 11, height: 11, border: '2px solid #0284c7', borderTopColor: 'transparent', borderRadius: '50%', animation: 'spin 0.7s linear infinite' } }),
                                " Creating\u2026")) : provisioned ? ('✓ Already Provisioned') : ('📁 Provision')));
                        })(),
                        React.createElement("a", { href: getSpoVesselFolderUrl(host.props.siteUrl, vessel.name), target: "_blank", rel: "noopener noreferrer", onClick: function (e) { return e.stopPropagation(); }, title: "Open in SharePoint", style: {
                                width: 32, display: 'flex', alignItems: 'center', justifyContent: 'center',
                                border: '1px solid #cbd5e1', borderRadius: 7, background: '#f8fafc',
                                color: '#0078d4', fontSize: 13, textDecoration: 'none',
                            } }, "\u2197")))));
        }))),
        (vesselLevelFolders.length > 0 || vesselLevelFiles.length > 0) && (React.createElement("div", { style: { marginTop: 36, marginBottom: 0 } },
            React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 12 } },
                React.createElement("div", { style: { flex: 1, height: 1, background: 'linear-gradient(to right, #f59e0b, transparent)' } }),
                React.createElement("span", { style: { fontSize: 12, fontWeight: 700, color: '#92400e', background: '#fef3c7', padding: '4px 14px', borderRadius: 20, border: '1px solid #f59e0b', whiteSpace: 'nowrap' } }, "\u26A0\uFE0F SharePoint Items Requiring Attention"),
                React.createElement("div", { style: { flex: 1, height: 1, background: 'linear-gradient(to left, #f59e0b, transparent)' } })))),
        vesselLevelFolders.length > 0 && renderUnrecognisedFolders(host, vesselLevelFolders),
        vesselLevelFiles.length > 0 && renderVesselLevelFiles(host, vesselLevelFiles),
        renderNormalFoldersSection(host),
        React.createElement("style", null, "\n          @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }\n        "),
        renderClassifyDialog(host),
        renderProvisionDialog(host),
        modal === 'create' && host._renderVesselForm('create'),
        modal === 'edit' && host._renderVesselForm('edit'),
        modal === 'delete' && host._renderDeleteModal()));
}
