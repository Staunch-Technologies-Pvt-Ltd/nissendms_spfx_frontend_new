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
exports.renderBentoComposeModal = renderBentoComposeModal;
var React = __importStar(require("react"));
var constants_1 = require("../constants");
function renderBentoComposeModal(host) {
    var _this = this;
    var _a = host.state, bentoComposeOpen = _a.bentoComposeOpen, bentoComposeForm = _a.bentoComposeForm, bentoComposeBusy = _a.bentoComposeBusy, bentoComposeMsg = _a.bentoComposeMsg, bentoComposeErr = _a.bentoComposeErr, vessels = _a.vessels, bentoLogs = _a.bentoLogs, bentoApprovedFiles = _a.bentoApprovedFiles;
    if (!bentoComposeOpen)
        return null;
    var FIXED_RECIPIENT = host.state.bentoConfigRecipient || host.props.userEmail || '';
    var setForm = function (patch, autoSubject) {
        var merged = __assign(__assign({}, bentoComposeForm), patch);
        // Auto-fill subject whenever vessel or tag changes
        if (autoSubject) {
            merged.subject_text = host._buildAutoSubject(merged.vessel_name, merged.datasource_tag, merged.subject_suffix || 'Subject');
        }
        host.setState({ bentoComposeForm: merged });
    };
    // Approved files available for the selected vessel
    var approvedFiles = host._getApprovedFilesForVesselTag(bentoComposeForm.vessel_name, bentoComposeForm.datasource_tag);
    var vLowerKey = bentoComposeForm.vessel_name.trim().toLowerCase();
    var isLoadingFiles = !!bentoComposeForm.vessel_name && !(vLowerKey in (bentoApprovedFiles || {}));
    var handleSend = function () { return __awaiter(_this, void 0, void 0, function () {
        var tempId, pendingLog, fd, attachFileId, fileRes, blob, _a, res, data_1, e_1;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    if (!bentoComposeForm.subject_text.trim()) {
                        host.setState({ bentoComposeErr: 'Subject is required.' });
                        return [2 /*return*/];
                    }
                    host.setState({ bentoComposeBusy: true, bentoComposeErr: null, bentoComposeMsg: null });
                    tempId = Date.now();
                    pendingLog = {
                        id: tempId,
                        datasource_tag_used: bentoComposeForm.datasource_tag,
                        tag_label: constants_1.DATASOURCE_TAGS_MAP[bentoComposeForm.datasource_tag],
                        tag_was_valid: true,
                        vessel_name: bentoComposeForm.vessel_name || undefined,
                        subject: bentoComposeForm.subject_text,
                        body: bentoComposeForm.body,
                        recipient: FIXED_RECIPIENT,
                        status: 'pending',
                        attachments_count: bentoComposeForm.existing_attachment ? 1 : 0,
                        attachment_names: bentoComposeForm.existing_attachment ? [bentoComposeForm.existing_attachment] : [],
                        created_at: new Date().toISOString(),
                    };
                    host.setState({ bentoLogs: __spreadArray([pendingLog], bentoLogs, true) });
                    _b.label = 1;
                case 1:
                    _b.trys.push([1, 11, , 12]);
                    fd = new FormData();
                    fd.append('vessel_name', bentoComposeForm.vessel_name);
                    fd.append('datasource_tag', bentoComposeForm.datasource_tag);
                    fd.append('subject_text', bentoComposeForm.subject_text);
                    fd.append('body', bentoComposeForm.body);
                    fd.append('recipient', FIXED_RECIPIENT);
                    if (!bentoComposeForm.existing_attachment) return [3 /*break*/, 8];
                    attachFileId = host._getFileIdForAttachment(bentoComposeForm.vessel_name, bentoComposeForm.existing_attachment);
                    if (!attachFileId) {
                        host.setState(function (prev) { return ({
                            bentoComposeBusy: false,
                            bentoComposeErr: "Couldn't locate \"".concat(bentoComposeForm.existing_attachment, "\" as a real file \u2014 please reselect it from the list."),
                            bentoLogs: prev.bentoLogs.filter(function (l) { return l.id !== tempId; }),
                        }); });
                        return [2 /*return*/];
                    }
                    fd.append('file_id', attachFileId);
                    _b.label = 2;
                case 2:
                    _b.trys.push([2, 7, , 8]);
                    return [4 /*yield*/, fetch("".concat(host._base(), "/api/files/").concat(encodeURIComponent(attachFileId), "/content"), {
                            headers: host._uploadHeaders(),
                        })];
                case 3:
                    fileRes = _b.sent();
                    if (!fileRes.ok) return [3 /*break*/, 5];
                    return [4 /*yield*/, fileRes.blob()];
                case 4:
                    blob = _b.sent();
                    fd.append('attachment', blob, bentoComposeForm.existing_attachment);
                    return [3 /*break*/, 6];
                case 5:
                    fd.append('attachment_name', bentoComposeForm.existing_attachment);
                    _b.label = 6;
                case 6: return [3 /*break*/, 8];
                case 7:
                    _a = _b.sent();
                    fd.append('attachment_name', bentoComposeForm.existing_attachment);
                    return [3 /*break*/, 8];
                case 8: return [4 /*yield*/, fetch("".concat(host._base(), "/api/bento/dispatch"), {
                        method: 'POST', headers: host._uploadHeaders(), body: fd,
                    })];
                case 9:
                    res = _b.sent();
                    if (!res.ok)
                        throw new Error("HTTP ".concat(res.status));
                    return [4 /*yield*/, res.json().catch(function () { return ({}); })];
                case 10:
                    data_1 = _b.sent();
                    // Update the pending log to completed
                    host.setState(function (prev) { return ({
                        bentoLogs: prev.bentoLogs.map(function (l) {
                            return l.id === tempId
                                ? __assign(__assign({}, l), { id: (data_1 === null || data_1 === void 0 ? void 0 : data_1.id) || tempId, status: 'completed', sent_at: new Date().toISOString(), display_status: 'Sent' }) : l;
                        }),
                        bentoComposeBusy: false,
                        bentoComposeMsg: "\u2705 Email dispatched to ".concat(FIXED_RECIPIENT),
                        bentoComposeForm: { vessel_name: '', datasource_tag: 'mail', subject_text: '[DataSource:mail]', body: '', file: null, existing_attachment: '', recipient: FIXED_RECIPIENT },
                    }); });
                    setTimeout(function () { return host.setState({ bentoComposeOpen: false, bentoComposeMsg: null }); }, 2000);
                    return [3 /*break*/, 12];
                case 11:
                    e_1 = _b.sent();
                    // Mark log as failed
                    host.setState(function (prev) { return ({
                        bentoLogs: prev.bentoLogs.map(function (l) {
                            return l.id === tempId ? __assign(__assign({}, l), { status: 'failed', error_message: e_1 === null || e_1 === void 0 ? void 0 : e_1.message }) : l;
                        }),
                        bentoComposeBusy: false,
                        bentoComposeErr: "Failed: ".concat((e_1 === null || e_1 === void 0 ? void 0 : e_1.message) || 'Unknown error'),
                    }); });
                    return [3 /*break*/, 12];
                case 12: return [2 /*return*/];
            }
        });
    }); };
    return (React.createElement("div", { style: { position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999 }, onClick: function () { return !bentoComposeBusy && host.setState({ bentoComposeOpen: false }); } },
        React.createElement("div", { style: { background: '#fff', borderRadius: 12, padding: '28px 32px', width: 540, maxWidth: '95vw', maxHeight: '90vh', overflowY: 'auto', boxShadow: '0 12px 40px rgba(0,0,0,0.25)' }, onClick: function (e) { return e.stopPropagation(); } },
            React.createElement("div", { style: { fontSize: 17, fontWeight: 800, color: '#0f172a', marginBottom: 20, display: 'flex', alignItems: 'center', gap: 8 } }, "\u2709 Compose & Dispatch Email"),
            bentoComposeMsg && React.createElement("div", { style: { background: '#dff6dd', color: '#107c10', padding: '8px 12px', borderRadius: 6, fontSize: 12, marginBottom: 12 } }, bentoComposeMsg),
            bentoComposeErr && React.createElement("div", { style: { background: '#fde7e9', color: '#a4262c', padding: '8px 12px', borderRadius: 6, fontSize: 12, marginBottom: 12 } }, bentoComposeErr),
            React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 12 } },
                React.createElement("div", null,
                    React.createElement("label", { style: { display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 } }, "Recipient (Fixed)"),
                    React.createElement("input", { type: "text", value: FIXED_RECIPIENT, readOnly: true, style: { width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, background: '#f8fafc', color: '#64748b', boxSizing: 'border-box', cursor: 'not-allowed' } })),
                React.createElement("div", null,
                    React.createElement("label", { style: { display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 } }, "Vessel"),
                    React.createElement("select", { value: bentoComposeForm.vessel_name, onChange: function (e) {
                            var selectedVesselName = e.target.value;
                            setForm({ vessel_name: selectedVesselName, existing_attachment: '' }, true);
                            if (selectedVesselName) {
                                // Fetch approved files directly from backend
                                host._fetchApprovedFilesForVessel(selectedVesselName).catch(function () { return undefined; });
                                // Also load via existing mechanism as fallback
                                if (!host._filesLoadedForVessels.has(selectedVesselName)) {
                                    host._filesLoadedForVessels.add(selectedVesselName);
                                    host._loadFilesForVessel(selectedVesselName).catch(function () { return undefined; });
                                }
                            }
                        }, disabled: bentoComposeBusy, style: { width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, background: '#fff', outline: 'none' } },
                        React.createElement("option", { value: "" }, "\u2014 Select vessel \u2014"),
                        vessels.map(function (v) { return React.createElement("option", { key: v.id, value: v.name }, v.name); }))),
                React.createElement("div", null,
                    React.createElement("label", { style: { display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 } }, "Document Tag"),
                    React.createElement("select", { value: bentoComposeForm.datasource_tag, onChange: function (e) { return setForm({ datasource_tag: e.target.value }, true); }, disabled: bentoComposeBusy, style: { width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, background: '#fff', outline: 'none' } }, Object.keys(constants_1.DATASOURCE_TAGS_MAP).map(function (k) { return (React.createElement("option", { key: k, value: k },
                        constants_1.DATASOURCE_TAGS_MAP[k],
                        " (",
                        k,
                        ")")); }))),
                React.createElement("div", null,
                    React.createElement("label", { style: { display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 } },
                        "Subject * ",
                        React.createElement("span", { style: { fontWeight: 400, color: '#94a3b8', fontSize: 11 } }, "(auto-filled, editable)")),
                    React.createElement("input", { type: "text", value: bentoComposeForm.subject_text, onChange: function (e) {
                            var val = e.target.value;
                            var suffix = val;
                            var slashIdx = val.indexOf(' / ');
                            if (slashIdx !== -1) {
                                suffix = val.substring(slashIdx + 3);
                            }
                            setForm({ subject_text: val, subject_suffix: suffix });
                        }, disabled: bentoComposeBusy, placeholder: "[DataSource:TAG] Vessel Name / Subject...", style: { width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #0284c7', fontSize: 13, outline: 'none', boxSizing: 'border-box', background: '#f0f9ff' } })),
                React.createElement("div", null,
                    React.createElement("label", { style: { display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 } }, "Body"),
                    React.createElement("textarea", { value: bentoComposeForm.body, onChange: function (e) { return setForm({ body: e.target.value }); }, disabled: bentoComposeBusy, placeholder: "Email body...", rows: 4, style: { width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, outline: 'none', boxSizing: 'border-box', resize: 'vertical' } })),
                React.createElement("div", null,
                    React.createElement("label", { style: { display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 } },
                        "Attach Approved File",
                        React.createElement("span", { style: { fontWeight: 400, color: '#94a3b8', fontSize: 11, marginLeft: 6 } }, "(from selected vessel documents)")),
                    React.createElement("select", { value: bentoComposeForm.existing_attachment || '', onChange: function (e) { return setForm({ existing_attachment: e.target.value }); }, disabled: bentoComposeBusy || !bentoComposeForm.vessel_name || isLoadingFiles, style: { width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, background: bentoComposeForm.vessel_name ? '#fff' : '#f8fafc', outline: 'none' } }, !bentoComposeForm.vessel_name ? (React.createElement("option", { value: "" }, "\u2014 Select Vessel First \u2014")) : isLoadingFiles ? (React.createElement("option", { value: "" }, "\u23F3 Loading files... \u2014")) : approvedFiles.length > 0 ? (React.createElement(React.Fragment, null,
                        React.createElement("option", { value: "" }, "\u2014 Select Attached File \u2014"),
                        approvedFiles.map(function (f) { return React.createElement("option", { key: f, value: f },
                            "\uD83D\uDCCE ",
                            f); }))) : (React.createElement("option", { value: "" }, "\u2014 No approved files found for host vessel \u2014")))),
                React.createElement("div", { style: { background: '#f8fafc', borderRadius: 6, padding: '8px 12px', fontSize: 11, color: '#64748b', border: '1px solid #e2e8f0' } },
                    "\uD83D\uDCCA Status flow: ",
                    React.createElement("strong", null, "Pending"),
                    " (on upload) \u2192 ",
                    React.createElement("strong", null, "Pending"),
                    " (awaiting send) \u2192 ",
                    React.createElement("strong", null, "Completed"),
                    " (after dispatch)")),
            React.createElement("div", { style: { display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 20 } },
                React.createElement("button", { onClick: function () { return host.setState({ bentoComposeOpen: false, bentoComposeErr: null, bentoComposeMsg: null }); }, disabled: bentoComposeBusy, style: { background: '#f1f5f9', border: '1px solid #cbd5e1', borderRadius: 6, padding: '8px 16px', fontSize: 13, cursor: 'pointer' } }, "Cancel"),
                React.createElement("button", { onClick: handleSend, disabled: bentoComposeBusy, style: { background: '#0284c7', color: '#fff', border: 'none', borderRadius: 6, padding: '8px 20px', fontSize: 13, fontWeight: 700, cursor: 'pointer' } }, bentoComposeBusy ? '⏳ Sending...' : '✉ Send Email')))));
}
