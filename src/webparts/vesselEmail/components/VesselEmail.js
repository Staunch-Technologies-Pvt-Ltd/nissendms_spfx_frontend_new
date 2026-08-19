"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
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
var React = __importStar(require("react"));
var vesselImagePool_1 = require("./vesselImagePool");
var graphFolderService_1 = require("./graphFolderService");
var deltaSync_1 = require("./deltaSync");
var constants_1 = require("./constants");
var Sidebar_1 = require("./pages/Sidebar");
var AppLayout_1 = require("./pages/AppLayout");
var DocPreviewDrawer_1 = require("./pages/DocPreviewDrawer");
var DashboardPage_1 = require("./pages/DashboardPage");
var DocumentsPage_1 = require("./pages/DocumentsPage");
var VesselsPage_1 = require("./pages/VesselsPage");
var TemplatesPage_1 = require("./pages/TemplatesPage");
var ApprovalsPage_1 = require("./pages/ApprovalsPage");
var NotificationsPage_1 = require("./pages/NotificationsPage");
var ReportsPage_1 = require("./pages/ReportsPage");
var UsersPage_1 = require("./pages/UsersPage");
var SettingsPage_1 = require("./pages/SettingsPage");
var BentoEmailDashboardPage_1 = require("./pages/BentoEmailDashboardPage");
var RecycleBinPage_1 = require("./pages/RecycleBinPage");
var ArchivePage_1 = require("./pages/ArchivePage");
var BentoComposeModal_1 = require("./modals/BentoComposeModal");
var VesselFormModal_1 = require("./modals/VesselFormModal");
var DeleteVesselModal_1 = require("./modals/DeleteVesselModal");
var BLANK_FORM = { name: '', imo: '', shipyard: '', hull_number: '', vessel_type: '' };
var PAGE_SIZE = 50;
// ── Component Definition ────────────────────────────────────────────────────
var VesselEmail = /** @class */ (function (_super) {
    __extends(VesselEmail, _super);
    function VesselEmail(props) {
        var _this = _super.call(this, props) || this;
        _this._abort = null;
        _this._filesLoadedForVessels = new Set();
        _this._syncScheduler = null;
        _this._isLoadingData = false;
        _this._deltaReloadTimer = null;
        _this._initialDocumentFolderRefreshDone = false;
        _this._handleResize = function () { _this.setState({ windowWidth: window.innerWidth }); };
        _this._dismissAnomaly = function (id) {
            var base = _this._base();
            _this._fetchJson("".concat(base, "/api/anomalies/").concat(id), {
                method: 'PATCH',
                body: JSON.stringify({ resolved: true }),
            }).catch(function () { return undefined; });
            _this.setState(function (prev) { return ({
                folderAnomalies: prev.folderAnomalies.filter(function (a) { return a.id !== id; }),
            }); });
        };
        /**
         * Dismiss anomaly AND add a corresponding entry to the Recycle Bin state
         * so the user can see it in the existing Recycle Bin page.
         */
        _this._moveAnomalyToRecycleBin = function (anomaly) {
            // Resolve the anomaly on the backend (mark resolved = true)
            _this._dismissAnomaly(anomaly.id);
            // Build a DeletedNode from the anomaly so it appears in Recycle Bin
            var deletedNode = {
                id: "anomaly_".concat(anomaly.drive_item_id),
                name: anomaly.name,
                kind: anomaly.item_type === 'file' ? 'file' : 'folder',
                item_type: anomaly.item_type,
                original_path: anomaly.spo_path,
                main_folder: anomaly.department,
                deleted_at: new Date().toISOString(),
                ext: anomaly.item_type === 'file' ? anomaly.name.split('.').pop() : undefined,
            };
            _this.setState(function (prev) { return ({
                recycleBin: __spreadArray(__spreadArray([], prev.recycleBin, true), [deletedNode], false),
                spoAnomalyDismissConfirm: null,
            }); });
        };
        _this._fetchAnomalies = function (signal) {
            var base = _this._base();
            _this._fetchJson("".concat(base, "/api/anomalies"), signal)
                .then(function (res) {
                if (Array.isArray(res)) {
                    _this.setState({ folderAnomalies: res });
                }
            })
                .catch(function () { return undefined; });
        };
        _this._triggerAnomalyScan = function () {
            var base = _this._base();
            _this._fetchJson("".concat(base, "/api/anomalies/scan"), {
                method: 'POST',
            })
                .then(function (res) {
                if (Array.isArray(res)) {
                    _this.setState({ folderAnomalies: res });
                }
            })
                .catch(function () { return undefined; });
        };
        _this._loadNormalFolders = function () {
            var base = _this._base();
            if (!base)
                return;
            _this._fetchJson("".concat(base, "/api/normal-folders"))
                .then(function (res) {
                if (Array.isArray(res)) {
                    _this.setState({ normalFolders: res });
                }
            })
                .catch(function () { return undefined; });
        };
        _this._saveNormalFolder = function (anomaly) { return __awaiter(_this, void 0, void 0, function () {
            var base;
            return __generator(this, function (_a) {
                base = this._base();
                return [2 /*return*/, this._fetchJson("".concat(base, "/api/normal-folders"), {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            drive_item_id: anomaly.drive_item_id,
                            name: anomaly.name,
                            item_type: anomaly.item_type,
                            spo_path: anomaly.spo_path,
                            department: anomaly.department || '',
                            vessel_name: anomaly.vessel_name || null,
                        }),
                    })];
            });
        }); };
        _this._handleSignOut = function () {
            var base = _this._base();
            var sid = _this.props.sessionId || '';
            var email = _this.props.userEmail || '';
            fetch("".concat(base, "/api/auth/logout"), {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ email: email, session_id: sid }),
            }).catch(function () { return undefined; });
            // Reload the page to force re-authentication
            window.location.reload();
        };
        // ── Graph API Folder Walking (mirrors VesselListView.tsx from reference project) ────
        _this.GRAPH_FETCH_CONCURRENCY = 6;
        // Current SharePoint structure: Documents / Vessels / {Vessel} / {Main Folder}.
        // Legacy layouts are still checked below as fallbacks for existing content.
        _this.VESSEL_ROOT = 'Vessels';
        _this.MAIN_FOLDER_NAMES = [
            'Technical & Crewing',
            'Commercial & Chartering',
            'Insurance',
            'Kaizen - Knowledge Bank',
        ];
        // ── Navigation & Views ────────────────────────────────────────────────────
        _this._goToView = function (view) { return __awaiter(_this, void 0, void 0, function () {
            var data, _a, data, _b, data, _c, userEmail, base, myDataPromise, adminDataPromise, _d, myData, adminData, combined, seen, _i, _e, a, key, mapDate_1, mapped, backendIds_1, localOnly, _f, base, data, myData, combined, seen, _g, _h, a, key, notifications, _j, data, _k;
            return __generator(this, function (_l) {
                switch (_l.label) {
                    case 0:
                        this.setState({ view: view });
                        if (!(view === 'vessels')) return [3 /*break*/, 5];
                        this.setState({ panelLoading: true });
                        _l.label = 1;
                    case 1:
                        _l.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._fetchJson("".concat(this._base(), "/api/vessels"))];
                    case 2:
                        data = _l.sent();
                        if (data && Array.isArray(data)) {
                            this.setState({
                                vessels: data.map(function (v) { return (__assign(__assign({}, v), { name: (0, constants_1.cleanName)(v.name), status: v.status || 'Active' })); }),
                                panelLoading: false,
                            });
                        }
                        else {
                            // If REST API returned nothing, vessels already in state are still valid
                            this.setState({ panelLoading: false });
                        }
                        return [3 /*break*/, 4];
                    case 3:
                        _a = _l.sent();
                        this.setState({ panelLoading: false });
                        return [3 /*break*/, 4];
                    case 4: return [3 /*break*/, 30];
                    case 5:
                        if (!(view === 'recycle')) return [3 /*break*/, 10];
                        this.setState({ panelLoading: true, recycleBin: [] });
                        _l.label = 6;
                    case 6:
                        _l.trys.push([6, 8, , 9]);
                        return [4 /*yield*/, this._fetchJson("".concat(this._base(), "/api/recycle-bin/nodes"))];
                    case 7:
                        data = _l.sent();
                        this.setState({ recycleBin: (data || []).map(function (v) { return (__assign(__assign({}, v), { name: (0, constants_1.cleanName)(v.name || '') })); }), panelLoading: false });
                        return [3 /*break*/, 9];
                    case 8:
                        _b = _l.sent();
                        this.setState({ recycleBin: [], panelLoading: false });
                        return [3 /*break*/, 9];
                    case 9: return [3 /*break*/, 30];
                    case 10:
                        if (!(view === 'archive')) return [3 /*break*/, 15];
                        this.setState({ panelLoading: true, archiveList: [] });
                        _l.label = 11;
                    case 11:
                        _l.trys.push([11, 13, , 14]);
                        return [4 /*yield*/, this._fetchJson("".concat(this._base(), "/api/archive/nodes"))];
                    case 12:
                        data = _l.sent();
                        this.setState({ archiveList: (data || []).map(function (v) { return (__assign(__assign({}, v), { name: (0, constants_1.cleanName)(v.name || '') })); }), panelLoading: false });
                        return [3 /*break*/, 14];
                    case 13:
                        _c = _l.sent();
                        this.setState({ archiveList: [], panelLoading: false });
                        return [3 /*break*/, 14];
                    case 14: return [3 /*break*/, 30];
                    case 15:
                        if (!(view === 'approvals')) return [3 /*break*/, 20];
                        this.setState({ panelLoading: true });
                        _l.label = 16;
                    case 16:
                        _l.trys.push([16, 18, , 19]);
                        userEmail = this.props.userEmail || '';
                        base = this._base();
                        myDataPromise = userEmail
                            ? this._fetchJson("".concat(base, "/api/my-approvals")).catch(function () { return null; })
                            : Promise.resolve(null);
                        adminDataPromise = userEmail
                            ? this._fetchJson("".concat(base, "/api/approvals?admin=").concat(encodeURIComponent(userEmail))).catch(function () { return null; })
                            : Promise.resolve(null);
                        return [4 /*yield*/, Promise.all([myDataPromise, adminDataPromise])];
                    case 17:
                        _d = _l.sent(), myData = _d[0], adminData = _d[1];
                        combined = [];
                        seen = new Set();
                        for (_i = 0, _e = __spreadArray(__spreadArray([], (Array.isArray(adminData) ? adminData : []), true), (Array.isArray(myData) ? myData : []), true); _i < _e.length; _i++) {
                            a = _e[_i];
                            key = String(a.id || a._id || '');
                            if (!seen.has(key)) {
                                seen.add(key);
                                combined.push(a);
                            }
                        }
                        mapDate_1 = function (v) {
                            var raw = (v === null || v === void 0 ? void 0 : v.created_at) || (v === null || v === void 0 ? void 0 : v.uploaded_at) || (v === null || v === void 0 ? void 0 : v.requestedOn) || (v === null || v === void 0 ? void 0 : v.requested_on);
                            if (!raw)
                                return '—';
                            try {
                                return new Date(raw).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
                            }
                            catch (_a) {
                                return raw;
                            }
                        };
                        mapped = combined.map(function (a) { return ({
                            id: String(a.id || a._id || Date.now()),
                            documentName: a.filename || a.file_name || a.fileName || a.document_name || a.documentName || a.name || 'Unknown',
                            vessel: a.vessel_name || a.vesselName || a.vessel || '—',
                            requestedBy: a.uploaded_by_email || a.uploaded_by_name || a.requested_by || a.requestedBy || a.uploader_email || a.uploader || a.user_email || '—',
                            requestedOn: mapDate_1(a),
                            status: a.status === 'approved' ? 'Approved' : a.status === 'rejected' ? 'Rejected' : 'Pending',
                        }); });
                        backendIds_1 = new Set(mapped.map(function (m) { return m.id; }));
                        localOnly = this.state.approvalsList.filter(function (a) { return !backendIds_1.has(a.id); });
                        this.setState({ approvalsList: __spreadArray(__spreadArray([], mapped, true), localOnly, true), panelLoading: false });
                        return [3 /*break*/, 19];
                    case 18:
                        _f = _l.sent();
                        this.setState({ panelLoading: false });
                        return [3 /*break*/, 19];
                    case 19: return [3 /*break*/, 30];
                    case 20:
                        if (!(view === 'notifications')) return [3 /*break*/, 26];
                        this.setState({ panelLoading: true });
                        _l.label = 21;
                    case 21:
                        _l.trys.push([21, 24, , 25]);
                        base = this._base();
                        return [4 /*yield*/, this._fetchJson("".concat(base, "/api/approvals?admin=").concat(encodeURIComponent(this.props.userEmail || ''))).catch(function () { return null; })];
                    case 22:
                        data = _l.sent();
                        return [4 /*yield*/, this._fetchJson("".concat(base, "/api/my-approvals")).catch(function () { return null; })];
                    case 23:
                        myData = _l.sent();
                        combined = [];
                        seen = new Set();
                        for (_g = 0, _h = __spreadArray(__spreadArray([], (Array.isArray(data) ? data : []), true), (Array.isArray(myData) ? myData : []), true); _g < _h.length; _g++) {
                            a = _h[_g];
                            key = String(a.id || '');
                            if (!seen.has(key)) {
                                seen.add(key);
                                combined.push(a);
                            }
                        }
                        if (combined.length > 0) {
                            notifications = combined.map(function (a) {
                                var fname = a.filename || a.file_name || a.name || 'Document';
                                var vessel = a.vessel_name || a.vesselName || '';
                                var uploader = a.uploaded_by_email || a.uploaded_by_name || '';
                                var raw = a.uploaded_at || a.created_at || '';
                                var ts = raw ? new Date(raw).toLocaleString('en-US', { month: 'short', day: 'numeric', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '—';
                                var status = a.status || 'pending';
                                var type = status === 'approved' ? 'success' : status === 'rejected' ? 'alert' : 'warning';
                                var priority = status === 'pending' ? 'High' : 'Low';
                                var title = status === 'approved' ? 'Document Approved' : status === 'rejected' ? 'Document Rejected' : 'Approval Requested';
                                var message = status === 'approved'
                                    ? "Document approved: ".concat(fname).concat(vessel ? " for ".concat(vessel) : '')
                                    : status === 'rejected'
                                        ? "Document rejected: ".concat(fname).concat(vessel ? " for ".concat(vessel) : '')
                                        : "Approval requested: ".concat(fname).concat(vessel ? " for ".concat(vessel) : '').concat(uploader ? " by ".concat(uploader) : '');
                                return { id: String(a.id), title: title, message: message, timestamp: ts, priority: priority, read: status !== 'pending', type: type };
                            });
                            this.setState({ notificationsList: notifications, panelLoading: false });
                        }
                        else {
                            this.setState({ panelLoading: false });
                        }
                        return [3 /*break*/, 25];
                    case 24:
                        _j = _l.sent();
                        this.setState({ panelLoading: false });
                        return [3 /*break*/, 25];
                    case 25: return [3 /*break*/, 30];
                    case 26:
                        if (!(view === 'bento_email' || view === 'email_notify')) return [3 /*break*/, 30];
                        this.setState({ panelLoading: true, bentoLogs: [] });
                        _l.label = 27;
                    case 27:
                        _l.trys.push([27, 29, , 30]);
                        return [4 /*yield*/, this._fetchJson("".concat(this._base(), "/api/email-logs"))];
                    case 28:
                        data = _l.sent();
                        this.setState({ bentoLogs: data || [], panelLoading: false });
                        return [3 /*break*/, 30];
                    case 29:
                        _k = _l.sent();
                        this.setState({ bentoLogs: [], panelLoading: false });
                        return [3 /*break*/, 30];
                    case 30: return [2 /*return*/];
                }
            });
        }); };
        // ── CRUD Handlers ─────────────────────────────────────────────────────────
        _this._openCreate = function () {
            _this.setState({ modal: 'create', selectedVessel: null, form: __assign({}, BLANK_FORM), modalMsg: null, modalError: null });
        };
        _this._openEditVessel = function (v) {
            _this.setState({
                modal: 'edit',
                selectedVessel: v,
                form: {
                    name: v.name, imo: v.imo || '',
                    shipyard: v.shipyard || '', hull_number: v.hull_number || '',
                    vessel_type: v.vessel_type || '',
                },
                modalMsg: null, modalError: null,
            });
        };
        _this._openDeleteVessel = function (v) {
            var selected = v || _this.state.selectedVessel;
            _this.setState({
                modal: 'delete',
                selectedVessel: selected || null,
                deleteVesselIds: selected ? new Set([selected.id]) : new Set(),
                modalMsg: null,
                modalError: null,
            });
        };
        _this._closeModal = function () {
            if (!_this.state.modalBusy)
                _this.setState({ modal: 'none', modalMsg: null, modalError: null });
        };
        _this._submitCreate = function () { return __awaiter(_this, void 0, void 0, function () {
            var form, vessels, normalizedName, newVesselRecord, res, data, serverMsg, e_1;
            var _a, _b;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        form = this.state.form;
                        vessels = this.state.vessels;
                        normalizedName = function (value) { return value.replace(/[ _'\"]+/g, '').toLowerCase(); };
                        if (!form.name.trim()) {
                            this.setState({ modalError: null, formFieldErrors: __assign(__assign({}, this.state.formFieldErrors), { name: 'Vessel name is required.' }) });
                            return [2 /*return*/];
                        }
                        if (vessels.some(function (v) { return normalizedName(v.name) === normalizedName(form.name); })) {
                            this.setState({ modalError: null, formFieldErrors: __assign(__assign({}, this.state.formFieldErrors), { name: 'Vessel name already exists.' }) });
                            return [2 /*return*/];
                        }
                        if (!form.imo.trim()) {
                            this.setState({ modalError: 'IMO number is required.' });
                            return [2 /*return*/];
                        }
                        if (!/^\d{7}$/.test(form.imo.trim())) {
                            this.setState({ modalError: 'IMO number must be exactly 7 digits.' });
                            return [2 /*return*/];
                        }
                        this.setState({ modalBusy: true, modalError: null, modalMsg: null });
                        newVesselRecord = {
                            id: "v_".concat(Date.now()),
                            name: form.name.trim(),
                            imo: form.imo.trim(),
                            shipyard: form.shipyard.trim() || undefined,
                            hull_number: form.hull_number.trim() || undefined,
                            vessel_type: form.vessel_type || undefined,
                            status: 'Active',
                            image_url: (0, vesselImagePool_1.pickRandomVesselImage)(form.vessel_type),
                        };
                        _c.label = 1;
                    case 1:
                        _c.trys.push([1, 4, , 5]);
                        return [4 /*yield*/, fetch("".concat(this._base(), "/api/vessels"), {
                                method: 'POST', headers: this._headers(),
                                body: JSON.stringify({ name: form.name.trim(), imo: form.imo.trim(), shipyard: form.shipyard.trim() || null, hull_number: form.hull_number.trim() || null, vessel_type: form.vessel_type || null }),
                            })];
                    case 2:
                        res = _c.sent();
                        return [4 /*yield*/, res.json()];
                    case 3:
                        data = _c.sent();
                        if (!res.ok && res.status !== 202) {
                            serverMsg = (data === null || data === void 0 ? void 0 : data.detail) || (data === null || data === void 0 ? void 0 : data.message) || "Error ".concat(res.status);
                            if (/vessel name|already exists/i.test(serverMsg)) {
                                this.setState({ modalBusy: false, modalError: null, formFieldErrors: __assign(__assign({}, this.state.formFieldErrors), { name: serverMsg }) });
                            }
                            else if (/imo/i.test(serverMsg)) {
                                this.setState({ modalBusy: false, modalError: null, formFieldErrors: __assign(__assign({}, this.state.formFieldErrors), { imo: serverMsg }) });
                            }
                            else {
                                this.setState({ modalBusy: false, modalError: serverMsg });
                            }
                            return [2 /*return*/];
                        }
                        if (data.id || ((_b = data.result) === null || _b === void 0 ? void 0 : _b.id)) {
                            newVesselRecord.id = data.id || data.result.id;
                        }
                        // Add new vessel immediately to top of state grid
                        this.setState(function (prev) { return ({
                            vessels: __spreadArray([newVesselRecord], prev.vessels.filter(function (v) { return v.name.toLowerCase() !== newVesselRecord.name.toLowerCase(); }), true),
                        }); });
                        // Fire-and-forget folder provisioning — backend already handles SPO folder creation in background
                        void this._provisionVesselFolders(form.name.trim(), newVesselRecord.id).catch(function () { return undefined; });
                        // Transition modal to success screen immediately
                        this.setState({
                            modalBusy: false,
                            modalMsg: "\uD83C\uDF89 Vessel \"".concat(form.name, "\" Created & Provisioned Successfully!"),
                            modalError: null,
                        });
                        return [3 /*break*/, 5];
                    case 4:
                        e_1 = _c.sent();
                        // Offline / fallback creation
                        this.setState(function (prev) { return ({
                            vessels: __spreadArray([newVesselRecord], prev.vessels.filter(function (v) { return v.name.toLowerCase() !== newVesselRecord.name.toLowerCase(); }), true),
                            modalBusy: false,
                            modalMsg: "\uD83C\uDF89 Vessel \"".concat(form.name, "\" Created & Provisioned Successfully!"),
                            modalError: null,
                        }); });
                        return [3 /*break*/, 5];
                    case 5: return [2 /*return*/];
                }
            });
        }); };
        _this._submitEdit = function () { return __awaiter(_this, void 0, void 0, function () {
            var _a, form, selectedVessel, vessels, res, data, msg, e_2, updated;
            var _this = this;
            var _b;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        _a = this.state, form = _a.form, selectedVessel = _a.selectedVessel, vessels = _a.vessels;
                        if (!selectedVessel)
                            return [2 /*return*/];
                        this.setState({ modalBusy: true, modalError: null, modalMsg: null });
                        _c.label = 1;
                    case 1:
                        _c.trys.push([1, 4, , 5]);
                        return [4 /*yield*/, fetch("".concat(this._base(), "/api/vessels/").concat(selectedVessel.id), {
                                method: 'PATCH', headers: this._headers(),
                                body: JSON.stringify({ name: form.name.trim() || null, imo: form.imo.trim() || null, shipyard: form.shipyard.trim() || null, hull_number: form.hull_number.trim() || null, vessel_type: form.vessel_type || null }),
                            })];
                    case 2:
                        res = _c.sent();
                        return [4 /*yield*/, res.json()];
                    case 3:
                        data = _c.sent();
                        if (!res.ok && res.status !== 202)
                            throw new Error((_b = data === null || data === void 0 ? void 0 : data.message) !== null && _b !== void 0 ? _b : "Error ".concat(res.status));
                        msg = data.status === 'pending' ? '⏳ Vessel update submitted for approval.' : "\u2705 Vessel updated successfully.";
                        this.setState({ modalBusy: false, modalMsg: msg });
                        setTimeout(function () { return _this.setState({ modal: 'none', reloadKey: _this.state.reloadKey + 1 }); }, 1600);
                        return [3 /*break*/, 5];
                    case 4:
                        e_2 = _c.sent();
                        updated = vessels.map(function (v) { return v.id === selectedVessel.id ? __assign(__assign({}, v), { name: form.name.trim() || v.name, imo: form.imo.trim() || v.imo, shipyard: form.shipyard || v.shipyard, hull_number: form.hull_number || v.hull_number, vessel_type: form.vessel_type || v.vessel_type }) : v; });
                        this.setState({ vessels: updated, modalBusy: false, modalMsg: '✅ Vessel updated successfully.' });
                        setTimeout(function () { return _this.setState({ modal: 'none' }); }, 1200);
                        return [3 /*break*/, 5];
                    case 5: return [2 /*return*/];
                }
            });
        }); };
        _this._submitDelete = function () { return __awaiter(_this, void 0, void 0, function () {
            var _a, deleteVesselIds, vessels, selected, deletedIds, pendingNames, failures, _i, selected_1, vessel, res, raw, data, e_3, removedCount, messageParts, data, _b;
            var _this = this;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        _a = this.state, deleteVesselIds = _a.deleteVesselIds, vessels = _a.vessels;
                        selected = vessels.filter(function (v) { return deleteVesselIds.has(v.id); });
                        if (selected.length === 0) {
                            this.setState({ modalError: 'Select at least one vessel to delete.' });
                            return [2 /*return*/];
                        }
                        this.setState({ modalBusy: true, modalError: null });
                        deletedIds = [];
                        pendingNames = [];
                        failures = [];
                        _i = 0, selected_1 = selected;
                        _c.label = 1;
                    case 1:
                        if (!(_i < selected_1.length)) return [3 /*break*/, 7];
                        vessel = selected_1[_i];
                        _c.label = 2;
                    case 2:
                        _c.trys.push([2, 5, , 6]);
                        return [4 /*yield*/, fetch("".concat(this._base(), "/api/vessels/").concat(vessel.id, "?vessel_name=").concat(encodeURIComponent(vessel.name)), {
                                method: 'DELETE', headers: this._headers(),
                            })];
                    case 3:
                        res = _c.sent();
                        return [4 /*yield*/, res.text()];
                    case 4:
                        raw = _c.sent();
                        data = {};
                        try {
                            data = raw ? JSON.parse(raw) : {};
                        }
                        catch ( /* response is not JSON */_d) { /* response is not JSON */ }
                        if (!res.ok && res.status !== 202)
                            throw new Error((data === null || data === void 0 ? void 0 : data.message) || raw || "Error ".concat(res.status));
                        if (data.status === 'pending' || res.status === 202) {
                            pendingNames.push(vessel.name);
                        }
                        else {
                            deletedIds.push(vessel.id);
                        }
                        return [3 /*break*/, 6];
                    case 5:
                        e_3 = _c.sent();
                        failures.push("".concat(vessel.name, ": ").concat((e_3 === null || e_3 === void 0 ? void 0 : e_3.message) || 'Delete failed'));
                        return [3 /*break*/, 6];
                    case 6:
                        _i++;
                        return [3 /*break*/, 1];
                    case 7:
                        removedCount = deletedIds.length;
                        messageParts = [];
                        if (removedCount)
                            messageParts.push("".concat(removedCount, " vessel").concat(removedCount === 1 ? '' : 's', " moved to Recycle Bin."));
                        if (pendingNames.length)
                            messageParts.push("".concat(pendingNames.length, " deletion request").concat(pendingNames.length === 1 ? '' : 's', " submitted for approval."));
                        this.setState(function (prev) {
                            var _a;
                            return ({
                                modalBusy: false,
                                modalMsg: messageParts.join(' '),
                                modalError: failures.length ? failures.join(' | ') : null,
                                vessels: prev.vessels.filter(function (v) { return !deletedIds.includes(v.id); }),
                                selectedVessel: deletedIds.includes(((_a = prev.selectedVessel) === null || _a === void 0 ? void 0 : _a.id) || '') ? null : prev.selectedVessel,
                                deleteVesselIds: new Set(pendingNames.length || failures.length ? selected.filter(function (v) { return !deletedIds.includes(v.id); }).map(function (v) { return v.id; }) : []),
                            });
                        });
                        if (!removedCount) return [3 /*break*/, 11];
                        _c.label = 8;
                    case 8:
                        _c.trys.push([8, 10, , 11]);
                        return [4 /*yield*/, this._fetchJson("".concat(this._base(), "/api/recycle-bin/nodes"))];
                    case 9:
                        data = _c.sent();
                        this.setState({ recycleBin: (data || []).map(function (item) { return (__assign(__assign({}, item), { name: (0, constants_1.cleanName)(item.name || '') })); }) });
                        return [3 /*break*/, 11];
                    case 10:
                        _b = _c.sent();
                        return [3 /*break*/, 11];
                    case 11:
                        if (!failures.length && !pendingNames.length) {
                            setTimeout(function () { return _this.setState({ modal: 'none', selectedVessel: null, deleteVesselIds: new Set() }); }, 1600);
                        }
                        return [2 /*return*/];
                }
            });
        }); };
        _this._filesLoadedForFolders = new Set();
        _this._uploadSuccessTimer = null;
        _this._startUploadSuccessTimer = function () {
            if (_this._uploadSuccessTimer)
                clearInterval(_this._uploadSuccessTimer);
            _this._uploadSuccessTimer = setInterval(function () {
                _this.setState(function (prev) {
                    if (!prev.uploadSuccessPopup) {
                        clearInterval(_this._uploadSuccessTimer);
                        return null;
                    }
                    var next = prev.uploadSuccessPopup.secondsLeft - 1;
                    if (next <= 0) {
                        clearInterval(_this._uploadSuccessTimer);
                        return { uploadSuccessPopup: null };
                    }
                    return { uploadSuccessPopup: __assign(__assign({}, prev.uploadSuccessPopup), { secondsLeft: next }) };
                });
            }, 1000);
        };
        _this._handleUpload = function (row, files) { return __awaiter(_this, void 0, void 0, function () {
            var done, failed, pending, lastFolderId, lastWebUrl, lastDestPath, _i, files_1, file, result, error_1, detail, firstName, refreshFolderId_1;
            var _this = this;
            var _a, _b;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        if (!files.length)
                            return [2 /*return*/];
                        this.setState({ uploadingGroupKey: row.groupKey, uploadError: null, uploadInfo: "Uploading ".concat(files.length === 1 ? files[0].name : "".concat(files.length, " files"), "\u2026") });
                        done = 0, failed = 0, pending = 0;
                        lastFolderId = null;
                        lastWebUrl = '';
                        lastDestPath = row.subFolderPath || row.subCategory || row.category;
                        _i = 0, files_1 = files;
                        _c.label = 1;
                    case 1:
                        if (!(_i < files_1.length)) return [3 /*break*/, 6];
                        file = files_1[_i];
                        _c.label = 2;
                    case 2:
                        _c.trys.push([2, 4, , 5]);
                        return [4 /*yield*/, this._uploadFileToFolder(row.uploadFolderId, row.subFolderPath, row.vesselName, file, row.monthDriven)];
                    case 3:
                        result = _c.sent();
                        if (result.statusPending)
                            pending++;
                        else
                            done++;
                        lastFolderId = result.folderId || lastFolderId;
                        return [3 /*break*/, 5];
                    case 4:
                        error_1 = _c.sent();
                        failed++;
                        detail = error_1 instanceof Error ? error_1.message : 'Upload failed';
                        this.setState({ uploadError: detail });
                        return [3 /*break*/, 5];
                    case 5:
                        _i++;
                        return [3 /*break*/, 1];
                    case 6:
                        this.setState({ uploadingGroupKey: null, uploadInfo: null, uploadError: failed > 0 ? "\u274C ".concat(failed, " file(s) failed.") : null });
                        if (done > 0 || pending > 0) {
                            firstName = ((_a = files[0]) === null || _a === void 0 ? void 0 : _a.name) || 'file';
                            this.setState({
                                uploadSuccessPopup: {
                                    fileName: firstName,
                                    destinationPath: lastDestPath,
                                    webUrl: lastWebUrl,
                                    isPending: pending > 0 && done === 0,
                                    secondsLeft: 10,
                                },
                            });
                            this._startUploadSuccessTimer();
                        }
                        if (!(done > 0 || pending > 0)) return [3 /*break*/, 8];
                        refreshFolderId_1 = lastFolderId || row.uploadFolderId;
                        return [4 /*yield*/, this._refreshFolderFiles(refreshFolderId_1, row.groupKey, true)];
                    case 7:
                        _c.sent();
                        // Trigger delta sync immediately so SPO changes propagate
                        void ((_b = this._syncScheduler) === null || _b === void 0 ? void 0 : _b.triggerNow().catch(function () { return undefined; }));
                        // Retry only the same live folder after Graph's consistency window.
                        setTimeout(function () { void _this._refreshFolderFiles(refreshFolderId_1, row.groupKey, true); }, 1500);
                        _c.label = 8;
                    case 8: return [2 /*return*/];
                }
            });
        }); };
        // ── Existing Views: Recycle Bin ───────────────────────────────────────────
        _this._restoreFromRecycleBin = function (item) { return __awaiter(_this, void 0, void 0, function () {
            var query, res, e_4, base, _a, restoredVessel_1;
            var _b;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        _c.trys.push([0, 2, , 3]);
                        query = new URLSearchParams({
                            type: item.kind === 'file' ? 'file' : 'folder',
                            item_name: item.name,
                            department: item.main_folder || '',
                            vessel_name: item.vessel_name || '',
                        });
                        return [4 /*yield*/, fetch("".concat(this._base(), "/api/recycle-bin/restore/").concat(encodeURIComponent(item.id), "?").concat(query.toString()), {
                                method: 'POST', headers: this._headers(),
                            })];
                    case 1:
                        res = _c.sent();
                        if (!res.ok && res.status !== 202)
                            throw new Error("Restore failed (".concat(res.status, ")"));
                        if (res.status === 202) {
                            window.alert("Restore request for \"".concat(item.name, "\" was submitted for approval."));
                            return [2 /*return*/];
                        }
                        return [3 /*break*/, 3];
                    case 2:
                        e_4 = _c.sent();
                        window.alert((e_4 === null || e_4 === void 0 ? void 0 : e_4.message) || "Could not restore \"".concat(item.name, "\". Please refresh the Recycle Bin and try again."));
                        return [2 /*return*/];
                    case 3:
                        if (!(item.kind === 'vessel' || item.item_type === 'vessel')) return [3 /*break*/, 8];
                        base = this._base();
                        _c.label = 4;
                    case 4:
                        _c.trys.push([4, 6, , 7]);
                        return [4 /*yield*/, this._fetchJson("".concat(base, "/api/vessels"), {
                                method: 'POST',
                                headers: { 'Content-Type': 'application/json' },
                                body: JSON.stringify({
                                    name: item.name,
                                    imo: item.imo && item.imo !== '—' ? item.imo : null,
                                    shipyard: item.shipyard || 'Restored',
                                    vessel_type: item.vessel_type || 'Bulk Carrier',
                                }),
                            })];
                    case 5:
                        _c.sent();
                        return [3 /*break*/, 7];
                    case 6:
                        _a = _c.sent();
                        return [3 /*break*/, 7];
                    case 7:
                        restoredVessel_1 = {
                            id: item.id.replace(/^vessel_/, ''),
                            name: (0, constants_1.cleanName)(item.name),
                            imo: item.imo || '—',
                            shipyard: item.shipyard || 'Restored',
                            hull_number: item.hull_number || '',
                            vessel_type: item.vessel_type || 'Bulk Carrier',
                            status: 'Active',
                            image_url: (0, vesselImagePool_1.pickRandomVesselImage)(item.vessel_type || 'Bulk Carrier'),
                        };
                        this.setState(function (prev) { return ({
                            vessels: __spreadArray([restoredVessel_1], prev.vessels.filter(function (v) { return v.name.toLowerCase() !== restoredVessel_1.name.toLowerCase(); }), true),
                        }); });
                        _c.label = 8;
                    case 8:
                        this.setState(function (prev) { return ({ recycleBin: prev.recycleBin.filter(function (r) { return r.id !== item.id; }) }); });
                        void ((_b = this._syncScheduler) === null || _b === void 0 ? void 0 : _b.triggerNow().catch(function () { return undefined; }));
                        return [2 /*return*/];
                }
            });
        }); };
        _this._permanentDeleteFromRecycleBin = function (item, skipConfirm) { return __awaiter(_this, void 0, void 0, function () {
            var query, res, data, e_5;
            var _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (!skipConfirm && !window.confirm("Permanently delete \"".concat(item.name, "\"? This action cannot be undone.")))
                            return [2 /*return*/, { ok: false, message: 'Deletion cancelled.' }];
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 4, , 5]);
                        query = new URLSearchParams({
                            type: item.kind === 'file' ? 'file' : 'folder',
                            item_name: item.name,
                            department: item.main_folder || '',
                            vessel_name: item.vessel_name || '',
                        });
                        return [4 /*yield*/, fetch("".concat(this._base(), "/api/recycle-bin/").concat(encodeURIComponent(item.id), "?").concat(query.toString()), {
                                method: 'DELETE', headers: this._headers(),
                            })];
                    case 2:
                        res = _b.sent();
                        return [4 /*yield*/, res.json().catch(function () { return ({}); })];
                    case 3:
                        data = _b.sent();
                        if (!res.ok && res.status !== 202)
                            throw new Error((data === null || data === void 0 ? void 0 : data.message) || "Permanent delete failed (".concat(res.status, ")"));
                        if (res.status === 202) {
                            return [2 /*return*/, { ok: false, message: (data === null || data === void 0 ? void 0 : data.message) || 'Deletion was submitted for approval.' }];
                        }
                        return [3 /*break*/, 5];
                    case 4:
                        e_5 = _b.sent();
                        return [2 /*return*/, { ok: false, message: (e_5 === null || e_5 === void 0 ? void 0 : e_5.message) || 'Could not permanently delete this item.' }];
                    case 5:
                        this.setState(function (prev) { return ({ recycleBin: prev.recycleBin.filter(function (r) { return r.id !== item.id; }) }); });
                        void ((_a = this._syncScheduler) === null || _a === void 0 ? void 0 : _a.triggerNow().catch(function () { return undefined; }));
                        return [2 /*return*/, { ok: true }];
                }
            });
        }); };
        _this.state = {
            rows: [],
            vessels: [],
            loading: false, error: null, reloadKey: 0,
            textFilter: '', vesselFilter: 'all', groupFilter: 'all', catFilter: 'all',
            sort: 'default', uploadingGroupKey: null, uploadInfo: null, uploadError: null,
            selectedFileIds: new Set(), page: 0,
            modal: 'none', selectedVessel: null, deleteVesselIds: new Set(), form: __assign({}, BLANK_FORM),
            modalBusy: false, modalMsg: null, modalError: null,
            view: 'dashboard',
            recycleBin: [], archiveList: [], panelLoading: false,
            vesselsSearch: '', vesselStatusFilter: 'all', vesselTypeFilter: 'all',
            documentsList: constants_1.INITIAL_MOCK_DOCUMENTS,
            docViewMode: 'folder',
            docMainFolder: null,
            showAllVesselsInFolderView: false,
            docListPage: 0,
            docListSort: 'default',
            docGroupFilter: 'all',
            documentVesselCount: 4,
            documentVesselsLoadingMore: false,
            docUploadRowKey: null,
            docUploadBusy: false,
            docUploadMsg: null,
            folderPathStack: [],
            uploadedFilesByFolder: {},
            selectedDocPreview: null,
            templatesList: constants_1.INITIAL_MOCK_TEMPLATES,
            approvalsList: [],
            approvalTab: 'Pending',
            notificationsList: constants_1.INITIAL_MOCK_NOTIFICATIONS,
            notificationFilter: 'all',
            usersList: constants_1.INITIAL_MOCK_USERS,
            userSearch: '', userRoleFilter: 'all',
            settingsTab: 'General',
            settingsForm: {
                siteTitle: 'Vessel Documents Management',
                siteDescription: 'Manage and track all vessel related documents efficiently.',
                dateFormat: 'MM/DD/YYYY',
                timeZone: '(UTC+05:30) Chennai, Kolkata, Mumbai, New Delhi',
            },
            settingsSavedMsg: false,
            bentoConfigRecipient: '',
            bentoLogs: [], bentoStatusFilter: 'all', bentoSearch: '',
            bentoComposeOpen: false,
            bentoComposeForm: { vessel_name: '', datasource_tag: 'mail', subject_text: '', body: '', file: null, existing_attachment: '', recipient: '' },
            bentoComposeBusy: false, bentoComposeMsg: null, bentoComposeErr: null,
            bentoDetailLog: null,
            bentoUploadFile: null, bentoUploadVessel: '', bentoUploadTag: 'mail',
            bentoUploadBusy: false, bentoUploadMsg: null, bentoUploadErr: null,
            bentoApprovedFiles: {},
            bentoApprovedFileIds: {},
            folderCreationBusy: false,
            folderCreationResults: null,
            folderCreationError: null,
            folderProvisioningVesselId: null,
            provisionedVesselIds: new Set(),
            spoFolderMap: new Map(),
            lastDeltaSync: null,
            sessionExpired: false,
            sessionReady: false,
            folderAnomalies: [],
            normalFolders: [],
            spoClassifyDialog: null,
            spoProvisionDialog: null,
            spoAnomalyDismissConfirm: null,
            sidebarCollapsed: false,
            windowWidth: typeof window !== 'undefined' ? window.innerWidth : 1200,
            uploadSuccessPopup: null,
        };
        return _this;
    }
    VesselEmail.prototype.componentDidMount = function () {
        var _this = this;
        if (this.props.sessionId) {
            this.setState({ sessionReady: true });
            this._loadData();
        }
        else if (this.props.sessionInitialized) {
            // Session flow already settled with no session_id (stub/no-DB backend) —
            // do NOT wait for a sessionId that will never arrive. Load data now.
            this._loadData();
        }
        else {
            // Session not yet available (bypass-login still in flight in the webpart).
            // Retry once after 2 s — by then _ensureSession() will have completed
            // one way or another (with or without a session_id).
            setTimeout(function () {
                if (_this.props.sessionId) {
                    _this.setState({ sessionReady: true });
                    _this._loadData();
                }
                else if (_this.props.sessionInitialized) {
                    _this._loadData();
                }
                else {
                    // Last-resort fallback: even if the webpart's onInit somehow never
                    // settled (e.g. an unhandled error before the finally{} block), still
                    // attempt to load data so the vessel list isn't blank forever.
                    _this._loadData();
                }
            }, 2000);
        }
        this._startDeltaSync();
        this._loadBentoConfig();
        window.addEventListener('resize', this._handleResize);
    };
    VesselEmail.prototype._loadBentoConfig = function () {
        var _this = this;
        this._fetchJson("".concat(this._base(), "/api/email-notification/config"))
            .then(function (cfg) {
            var recipient = (cfg === null || cfg === void 0 ? void 0 : cfg.ai_bento_recipient) || (cfg === null || cfg === void 0 ? void 0 : cfg.graph_sender_mailbox) || '';
            _this.setState({ bentoConfigRecipient: recipient });
        })
            .catch(function () { return undefined; });
    };
    VesselEmail.prototype.componentDidUpdate = function (pp, ps) {
        // Mark session as ready once we receive a valid session_id prop
        if (!this.state.sessionReady && this.props.sessionId) {
            this.setState({ sessionReady: true });
        }
        // Trigger initial load when sessionId first arrives (was missing on mount)
        if (!pp.sessionId && this.props.sessionId) {
            this._filesLoadedForVessels.clear();
            this._filesLoadedForFolders.clear();
            this._loadData();
            return;
        }
        // Trigger initial load when session init settles WITHOUT a sessionId
        // (stub/no-DB backend) — otherwise componentDidMount's 2s fallback is the
        // only chance to load, and only if timing lines up.
        if (!pp.sessionInitialized && this.props.sessionInitialized && !this.props.sessionId && this.state.rows.length === 0) {
            this._loadData();
            return;
        }
        // ONLY reload data when the API URL changes or an explicit reload is requested.
        // Do NOT reload on any other state changes – that would create infinite loops.
        if (pp.apiBaseUrl !== this.props.apiBaseUrl) {
            this._filesLoadedForVessels.clear();
            this._filesLoadedForFolders.clear();
            this._loadData();
        }
        else if (ps.reloadKey !== this.state.reloadKey) {
            this._filesLoadedForVessels.clear();
            this._filesLoadedForFolders.clear();
            this._loadData();
        }
        // Reset pagination when filters change
        if (ps.textFilter !== this.state.textFilter || ps.vesselFilter !== this.state.vesselFilter ||
            ps.groupFilter !== this.state.groupFilter || ps.catFilter !== this.state.catFilter) {
            // Only call setState if page is not already 0 to avoid unnecessary re-render
            if (this.state.page !== 0)
                this.setState({ page: 0 });
        }
        // On-demand vessel row loading when user selects a specific vessel in the Documents list view.
        // Call the vessel-specific flat-tree endpoint immediately — returns in ~50ms from DB only.
        var vesselFilter = this.state.vesselFilter;
        if (ps.vesselFilter !== vesselFilter && vesselFilter !== 'all') {
            if (!this._filesLoadedForVessels.has(vesselFilter)) {
                this._filesLoadedForVessels.add(vesselFilter);
                this._loadVesselRowsFromApi(vesselFilter).catch(function () { return undefined; });
            }
        }
    };
    VesselEmail.prototype.componentWillUnmount = function () {
        var _a, _b;
        (_a = this._abort) === null || _a === void 0 ? void 0 : _a.abort();
        (_b = this._syncScheduler) === null || _b === void 0 ? void 0 : _b.stop();
        if (this._deltaReloadTimer)
            clearTimeout(this._deltaReloadTimer);
        window.removeEventListener('resize', this._handleResize);
    };
    // ── Delta Sync ────────────────────────────────────────────────────────────
    VesselEmail.prototype._startDeltaSync = function () {
        var _this = this;
        var _a = this.props, graphClient = _a.graphClient, siteId = _a.siteId, driveId = _a.driveId;
        if (!graphClient || !siteId || !driveId)
            return;
        this._syncScheduler = (0, deltaSync_1.createSyncScheduler)(graphClient, siteId, driveId, function (result) { return _this._applyDeltaResult(result); }, 30000, function (err) { return console.warn('[VesselDMS] Delta sync error:', err); });
        this._syncScheduler.start();
        // Run immediately on mount
        void this._syncScheduler.triggerNow().catch(function () { return undefined; });
    };
    VesselEmail.prototype._applyDeltaResult = function (result) {
        var _this = this;
        var _a = this.props, graphClient = _a.graphClient, siteId = _a.siteId, driveId = _a.driveId;
        if (!graphClient || !siteId || !driveId)
            return;
        this.setState(function (prev) {
            var map = new Map(prev.spoFolderMap);
            // Process deletions first
            for (var _i = 0, _a = result.deleted; _i < _a.length; _i++) {
                var id = _a[_i];
                (0, deltaSync_1.removeNodeFromMap)(map, id);
            }
            // Process additions/updates
            var missingParentIds = new Set();
            for (var _b = 0, _c = result.added; _b < _c.length; _b++) {
                var node = _c[_b];
                var missingParentId = (0, deltaSync_1.mergeNodeIntoMap)(map, node).missingParentId;
                if (missingParentId)
                    missingParentIds.add(missingParentId);
            }
            // For any node whose parent branch isn't loaded, trigger a targeted re-fetch
            // (done outside setState to avoid async inside setState)
            if (missingParentIds.size > 0) {
                setTimeout(function () {
                    missingParentIds.forEach(function (parentId) {
                        var parentNode = map.get(parentId);
                        if (parentNode) {
                            (0, deltaSync_1.fetchFolderChildren)(graphClient, siteId, driveId, parentNode.serverRelativePath)
                                .then(function (children) {
                                _this.setState(function (prev2) {
                                    var m2 = new Map(prev2.spoFolderMap);
                                    children.forEach(function (c) { return (0, deltaSync_1.mergeNodeIntoMap)(m2, c); });
                                    return { spoFolderMap: m2, lastDeltaSync: new Date() };
                                });
                            })
                                .catch(function () { return undefined; });
                        }
                    });
                }, 0);
            }
            return { spoFolderMap: map, lastDeltaSync: new Date() };
        }, function () {
            // The delta map is the source of truth for folders in this Documents
            // drive. Do the first live-file merge only after it is available.
            if (!_this._initialDocumentFolderRefreshDone && _this.state.spoFolderMap.size > 0 && _this.state.rows.length > 0) {
                _this._initialDocumentFolderRefreshDone = true;
                var count = Math.max(1, _this.state.documentVesselCount);
                _this._refreshDocumentVesselFiles(_this.state.vessels.slice(0, count).map(function (v) { return v.name; }));
            }
            // Run file refreshes after React has committed the new delta map. This
            // lets the live parent ID be matched to its Documents list row.
            var parentIds = Array.from(new Set(result.added.filter(function (n) { return !n.isFolder && n.parentId; }).map(function (n) { return n.parentId; }).filter(Boolean)));
            parentIds.forEach(function (parentId) {
                var groupKey = _this._getGroupKeyForLiveFolderId(parentId);
                void _this._refreshFolderFiles(parentId, groupKey, true).catch(function () { return undefined; });
            });
        });
        if ((result.added && result.added.length > 0) || (result.deleted && result.deleted.length > 0)) {
            // Graph is the live document source. Do not reload the backend flat tree
            // after a delta event: it can contain stale folder IDs and overwrite the
            // just-refreshed SharePoint files. The updated delta map is used for
            // subsequent navigation and uploads.
            if (this._deltaReloadTimer) {
                clearTimeout(this._deltaReloadTimer);
                this._deltaReloadTimer = null;
            }
        }
    };
    // ── HTTP Helpers ──────────────────────────────────────────────────────────
    VesselEmail.prototype._headers = function () {
        var h = { 'Content-Type': 'application/json' };
        var sid = this.props.sessionId;
        if (sid) {
            h['Authorization'] = "Bearer ".concat(sid);
            h['X-Session-ID'] = sid;
        }
        if (this.props.userEmail)
            h['X-User-Email'] = this.props.userEmail;
        return h;
    };
    // Upload headers – DO NOT set Content-Type; the browser must set it with the multipart boundary
    VesselEmail.prototype._uploadHeaders = function () {
        var h = {};
        var sid = this.props.sessionId;
        if (sid) {
            h['Authorization'] = "Bearer ".concat(sid);
            h['X-Session-ID'] = sid;
        }
        if (this.props.userEmail) {
            h['X-User-Email'] = this.props.userEmail;
        }
        return h;
    };
    VesselEmail.prototype._fetchJson = function (url, optionsOrSignal) {
        var _this = this;
        var opts = optionsOrSignal && 'aborted' in optionsOrSignal
            ? { signal: optionsOrSignal, headers: this._headers(), cache: 'no-store' }
            : __assign(__assign({ cache: 'no-store' }, (optionsOrSignal || {})), { headers: __assign(__assign({}, this._headers()), ((optionsOrSignal === null || optionsOrSignal === void 0 ? void 0 : optionsOrSignal.headers) || {})) });
        return fetch(url, opts)
            .then(function (r) {
            if (r.status === 401 || r.status === 403) {
                _this.setState({ sessionExpired: true });
                throw new Error('SESSION_EXPIRED');
            }
            if (r.status >= 500) {
                throw new Error("SERVER_ERROR_".concat(r.status));
            }
            if (!r.ok)
                throw new Error("HTTP ".concat(r.status));
            return r.json();
        })
            .catch(function (err) { return __awaiter(_this, void 0, void 0, function () {
            var fallbackUrl, r, localErr_1;
            var _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (((_a = opts.signal) === null || _a === void 0 ? void 0 : _a.aborted) || (err === null || err === void 0 ? void 0 : err.message) === 'SESSION_EXPIRED')
                            throw err;
                        if (!url.includes('nk-dms-dev.sg-nissenkaiun.com')) return [3 /*break*/, 5];
                        VesselEmail._remoteServerDown = true;
                        fallbackUrl = url.replace('https://nk-dms-dev.sg-nissenkaiun.com', 'http://localhost:8000');
                        console.warn("[VesselDMS] Remote API 502/Network error (".concat(err === null || err === void 0 ? void 0 : err.message, ") \u2014 auto-switched primary base to local backend: ").concat(fallbackUrl));
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 4, , 5]);
                        return [4 /*yield*/, fetch(fallbackUrl, opts)];
                    case 2:
                        r = _b.sent();
                        if (r.status === 401 || r.status === 403) {
                            this.setState({ sessionExpired: true });
                            throw new Error('SESSION_EXPIRED');
                        }
                        if (!r.ok)
                            throw new Error("HTTP ".concat(r.status));
                        return [4 /*yield*/, r.json()];
                    case 3: return [2 /*return*/, _b.sent()];
                    case 4:
                        localErr_1 = _b.sent();
                        if ((localErr_1 === null || localErr_1 === void 0 ? void 0 : localErr_1.message) === 'SESSION_EXPIRED')
                            throw localErr_1;
                        console.warn('[VesselDMS] Local fallback fetch failed:', localErr_1);
                        return [2 /*return*/, null];
                    case 5: throw err;
                }
            });
        }); });
    };
    VesselEmail.prototype._base = function () {
        var url = (this.props.apiBaseUrl || '').replace(/\/$/, '');
        if (VesselEmail._remoteServerDown && (!url || url.includes('nk-dms-dev.sg-nissenkaiun.com'))) {
            return 'http://localhost:8000';
        }
        return url || 'https://nk-dms-dev.sg-nissenkaiun.com';
    };
    // ── Data Loading ──────────────────────────────────────────────────────────
    VesselEmail.prototype._loadData = function () {
        return __awaiter(this, arguments, void 0, function (force) {
            var signal, base, vesselList, _a, vessels_1, _b, graphClient, siteId, driveId, vesselNodes, flatTreeParams, flatTreeUrl, flatTree, initialRows, initialVessels, rowVesselNames_1, missingVessels, supplements, initialVesselNames_1, graphRows, graphErr_1, rows, err_1;
            var _this = this;
            var _c, _d;
            if (force === void 0) { force = false; }
            return __generator(this, function (_e) {
                switch (_e.label) {
                    case 0:
                        if (this._isLoadingData && !force) {
                            return [2 /*return*/];
                        }
                        // Cancel any pending debounced reload since we're loading now
                        if (this._deltaReloadTimer) {
                            clearTimeout(this._deltaReloadTimer);
                            this._deltaReloadTimer = null;
                        }
                        this._isLoadingData = true;
                        (_c = this._abort) === null || _c === void 0 ? void 0 : _c.abort();
                        this._abort = new AbortController();
                        signal = this._abort.signal;
                        _e.label = 1;
                    case 1:
                        _e.trys.push([1, 19, 20, 21]);
                        base = this._base();
                        if (!base) {
                            this.setState({ loading: false, rows: [] });
                            return [2 /*return*/];
                        }
                        if (this.state.rows.length === 0) {
                            this.setState({ loading: true, error: null });
                        }
                        // Kick off non-critical background tasks
                        this._fetchAnomalies(signal);
                        this._loadNormalFolders();
                        vesselList = null;
                        _e.label = 2;
                    case 2:
                        _e.trys.push([2, 4, , 5]);
                        return [4 /*yield*/, this._fetchJson("".concat(base, "/api/vessels"), signal)];
                    case 3:
                        vesselList = _e.sent();
                        return [3 /*break*/, 5];
                    case 4:
                        _a = _e.sent();
                        return [3 /*break*/, 5];
                    case 5:
                        if (signal.aborted)
                            return [2 /*return*/];
                        vessels_1 = [];
                        if (vesselList && Array.isArray(vesselList) && vesselList.length > 0) {
                            vessels_1 = vesselList.map(function (v) { return (__assign(__assign({}, v), { name: (0, constants_1.cleanName)(v.name), status: v.status || 'Active' })); });
                        }
                        _b = this.props, graphClient = _b.graphClient, siteId = _b.siteId, driveId = _b.driveId;
                        if (!(vessels_1.length === 0 && graphClient && siteId && driveId)) return [3 /*break*/, 7];
                        return [4 /*yield*/, this._getGraphChildren(this.VESSEL_ROOT, signal).catch(function () { return []; })];
                    case 6:
                        vesselNodes = _e.sent();
                        if (!signal.aborted) {
                            vessels_1 = vesselNodes
                                .filter(function (node) { return node.isFolder && node.name && !/^pool-/i.test(node.name); })
                                .map(function (node) { return ({
                                id: node.id || node.name,
                                name: (0, constants_1.cleanName)(node.name),
                                status: 'Active',
                                is_provisioned: true,
                            }); });
                        }
                        _e.label = 7;
                    case 7:
                        if (graphClient && siteId && driveId) {
                            void (function () { return __awaiter(_this, void 0, void 0, function () {
                                var spoAnomalies_1, mainRootNodes, knownMainDepts, _i, mainRootNodes_1, node, cleanN, spoVesselNodes, stripVP_1, dbVNames, dbVStripped, knownCommon, _a, spoVesselNodes_1, node, cleanN, strippedN, spoErr_1;
                                return __generator(this, function (_b) {
                                    switch (_b.label) {
                                        case 0:
                                            _b.trys.push([0, 3, , 4]);
                                            spoAnomalies_1 = [];
                                            return [4 /*yield*/, this._getGraphChildren('', signal).catch(function () { return []; })];
                                        case 1:
                                            mainRootNodes = _b.sent();
                                            if (signal.aborted)
                                                return [2 /*return*/];
                                            knownMainDepts = new Set(['commercial & chartering', 'insurance', 'kaizen - knowledge bank', 'technical & crewing', 'knowledge bank']);
                                            for (_i = 0, mainRootNodes_1 = mainRootNodes; _i < mainRootNodes_1.length; _i++) {
                                                node = mainRootNodes_1[_i];
                                                if (!node.name)
                                                    continue;
                                                cleanN = (0, constants_1.cleanName)(node.name).trim().toLowerCase();
                                                if (knownMainDepts.has(cleanN))
                                                    continue;
                                                spoAnomalies_1.push({
                                                    id: Date.now() + Math.floor(Math.random() * 100000),
                                                    drive_item_id: node.id || "spo_main_".concat(node.name),
                                                    name: node.name,
                                                    item_type: node.isFolder ? 'folder' : 'file',
                                                    anomaly_type: 'main_folder_unmatched',
                                                    department: 'Main Root',
                                                    vessel_name: null,
                                                    spo_path: node.name,
                                                    resolved: false,
                                                    detected_at: new Date().toISOString(),
                                                });
                                            }
                                            return [4 /*yield*/, this._getGraphChildren('Technical & Crewing', signal).catch(function () { return []; })];
                                        case 2:
                                            spoVesselNodes = _b.sent();
                                            if (signal.aborted)
                                                return [2 /*return*/];
                                            stripVP_1 = function (name) {
                                                return (name || '').trim().toLowerCase().replace(/^(mv|m\/v|m\.v\.|mt|m\/t|m\.t\.)\s+/i, '');
                                            };
                                            dbVNames = new Set(vessels_1.map(function (v) { return (v.name || '').trim().toLowerCase(); }));
                                            dbVStripped = new Set(vessels_1.map(function (v) { return stripVP_1(v.name); }));
                                            knownCommon = new Set([
                                                'month end reports', 'service agreements', 'registration', 'drawings and manuals',
                                                'po & invoice', 'incidents', 'crewing', 'to be classified', 'common for all ships',
                                                'common (for all ships)', 'common agreements (not ship specific)', 'common (not ship specific)',
                                                'common agreements', 'common'
                                            ]);
                                            for (_a = 0, spoVesselNodes_1 = spoVesselNodes; _a < spoVesselNodes_1.length; _a++) {
                                                node = spoVesselNodes_1[_a];
                                                if (!node.name)
                                                    continue;
                                                cleanN = (0, constants_1.cleanName)(node.name).trim().toLowerCase();
                                                strippedN = stripVP_1(node.name);
                                                if (dbVNames.has(cleanN) || dbVStripped.has(strippedN) ||
                                                    cleanN.includes('common') || knownCommon.has(cleanN) || cleanN.startsWith('pool-'))
                                                    continue;
                                                spoAnomalies_1.push({
                                                    id: Date.now() + Math.floor(Math.random() * 100000),
                                                    drive_item_id: node.id || "spo_vessel_".concat(node.name),
                                                    name: node.name,
                                                    item_type: node.isFolder ? 'folder' : 'file',
                                                    anomaly_type: 'vessel_level_unmatched',
                                                    department: 'Technical & Crewing',
                                                    vessel_name: null,
                                                    spo_path: "Technical & Crewing/".concat(node.name),
                                                    resolved: false,
                                                    detected_at: new Date().toISOString(),
                                                });
                                            }
                                            if (spoAnomalies_1.length > 0) {
                                                this.setState(function (prev) {
                                                    var existing = new Set(prev.folderAnomalies.map(function (a) { return a.spo_path; }));
                                                    var newOnly = spoAnomalies_1.filter(function (a) { return !existing.has(a.spo_path); });
                                                    if (newOnly.length === 0)
                                                        return null;
                                                    return { folderAnomalies: __spreadArray(__spreadArray([], prev.folderAnomalies, true), newOnly, true) };
                                                });
                                            }
                                            return [3 /*break*/, 4];
                                        case 3:
                                            spoErr_1 = _b.sent();
                                            console.warn('[VesselDMS] SPO anomaly scanning warning:', spoErr_1);
                                            return [3 /*break*/, 4];
                                        case 4: return [2 /*return*/];
                                    }
                                });
                            }); })();
                        }
                        // Update vessel list in state immediately so UI shows vessels right away
                        if (vessels_1.length > 0) {
                            this.setState({ vessels: vessels_1 });
                        }
                        flatTreeParams = new URLSearchParams({ vessel_limit: '4', vessel_offset: '0' });
                        if (force)
                            flatTreeParams.set('force_refresh', 'true');
                        flatTreeUrl = "".concat(base, "/api/vessels/flat-tree?").concat(flatTreeParams.toString());
                        return [4 /*yield*/, this._fetchJson(flatTreeUrl, signal).catch(function () { return null; })];
                    case 8:
                        flatTree = _e.sent();
                        if (signal.aborted)
                            return [2 /*return*/];
                        if (!(flatTree && Array.isArray(flatTree) && flatTree.length > 0)) return [3 /*break*/, 11];
                        initialRows = this._normalize(flatTree);
                        initialVessels = vessels_1.slice(0, 4);
                        rowVesselNames_1 = new Set(initialRows.map(function (row) { return (0, constants_1.cleanName)(row.vesselName).trim().toLowerCase(); }));
                        missingVessels = initialVessels.filter(function (v) { return !rowVesselNames_1.has((0, constants_1.cleanName)(v.name).trim().toLowerCase()); });
                        if (!(missingVessels.length > 0)) return [3 /*break*/, 10];
                        return [4 /*yield*/, Promise.all(missingVessels.map(function (vessel) { return __awaiter(_this, void 0, void 0, function () {
                                var url, data;
                                return __generator(this, function (_a) {
                                    switch (_a.label) {
                                        case 0:
                                            url = "".concat(this._base(), "/api/vessels/flat-tree?vessel_name=").concat(encodeURIComponent(vessel.name));
                                            return [4 /*yield*/, this._fetchJson(url, signal).catch(function () { return null; })];
                                        case 1:
                                            data = _a.sent();
                                            return [2 /*return*/, Array.isArray(data) ? this._normalize(data) : []];
                                    }
                                });
                            }); }))];
                    case 9:
                        supplements = _e.sent();
                        initialRows = initialRows.concat.apply(initialRows, supplements);
                        _e.label = 10;
                    case 10:
                        initialVesselNames_1 = initialVessels.map(function (vessel) { return vessel.name; });
                        this.setState({
                            vessels: vessels_1,
                            rows: initialRows,
                            loading: false,
                            documentVesselCount: Math.min(4, vessels_1.length),
                            documentVesselsLoadingMore: false,
                            // Show all four recent vessels on the initial Documents load.
                            vesselFilter: 'all',
                            docGroupFilter: 'all',
                            catFilter: 'all',
                            docListPage: 0,
                        }, function () {
                            // Files uploaded through this SPFx web part go directly to the
                            // SharePoint library, so they are not necessarily represented by an
                            // approval/database row. Load folder contents after rendering the
                            // fast DB tree to merge those existing files into the list.
                            _this._refreshDocumentVesselFiles(initialVesselNames_1);
                        });
                        return [2 /*return*/];
                    case 11:
                        if (!(graphClient && siteId && driveId && vessels_1.length > 0)) return [3 /*break*/, 15];
                        _e.label = 12;
                    case 12:
                        _e.trys.push([12, 14, , 15]);
                        return [4 /*yield*/, this._flattenAllViaGraph(vessels_1, signal)];
                    case 13:
                        graphRows = _e.sent();
                        if (!signal.aborted) {
                            this.setState({ vessels: vessels_1, rows: this._normalize(graphRows), loading: false });
                            return [2 /*return*/];
                        }
                        return [3 /*break*/, 15];
                    case 14:
                        graphErr_1 = _e.sent();
                        console.warn('[VesselDMS] _loadData: Graph walk failed, falling back to REST API:', graphErr_1);
                        return [3 /*break*/, 15];
                    case 15:
                        if (signal.aborted)
                            return [2 /*return*/];
                        if (!(vessels_1.length > 0)) return [3 /*break*/, 17];
                        return [4 /*yield*/, this._flattenAll(vessels_1, signal)];
                    case 16:
                        rows = _e.sent();
                        if (!signal.aborted) {
                            this.setState({ vessels: vessels_1, rows: this._normalize(rows), loading: false });
                        }
                        return [3 /*break*/, 18];
                    case 17:
                        this.setState({ vessels: vessels_1, rows: [], loading: false });
                        _e.label = 18;
                    case 18: return [3 /*break*/, 21];
                    case 19:
                        err_1 = _e.sent();
                        if (!signal.aborted)
                            this.setState({ loading: false, error: (_d = err_1 === null || err_1 === void 0 ? void 0 : err_1.message) !== null && _d !== void 0 ? _d : 'Failed to load data.' });
                        return [3 /*break*/, 21];
                    case 20:
                        this._isLoadingData = false;
                        return [7 /*endfinally*/];
                    case 21: return [2 /*return*/];
                }
            });
        });
    };
    VesselEmail.prototype._normalize = function (raw) {
        return raw.map(function (r) { return (__assign(__assign({}, r), { vesselName: (0, constants_1.cleanName)(r.vesselName), group: (0, constants_1.cleanName)(r.group), category: (0, constants_1.cleanName)(r.category) })); });
    };
    /** Convert the UI breadcrumb to its location in this site's Documents library. */
    VesselEmail.prototype._sharePointFolderPath = function (subFolderPath, fallback) {
        if (fallback === void 0) { fallback = ''; }
        var parts = (subFolderPath || '').split('>').map(function (part) { return part.trim(); }).filter(Boolean);
        // UI breadcrumb and the new Documents hierarchy use the same order:
        // Vessel / Main folder / Category / Sub-category.
        return parts.length >= 2
            ? parts.join('/')
            : fallback.replace(/^\/+/, '');
    };
    /**
     * Immediately fetch rows for a specific vessel from the backend
     * (uses the vessel_name-filtered flat-tree endpoint, hits DB only, ~50ms).
     * Merges the result into `rows` so the list view updates right away.
     */
    VesselEmail.prototype._loadVesselRowsFromApi = function (vesselName) {
        return __awaiter(this, void 0, void 0, function () {
            var base, url, data, incoming_1, err_2;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        base = this._base();
                        if (!base)
                            return [2 /*return*/];
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        url = "".concat(base, "/api/vessels/flat-tree?vessel_name=").concat(encodeURIComponent(vesselName));
                        return [4 /*yield*/, this._fetchJson(url, new AbortController().signal).catch(function () { return null; })];
                    case 2:
                        data = _a.sent();
                        if (!data || !Array.isArray(data) || data.length === 0)
                            return [2 /*return*/];
                        incoming_1 = this._normalize(data);
                        this.setState(function (prev) {
                            // Replace/merge: remove old rows for this vessel, then prepend fresh ones
                            var normV = vesselName.trim().toLowerCase();
                            var kept = prev.rows.filter(function (r) { return (r.vesselName || '').trim().toLowerCase() !== normV; });
                            return { rows: __spreadArray(__spreadArray([], incoming_1, true), kept, true) };
                        }, function () {
                            // The targeted flat-tree call is DB-backed and can legitimately return
                            // a folder row with no files. Discover the actual SharePoint vessel
                            // tree, then merge its files into List view.
                            void _this._mergeLiveSharePointFiles([vesselName]);
                        });
                        return [3 /*break*/, 4];
                    case 3:
                        err_2 = _a.sent();
                        console.warn('[VesselDMS] _loadVesselRowsFromApi warning:', err_2);
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /** Load the next page of vessel document rows without replacing prior pages. */
    VesselEmail.prototype._loadMoreDocumentVessels = function () {
        return __awaiter(this, void 0, void 0, function () {
            var _a, documentVesselCount, documentVesselsLoadingMore, vessels, params, data, incoming_2, nextVesselNames_1, err_3;
            var _this = this;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _a = this.state, documentVesselCount = _a.documentVesselCount, documentVesselsLoadingMore = _a.documentVesselsLoadingMore, vessels = _a.vessels;
                        if (documentVesselsLoadingMore || documentVesselCount >= vessels.length)
                            return [2 /*return*/];
                        this.setState({ documentVesselsLoadingMore: true });
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 3, , 4]);
                        params = new URLSearchParams({
                            vessel_limit: '8',
                            vessel_offset: String(documentVesselCount),
                        });
                        return [4 /*yield*/, this._fetchJson("".concat(this._base(), "/api/vessels/flat-tree?").concat(params.toString()))];
                    case 2:
                        data = _b.sent();
                        if (!Array.isArray(data))
                            throw new Error('The next vessel page could not be loaded.');
                        incoming_2 = this._normalize(data);
                        nextVesselNames_1 = new Set(vessels.slice(documentVesselCount, documentVesselCount + 8)
                            .map(function (v) { return (0, constants_1.cleanName)(v.name).trim().toLowerCase(); }));
                        this.setState(function (prev) {
                            // A vessel can be selected independently while this request is in
                            // flight. Replace rows only for this page and retain all others.
                            var retained = prev.rows.filter(function (row) {
                                return !nextVesselNames_1.has((0, constants_1.cleanName)(row.vesselName).trim().toLowerCase());
                            });
                            return {
                                rows: __spreadArray(__spreadArray([], retained, true), incoming_2, true),
                                documentVesselCount: Math.min(prev.documentVesselCount + 8, prev.vessels.length),
                                documentVesselsLoadingMore: false,
                            };
                        }, function () {
                            _this._refreshDocumentVesselFiles(Array.from(new Set(incoming_2.map(function (row) { return row.vesselName; }))));
                        });
                        return [3 /*break*/, 4];
                    case 3:
                        err_3 = _b.sent();
                        console.warn('[VesselDMS] failed to load the next vessel page:', err_3);
                        this.setState({ documentVesselsLoadingMore: false });
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /** Merge existing SharePoint files for the vessels currently shown in Documents. */
    VesselEmail.prototype._refreshDocumentVesselFiles = function (vesselNames) {
        var uniqueVessels = Array.from(new Set(vesselNames.filter(Boolean)));
        if (uniqueVessels.length > 0)
            void this._mergeLiveSharePointFiles(uniqueVessels);
    };
    /**
     * Merge folders and files already present in SharePoint into the backend list rows.
     * This deliberately walks both the current Vessels/{vessel} layout and the
     * legacy Main Folder/{vessel} layout handled by _flattenVesselViaGraph.
     */
    VesselEmail.prototype._mergeLiveSharePointFiles = function (vesselNames) {
        return __awaiter(this, void 0, void 0, function () {
            var _a, graphClient, siteId, driveId, signal, liveFolderRows, keyFor;
            var _this = this;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _a = this.props, graphClient = _a.graphClient, siteId = _a.siteId, driveId = _a.driveId;
                        if (!graphClient || !siteId || !driveId || vesselNames.length === 0)
                            return [2 /*return*/];
                        signal = new AbortController().signal;
                        liveFolderRows = [];
                        return [4 /*yield*/, this._mapLimit(vesselNames.filter(Boolean), 2, function (vesselName) { return __awaiter(_this, void 0, void 0, function () {
                                var rows;
                                return __generator(this, function (_a) {
                                    switch (_a.label) {
                                        case 0: return [4 /*yield*/, this._flattenVesselViaGraph(vesselName, signal, function () { return undefined; }).catch(function () { return []; })];
                                        case 1:
                                            rows = _a.sent();
                                            // _walkGraphFolder emits one row for every uploadable leaf, even when
                                            // it has no files. Keeping those empty rows is essential: their
                                            // uploadFolderId is the actual SPO location used for later uploads.
                                            liveFolderRows.push.apply(liveFolderRows, rows.filter(function (row) { return Boolean(row.uploadFolderId); }));
                                            return [2 /*return*/];
                                    }
                                });
                            }); })];
                    case 1:
                        _b.sent();
                        if (liveFolderRows.length === 0)
                            return [2 /*return*/];
                        keyFor = function (row) {
                            return [row.vesselName, row.group, row.category, row.subCategory || row.category, row.subFolderPath]
                                .map(function (value) { return (0, constants_1.cleanName)(value || '').trim().toLowerCase(); })
                                .join('||');
                        };
                        this.setState(function (previous) {
                            var baseRows = new Map();
                            previous.rows.forEach(function (row) {
                                var key = keyFor(row);
                                if (!baseRows.has(key))
                                    baseRows.set(key, row);
                            });
                            var mergedFolderRows = [];
                            var affectedKeys = new Set();
                            var liveRowsByKey = new Map();
                            liveFolderRows.forEach(function (row) {
                                var key = keyFor(row);
                                var rows = liveRowsByKey.get(key) || [];
                                rows.push(row);
                                liveRowsByKey.set(key, rows);
                            });
                            liveRowsByKey.forEach(function (folderRows, semanticKey) {
                                var liveFolder = folderRows[0];
                                var baseRow = baseRows.get(semanticKey);
                                // Preserve backend metadata/group key where available, but always use
                                // the path obtained by listing the active SharePoint drive.
                                var rowBase = baseRow
                                    ? __assign(__assign({}, baseRow), { uploadFolderId: liveFolder.uploadFolderId, monthDriven: liveFolder.monthDriven, canUpload: liveFolder.canUpload, filePending: false }) : liveFolder;
                                var files = folderRows.filter(function (row) { return Boolean(row.fileName); });
                                affectedKeys.add(semanticKey);
                                if (files.length === 0) {
                                    mergedFolderRows.push(__assign(__assign({}, rowBase), { fileName: null, fileId: null, filePending: false }));
                                    return;
                                }
                                var seenFiles = new Set();
                                files.forEach(function (fileRow) {
                                    var fileKey = (fileRow.fileName || '').toLowerCase();
                                    if (seenFiles.has(fileKey))
                                        return;
                                    seenFiles.add(fileKey);
                                    mergedFolderRows.push(__assign(__assign({}, rowBase), { fileName: fileRow.fileName, fileId: fileRow.fileId, filePending: false }));
                                });
                            });
                            // Replace rows for folders confirmed by the live SPO walk, including
                            // empty folders, while preserving backend-only template rows.
                            var retained = previous.rows.filter(function (row) { return !affectedKeys.has(keyFor(row)); });
                            return { rows: __spreadArray(__spreadArray([], retained, true), mergedFolderRows, true) };
                        });
                        return [2 /*return*/];
                }
            });
        });
    };
    /** Open the exact SharePoint folder represented by a Documents list row. */
    VesselEmail.prototype._openSharePointFolder = function (row) {
        return __awaiter(this, void 0, void 0, function () {
            var target, _a, graphClient, siteId, driveId, normalisePath, sharePointPath, expectedPath, liveFolderId, _i, _b, _c, livePath, node, path, item, paths, lastError, _d, _e, candidatePath, encodedPath, candidate, pathError_1, err_4;
            return __generator(this, function (_f) {
                switch (_f.label) {
                    case 0:
                        target = window.open('', '_blank', 'noopener,noreferrer');
                        _a = this.props, graphClient = _a.graphClient, siteId = _a.siteId, driveId = _a.driveId;
                        normalisePath = function (path) {
                            return (path || '').replace(/^\/+/, '').replace(/\/+$/, '').replace(/\/vessels\//i, '/').toLocaleLowerCase();
                        };
                        sharePointPath = this._sharePointFolderPath(row.subFolderPath, '');
                        expectedPath = normalisePath(sharePointPath);
                        for (_i = 0, _b = Array.from(this.state.spoFolderMap.entries()); _i < _b.length; _i++) {
                            _c = _b[_i], livePath = _c[0], node = _c[1];
                            if (!node.isFolder)
                                continue;
                            path = normalisePath(livePath);
                            if (path === expectedPath || (expectedPath && path.endsWith("/".concat(expectedPath)))) {
                                liveFolderId = node.id;
                                break;
                            }
                        }
                        _f.label = 1;
                    case 1:
                        _f.trys.push([1, 11, , 12]);
                        if (!graphClient || !siteId || !driveId) {
                            throw new Error('SharePoint connection is not available for this web part.');
                        }
                        item = void 0;
                        if (!liveFolderId) return [3 /*break*/, 3];
                        return [4 /*yield*/, graphClient
                                .api("/sites/".concat(siteId, "/drives/").concat(driveId, "/items/").concat(liveFolderId, "?$select=id,folder,webUrl"))
                                .get()];
                    case 2:
                        item = _f.sent();
                        return [3 /*break*/, 10];
                    case 3:
                        paths = [sharePointPath, sharePointPath.replace(/\/Vessels\//i, '/')];
                        if (sharePointPath && !sharePointPath.toLocaleLowerCase().startsWith('vessel management/')) {
                            paths.push("Vessel Management/".concat(sharePointPath));
                        }
                        lastError = void 0;
                        _d = 0, _e = paths.filter(Boolean);
                        _f.label = 4;
                    case 4:
                        if (!(_d < _e.length)) return [3 /*break*/, 9];
                        candidatePath = _e[_d];
                        _f.label = 5;
                    case 5:
                        _f.trys.push([5, 7, , 8]);
                        encodedPath = candidatePath.split('/').map(function (part) { return encodeURIComponent(part); }).join('/');
                        return [4 /*yield*/, graphClient
                                .api("/sites/".concat(siteId, "/drives/").concat(driveId, "/root:/").concat(encodedPath, "?$select=id,folder,webUrl"))
                                .get()];
                    case 6:
                        candidate = _f.sent();
                        if ((candidate === null || candidate === void 0 ? void 0 : candidate.folder) && (candidate === null || candidate === void 0 ? void 0 : candidate.webUrl)) {
                            item = candidate;
                            return [3 /*break*/, 9];
                        }
                        return [3 /*break*/, 8];
                    case 7:
                        pathError_1 = _f.sent();
                        lastError = pathError_1;
                        return [3 /*break*/, 8];
                    case 8:
                        _d++;
                        return [3 /*break*/, 4];
                    case 9:
                        if (!item) {
                            throw lastError || new Error('The selected SharePoint folder was not found.');
                        }
                        _f.label = 10;
                    case 10:
                        if (!(item === null || item === void 0 ? void 0 : item.webUrl))
                            throw new Error('SharePoint did not return a folder link.');
                        if (target)
                            target.location.href = item.webUrl;
                        else
                            window.open(item.webUrl, '_blank', 'noopener,noreferrer');
                        return [3 /*break*/, 12];
                    case 11:
                        err_4 = _f.sent();
                        if (target)
                            target.close();
                        console.warn('[VesselDMS] failed to open SharePoint folder:', err_4);
                        alert((err_4 === null || err_4 === void 0 ? void 0 : err_4.message) || 'Unable to open this SharePoint folder.');
                        return [3 /*break*/, 12];
                    case 12: return [2 /*return*/];
                }
            });
        });
    };
    /** Resolve a UI breadcrumb to the matching folder in the active SPO drive. */
    VesselEmail.prototype._getLiveSharePointFolderId = function (subFolderPath) {
        var normalisePath = function (path) {
            return (path || '').replace(/^\/+/, '').replace(/\/+$/, '').replace(/\/vessels\//i, '/').toLocaleLowerCase();
        };
        var expectedPath = normalisePath(this._sharePointFolderPath(subFolderPath, ''));
        if (!expectedPath)
            return null;
        for (var _i = 0, _a = Array.from(this.state.spoFolderMap.entries()); _i < _a.length; _i++) {
            var _b = _a[_i], rawPath = _b[0], node = _b[1];
            if (!node.isFolder)
                continue;
            var livePath = normalisePath(rawPath);
            if (livePath === expectedPath || livePath.endsWith("/".concat(expectedPath)))
                return node.id;
        }
        return null;
    };
    /**
     * Upload a file to a SharePoint folder via Graph API (direct SPO upload) with REST API fallback.
     * Prevents FastAPI 404 path routing errors when folder path strings containing slashes are used.
     */
    VesselEmail.prototype._uploadFileToFolder = function (uploadFolderId_1, subFolderPath_1, vesselName_1, file_1) {
        return __awaiter(this, arguments, void 0, function (uploadFolderId, subFolderPath, vesselName, file, monthDriven) {
            var base, sharePointFolderPath, liveFolderId, resolvedSharePointFolderPath, _a, graphClient, siteId, driveId, encodedPath, folderUrl, folder, uploadUrl, item, graphErr_2, detail, form, endpoint, resp, errText, data;
            var _b;
            if (monthDriven === void 0) { monthDriven = false; }
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        base = this._base();
                        if (!base)
                            return [2 /*return*/, { fileId: null, statusPending: false, folderId: null }];
                        sharePointFolderPath = this._sharePointFolderPath(subFolderPath, uploadFolderId);
                        liveFolderId = this._getLiveSharePointFolderId(subFolderPath);
                        resolvedSharePointFolderPath = uploadFolderId.includes('/')
                            ? uploadFolderId.replace(/^\/+|\/+$/g, '')
                            : sharePointFolderPath;
                        _a = this.props, graphClient = _a.graphClient, siteId = _a.siteId, driveId = _a.driveId;
                        if (!(!monthDriven && graphClient && siteId && driveId)) return [3 /*break*/, 5];
                        _c.label = 1;
                    case 1:
                        _c.trys.push([1, 4, , 5]);
                        encodedPath = resolvedSharePointFolderPath.split('/').map(function (s) { return encodeURIComponent(s); }).join('/');
                        folderUrl = liveFolderId
                            ? "/sites/".concat(siteId, "/drives/").concat(driveId, "/items/").concat(liveFolderId, "?$select=id,name,folder,webUrl")
                            : "/sites/".concat(siteId, "/drives/").concat(driveId, "/root:/").concat(encodedPath, "?$select=id,name,folder,webUrl");
                        return [4 /*yield*/, graphClient
                                .api(folderUrl)
                                .get()];
                    case 2:
                        folder = _c.sent();
                        if (!(folder === null || folder === void 0 ? void 0 : folder.id) || !(folder === null || folder === void 0 ? void 0 : folder.folder)) {
                            throw new Error('The selected SharePoint sub-category folder was not found.');
                        }
                        uploadUrl = "/sites/".concat(siteId, "/drives/").concat(driveId, "/items/").concat(folder.id, ":/").concat(encodeURIComponent(file.name), ":/content");
                        return [4 /*yield*/, graphClient.api(uploadUrl).put(file)];
                    case 3:
                        item = _c.sent();
                        if (!(item === null || item === void 0 ? void 0 : item.id))
                            throw new Error('SharePoint did not return an uploaded file ID.');
                        if (((_b = item.parentReference) === null || _b === void 0 ? void 0 : _b.id) && item.parentReference.id !== folder.id) {
                            throw new Error('SharePoint saved the file outside of the selected sub-category folder.');
                        }
                        return [2 /*return*/, { fileId: item.id, statusPending: false, folderId: folder.id }];
                    case 4:
                        graphErr_2 = _c.sent();
                        detail = graphErr_2 instanceof Error ? graphErr_2.message : String(graphErr_2);
                        throw new Error("Could not upload to the SharePoint folder \"".concat(resolvedSharePointFolderPath, "\": ").concat(detail));
                    case 5:
                        form = new FormData();
                        form.append('file', file);
                        if (this.props.userEmail)
                            form.append('uploader_email', this.props.userEmail);
                        if (uploadFolderId.includes('/')) {
                            endpoint = "".concat(base, "/api/folders/upload-by-path?path=").concat(encodeURIComponent(uploadFolderId));
                        }
                        else {
                            endpoint = monthDriven
                                ? "".concat(base, "/api/folders/").concat(encodeURIComponent(uploadFolderId), "/month-upload")
                                : "".concat(base, "/api/folders/").concat(encodeURIComponent(uploadFolderId), "/upload");
                        }
                        return [4 /*yield*/, fetch(endpoint, {
                                method: 'POST',
                                headers: this._uploadHeaders(),
                                body: form,
                            })];
                    case 6:
                        resp = _c.sent();
                        if (!!resp.ok) return [3 /*break*/, 8];
                        return [4 /*yield*/, resp.text().catch(function () { return resp.statusText; })];
                    case 7:
                        errText = _c.sent();
                        throw new Error("HTTP ".concat(resp.status, ": ").concat(errText));
                    case 8: return [4 /*yield*/, resp.json().catch(function () { return ({}); })];
                    case 9:
                        data = _c.sent();
                        return [2 /*return*/, { fileId: (data === null || data === void 0 ? void 0 : data.id) || null, statusPending: (data === null || data === void 0 ? void 0 : data.status) === 'pending', folderId: uploadFolderId || null }];
                }
            });
        });
    };
    /**
     * List children of a SharePoint drive path via Graph API.
     * Returns an array of { id, name, isFolder, monthDriven, upload } items.
     */
    VesselEmail.prototype._getGraphChildren = function (folderPath, signal) {
        return __awaiter(this, void 0, void 0, function () {
            var _a, graphClient, siteId, driveId, fetchChildrenForPath, kids, vmKids, err_5, status_1, _b;
            var _this = this;
            var _c, _d;
            return __generator(this, function (_e) {
                switch (_e.label) {
                    case 0:
                        _a = this.props, graphClient = _a.graphClient, siteId = _a.siteId, driveId = _a.driveId;
                        if (!graphClient || !siteId || !driveId)
                            return [2 /*return*/, []];
                        fetchChildrenForPath = function (p) { return __awaiter(_this, void 0, void 0, function () {
                            var cleanP, encodedPath, itemPath, url, result;
                            var _a;
                            return __generator(this, function (_b) {
                                switch (_b.label) {
                                    case 0:
                                        cleanP = p.replace(/^\/+/, '');
                                        encodedPath = cleanP.split('/').map(function (s) { return encodeURIComponent(s); }).join('/');
                                        itemPath = cleanP ? "root:/".concat(encodedPath, ":/children") : 'root/children';
                                        url = "/sites/".concat(siteId, "/drives/").concat(driveId, "/").concat(itemPath) +
                                            "?$select=id,name,folder,file&$top=200";
                                        console.log("[VesselDMS] _getGraphChildren \u2192 ".concat(url));
                                        return [4 /*yield*/, graphClient.api(url).get()];
                                    case 1:
                                        result = _b.sent();
                                        if (signal.aborted)
                                            return [2 /*return*/, []];
                                        return [2 /*return*/, ((_a = result.value) !== null && _a !== void 0 ? _a : []).map(function (item) { return ({
                                                id: item.id,
                                                name: item.name,
                                                isFolder: !!item.folder,
                                                monthDriven: false,
                                                upload: !!item.folder,
                                            }); })];
                                }
                            });
                        }); };
                        _e.label = 1;
                    case 1:
                        _e.trys.push([1, 5, , 10]);
                        return [4 /*yield*/, fetchChildrenForPath(folderPath)];
                    case 2:
                        kids = _e.sent();
                        if (kids.length > 0)
                            return [2 /*return*/, kids];
                        if (!!folderPath.startsWith('Vessel Management')) return [3 /*break*/, 4];
                        return [4 /*yield*/, fetchChildrenForPath("Vessel Management/".concat(folderPath)).catch(function () { return []; })];
                    case 3:
                        vmKids = _e.sent();
                        if (vmKids.length > 0)
                            return [2 /*return*/, vmKids];
                        _e.label = 4;
                    case 4: return [2 /*return*/, kids];
                    case 5:
                        err_5 = _e.sent();
                        status_1 = (_d = (_c = err_5 === null || err_5 === void 0 ? void 0 : err_5.statusCode) !== null && _c !== void 0 ? _c : err_5 === null || err_5 === void 0 ? void 0 : err_5.code) !== null && _d !== void 0 ? _d : 0;
                        if (!(status_1 === 404 && !folderPath.startsWith('Vessel Management'))) return [3 /*break*/, 9];
                        _e.label = 6;
                    case 6:
                        _e.trys.push([6, 8, , 9]);
                        return [4 /*yield*/, fetchChildrenForPath("Vessel Management/".concat(folderPath))];
                    case 7: return [2 /*return*/, _e.sent()];
                    case 8:
                        _b = _e.sent();
                        return [2 /*return*/, []];
                    case 9:
                        if (status_1 === 404)
                            return [2 /*return*/, []];
                        throw err_5;
                    case 10: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Recursively walk a Graph folder path and emit FlatRow entries.
     * Mirrors the walkFolder() + leafToRows() pattern from the reference project's VesselListView.tsx.
     */
    VesselEmail.prototype._walkGraphFolder = function (folderPath, pathParts, vesselName, signal, srCounter, onRows) {
        return __awaiter(this, void 0, void 0, function () {
            var kids, _a, files, subFolders, stripPrefix, group, category, subCategory, breadcrumbParts, subPath, groupKey, canUpload, baseSr_1, suffixes_1, leafRows;
            var _this = this;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (signal.aborted)
                            return [2 /*return*/];
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._getGraphChildren(folderPath, signal)];
                    case 2:
                        kids = _b.sent();
                        return [3 /*break*/, 4];
                    case 3:
                        _a = _b.sent();
                        return [2 /*return*/];
                    case 4:
                        if (signal.aborted)
                            return [2 /*return*/];
                        files = kids.filter(function (k) { return !k.isFolder; });
                        subFolders = kids.filter(function (k) { return k.isFolder; });
                        stripPrefix = function (s) { return s.replace(/^Folder-\d+\s+/i, ''); };
                        group = pathParts.length >= 1 ? stripPrefix(pathParts[0]) : '';
                        category = pathParts.length >= 3 ? stripPrefix(pathParts[2]) : (pathParts.length >= 2 ? stripPrefix(pathParts[pathParts.length - 1]) : group);
                        subCategory = pathParts.length >= 4 ? stripPrefix(pathParts[pathParts.length - 1]) : category;
                        breadcrumbParts = [vesselName, group, category];
                        if (subCategory && subCategory !== category)
                            breadcrumbParts.push(subCategory);
                        subPath = breadcrumbParts.join(' > ');
                        groupKey = "".concat(vesselName, "||").concat(group, "||").concat(category, "||").concat(subCategory, "||").concat(subPath);
                        canUpload = subFolders.length === 0;
                        // Emit a leaf row (matches reference leafToRows pattern)
                        if (files.length > 0 || subFolders.length === 0) {
                            srCounter.value += 1;
                            baseSr_1 = String(srCounter.value);
                            suffixes_1 = 'abcdefghijklmnopqrstuvwxyz';
                            if (files.length === 0) {
                                onRows([{
                                        srNo: baseSr_1,
                                        vesselName: vesselName,
                                        group: group,
                                        category: category,
                                        subCategory: subCategory,
                                        subFolderPath: subPath, fileName: null, fileId: null,
                                        canUpload: canUpload,
                                        groupKey: groupKey,
                                        uploadFolderId: folderPath,
                                        monthDriven: false,
                                    }]);
                            }
                            else {
                                leafRows = files.map(function (f, idx) {
                                    var _a;
                                    return ({
                                        srNo: idx === 0 ? baseSr_1 : "".concat(baseSr_1).concat((_a = suffixes_1[idx - 1]) !== null && _a !== void 0 ? _a : idx),
                                        vesselName: vesselName,
                                        group: group,
                                        category: category,
                                        subCategory: subCategory,
                                        subFolderPath: subPath,
                                        fileName: f.name, fileId: f.id,
                                        canUpload: canUpload,
                                        groupKey: groupKey,
                                        uploadFolderId: folderPath,
                                        monthDriven: false,
                                    });
                                });
                                onRows(leafRows);
                            }
                        }
                        if (!(subFolders.length > 0 && !signal.aborted)) return [3 /*break*/, 6];
                        return [4 /*yield*/, this._mapLimit(subFolders, this.GRAPH_FETCH_CONCURRENCY, function (sf) { return __awaiter(_this, void 0, void 0, function () {
                                return __generator(this, function (_a) {
                                    return [2 /*return*/, this._walkGraphFolder("".concat(folderPath, "/").concat(sf.name), __spreadArray(__spreadArray([], pathParts, true), [sf.name], false), vesselName, signal, srCounter, onRows)];
                                });
                            }); })];
                    case 5:
                        _b.sent();
                        _b.label = 6;
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    /** Bounded-concurrency async map (mirrors mapLimit from reference project). */
    VesselEmail.prototype._mapLimit = function (items, limit, mapper) {
        return __awaiter(this, void 0, void 0, function () {
            var results, nextIndex, runWorker, workers;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (items.length === 0)
                            return [2 /*return*/, []];
                        results = new Array(items.length);
                        nextIndex = 0;
                        runWorker = function () { return __awaiter(_this, void 0, void 0, function () {
                            var i, _a, _b;
                            return __generator(this, function (_c) {
                                switch (_c.label) {
                                    case 0:
                                        i = nextIndex++;
                                        _c.label = 1;
                                    case 1:
                                        if (!(i < items.length)) return [3 /*break*/, 3];
                                        _a = results;
                                        _b = i;
                                        return [4 /*yield*/, mapper(items[i], i)];
                                    case 2:
                                        _a[_b] = _c.sent();
                                        i = nextIndex++;
                                        return [3 /*break*/, 1];
                                    case 3: return [2 /*return*/];
                                }
                            });
                        }); };
                        workers = Array.from({ length: Math.min(limit, items.length) }, runWorker);
                        return [4 /*yield*/, Promise.all(workers)];
                    case 1:
                        _a.sent();
                        return [2 /*return*/, results];
                }
            });
        });
    };
    /**
     * Flatten a single vessel's SPO folder tree via Graph API.
     * Calls onChunk with each batch of rows as they arrive (progressive rendering).
     */
    VesselEmail.prototype._flattenVesselViaGraph = function (vesselName, signal, onChunk) {
        return __awaiter(this, void 0, void 0, function () {
            var allRows, srCounter, normaliseName, knownMainFolders, discoveredMainFolders, foundSupportedStructure, basePaths, _loop_1, this_1, _i, basePaths_1, basePath, state_1;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        allRows = [];
                        srCounter = { value: 0 };
                        normaliseName = function (name) { return (0, constants_1.cleanName)(name || '')
                            .replace(/^Folder-\d+\s+/i, '')
                            .trim()
                            .toLowerCase(); };
                        knownMainFolders = new Set(this.MAIN_FOLDER_NAMES.map(normaliseName));
                        discoveredMainFolders = [];
                        foundSupportedStructure = false;
                        basePaths = ['', 'Vessel Management'];
                        _loop_1 = function (basePath) {
                            var baseChildren, prefix, vesselsRoot, vesselNodes, vesselNode, vesselPath_1, mainNodes, mainNodes, _b, mainNodes_1, mainNode, mainPath, mainChildren, vesselNode, vesselPath, legacyVesselsRoot, vesselNodes;
                            return __generator(this, function (_c) {
                                switch (_c.label) {
                                    case 0:
                                        if (signal.aborted || discoveredMainFolders.length > 0)
                                            return [2 /*return*/, "break"];
                                        return [4 /*yield*/, this_1._getGraphChildren(basePath, signal).catch(function () { return []; })];
                                    case 1:
                                        baseChildren = _c.sent();
                                        prefix = basePath ? "".concat(basePath, "/") : '';
                                        vesselsRoot = baseChildren.find(function (node) { return node.isFolder && normaliseName(node.name) === 'vessels'; });
                                        if (!vesselsRoot) return [3 /*break*/, 4];
                                        foundSupportedStructure = true;
                                        return [4 /*yield*/, this_1._getGraphChildren("".concat(prefix).concat(vesselsRoot.name), signal).catch(function () { return []; })];
                                    case 2:
                                        vesselNodes = _c.sent();
                                        vesselNode = vesselNodes.find(function (node) { return node.isFolder && normaliseName(node.name) === normaliseName(vesselName); });
                                        if (!vesselNode) return [3 /*break*/, 4];
                                        vesselPath_1 = "".concat(prefix).concat(vesselsRoot.name, "/").concat(vesselNode.name);
                                        return [4 /*yield*/, this_1._getGraphChildren(vesselPath_1, signal).catch(function () { return []; })];
                                    case 3:
                                        mainNodes = _c.sent();
                                        mainNodes.filter(function (node) { return node.isFolder; }).forEach(function (node) {
                                            discoveredMainFolders.push({ path: "".concat(vesselPath_1, "/").concat(node.name), group: node.name });
                                        });
                                        _c.label = 4;
                                    case 4:
                                        if (!(discoveredMainFolders.length === 0)) return [3 /*break*/, 10];
                                        mainNodes = baseChildren.filter(function (node) { return node.isFolder && knownMainFolders.has(normaliseName(node.name)); });
                                        if (mainNodes.length > 0)
                                            foundSupportedStructure = true;
                                        _b = 0, mainNodes_1 = mainNodes;
                                        _c.label = 5;
                                    case 5:
                                        if (!(_b < mainNodes_1.length)) return [3 /*break*/, 10];
                                        mainNode = mainNodes_1[_b];
                                        if (signal.aborted)
                                            return [3 /*break*/, 10];
                                        mainPath = "".concat(prefix).concat(mainNode.name);
                                        return [4 /*yield*/, this_1._getGraphChildren(mainPath, signal).catch(function () { return []; })];
                                    case 6:
                                        mainChildren = _c.sent();
                                        vesselNode = mainChildren.find(function (node) { return node.isFolder && normaliseName(node.name) === normaliseName(vesselName); });
                                        vesselPath = vesselNode ? "".concat(mainPath, "/").concat(vesselNode.name) : '';
                                        if (!!vesselNode) return [3 /*break*/, 8];
                                        legacyVesselsRoot = mainChildren.find(function (node) { return node.isFolder && normaliseName(node.name) === 'vessels'; });
                                        if (!legacyVesselsRoot) return [3 /*break*/, 8];
                                        return [4 /*yield*/, this_1._getGraphChildren("".concat(mainPath, "/").concat(legacyVesselsRoot.name), signal).catch(function () { return []; })];
                                    case 7:
                                        vesselNodes = _c.sent();
                                        vesselNode = vesselNodes.find(function (node) { return node.isFolder && normaliseName(node.name) === normaliseName(vesselName); });
                                        vesselPath = vesselNode ? "".concat(mainPath, "/").concat(legacyVesselsRoot.name, "/").concat(vesselNode.name) : '';
                                        _c.label = 8;
                                    case 8:
                                        if (vesselNode && vesselPath)
                                            discoveredMainFolders.push({ path: vesselPath, group: mainNode.name });
                                        _c.label = 9;
                                    case 9:
                                        _b++;
                                        return [3 /*break*/, 5];
                                    case 10: return [2 /*return*/];
                                }
                            });
                        };
                        this_1 = this;
                        _i = 0, basePaths_1 = basePaths;
                        _a.label = 1;
                    case 1:
                        if (!(_i < basePaths_1.length)) return [3 /*break*/, 4];
                        basePath = basePaths_1[_i];
                        return [5 /*yield**/, _loop_1(basePath)];
                    case 2:
                        state_1 = _a.sent();
                        if (state_1 === "break")
                            return [3 /*break*/, 4];
                        _a.label = 3;
                    case 3:
                        _i++;
                        return [3 /*break*/, 1];
                    case 4:
                        if (!(discoveredMainFolders.length > 0)) return [3 /*break*/, 6];
                        return [4 /*yield*/, this._mapLimit(discoveredMainFolders, 3, function (folder) { return __awaiter(_this, void 0, void 0, function () {
                                return __generator(this, function (_a) {
                                    switch (_a.label) {
                                        case 0: return [4 /*yield*/, this._walkGraphFolder(folder.path, [folder.group, vesselName], vesselName, signal, srCounter, function (rows) {
                                                allRows.push.apply(allRows, rows);
                                                onChunk(rows);
                                            })];
                                        case 1:
                                            _a.sent();
                                            return [2 /*return*/];
                                    }
                                });
                            }); })];
                    case 5:
                        _a.sent();
                        return [2 /*return*/, allRows];
                    case 6:
                        // A recognised vessel layout exists, but this vessel does not exist in
                        // it. Do not probe every template path: those requests only create noisy
                        // Graph 404s and cannot produce files for the requested vessel.
                        if (foundSupportedStructure)
                            return [2 /*return*/, allRows];
                        // Walk each main folder in parallel
                        return [4 /*yield*/, this._mapLimit(this.MAIN_FOLDER_NAMES, 3, function (mainFolderName) { return __awaiter(_this, void 0, void 0, function () {
                                var vesselPath, topCats, legacyPaths, _i, legacyPaths_1, legacyPath, legacyCats, folders, files, displayMainFolder, templateTree, cats, fallbackRows, _a, cats_1, item, _b, _c, subCat, breadcrumb, logicalPath;
                                var _this = this;
                                return __generator(this, function (_d) {
                                    switch (_d.label) {
                                        case 0:
                                            if (signal.aborted)
                                                return [2 /*return*/];
                                            vesselPath = "".concat(this.VESSEL_ROOT, "/").concat(vesselName, "/").concat(mainFolderName);
                                            return [4 /*yield*/, this._getGraphChildren(vesselPath, signal).catch(function () { return []; })];
                                        case 1:
                                            topCats = _d.sent();
                                            if (!(topCats.length === 0)) return [3 /*break*/, 5];
                                            legacyPaths = [
                                                "".concat(mainFolderName, "/Vessels/").concat(vesselName),
                                                "".concat(mainFolderName, "/").concat(vesselName),
                                            ];
                                            _i = 0, legacyPaths_1 = legacyPaths;
                                            _d.label = 2;
                                        case 2:
                                            if (!(_i < legacyPaths_1.length)) return [3 /*break*/, 5];
                                            legacyPath = legacyPaths_1[_i];
                                            return [4 /*yield*/, this._getGraphChildren(legacyPath, signal).catch(function () { return []; })];
                                        case 3:
                                            legacyCats = _d.sent();
                                            if (legacyCats.length > 0) {
                                                vesselPath = legacyPath;
                                                topCats = legacyCats;
                                                return [3 /*break*/, 5];
                                            }
                                            _d.label = 4;
                                        case 4:
                                            _i++;
                                            return [3 /*break*/, 2];
                                        case 5:
                                            if (signal.aborted)
                                                return [2 /*return*/];
                                            folders = topCats.filter(function (c) { return c.isFolder; });
                                            files = topCats.filter(function (c) { return !c.isFolder; });
                                            if (folders.length === 0) {
                                                displayMainFolder = mainFolderName.replace(/^Folder-\d+\s+/i, '');
                                                templateTree = {
                                                    'Technical & Crewing': [
                                                        { category: 'Month End Reports', subCats: ['Main Engine', 'Aux Engine', 'Cooling Water', 'Inspection Reports', 'Defect Reports', 'Guarantee Claims', 'To be Classified'] },
                                                        { category: 'Service Agreements', subCats: ['Technical Management', 'Crew Management', 'Vendor & Service Provider', 'To be Classified'] },
                                                        { category: 'Registration', subCats: ['Flag & MPA', 'Ship Builder', 'Radio & Telecom', 'Crewing & SMOU', 'Novation', 'To be Classified'] },
                                                        { category: 'Drawings and Manuals', subCats: ['Drawing', 'Manual', 'To be Classified'] },
                                                        { category: 'PO & Invoice', subCats: ['Purchase Order and Vendor Invoice'] },
                                                        { category: 'Incidents', subCats: ['Incidents'] },
                                                        { category: 'Crewing', subCats: ['Crewing'] },
                                                        { category: 'To be Classified', subCats: ['To be Classified'] },
                                                    ],
                                                    'Commercial & Chartering': [
                                                        { category: 'Agreements', subCats: ['Charter party', 'Pool Agreement', 'Commission Agreement', 'To be Classified'] },
                                                        { category: 'Invoices & Payments', subCats: ['Invoice', 'Payments', 'To be Classified'] },
                                                        { category: 'Claims & Disputes', subCats: ['Disputes', 'Claims', 'To be Classified'] },
                                                        { category: 'To be Classified', subCats: ['To be Classified'] },
                                                    ],
                                                    'Insurance': [
                                                        { category: 'P&I', subCats: ['P&I'] },
                                                        { category: 'H&M', subCats: ['H&M'] },
                                                        { category: 'War Risk', subCats: ['War Risk'] },
                                                        { category: 'Flag & MPA', subCats: ['Flag & MPA'] },
                                                        { category: 'USA Related', subCats: ['USA Related'] },
                                                    ],
                                                    'Kaizen - Knowledge Bank': [
                                                        { category: 'Circulars and Guidance', subCats: ['Equipment Maker', 'Class', 'Flag / Port State', 'SIRE/OCIMF/RightShip', 'Shipyard'] },
                                                        { category: 'Lessons Learned', subCats: ['Lessons Learned'] },
                                                        { category: 'Procedures and Work Instructions', subCats: ['Procedures and Work Instructions'] },
                                                        { category: 'Templates', subCats: ['Templates'] },
                                                    ],
                                                };
                                                cats = templateTree[displayMainFolder] || [
                                                    { category: displayMainFolder, subCats: [displayMainFolder] }
                                                ];
                                                fallbackRows = [];
                                                for (_a = 0, cats_1 = cats; _a < cats_1.length; _a++) {
                                                    item = cats_1[_a];
                                                    for (_b = 0, _c = item.subCats; _b < _c.length; _b++) {
                                                        subCat = _c[_b];
                                                        srCounter.value += 1;
                                                        breadcrumb = subCat !== item.category
                                                            ? "".concat(vesselName, " > ").concat(displayMainFolder, " > ").concat(item.category, " > ").concat(subCat)
                                                            : "".concat(vesselName, " > ").concat(displayMainFolder, " > ").concat(item.category);
                                                        logicalPath = "".concat(vesselName, "/").concat(displayMainFolder, "/").concat(item.category, "/").concat(subCat);
                                                        fallbackRows.push({
                                                            srNo: String(srCounter.value),
                                                            vesselName: vesselName,
                                                            group: displayMainFolder,
                                                            category: item.category,
                                                            subCategory: subCat,
                                                            subFolderPath: breadcrumb,
                                                            fileName: null,
                                                            fileId: null,
                                                            canUpload: true,
                                                            groupKey: "".concat(vesselName, "||").concat(displayMainFolder, "||").concat(item.category, "||").concat(subCat, "||").concat(breadcrumb),
                                                            uploadFolderId: logicalPath,
                                                            monthDriven: false,
                                                        });
                                                    }
                                                }
                                                allRows.push.apply(allRows, fallbackRows);
                                                onChunk(fallbackRows);
                                                return [2 /*return*/];
                                            }
                                            // Walk each category folder with bounded concurrency
                                            return [4 /*yield*/, this._mapLimit(folders, this.GRAPH_FETCH_CONCURRENCY, function (cat) { return __awaiter(_this, void 0, void 0, function () {
                                                    var chunkRows;
                                                    return __generator(this, function (_a) {
                                                        switch (_a.label) {
                                                            case 0:
                                                                if (signal.aborted)
                                                                    return [2 /*return*/];
                                                                chunkRows = [];
                                                                return [4 /*yield*/, this._walkGraphFolder("".concat(vesselPath, "/").concat(cat.name), [mainFolderName, vesselName, cat.name], vesselName, signal, srCounter, function (rows) {
                                                                        chunkRows.push.apply(chunkRows, rows);
                                                                        allRows.push.apply(allRows, rows);
                                                                    })];
                                                            case 1:
                                                                _a.sent();
                                                                if (chunkRows.length > 0)
                                                                    onChunk(chunkRows);
                                                                return [2 /*return*/];
                                                        }
                                                    });
                                                }); })];
                                        case 6:
                                            // Walk each category folder with bounded concurrency
                                            _d.sent();
                                            return [2 /*return*/];
                                    }
                                });
                            }); })];
                    case 7:
                        // Walk each main folder in parallel
                        _a.sent();
                        return [2 /*return*/, allRows];
                }
            });
        });
    };
    /**
     * Flatten ALL vessels via Graph API (called from _loadData when Graph context is available).
     * Streams rows progressively into state via onChunk for a responsive UI.
     */
    VesselEmail.prototype._flattenAllViaGraph = function (vessels, signal) {
        return __awaiter(this, void 0, void 0, function () {
            var allRows;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        allRows = [];
                        return [4 /*yield*/, this._mapLimit(vessels, 3, function (vessel) { return __awaiter(_this, void 0, void 0, function () {
                                return __generator(this, function (_a) {
                                    switch (_a.label) {
                                        case 0:
                                            if (signal.aborted)
                                                return [2 /*return*/];
                                            return [4 /*yield*/, this._flattenVesselViaGraph(vessel.name, signal, function (chunk) {
                                                    allRows.push.apply(allRows, chunk);
                                                })];
                                        case 1:
                                            _a.sent();
                                            return [2 /*return*/];
                                    }
                                });
                            }); })];
                    case 1:
                        _a.sent();
                        return [2 /*return*/, allRows];
                }
            });
        });
    };
    VesselEmail.prototype._flattenAll = function (vessels, signal) {
        return __awaiter(this, void 0, void 0, function () {
            var base, out, sr, _i, vessels_2, v, mains, _a, mains_1, m, group, cats, _b, cats_2, cat, category, subCats, leaves, _c, leaves_1, leaf, subCategory, sp, sp, _d, sp, _e, _f;
            return __generator(this, function (_g) {
                switch (_g.label) {
                    case 0:
                        base = this._base();
                        out = [];
                        sr = 0;
                        _i = 0, vessels_2 = vessels;
                        _g.label = 1;
                    case 1:
                        if (!(_i < vessels_2.length)) return [3 /*break*/, 18];
                        v = vessels_2[_i];
                        if (signal.aborted)
                            return [3 /*break*/, 18];
                        _g.label = 2;
                    case 2:
                        _g.trys.push([2, 16, , 17]);
                        return [4 /*yield*/, this._fetchJson("".concat(base, "/api/vessels/").concat(v.id, "/mains"), signal)];
                    case 3:
                        mains = _g.sent();
                        _a = 0, mains_1 = mains;
                        _g.label = 4;
                    case 4:
                        if (!(_a < mains_1.length)) return [3 /*break*/, 15];
                        m = mains_1[_a];
                        if (signal.aborted)
                            return [3 /*break*/, 15];
                        group = m.name;
                        _g.label = 5;
                    case 5:
                        _g.trys.push([5, 13, , 14]);
                        return [4 /*yield*/, this._fetchJson("".concat(base, "/api/folders/").concat(m.id, "/children"), signal)];
                    case 6:
                        cats = _g.sent();
                        _b = 0, cats_2 = cats;
                        _g.label = 7;
                    case 7:
                        if (!(_b < cats_2.length)) return [3 /*break*/, 12];
                        cat = cats_2[_b];
                        if ((cat === null || cat === void 0 ? void 0 : cat.kind) === 'file')
                            return [3 /*break*/, 11];
                        category = cat.name;
                        _g.label = 8;
                    case 8:
                        _g.trys.push([8, 10, , 11]);
                        return [4 /*yield*/, this._fetchJson("".concat(base, "/api/folders/").concat(cat.id, "/children"), signal)];
                    case 9:
                        subCats = _g.sent();
                        leaves = subCats.filter(function (s) { return (s === null || s === void 0 ? void 0 : s.kind) !== 'file'; });
                        if (leaves.length > 0) {
                            for (_c = 0, leaves_1 = leaves; _c < leaves_1.length; _c++) {
                                leaf = leaves_1[_c];
                                if ((leaf === null || leaf === void 0 ? void 0 : leaf.kind) === 'file')
                                    continue;
                                sr++;
                                subCategory = leaf.name;
                                sp = "".concat(v.name, " > ").concat(group, " > ").concat(category, " > ").concat(subCategory);
                                out.push({ srNo: String(sr), vesselName: v.name, group: group, category: category, subCategory: subCategory, subFolderPath: sp, fileName: null, fileId: null, canUpload: true, groupKey: "".concat(v.id, "||").concat(group, "||").concat(category, "||").concat(subCategory, "||").concat(sp), uploadFolderId: leaf.id, monthDriven: Boolean(leaf.month_driven) });
                            }
                        }
                        else {
                            // No sub-folders → cat itself is a leaf
                            sr++;
                            sp = "".concat(v.name, " > ").concat(group, " > ").concat(category);
                            out.push({ srNo: String(sr), vesselName: v.name, group: group, category: category, subCategory: category, subFolderPath: sp, fileName: null, fileId: null, canUpload: true, groupKey: "".concat(v.id, "||").concat(group, "||").concat(category, "||").concat(category, "||").concat(sp), uploadFolderId: cat.id, monthDriven: Boolean(cat.month_driven) });
                        }
                        return [3 /*break*/, 11];
                    case 10:
                        _d = _g.sent();
                        sr++;
                        sp = "".concat(v.name, " > ").concat(group, " > ").concat(category);
                        out.push({ srNo: String(sr), vesselName: v.name, group: group, category: category, subCategory: category, subFolderPath: sp, fileName: null, fileId: null, canUpload: true, groupKey: "".concat(v.id, "||").concat(group, "||").concat(category, "||").concat(category, "||").concat(sp), uploadFolderId: cat.id, monthDriven: Boolean(cat.month_driven) });
                        return [3 /*break*/, 11];
                    case 11:
                        _b++;
                        return [3 /*break*/, 7];
                    case 12: return [3 /*break*/, 14];
                    case 13:
                        _e = _g.sent();
                        return [3 /*break*/, 14];
                    case 14:
                        _a++;
                        return [3 /*break*/, 4];
                    case 15: return [3 /*break*/, 17];
                    case 16:
                        _f = _g.sent();
                        return [3 /*break*/, 17];
                    case 17:
                        _i++;
                        return [3 /*break*/, 1];
                    case 18: return [2 /*return*/, out];
                }
            });
        });
    };
    // SharePoint Folder Provisioning ─────────────────────────────────────────
    VesselEmail.prototype._provisionVesselFolders = function (vesselName, vesselId) {
        return __awaiter(this, void 0, void 0, function () {
            var _a, graphClient, siteId, driveId, result, provisionedSet, err_6;
            var _b, _c;
            return __generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        _a = this.props, graphClient = _a.graphClient, siteId = _a.siteId, driveId = _a.driveId;
                        console.log("[VesselDMS] _provisionVesselFolders vessel=\"".concat(vesselName, "\" graphClient=").concat(!!graphClient, " siteId=\"").concat(siteId, "\" driveId=\"").concat(driveId, "\""));
                        if (!graphClient || !siteId || !driveId) {
                            console.error('[VesselDMS] _provisionVesselFolders aborted — missing graphClient/siteId/driveId. Props:', { graphClient: !!graphClient, siteId: siteId, driveId: driveId });
                            this.setState({ folderCreationError: 'SharePoint context not ready. Please refresh the page and try again.' });
                            return [2 /*return*/];
                        }
                        this.setState({ folderCreationBusy: true, folderCreationResults: null, folderCreationError: null, folderProvisioningVesselId: vesselId || null });
                        _d.label = 1;
                    case 1:
                        _d.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, (0, graphFolderService_1.createVesselFolders)(graphClient, siteId, driveId, vesselName)];
                    case 2:
                        result = _d.sent();
                        provisionedSet = new Set(this.state.provisionedVesselIds);
                        if (vesselId) {
                            provisionedSet.add(vesselId);
                            // Persist provisioned status in database
                            void fetch("".concat(this._base(), "/api/vessels/").concat(vesselId, "/provision"), {
                                method: 'POST',
                                headers: this._headers(),
                            }).catch(function () { return undefined; });
                        }
                        this.setState({
                            folderCreationBusy: false,
                            folderProvisioningVesselId: null,
                            folderCreationResults: result.results,
                            folderCreationError: result.success ? null : 'Some folders could not be created. Check the creation log.',
                            provisionedVesselIds: provisionedSet,
                        });
                        // Trigger an immediate delta sync so the new folders appear in the tree
                        void ((_b = this._syncScheduler) === null || _b === void 0 ? void 0 : _b.triggerNow().catch(function () { return undefined; }));
                        return [3 /*break*/, 4];
                    case 3:
                        err_6 = _d.sent();
                        this.setState({
                            folderCreationBusy: false,
                            folderProvisioningVesselId: null,
                            folderCreationError: (_c = err_6 === null || err_6 === void 0 ? void 0 : err_6.message) !== null && _c !== void 0 ? _c : 'Folder provisioning failed.',
                        });
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    // ── Load Files for Vessel ─────────────────────────────────────────────────
    /** Find the Documents list row represented by a live SharePoint folder ID. */
    VesselEmail.prototype._getGroupKeyForLiveFolderId = function (folderId) {
        var _this = this;
        var node = this.state.spoFolderMap.get(folderId);
        if (!(node === null || node === void 0 ? void 0 : node.isFolder) || !node.serverRelativePath)
            return '';
        var normalisePath = function (path) {
            return (path || '').replace(/^\/+/, '').replace(/\/+$/, '').replace(/\/vessels\//i, '/').toLocaleLowerCase();
        };
        var livePath = normalisePath(node.serverRelativePath);
        var row = this.state.rows.find(function (candidate) {
            var expectedPath = normalisePath(_this._sharePointFolderPath(candidate.subFolderPath, ''));
            return Boolean(expectedPath) &&
                (livePath === expectedPath || livePath.endsWith("/".concat(expectedPath)));
        });
        return (row === null || row === void 0 ? void 0 : row.groupKey) || '';
    };
    VesselEmail.prototype._loadFilesForVessel = function (vesselName) {
        return __awaiter(this, void 0, void 0, function () {
            var rows, vesselRows, normalisePath, liveFolderIdsByPath, liveFolders, BATCH, i, batch, mainFolder, folderPath, normalisedVesselPath, hasLiveVesselFolder, signal, children, legacyPath, expectedCats, expected, newAnomalies_1, _i, children_1, item, err_7;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        rows = this.state.rows;
                        vesselRows = rows.filter(function (r) { return r.vesselName === vesselName && r.uploadFolderId; });
                        normalisePath = function (path) {
                            return (path || '').replace(/^\/+/, '').replace(/\/+$/, '').replace(/\/vessels\//i, '/').toLocaleLowerCase();
                        };
                        liveFolderIdsByPath = new Map();
                        this.state.spoFolderMap.forEach(function (node) {
                            if (!node.deleted && node.isFolder && node.serverRelativePath) {
                                liveFolderIdsByPath.set(normalisePath(node.serverRelativePath), node.id);
                            }
                        });
                        liveFolders = vesselRows.map(function (row) {
                            var expectedPath = normalisePath(_this._sharePointFolderPath(row.subFolderPath, ''));
                            var id = liveFolderIdsByPath.get(expectedPath);
                            if (!id && expectedPath) {
                                // Some libraries keep the DMS tree under a top-level container such
                                // as "Vessel Management". The backend breadcrumb omits that container,
                                // so match the live folder path by its unambiguous suffix.
                                for (var _i = 0, _a = Array.from(liveFolderIdsByPath.entries()); _i < _a.length; _i++) {
                                    var _b = _a[_i], livePath = _b[0], liveId = _b[1];
                                    if (livePath.endsWith("/".concat(expectedPath))) {
                                        id = liveId;
                                        break;
                                    }
                                }
                            }
                            return { groupKey: row.groupKey, id: id };
                        }).filter(function (item) { return Boolean(item.id); });
                        if (!(liveFolders.length > 0)) return [3 /*break*/, 4];
                        BATCH = 5;
                        i = 0;
                        _a.label = 1;
                    case 1:
                        if (!(i < liveFolders.length)) return [3 /*break*/, 4];
                        batch = liveFolders.slice(i, i + BATCH);
                        return [4 /*yield*/, Promise.all(batch.map(function (folder) { return _this._refreshFolderFiles(folder.id, folder.groupKey, true); }))];
                    case 2:
                        _a.sent();
                        _a.label = 3;
                    case 3:
                        i += BATCH;
                        return [3 /*break*/, 1];
                    case 4:
                        mainFolder = this.state.docMainFolder || 'Technical & Crewing';
                        folderPath = "".concat(vesselName, "/").concat(mainFolder);
                        normalisedVesselPath = normalisePath(folderPath);
                        hasLiveVesselFolder = liveFolderIdsByPath.has(normalisedVesselPath) ||
                            Array.from(liveFolderIdsByPath.keys()).some(function (path) { return path.endsWith("/".concat(normalisedVesselPath)); });
                        if (!hasLiveVesselFolder)
                            return [2 /*return*/];
                        signal = new AbortController().signal;
                        _a.label = 5;
                    case 5:
                        _a.trys.push([5, 9, , 10]);
                        return [4 /*yield*/, this._getGraphChildren(folderPath, signal)];
                    case 6:
                        children = _a.sent();
                        if (!(!children || children.length === 0)) return [3 /*break*/, 8];
                        legacyPath = "".concat(mainFolder, "/").concat(vesselName);
                        return [4 /*yield*/, this._getGraphChildren(legacyPath, signal)];
                    case 7:
                        children = _a.sent();
                        if (children.length > 0)
                            folderPath = legacyPath;
                        _a.label = 8;
                    case 8:
                        if (!children || children.length === 0)
                            return [2 /*return*/];
                        expectedCats = {
                            'Technical & Crewing': ['Month End Reports', 'Service Agreements', 'Registration', 'Drawings and Manuals', 'PO & Invoice', 'Incidents', 'Crewing', 'To be Classified'],
                            'Commercial & Chartering': ['Agreements', 'Invoices & Payments', 'Claims & Disputes', 'To be Classified'],
                            'Insurance': ['P&I', 'H&M', 'War Risk', 'Flag and MPA'],
                            'Kaizen - Knowledge Bank': ['Templates', 'Procedures and Work Instructions', 'Lessons Learned', 'Circulars and Guidance'],
                            'Knowledge Bank': ['Templates', 'Procedures and Work Instructions', 'Lessons Learned', 'Circulars and Guidance'],
                        };
                        expected = expectedCats[mainFolder] || [];
                        newAnomalies_1 = [];
                        for (_i = 0, children_1 = children; _i < children_1.length; _i++) {
                            item = children_1[_i];
                            if (!expected.includes(item.name) && item.name !== 'To be Classified') {
                                newAnomalies_1.push({
                                    id: Date.now() + Math.floor(Math.random() * 10000),
                                    drive_item_id: item.id,
                                    name: item.name,
                                    item_type: item.isFolder ? 'folder' : 'file',
                                    anomaly_type: 'subfolder_unmatched',
                                    department: mainFolder,
                                    vessel_name: vesselName,
                                    spo_path: "".concat(folderPath, "/").concat(item.name),
                                    resolved: false,
                                    detected_at: new Date().toISOString(),
                                });
                            }
                        }
                        if (newAnomalies_1.length > 0) {
                            this.setState(function (prev) {
                                var existingPaths = new Set(prev.folderAnomalies.map(function (a) { return a.spo_path; }));
                                var filtered = newAnomalies_1.filter(function (a) { return !existingPaths.has(a.spo_path); });
                                if (filtered.length === 0)
                                    return null;
                                return { folderAnomalies: __spreadArray(__spreadArray([], prev.folderAnomalies, true), filtered, true) };
                            });
                        }
                        return [3 /*break*/, 10];
                    case 9:
                        err_7 = _a.sent();
                        console.warn('[VesselDMS] _loadFilesForVessel warning:', err_7);
                        return [3 /*break*/, 10];
                    case 10: return [2 /*return*/];
                }
            });
        });
    };
    VesselEmail.prototype._refreshFolderFiles = function (uploadFolderId_1, groupKey_1) {
        return __awaiter(this, arguments, void 0, function (uploadFolderId, groupKey, force) {
            var resolvedId, resolved, _a, _b, graphClient, siteId, driveId, fileItems, fetchedFromGraph, matchingRow, sharePointFolderPath, isLiveDriveFolderId, encodedPath, pathUrl, result, items, pathLookupError_1, url, result, items, idLookupError_1, children, err_8, parsedUploads;
            var _c, _d;
            if (force === void 0) { force = false; }
            return __generator(this, function (_e) {
                switch (_e.label) {
                    case 0:
                        if (!uploadFolderId)
                            return [2 /*return*/];
                        if (/^f\d+$/.test(uploadFolderId))
                            return [2 /*return*/];
                        if (!force && this._filesLoadedForFolders.has(uploadFolderId))
                            return [2 /*return*/];
                        resolvedId = uploadFolderId;
                        if (!uploadFolderId.includes('/')) return [3 /*break*/, 4];
                        _e.label = 1;
                    case 1:
                        _e.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._fetchJson("".concat(this._base(), "/api/folders/upload-by-path?path=").concat(encodeURIComponent(uploadFolderId), "&resolve_only=true")).catch(function () { return null; })];
                    case 2:
                        resolved = _e.sent();
                        if (resolved === null || resolved === void 0 ? void 0 : resolved.folder_id) {
                            resolvedId = resolved.folder_id;
                        }
                        else {
                            return [2 /*return*/];
                        }
                        return [3 /*break*/, 4];
                    case 3:
                        _a = _e.sent();
                        return [2 /*return*/];
                    case 4:
                        _b = this.props, graphClient = _b.graphClient, siteId = _b.siteId, driveId = _b.driveId;
                        fileItems = [];
                        fetchedFromGraph = false;
                        matchingRow = this.state.rows.find(function (r) {
                            return r.uploadFolderId === uploadFolderId ||
                                r.uploadFolderId === resolvedId ||
                                (groupKey && r.groupKey === groupKey);
                        });
                        sharePointFolderPath = matchingRow
                            ? this._sharePointFolderPath(matchingRow.subFolderPath, '')
                            : '';
                        isLiveDriveFolderId = this.state.spoFolderMap.has(uploadFolderId);
                        if (!(graphClient && siteId && driveId)) return [3 /*break*/, 12];
                        if (!(sharePointFolderPath && !isLiveDriveFolderId)) return [3 /*break*/, 8];
                        _e.label = 5;
                    case 5:
                        _e.trys.push([5, 7, , 8]);
                        // The stored breadcrumb starts at the vessel name; the
                        // live Documents hierarchy also requires both parent
                        // containers before that segment.
                        encodedPath = ("Vessels/Specific Vessels/" + sharePointFolderPath).split('/').map(function (s) { return encodeURIComponent(s); }).join('/');
                        pathUrl = "/sites/".concat(siteId, "/drives/").concat(driveId, "/root:/").concat(encodedPath, ":/children?$select=id,name,size,lastModifiedDateTime,file,folder&$top=200");
                        return [4 /*yield*/, graphClient.api(pathUrl).get()];
                    case 6:
                        result = _e.sent();
                        items = (_c = result === null || result === void 0 ? void 0 : result.value) !== null && _c !== void 0 ? _c : [];
                        fileItems = items.filter(function (i) { return !!i.file; });
                        fetchedFromGraph = true;
                        return [3 /*break*/, 8];
                    case 7:
                        pathLookupError_1 = _e.sent();
                        console.warn("[VesselDMS] failed to load SharePoint folder \"".concat(sharePointFolderPath, "\""), pathLookupError_1);
                        return [3 /*break*/, 8];
                    case 8:
                        if (!((isLiveDriveFolderId || !sharePointFolderPath) && !fetchedFromGraph && resolvedId && !resolvedId.includes('/'))) return [3 /*break*/, 12];
                        _e.label = 9;
                    case 9:
                        _e.trys.push([9, 11, , 12]);
                        url = "/sites/".concat(siteId, "/drives/").concat(driveId, "/items/").concat(resolvedId, "/children?$select=id,name,size,lastModifiedDateTime,file,folder&$top=200");
                        return [4 /*yield*/, graphClient.api(url).get()];
                    case 10:
                        result = _e.sent();
                        items = (_d = result === null || result === void 0 ? void 0 : result.value) !== null && _d !== void 0 ? _d : [];
                        fileItems = items.filter(function (i) { return !!i.file; }); // only files, not subfolders
                        fetchedFromGraph = true;
                        return [3 /*break*/, 12];
                    case 11:
                        idLookupError_1 = _e.sent();
                        console.warn("[VesselDMS] failed to load backend folder ID \"".concat(resolvedId, "\""), idLookupError_1);
                        return [3 /*break*/, 12];
                    case 12:
                        if (!!fetchedFromGraph) return [3 /*break*/, 16];
                        // Skip REST fallback for numeric-only IDs — these are DB row IDs, not SharePoint drive item IDs.
                        if (/^\d+$/.test(resolvedId)) {
                            console.warn("[VesselDMS] _refreshFolderFiles: skipping REST fallback for numeric DB ID \"".concat(resolvedId, "\" \u2014 not a valid drive item ID"));
                            return [2 /*return*/];
                        }
                        _e.label = 13;
                    case 13:
                        _e.trys.push([13, 15, , 16]);
                        return [4 /*yield*/, this._fetchJson("".concat(this._base(), "/api/folders/").concat(encodeURIComponent(resolvedId), "/children"))];
                    case 14:
                        children = _e.sent();
                        fileItems = (children || []).filter(function (c) { return c.kind === 'file'; });
                        return [3 /*break*/, 16];
                    case 15:
                        err_8 = _e.sent();
                        console.warn("[VesselDMS] failed to load children for folder \"".concat(resolvedId, "\""), err_8);
                        return [2 /*return*/];
                    case 16:
                        // Keep empty responses retryable during SPO consistency windows.
                        if (fileItems.length > 0)
                            this._filesLoadedForFolders.add(uploadFolderId);
                        parsedUploads = fileItems.map(function (f) { return ({
                            name: f.name || f.displayName,
                            size: f.size ? "".concat((f.size / 1024).toFixed(1), " KB") : '—',
                            date: f.lastModifiedDateTime ? new Date(f.lastModifiedDateTime).toLocaleDateString() : (f.modified || 'Today'),
                            pending: false,
                            id: f.id,
                            uploadedAt: f.lastModifiedDateTime ? Date.parse(f.lastModifiedDateTime) : undefined,
                        }); });
                        this.setState(function (prev) {
                            var _a;
                            // Merge fetched parsedUploads with any existing local uploads so we don't wipe out freshly uploaded files
                            var existingFolderUploads = __spreadArray(__spreadArray([], (prev.uploadedFilesByFolder[uploadFolderId] || []), true), (prev.uploadedFilesByFolder[resolvedId] || []), true);
                            var cachedByName = new Map();
                            existingFolderUploads.forEach(function (file) {
                                if ((file === null || file === void 0 ? void 0 : file.name) && !cachedByName.has(file.name.toLowerCase()))
                                    cachedByName.set(file.name.toLowerCase(), file);
                            });
                            var mergedUploads = parsedUploads.map(function (file) { return (__assign(__assign({}, file), { uploadedAt: ((_a = cachedByName.get(file.name.toLowerCase())) === null || _a === void 0 ? void 0 : _a.uploadedAt) })); });
                            var _loop_2 = function (ex) {
                                if (!mergedUploads.some(function (u) { return u.name === ex.name; })) {
                                    mergedUploads.push({
                                        name: ex.name,
                                        size: ex.size || '—',
                                        date: ex.date || 'Today',
                                        pending: Boolean(ex.pending),
                                        id: ex.id || ex.name,
                                        uploadedAt: ex.uploadedAt,
                                    });
                                }
                            };
                            for (var _i = 0, existingFolderUploads_1 = existingFolderUploads; _i < existingFolderUploads_1.length; _i++) {
                                var ex = existingFolderUploads_1[_i];
                                _loop_2(ex);
                            }
                            var updatedByFolder = __assign(__assign({}, prev.uploadedFilesByFolder), (_a = {}, _a[uploadFolderId] = mergedUploads, _a[resolvedId] = mergedUploads, _a));
                            var existingBaseRows = prev.rows.filter(function (r) {
                                return r.uploadFolderId === uploadFolderId || r.uploadFolderId === resolvedId ||
                                    Boolean(groupKey) && r.groupKey === groupKey;
                            });
                            var newRows = prev.rows;
                            if (existingBaseRows.length > 0) {
                                var baseRow_1 = existingBaseRows[0];
                                var existingGroupUploads = prev.uploadedFilesByFolder[baseRow_1.groupKey] || [];
                                var mergedGroupUploads = __spreadArray([], mergedUploads, true);
                                var _loop_3 = function (ex) {
                                    if (!mergedGroupUploads.some(function (u) { return u.name === ex.name; })) {
                                        mergedGroupUploads.push({
                                            name: ex.name,
                                            size: ex.size || '—',
                                            date: ex.date || 'Today',
                                            pending: Boolean(ex.pending),
                                            id: ex.id || ex.name,
                                        });
                                    }
                                };
                                for (var _b = 0, existingGroupUploads_1 = existingGroupUploads; _b < existingGroupUploads_1.length; _b++) {
                                    var ex = existingGroupUploads_1[_b];
                                    _loop_3(ex);
                                }
                                updatedByFolder[baseRow_1.groupKey] = mergedGroupUploads;
                                var otherFolderRows = prev.rows.filter(function (r) {
                                    return r.uploadFolderId !== uploadFolderId && r.uploadFolderId !== resolvedId &&
                                        (!groupKey || r.groupKey !== groupKey);
                                });
                                // If fileItems is empty but we have local/existing uploads, keep them!
                                var activeFiles = mergedGroupUploads.length > 0 ? mergedGroupUploads : parsedUploads;
                                if (activeFiles.length === 0) {
                                    // Graph may not be consistent yet - never overwrite rows when we have nothing to show.
                                    return { rows: prev.rows, uploadedFilesByFolder: updatedByFolder };
                                }
                                else {
                                    var mappedRows = activeFiles.map(function (f, i) { return (__assign(__assign({}, baseRow_1), { srNo: i === 0 ? baseRow_1.srNo : "".concat(baseRow_1.srNo, ".").concat(i + 1), fileName: f.name || f.displayName || null, fileId: f.id || null, filePending: Boolean(f.pending) })); });
                                    newRows = __spreadArray(__spreadArray([], otherFolderRows, true), mappedRows, true);
                                }
                            }
                            return { rows: newRows, uploadedFilesByFolder: updatedByFolder };
                        });
                        return [2 /*return*/];
                }
            });
        });
    };
    VesselEmail.prototype._renderUploadSuccessPopup = function () {
        var _this = this;
        var p = this.state.uploadSuccessPopup;
        if (!p)
            return null;
        return (React.createElement("div", { style: {
                position: 'fixed', bottom: 28, right: 28, zIndex: 99999,
                background: '#fff', borderRadius: 16, boxShadow: '0 8px 40px rgba(0,0,0,0.18)',
                border: '1.5px solid #86efac', padding: '20px 24px 18px', minWidth: 340, maxWidth: 420,
                fontFamily: "'Segoe UI', sans-serif", animation: 'slideInRight 0.3s ease',
            } },
            React.createElement("style", null, "\n          @keyframes slideInRight { from { opacity:0; transform:translateX(40px); } to { opacity:1; transform:translateX(0); } }\n          @keyframes countdownShrink { from { width:100%; } to { width:0%; } }\n        "),
            React.createElement("div", { style: { display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', gap: 12, marginBottom: 10 } },
                React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 10 } },
                    React.createElement("div", { style: {
                            width: 36, height: 36, borderRadius: 10,
                            background: p.isPending ? '#fef3c7' : '#dcfce7',
                            display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 18, flexShrink: 0,
                        } }, p.isPending ? '⏳' : '✅'),
                    React.createElement("div", null,
                        React.createElement("div", { style: { fontWeight: 700, fontSize: 14, color: '#0f172a' } }, p.isPending ? 'Submitted for Approval' : 'Upload Successful!'),
                        React.createElement("div", { style: { fontSize: 11, color: '#64748b', marginTop: 2 } },
                            "Auto-closes in ",
                            p.secondsLeft,
                            "s"))),
                React.createElement("button", { onClick: function () { clearInterval(_this._uploadSuccessTimer); _this.setState({ uploadSuccessPopup: null }); }, style: { border: 'none', background: 'none', cursor: 'pointer', color: '#94a3b8', fontSize: 18, padding: '0 2px', lineHeight: 1, flexShrink: 0 }, title: "Close" }, "\u2715")),
            React.createElement("div", { style: { background: '#f8fafc', borderRadius: 8, padding: '8px 12px', marginBottom: 12 } },
                React.createElement("div", { style: { fontSize: 12, color: '#64748b', marginBottom: 3 } }, "\uD83D\uDCC4 File"),
                React.createElement("div", { style: { fontSize: 13, fontWeight: 600, color: '#0f172a', wordBreak: 'break-all' } }, p.fileName),
                React.createElement("div", { style: { fontSize: 11, color: '#64748b', marginTop: 6 } }, "\uD83D\uDCC1 Path"),
                React.createElement("div", { style: { fontSize: 12, color: '#334155', marginTop: 2, wordBreak: 'break-all' } }, p.destinationPath)),
            React.createElement("div", { style: { display: 'flex', gap: 8 } },
                p.webUrl && (React.createElement("a", { href: p.webUrl, target: "_blank", rel: "noopener noreferrer", style: {
                        flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6,
                        background: '#0078d4', color: '#fff', borderRadius: 8,
                        padding: '8px 12px', fontSize: 12, fontWeight: 600, textDecoration: 'none',
                        cursor: 'pointer',
                    } }, "\uD83D\uDD17 Open in SharePoint")),
                React.createElement("button", { onClick: function () { clearInterval(_this._uploadSuccessTimer); _this.setState({ uploadSuccessPopup: null }); }, style: {
                        flex: 1, background: '#f1f5f9', color: '#334155', border: 'none', borderRadius: 8,
                        padding: '8px 12px', fontSize: 12, fontWeight: 600, cursor: 'pointer',
                    } }, "Close")),
            React.createElement("div", { style: { marginTop: 12, height: 3, background: '#f1f5f9', borderRadius: 2, overflow: 'hidden' } },
                React.createElement("div", { style: {
                        height: '100%', borderRadius: 2,
                        background: p.isPending ? '#f59e0b' : '#16a34a',
                        width: "".concat((p.secondsLeft / 10) * 100, "%"),
                        transition: 'width 1s linear',
                    } }))));
    };
    VesselEmail.prototype._getGrouped = function () {
        var _a = this.state, rows = _a.rows, textFilter = _a.textFilter, vesselFilter = _a.vesselFilter, groupFilter = _a.groupFilter, catFilter = _a.catFilter, sort = _a.sort;
        var q = textFilter.toLowerCase();
        var f = rows;
        if (vesselFilter !== 'all')
            f = f.filter(function (r) { return r.vesselName === vesselFilter; });
        if (groupFilter !== 'all')
            f = f.filter(function (r) { return r.group === groupFilter; });
        if (catFilter !== 'all')
            f = f.filter(function (r) { return r.category === catFilter; });
        if (q)
            f = f.filter(function (r) { var _a; return [r.vesselName, r.group, r.category, r.subFolderPath, (_a = r.fileName) !== null && _a !== void 0 ? _a : ''].some(function (s) { return s.toLowerCase().indexOf(q) !== -1; }); });
        var map = new Map();
        for (var _i = 0, f_1 = f; _i < f_1.length; _i++) {
            var row = f_1[_i];
            var ex = map.get(row.groupKey);
            if (!ex)
                map.set(row.groupKey, { srNo: row.srNo, vesselName: row.vesselName, group: row.group, category: row.category, subCategory: row.subCategory || row.category, subFolderPath: row.subFolderPath, groupKey: row.groupKey, uploadFolderId: row.uploadFolderId, monthDriven: row.monthDriven, canUpload: row.canUpload, files: row.fileId && row.fileName ? [{ id: row.fileId, name: row.fileName }] : [] });
            else if (row.fileId && row.fileName)
                ex.files.push({ id: row.fileId, name: row.fileName });
        }
        var grouped = Array.from(map.values());
        if (sort === 'name_asc')
            grouped.sort(function (a, b) { return a.vesselName.localeCompare(b.vesselName); });
        else if (sort === 'name_desc')
            grouped.sort(function (a, b) { return b.vesselName.localeCompare(a.vesselName); });
        return grouped;
    };
    // ── Render Sidebar Navigation ─────────────────────────────────────────────
    VesselEmail.prototype._renderSidebar = function () {
        return (0, Sidebar_1.renderSidebar)(this);
    };
    VesselEmail.prototype._renderLayout = function (content) {
        return (0, AppLayout_1.renderLayout)(this, content);
    };
    VesselEmail.prototype._renderDocPreviewDrawer = function () {
        return (0, DocPreviewDrawer_1.renderDocPreviewDrawer)(this);
    };
    VesselEmail.prototype._renderDashboard = function () {
        return (0, DashboardPage_1.renderDashboard)(this);
    };
    VesselEmail.prototype._renderDocumentsPage = function () {
        return (0, DocumentsPage_1.renderDocumentsPage)(this);
    };
    VesselEmail.prototype._renderVesselsPage = function () {
        return (0, VesselsPage_1.renderVesselsPage)(this);
    };
    VesselEmail.prototype._renderTemplatesPage = function () {
        return (0, TemplatesPage_1.renderTemplatesPage)(this);
    };
    VesselEmail.prototype._renderApprovalsPage = function () {
        return (0, ApprovalsPage_1.renderApprovalsPage)(this);
    };
    VesselEmail.prototype._renderNotificationsPage = function () {
        return (0, NotificationsPage_1.renderNotificationsPage)(this);
    };
    VesselEmail.prototype._renderReportsPage = function () {
        return (0, ReportsPage_1.renderReportsPage)(this);
    };
    VesselEmail.prototype._renderUsersPage = function () {
        return (0, UsersPage_1.renderUsersPage)(this);
    };
    VesselEmail.prototype._renderSettingsPage = function () {
        return (0, SettingsPage_1.renderSettingsPage)(this);
    };
    VesselEmail.prototype._renderBentoEmailDashboardPage = function () {
        return (0, BentoEmailDashboardPage_1.renderBentoEmailDashboardPage)(this);
    };
    // ── Bento Compose Modal ─────────────────────────────────────────────────
    // Build auto subject from vessel + tag ([DataSource:TAG] Vessel Name format)
    VesselEmail.prototype._buildAutoSubject = function (vesselName, tag, customSubjectText) {
        if (customSubjectText === void 0) { customSubjectText = ''; }
        var tagKey = tag || 'mail';
        var vName = (vesselName || '').trim();
        var text = (customSubjectText || '').trim();
        if (vName && text && text !== 'Subject') {
            return "[DataSource:".concat(tagKey, "] ").concat(vName, " / ").concat(text);
        }
        else if (vName) {
            return "[DataSource:".concat(tagKey, "] ").concat(vName);
        }
        return text && text !== 'Subject' ? "[DataSource:".concat(tagKey, "] ").concat(text) : "[DataSource:".concat(tagKey, "]");
    };
    // Collect approved file names from rows & documents list for the selected vessel
    VesselEmail.prototype._getApprovedFilesForVesselTag = function (vesselName, tag) {
        var _a = this.state, rows = _a.rows, uploadedFilesByFolder = _a.uploadedFilesByFolder, documentsList = _a.documentsList, bentoApprovedFiles = _a.bentoApprovedFiles;
        var names = [];
        if (!vesselName || !vesselName.trim())
            return names;
        var vLower = vesselName.trim().toLowerCase();
        // 1. From directly fetched approved files (most reliable)
        if (bentoApprovedFiles && bentoApprovedFiles[vLower]) {
            bentoApprovedFiles[vLower].forEach(function (f) {
                if (names.indexOf(f) === -1)
                    names.push(f);
            });
        }
        // 2. From flat rows
        rows.forEach(function (r) {
            if (r.vesselName.trim().toLowerCase() !== vLower)
                return;
            if (r.fileName && !r.filePending && names.indexOf(r.fileName) === -1) {
                names.push(r.fileName);
            }
            var uploads = uploadedFilesByFolder[r.groupKey] || [];
            (uploads || []).filter(function (f) { return !f.pending; }).forEach(function (f) {
                if (names.indexOf(f.name) === -1)
                    names.push(f.name);
            });
        });
        // 3. From documents list
        (documentsList || []).forEach(function (doc) {
            if (doc.vessel && doc.vessel.trim().toLowerCase() === vLower && !doc.isFolder) {
                if (names.indexOf(doc.name) === -1)
                    names.push(doc.name);
            }
        });
        return names;
    };
    // Fetch all approved files for a vessel directly from the backend
    VesselEmail.prototype._fetchApprovedFilesForVessel = function (vesselName) {
        return __awaiter(this, void 0, void 0, function () {
            var vLower, base_1, names_1, approvals, fileIdUpdates_1, _a, rows, uploadedFilesByFolder_1, _b;
            var _this = this;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        if (!vesselName)
                            return [2 /*return*/];
                        vLower = vesselName.trim().toLowerCase();
                        _c.label = 1;
                    case 1:
                        _c.trys.push([1, 5, , 6]);
                        base_1 = this._base();
                        names_1 = [];
                        if (!!this._filesLoadedForVessels.has(vesselName)) return [3 /*break*/, 3];
                        this._filesLoadedForVessels.add(vesselName);
                        return [4 /*yield*/, this._mergeLiveSharePointFiles([vesselName]).catch(function () { return undefined; })];
                    case 2:
                        _c.sent();
                        _c.label = 3;
                    case 3: return [4 /*yield*/, this._fetchJson("".concat(base_1, "/api/my-approvals?status=approved")).catch(function () {
                            return _this._fetchJson("".concat(base_1, "/api/approvals?status=approved")).catch(function () { return []; });
                        })];
                    case 4:
                        approvals = _c.sent();
                        fileIdUpdates_1 = {};
                        if (Array.isArray(approvals)) {
                            approvals.forEach(function (a) {
                                var aVessel = (a.vessel_name || a.vesselName || '').trim().toLowerCase();
                                var fname = a.file_name || a.fileName || a.filename || a.name;
                                if (fname && (aVessel === vLower || !aVessel) && names_1.indexOf(fname) === -1) {
                                    names_1.push(fname);
                                }
                                var realFileId = a.drive_item_id || a.driveItemId || a.file_id || a.fileId;
                                if (fname && realFileId) {
                                    fileIdUpdates_1["".concat(vLower, "||").concat(fname)] = realFileId;
                                }
                            });
                        }
                        _a = this.state, rows = _a.rows, uploadedFilesByFolder_1 = _a.uploadedFilesByFolder;
                        rows.forEach(function (r) {
                            if (r.vesselName.trim().toLowerCase() !== vLower)
                                return;
                            if (r.fileName && !r.filePending && names_1.indexOf(r.fileName) === -1) {
                                names_1.push(r.fileName);
                            }
                            (uploadedFilesByFolder_1[r.groupKey] || []).filter(function (f) { return !f.pending; }).forEach(function (f) {
                                if (names_1.indexOf(f.name) === -1)
                                    names_1.push(f.name);
                            });
                        });
                        // Always update state to populate dropdown
                        this.setState(function (prev) {
                            var _a;
                            return ({
                                bentoApprovedFiles: __assign(__assign({}, (prev.bentoApprovedFiles || {})), (_a = {}, _a[vLower] = names_1, _a)),
                                bentoApprovedFileIds: __assign(__assign({}, (prev.bentoApprovedFileIds || {})), fileIdUpdates_1),
                            });
                        });
                        return [3 /*break*/, 6];
                    case 5:
                        _b = _c.sent();
                        this.setState(function (prev) {
                            var _a;
                            return ({
                                bentoApprovedFiles: __assign(__assign({}, (prev.bentoApprovedFiles || {})), (_a = {}, _a[vLower] = [], _a)),
                            });
                        });
                        return [3 /*break*/, 6];
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    VesselEmail.prototype._getFileIdForAttachment = function (vesselName, fileName) {
        var _a = this.state, rows = _a.rows, bentoApprovedFileIds = _a.bentoApprovedFileIds;
        var vLower = vesselName.trim().toLowerCase();
        for (var _i = 0, rows_1 = rows; _i < rows_1.length; _i++) {
            var r = rows_1[_i];
            if (r.vesselName.trim().toLowerCase() === vLower && r.fileName === fileName && r.fileId) {
                return r.fileId;
            }
        }
        var approvedId = (bentoApprovedFileIds || {})["".concat(vLower, "||").concat(fileName)];
        return approvedId || null;
    };
    VesselEmail.prototype._renderBentoComposeModal = function () {
        return (0, BentoComposeModal_1.renderBentoComposeModal)(this);
    };
    VesselEmail.prototype._renderRecycleBinPage = function () {
        return (0, RecycleBinPage_1.renderRecycleBinPage)(this);
    };
    VesselEmail.prototype._renderArchivePage = function () {
        return (0, ArchivePage_1.renderArchivePage)(this);
    };
    VesselEmail.prototype._renderVesselForm = function (mode) {
        return (0, VesselFormModal_1.renderVesselForm)(this, mode);
    };
    VesselEmail.prototype._renderDeleteModal = function () {
        return (0, DeleteVesselModal_1.renderDeleteModal)(this);
    };
    VesselEmail.prototype.render = function () {
        var _a = this.state, sessionExpired = _a.sessionExpired, view = _a.view;
        if (sessionExpired) {
            return (React.createElement("div", { style: {
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                    minHeight: '100vh', background: '#f8fafc', fontFamily: "'Segoe UI', sans-serif",
                } },
                React.createElement("div", { style: {
                        background: '#fff', borderRadius: 12, padding: '40px 48px', maxWidth: 420,
                        textAlign: 'center', boxShadow: '0 4px 24px rgba(0,0,0,0.10)', border: '1px solid #e2e8f0',
                    } },
                    React.createElement("div", { style: { fontSize: 48, marginBottom: 16 } }, "\uD83D\uDD12"),
                    React.createElement("h2", { style: { margin: '0 0 8px', fontSize: 20, fontWeight: 700, color: '#0f172a' } }, "Session Expired"),
                    React.createElement("p", { style: { margin: '0 0 24px', fontSize: 14, color: '#64748b', lineHeight: 1.6 } }, "Your session has expired or is no longer valid. Please sign out and sign back in to continue."),
                    React.createElement("button", { onClick: this._handleSignOut, style: {
                            background: '#0078d4', color: '#fff', border: 'none', borderRadius: 8,
                            padding: '10px 28px', fontSize: 14, fontWeight: 600, cursor: 'pointer',
                        } }, "Sign Out & Reload"))));
        }
        var content;
        switch (view) {
            case 'dashboard':
                content = this._renderDashboard();
                break;
            case 'list':
                content = this._renderDocumentsPage();
                break;
            case 'vessels':
                content = this._renderVesselsPage();
                break;
            case 'templates':
                content = this._renderTemplatesPage();
                break;
            case 'approvals':
                content = this._renderApprovalsPage();
                break;
            case 'notifications':
                content = this._renderNotificationsPage();
                break;
            case 'reports':
                content = this._renderReportsPage();
                break;
            case 'users':
                content = this._renderUsersPage();
                break;
            case 'settings':
                content = this._renderSettingsPage();
                break;
            case 'bento_email':
            case 'email_notify':
                content = this._renderBentoEmailDashboardPage();
                break;
            case 'recycle':
                content = this._renderRecycleBinPage();
                break;
            case 'archive':
                content = this._renderArchivePage();
                break;
            default:
                content = this._renderDashboard();
        }
        return (React.createElement(React.Fragment, null,
            this._renderLayout(content),
            this._renderBentoComposeModal(),
            this._renderUploadSuccessPopup()));
    };
    VesselEmail._remoteServerDown = false;
    return VesselEmail;
}(React.Component));
exports.default = VesselEmail;
