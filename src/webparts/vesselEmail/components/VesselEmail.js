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
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
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
exports.__esModule = true;
var React = require("react");
var vesselImagePool_1 = require("./vesselImagePool");
var graphFolderService_1 = require("./graphFolderService");
var deltaSync_1 = require("./deltaSync");
// ── Helpers & Constants ─────────────────────────────────────────────────────
function cleanName(name) {
    return (name || '').replace(/^_+|_+$/g, '').trim();
}
var VESSEL_TYPES = [
    'Bulk Carrier', 'Container Ship', 'Gas Carrier',
    'Oil Tanker', 'Chemical Tanker', 'General Cargo', 'Offshore Support', 'Other Cargo Ships',
];
var DATASOURCE_TAGS_MAP = {
    contract: 'CP',
    other_contract: 'Other Contracts',
    vessels_certificate: 'Certificate',
    vessels_drawing: 'Drawing',
    mail: 'Mail',
    imo: 'IMO',
    uscg: 'USCG',
    msib: 'USCG MSIB',
    imo_flag_country_others: 'IMO Flag Country Others',
    panama_flag_circular: 'Panama Flag Circular',
    imo_flag_country_flag: 'Flag',
    nk: 'NK',
    japan_p_and_i: 'Japan P&I',
    ukpandi: 'UK P&I',
    gard: 'GARD',
    scmg: 'Standard Club',
    britannia_p_and_i: 'Britannia P&I',
    bimco: 'BIMCO',
    security_information: 'Security Information',
    omc_kaikoumu: 'OMC Marine & Tech. Support Center',
    ice_information: 'Ice Information',
    right_ship: 'RightShip',
    others: 'Others'
};
function suggestTagFromFilename(filename) {
    var s = (filename || '').toLowerCase();
    if (['cp', 'charter', 'contract', 'agreement', 'party'].some(function (k) { return s.indexOf(k) !== -1; }))
        return 'contract';
    if (['vendor', 'supplier', 'subcontract', 'other_contract'].some(function (k) { return s.indexOf(k) !== -1; }))
        return 'other_contract';
    if (['cert', 'certificate', 'class', 'survey', 'statutory', 'audit'].some(function (k) { return s.indexOf(k) !== -1; }))
        return 'vessels_certificate';
    if (['draw', 'plan', 'schematic', 'manual', 'diagram', 'blueprint', 'ga_plan'].some(function (k) { return s.indexOf(k) !== -1; }))
        return 'vessels_drawing';
    if (s.indexOf('msib') !== -1)
        return 'msib';
    if (['uscg', 'coastguard', 'coast guard'].some(function (k) { return s.indexOf(k) !== -1; }))
        return 'uscg';
    if (['panama', 'circular'].some(function (k) { return s.indexOf(k) !== -1; }))
        return 'panama_flag_circular';
    if (['nk', 'nippon'].some(function (k) { return s.indexOf(k) !== -1; }))
        return 'nk';
    if (['ukpandi', 'uk_p_and_i', 'uk p&i'].some(function (k) { return s.indexOf(k) !== -1; }))
        return 'ukpandi';
    if (s.indexOf('gard') !== -1)
        return 'gard';
    if (['scmg', 'standard_club', 'standard club'].some(function (k) { return s.indexOf(k) !== -1; }))
        return 'scmg';
    if (s.indexOf('britannia') !== -1)
        return 'britannia_p_and_i';
    if (['pandi', 'p&i', 'pi', 'japan_p_and_i', 'protection'].some(function (k) { return s.indexOf(k) !== -1; }))
        return 'japan_p_and_i';
    if (s.indexOf('bimco') !== -1)
        return 'bimco';
    if (['security', 'isps', 'sec_info'].some(function (k) { return s.indexOf(k) !== -1; }))
        return 'security_information';
    if (['kaikoumu', 'omc', 'tech_support'].some(function (k) { return s.indexOf(k) !== -1; }))
        return 'omc_kaikoumu';
    if (['ice', 'ice_info'].some(function (k) { return s.indexOf(k) !== -1; }))
        return 'ice_information';
    if (['rightship', 'right_ship', 'right ship'].some(function (k) { return s.indexOf(k) !== -1; }))
        return 'right_ship';
    if (['flag', 'flag_state'].some(function (k) { return s.indexOf(k) !== -1; }))
        return 'imo_flag_country_flag';
    if (s.indexOf('imo') !== -1)
        return 'imo';
    if (['mail', 'email', 'letter', 'memo', 'msg', 'eml'].some(function (k) { return s.indexOf(k) !== -1; }))
        return 'mail';
    return 'mail';
}
function badge(color, text) {
    var map = {
        blue: ['#e1efff', '#0078d4'], orange: ['#fff4ce', '#8a5700'],
        purple: ['#ede8f5', '#5c2d91'], green: ['#dff6dd', '#107c10'],
        red: ['#fde7e9', '#a4262c'], "default": ['#f3f2f1', '#323130']
    };
    var pair = map[color] || map['default'];
    return <span style={{ display: 'inline-block', borderRadius: 12, padding: '3px 10px', fontSize: 11, fontWeight: 600, background: pair[0], color: pair[1], whiteSpace: 'nowrap' }}>{text}</span>;
}
var GROUP_COLORS = {
    Agreements: 'blue', 'Claims & Disputes': 'orange', Crewing: 'blue',
    Electrical: 'orange', Hull: 'blue', Incidents: 'red', Insurance: 'orange',
    Machinery: 'purple', Safety: 'green', 'Technical & Crewing': 'orange',
    'Commercial & Chartering': 'green', 'To be Classified': 'default'
};
// ── Default Mock Data ────────────────────────────────────────────────────────
var INITIAL_MOCK_DOCUMENTS = [
    { id: 'd1', name: 'Certificates', vessel: '-', type: 'Folder', expiryDate: '-', status: 'Valid', modified: 'May 12, 2024', isFolder: true },
    { id: 'd2', name: 'Crew Documents', vessel: '-', type: 'Folder', expiryDate: '-', status: 'Valid', modified: 'May 11, 2024', isFolder: true },
    { id: 'd3', name: 'Insurance', vessel: '-', type: 'Folder', expiryDate: '-', status: 'Valid', modified: 'May 10, 2024', isFolder: true },
    { id: 'd4', name: 'Insurance Certificate.pdf', vessel: 'Ocean Star', type: 'Insurance', expiryDate: 'May 25, 2024', status: 'Expiring Soon', modified: 'May 12, 2024' },
    { id: 'd5', name: 'Crew List.docx', vessel: 'Ocean Star', type: 'Crew', expiryDate: 'Jun 10, 2024', status: 'Valid', modified: 'May 11, 2024' },
    { id: 'd6', name: 'Maintenance Log.xlsx', vessel: 'Sea Breeze', type: 'Maintenance', expiryDate: 'May 20, 2024', status: 'Valid', modified: 'May 10, 2024' },
    { id: 'd7', name: 'Safety Certificate.pdf', vessel: 'Blue Horizon', type: 'Certificate', expiryDate: 'Apr 15, 2024', status: 'Expired', modified: 'May 09, 2024' },
];
var INITIAL_MOCK_TEMPLATES = [
    { id: 't1', name: 'Insurance Certificate Template.docx', type: 'Insurance', description: 'Template for Insurance Certificate', modified: 'May 12, 2024' },
    { id: 't2', name: 'Crew List Template.docx', type: 'Crew', description: 'Template for Crew List', modified: 'May 11, 2024' },
    { id: 't3', name: 'Maintenance Report Template.docx', type: 'Maintenance', description: 'Template for Maintenance Report', modified: 'May 10, 2024' },
    { id: 't4', name: 'Safety Certificate Template.docx', type: 'Certificate', description: 'Template for Safety Certificate', modified: 'May 09, 2024' },
    { id: 't5', name: 'Survey Report Template.docx', type: 'Survey', description: 'Template for Survey Report', modified: 'May 08, 2024' },
];
var INITIAL_MOCK_APPROVALS = [
    { id: 'a1', documentName: 'Insurance Certificate.pdf', vessel: 'Ocean Star', requestedBy: 'John Doe', requestedOn: 'May 12, 2024', status: 'Pending' },
    { id: 'a2', documentName: 'Crew List.docx', vessel: 'Ocean Star', requestedBy: 'Priya Sharma', requestedOn: 'May 11, 2024', status: 'Pending' },
    { id: 'a3', documentName: 'Maintenance Log.xlsx', vessel: 'Sea Breeze', requestedBy: 'Rohit Kumar', requestedOn: 'May 10, 2024', status: 'Pending' },
    { id: 'a4', documentName: 'Safety Certificate.pdf', vessel: 'Blue Horizon', requestedBy: 'Ankita Verma', requestedOn: 'May 09, 2024', status: 'Pending' },
    { id: 'a5', documentName: 'Survey Report.pdf', vessel: 'Pacific Dawn', requestedBy: 'Vikram Singh', requestedOn: 'May 08, 2024', status: 'Pending' },
    { id: 'a6', documentName: 'Cargo Manifest.pdf', vessel: 'Golden Pearl', requestedBy: 'John Doe', requestedOn: 'May 05, 2024', status: 'Approved' },
    { id: 'a7', documentName: 'De-ballasting Plan.docx', vessel: 'Atlantic Wave', requestedBy: 'Rohit Kumar', requestedOn: 'May 02, 2024', status: 'Rejected' },
];
var INITIAL_MOCK_NOTIFICATIONS = [
    { id: 'n1', title: 'Document Expired', message: 'Document expired: Safety Certificate.pdf for Blue Horizon', timestamp: 'May 12, 2024 10:30 AM', priority: 'High', read: false, type: 'alert' },
    { id: 'n2', title: 'Document Expiring Soon', message: 'Document expiring in 7 days: Insurance Certificate.pdf for Ocean Star', timestamp: 'May 12, 2024 09:15 AM', priority: 'Medium', read: false, type: 'warning' },
    { id: 'n3', title: 'Document Approved', message: 'Document approved: Crew List.docx for Ocean Star', timestamp: 'May 11, 2024 04:20 PM', priority: 'Low', read: true, type: 'success' },
    { id: 'n4', title: 'New Document Uploaded', message: 'New document uploaded: Maintenance Log.xlsx for Sea Breeze', timestamp: 'May 11, 2024 11:05 AM', priority: 'Low', read: true, type: 'info' },
    { id: 'n5', title: 'Approval Requested', message: 'Approval requested: Survey Report.pdf for Pacific Dawn', timestamp: 'May 10, 2024 02:45 PM', priority: 'Medium', read: false, type: 'warning' },
];
var INITIAL_MOCK_USERS = [
    { id: 'u1', name: 'Priya Sharma', email: 'priya.sharma@company.com', role: 'Administrator', status: 'Active', lastLogin: 'May 12, 2024' },
    { id: 'u2', name: 'John Doe', email: 'john.doe@company.com', role: 'Manager', status: 'Active', lastLogin: 'May 12, 2024' },
    { id: 'u3', name: 'Rohit Kumar', email: 'rohit.kumar@company.com', role: 'User', status: 'Active', lastLogin: 'May 11, 2024' },
    { id: 'u4', name: 'Ankita Verma', email: 'ankita.verma@company.com', role: 'User', status: 'Active', lastLogin: 'May 10, 2024' },
    { id: 'u5', name: 'Vikram Singh', email: 'vikram.singh@company.com', role: 'Reviewer', status: 'Inactive', lastLogin: 'May 08, 2024' },
];
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
        _this._handleSignOut = function () {
            var base = _this._base();
            var sid = _this.props.sessionId || '';
            var email = _this.props.userEmail || '';
            fetch("".concat(base, "/api/auth/logout"), {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ email: email, session_id: sid })
            })["catch"](function () { return undefined; });
            // Reload the page to force re-authentication
            window.location.reload();
        };
        // ── Graph API Folder Walking (mirrors VesselListView.tsx from reference project) ────
        _this.GRAPH_FETCH_CONCURRENCY = 6;
        _this.VESSEL_ROOT = 'Vessel Management';
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
                                vessels: data.map(function (v) { return (__assign(__assign({}, v), { name: cleanName(v.name), status: v.status || 'Active' })); }),
                                panelLoading: false
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
                        this.setState({ recycleBin: (data || []).map(function (v) { return (__assign(__assign({}, v), { name: cleanName(v.name || '') })); }), panelLoading: false });
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
                        this.setState({ archiveList: (data || []).map(function (v) { return (__assign(__assign({}, v), { name: cleanName(v.name || '') })); }), panelLoading: false });
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
                            ? this._fetchJson("".concat(base, "/api/my-approvals"))["catch"](function () { return null; })
                            : Promise.resolve(null);
                        adminDataPromise = userEmail
                            ? this._fetchJson("".concat(base, "/api/approvals?admin=").concat(encodeURIComponent(userEmail)))["catch"](function () { return null; })
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
                            status: a.status === 'approved' ? 'Approved' : a.status === 'rejected' ? 'Rejected' : 'Pending'
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
                        return [4 /*yield*/, this._fetchJson("".concat(base, "/api/approvals?admin=").concat(encodeURIComponent(this.props.userEmail || '')))["catch"](function () { return null; })];
                    case 22:
                        data = _l.sent();
                        return [4 /*yield*/, this._fetchJson("".concat(base, "/api/my-approvals"))["catch"](function () { return null; })];
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
                    vessel_type: v.vessel_type || ''
                },
                modalMsg: null, modalError: null
            });
        };
        _this._openDeleteVessel = function (v) {
            _this.setState({ modal: 'delete', selectedVessel: v, modalMsg: null, modalError: null });
        };
        _this._closeModal = function () {
            if (!_this.state.modalBusy)
                _this.setState({ modal: 'none', modalMsg: null, modalError: null });
        };
        _this._submitCreate = function () { return __awaiter(_this, void 0, void 0, function () {
            var _a, form, vessels, res, data, msg, newVesselRecord, e_1, newVessel;
            var _this = this;
            var _b, _c;
            return __generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        _a = this.state, form = _a.form, vessels = _a.vessels;
                        if (!form.name.trim()) {
                            this.setState({ modalError: 'Vessel name is required.' });
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
                        _d.label = 1;
                    case 1:
                        _d.trys.push([1, 4, , 5]);
                        return [4 /*yield*/, fetch("".concat(this._base(), "/api/vessels"), {
                                method: 'POST', headers: this._headers(),
                                body: JSON.stringify({ name: form.name.trim(), imo: form.imo.trim(), shipyard: form.shipyard.trim() || null, hull_number: form.hull_number.trim() || null, vessel_type: form.vessel_type || null })
                            })];
                    case 2:
                        res = _d.sent();
                        return [4 /*yield*/, res.json()];
                    case 3:
                        data = _d.sent();
                        if (!res.ok && res.status !== 202)
                            throw new Error((_b = data === null || data === void 0 ? void 0 : data.message) !== null && _b !== void 0 ? _b : "Error ".concat(res.status));
                        msg = data.status === 'pending' ? '⏳ Vessel creation submitted for approval.' : "\u2705 Vessel \"".concat(form.name, "\" created successfully.");
                        newVesselRecord = {
                            id: data.id || ((_c = data.result) === null || _c === void 0 ? void 0 : _c.id) || "v_".concat(Date.now()),
                            name: form.name.trim(),
                            imo: form.imo.trim(),
                            shipyard: form.shipyard.trim() || undefined,
                            hull_number: form.hull_number.trim() || undefined,
                            vessel_type: form.vessel_type || undefined,
                            status: 'Active',
                            image_url: (0, vesselImagePool_1.pickRandomVesselImage)(form.vessel_type)
                        };
                        this.setState({ modalBusy: false, modalMsg: msg, modalError: null, vessels: __spreadArray(__spreadArray([], this.state.vessels, true), [newVesselRecord], false) });
                        // Kick off SharePoint folder provisioning in the background (non-blocking)
                        this._provisionVesselFolders(form.name.trim())["catch"](function () { return undefined; });
                        setTimeout(function () { return _this.setState({ modal: 'none', reloadKey: _this.state.reloadKey + 1 }); }, 1600);
                        return [3 /*break*/, 5];
                    case 4:
                        e_1 = _d.sent();
                        newVessel = {
                            id: "v_".concat(Date.now()),
                            name: form.name.trim(),
                            imo: form.imo.trim(),
                            shipyard: form.shipyard.trim(),
                            hull_number: form.hull_number.trim(),
                            vessel_type: form.vessel_type || 'Bulk Carrier',
                            status: 'Active',
                            image_url: (0, vesselImagePool_1.pickRandomVesselImage)(form.vessel_type)
                        };
                        this.setState({
                            vessels: __spreadArray(__spreadArray([], vessels, true), [newVessel], false),
                            modalBusy: false,
                            modalMsg: "\u2705 Vessel \"".concat(form.name, "\" created successfully.")
                        });
                        // Kick off SharePoint folder provisioning in the background (non-blocking)
                        this._provisionVesselFolders(form.name.trim())["catch"](function () { return undefined; });
                        setTimeout(function () { return _this.setState({ modal: 'none' }); }, 1200);
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
                                body: JSON.stringify({ name: form.name.trim() || null, imo: form.imo.trim() || null, shipyard: form.shipyard.trim() || null, hull_number: form.hull_number.trim() || null, vessel_type: form.vessel_type || null })
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
            var _a, selectedVessel, vessels, res, data, updated, msg, e_3, updated;
            var _this = this;
            var _b;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        _a = this.state, selectedVessel = _a.selectedVessel, vessels = _a.vessels;
                        if (!selectedVessel)
                            return [2 /*return*/];
                        this.setState({ modalBusy: true, modalError: null });
                        _c.label = 1;
                    case 1:
                        _c.trys.push([1, 4, , 5]);
                        return [4 /*yield*/, fetch("".concat(this._base(), "/api/vessels/").concat(selectedVessel.id, "?vessel_name=").concat(encodeURIComponent(selectedVessel.name)), {
                                method: 'DELETE', headers: this._headers()
                            })];
                    case 2:
                        res = _c.sent();
                        return [4 /*yield*/, res.json()];
                    case 3:
                        data = _c.sent();
                        if (!res.ok && res.status !== 202)
                            throw new Error((_b = data === null || data === void 0 ? void 0 : data.message) !== null && _b !== void 0 ? _b : "Error ".concat(res.status));
                        updated = vessels.filter(function (v) { return v.id !== selectedVessel.id; });
                        msg = data.status === 'pending'
                            ? "Delete submitted for approval. \"".concat(selectedVessel.name, "\" removed from list.")
                            : "\"".concat(selectedVessel.name, "\" moved to Recycle Bin.");
                        this.setState({ modalBusy: false, modalMsg: msg, vessels: updated });
                        setTimeout(function () { return _this.setState({ modal: 'none', selectedVessel: null }); }, 1800);
                        return [3 /*break*/, 5];
                    case 4:
                        e_3 = _c.sent();
                        updated = vessels.filter(function (v) { return v.id !== selectedVessel.id; });
                        this.setState({ vessels: updated, modalBusy: false, modalMsg: "\"".concat(selectedVessel.name, "\" deleted.") });
                        setTimeout(function () { return _this.setState({ modal: 'none', selectedVessel: null }); }, 1200);
                        return [3 /*break*/, 5];
                    case 5: return [2 /*return*/];
                }
            });
        }); };
        _this._filesLoadedForFolders = new Set();
        _this._handleUpload = function (row, files) { return __awaiter(_this, void 0, void 0, function () {
            var done, failed, pending, _i, files_1, file, fd, endpoint, h, res, data, _a, _b, infoMsg, errMsg;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        if (!files.length)
                            return [2 /*return*/];
                        this.setState({ uploadingGroupKey: row.groupKey, uploadError: null, uploadInfo: "Uploading ".concat(files.length === 1 ? files[0].name : "".concat(files.length, " files"), "\u2026") });
                        done = 0, failed = 0, pending = 0;
                        _i = 0, files_1 = files;
                        _c.label = 1;
                    case 1:
                        if (!(_i < files_1.length)) return [3 /*break*/, 12];
                        file = files_1[_i];
                        _c.label = 2;
                    case 2:
                        _c.trys.push([2, 10, , 11]);
                        fd = new FormData();
                        fd.append('file', file);
                        endpoint = row.monthDriven
                            ? "".concat(this._base(), "/api/folders/").concat(row.uploadFolderId, "/month-upload")
                            : "".concat(this._base(), "/api/folders/").concat(row.uploadFolderId, "/upload");
                        h = this._uploadHeaders();
                        return [4 /*yield*/, fetch(endpoint, { method: 'POST', body: fd, headers: h })];
                    case 3:
                        res = _c.sent();
                        if (!(res.ok || res.status === 202)) return [3 /*break*/, 8];
                        _c.label = 4;
                    case 4:
                        _c.trys.push([4, 6, , 7]);
                        return [4 /*yield*/, res.json()];
                    case 5:
                        data = _c.sent();
                        if ((data === null || data === void 0 ? void 0 : data.status) === 'pending') {
                            pending++;
                        }
                        else {
                            done++;
                        }
                        return [3 /*break*/, 7];
                    case 6:
                        _a = _c.sent();
                        done++;
                        return [3 /*break*/, 7];
                    case 7: return [3 /*break*/, 9];
                    case 8:
                        failed++;
                        _c.label = 9;
                    case 9: return [3 /*break*/, 11];
                    case 10:
                        _b = _c.sent();
                        failed++;
                        return [3 /*break*/, 11];
                    case 11:
                        _i++;
                        return [3 /*break*/, 1];
                    case 12:
                        infoMsg = null;
                        errMsg = null;
                        if (failed > 0) {
                            errMsg = "\u274C ".concat(failed, " file(s) failed. ").concat(done, " succeeded.");
                        }
                        else if (pending > 0 && done === 0) {
                            infoMsg = "\u23F3 ".concat(pending, " file(s) submitted for approval \u2014 status set to Pending.");
                        }
                        else if (pending > 0) {
                            infoMsg = "\u2705 ".concat(done, " uploaded. \u23F3 ").concat(pending, " pending approval.");
                        }
                        else {
                            infoMsg = "\u2705 ".concat(done, " file(s) uploaded successfully.");
                        }
                        this.setState({ uploadingGroupKey: null, uploadInfo: infoMsg, uploadError: errMsg });
                        if (!(done > 0)) return [3 /*break*/, 14];
                        return [4 /*yield*/, this._refreshFolderFiles(row.uploadFolderId, row.groupKey, true)];
                    case 13:
                        _c.sent();
                        _c.label = 14;
                    case 14: return [2 /*return*/];
                }
            });
        }); };
        // ── Existing Views: Recycle Bin ───────────────────────────────────────────
        _this._restoreFromRecycleBin = function (item) { return __awaiter(_this, void 0, void 0, function () {
            var _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, fetch("".concat(this._base(), "/api/recycle-bin/nodes/").concat(item.id, "/restore"), {
                                method: 'POST', headers: this._headers()
                            })];
                    case 1:
                        _b.sent();
                        return [3 /*break*/, 3];
                    case 2:
                        _a = _b.sent();
                        return [3 /*break*/, 3];
                    case 3:
                        this.setState(function (prev) { return ({ recycleBin: prev.recycleBin.filter(function (r) { return r.id !== item.id; }) }); });
                        if (item.kind === 'vessel' || item.item_type === 'vessel') {
                            this.setState(function (prev) { return ({ reloadKey: prev.reloadKey + 1 }); });
                        }
                        return [2 /*return*/];
                }
            });
        }); };
        _this._permanentDeleteFromRecycleBin = function (item) { return __awaiter(_this, void 0, void 0, function () {
            var _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (!window.confirm("Permanently delete \"".concat(item.name, "\"? This cannot be undone.")))
                            return [2 /*return*/];
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, fetch("".concat(this._base(), "/api/recycle-bin/nodes/").concat(item.id), {
                                method: 'DELETE', headers: this._headers()
                            })];
                    case 2:
                        _b.sent();
                        return [3 /*break*/, 4];
                    case 3:
                        _a = _b.sent();
                        return [3 /*break*/, 4];
                    case 4:
                        this.setState(function (prev) { return ({ recycleBin: prev.recycleBin.filter(function (r) { return r.id !== item.id; }) }); });
                        return [2 /*return*/];
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
            modal: 'none', selectedVessel: null, form: __assign({}, BLANK_FORM),
            modalBusy: false, modalMsg: null, modalError: null,
            view: 'dashboard',
            recycleBin: [], archiveList: [], panelLoading: false,
            vesselsSearch: '', vesselStatusFilter: 'all', vesselTypeFilter: 'all',
            documentsList: INITIAL_MOCK_DOCUMENTS,
            docViewMode: 'folder',
            docMainFolder: null,
            showAllVesselsInFolderView: false,
            docListPage: 0,
            docListSort: 'default',
            docGroupFilter: 'all',
            docUploadRowKey: null,
            docUploadBusy: false,
            docUploadMsg: null,
            folderPathStack: [],
            uploadedFilesByFolder: {},
            selectedDocPreview: null,
            templatesList: INITIAL_MOCK_TEMPLATES,
            approvalsList: [],
            approvalTab: 'Pending',
            notificationsList: INITIAL_MOCK_NOTIFICATIONS,
            notificationFilter: 'all',
            usersList: INITIAL_MOCK_USERS,
            userSearch: '', userRoleFilter: 'all',
            settingsTab: 'General',
            settingsForm: {
                siteTitle: 'Vessel Documents Management',
                siteDescription: 'Manage and track all vessel related documents efficiently.',
                dateFormat: 'MM/DD/YYYY',
                timeZone: '(UTC+05:30) Chennai, Kolkata, Mumbai, New Delhi'
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
            spoFolderMap: new Map(),
            lastDeltaSync: null,
            sessionExpired: false,
            sessionReady: false
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
    };
    VesselEmail.prototype._loadBentoConfig = function () {
        var _this = this;
        this._fetchJson("".concat(this._base(), "/api/email-notification/config"))
            .then(function (cfg) {
            var recipient = (cfg === null || cfg === void 0 ? void 0 : cfg.ai_bento_recipient) || (cfg === null || cfg === void 0 ? void 0 : cfg.graph_sender_mailbox) || '';
            _this.setState({ bentoConfigRecipient: recipient });
        })["catch"](function () { return undefined; });
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
        // On-demand file loading when user selects a specific vessel in the Documents list view
        var _a = this.state, vesselFilter = _a.vesselFilter, rows = _a.rows;
        if (ps.vesselFilter !== vesselFilter && vesselFilter !== 'all' && rows.length > 0) {
            if (!this._filesLoadedForVessels.has(vesselFilter)) {
                this._filesLoadedForVessels.add(vesselFilter);
                this._loadFilesForVessel(vesselFilter)["catch"](function () { return undefined; });
            }
        }
    };
    VesselEmail.prototype.componentWillUnmount = function () {
        var _a, _b;
        (_a = this._abort) === null || _a === void 0 ? void 0 : _a.abort();
        (_b = this._syncScheduler) === null || _b === void 0 ? void 0 : _b.stop();
    };
    // ── Delta Sync ────────────────────────────────────────────────────────────
    VesselEmail.prototype._startDeltaSync = function () {
        var _this = this;
        var _a = this.props, graphClient = _a.graphClient, siteId = _a.siteId, driveId = _a.driveId;
        if (!graphClient || !siteId || !driveId)
            return;
        this._syncScheduler = (0, deltaSync_1.createSyncScheduler)(graphClient, siteId, driveId, function (result) { return _this._applyDeltaResult(result); }, 45000, function (err) { return console.warn('[VesselDMS] Delta sync error:', err); });
        this._syncScheduler.start();
        // Run immediately on mount
        void this._syncScheduler.triggerNow()["catch"](function () { return undefined; });
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
                            })["catch"](function () { return undefined; });
                        }
                    });
                }, 0);
            }
            return { spoFolderMap: map, lastDeltaSync: new Date() };
        });
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
    VesselEmail.prototype._fetchJson = function (url, signal) {
        var _this = this;
        return fetch(url, { signal: signal, headers: this._headers() })
            .then(function (r) {
            if (r.status === 401 && _this.state.sessionReady) {
                // Only show expired screen if session was previously confirmed valid
                _this.setState({ sessionExpired: true });
                throw new Error('SESSION_EXPIRED');
            }
            if (!r.ok)
                throw new Error("HTTP ".concat(r.status));
            return r.json();
        });
    };
    VesselEmail.prototype._base = function () {
        var url = (this.props.apiBaseUrl || '').replace(/\/$/, '');
        return url || 'http://localhost:8000';
    };
    // ── Data Loading ──────────────────────────────────────────────────────────
    VesselEmail.prototype._loadData = function () {
        var _this = this;
        var _a;
        // NOTE: we intentionally do NOT require this.props.sessionId here.
        // A backend running without a database (settings.db_configured === false)
        // legitimately never issues a session_id, and require_session() on the
        // server allows unauthenticated calls in that mode. Gating on sessionId
        // alone left the vessel list permanently blank in that configuration.
        // _headers() already omits the Authorization/X-Session-ID headers when
        // sessionId is empty, so this is safe to call regardless.
        (_a = this._abort) === null || _a === void 0 ? void 0 : _a.abort();
        this._abort = new AbortController();
        var signal = this._abort.signal;
        var base = this._base();
        if (!base) {
            this.setState({ loading: false, rows: [] });
            return;
        }
        this.setState({ loading: true, error: null });
        // Step 1: Always fetch vessel list from backend database first
        this._fetchJson("".concat(base, "/api/vessels"), signal)["catch"](function () { return null; })
            .then(function (vesselList) { return __awaiter(_this, void 0, void 0, function () {
            var vessels, _a, graphClient, siteId, driveId, graphRows, graphErr_1, flatTree, rows;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (signal.aborted)
                            return [2 /*return*/];
                        vessels = [];
                        if (vesselList && Array.isArray(vesselList) && vesselList.length > 0) {
                            vessels = vesselList.map(function (v) { return (__assign(__assign({}, v), { name: cleanName(v.name), status: v.status || 'Active' })); });
                        }
                        // Update vessel list in state immediately so UI shows vessels right away
                        if (vessels.length > 0) {
                            this.setState({ vessels: vessels });
                        }
                        _a = this.props, graphClient = _a.graphClient, siteId = _a.siteId, driveId = _a.driveId;
                        if (!(graphClient && siteId && driveId && vessels.length > 0)) return [3 /*break*/, 4];
                        console.log('[VesselDMS] _loadData: Graph context available — using Graph API for folder tree walk');
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._flattenAllViaGraph(vessels, signal)];
                    case 2:
                        graphRows = _b.sent();
                        if (!signal.aborted) {
                            console.log("[VesselDMS] _loadData: Graph walk returned ".concat(graphRows.length, " rows"));
                            this.setState({ vessels: vessels, rows: this._normalize(graphRows), loading: false });
                            return [2 /*return*/];
                        }
                        return [3 /*break*/, 4];
                    case 3:
                        graphErr_1 = _b.sent();
                        console.warn('[VesselDMS] _loadData: Graph walk failed, falling back to REST API:', graphErr_1);
                        return [3 /*break*/, 4];
                    case 4:
                        if (signal.aborted)
                            return [2 /*return*/];
                        return [4 /*yield*/, this._fetchJson("".concat(base, "/api/vessels/flat-tree"), signal)["catch"](function () { return null; })];
                    case 5:
                        flatTree = _b.sent();
                        if (!!signal.aborted) return [3 /*break*/, 9];
                        if (!(flatTree && Array.isArray(flatTree) && flatTree.length > 0)) return [3 /*break*/, 6];
                        this.setState({ vessels: vessels, rows: this._normalize(flatTree), loading: false });
                        return [3 /*break*/, 9];
                    case 6:
                        if (!(vessels.length > 0)) return [3 /*break*/, 8];
                        return [4 /*yield*/, this._flattenAll(vessels, signal)];
                    case 7:
                        rows = _b.sent();
                        if (!signal.aborted) {
                            this.setState({ vessels: vessels, rows: this._normalize(rows), loading: false });
                        }
                        return [3 /*break*/, 9];
                    case 8:
                        this.setState({ vessels: vessels, rows: [], loading: false });
                        _b.label = 9;
                    case 9: return [2 /*return*/];
                }
            });
        }); })["catch"](function (err) {
            var _a;
            if (!signal.aborted)
                _this.setState({ loading: false, error: (_a = err === null || err === void 0 ? void 0 : err.message) !== null && _a !== void 0 ? _a : 'Failed to load data.' });
        });
    };
    VesselEmail.prototype._normalize = function (raw) {
        return raw.map(function (r) { return (__assign(__assign({}, r), { vesselName: cleanName(r.vesselName), group: cleanName(r.group), category: cleanName(r.category) })); });
    };
    /**
     * Upload a file to a SharePoint folder via Graph API (direct SPO upload) with REST API fallback.
     * Prevents FastAPI 404 path routing errors when folder path strings containing slashes are used.
     */
    VesselEmail.prototype._uploadFileToFolder = function (uploadFolderId, subFolderPath, vesselName, file, monthDriven) {
        if (monthDriven === void 0) { monthDriven = false; }
        return __awaiter(this, void 0, void 0, function () {
            var base, form, endpoint, resp, errText, data;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        base = this._base();
                        if (!base)
                            return [2 /*return*/, { fileId: null, statusPending: false }];
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
                                body: form
                            })];
                    case 1:
                        resp = _a.sent();
                        if (!!resp.ok) return [3 /*break*/, 3];
                        return [4 /*yield*/, resp.text()["catch"](function () { return resp.statusText; })];
                    case 2:
                        errText = _a.sent();
                        throw new Error("HTTP ".concat(resp.status, ": ").concat(errText));
                    case 3: return [4 /*yield*/, resp.json()["catch"](function () { return ({}); })];
                    case 4:
                        data = _a.sent();
                        return [2 /*return*/, { fileId: (data === null || data === void 0 ? void 0 : data.id) || null, statusPending: (data === null || data === void 0 ? void 0 : data.status) === 'pending' }];
                }
            });
        });
    };
    /**
     * List children of a SharePoint drive path via Graph API.
     * Returns an array of { id, name, isFolder, monthDriven, upload } items.
     */
    VesselEmail.prototype._getGraphChildren = function (folderPath, signal) {
        var _a, _b, _c;
        return __awaiter(this, void 0, void 0, function () {
            var _d, graphClient, siteId, driveId, encodedPath, url, result, err_1, status_1;
            return __generator(this, function (_e) {
                switch (_e.label) {
                    case 0:
                        _d = this.props, graphClient = _d.graphClient, siteId = _d.siteId, driveId = _d.driveId;
                        if (!graphClient || !siteId || !driveId)
                            return [2 /*return*/, []];
                        encodedPath = folderPath.split('/').map(function (s) { return encodeURIComponent(s); }).join('/');
                        url = "/sites/".concat(siteId, "/drives/").concat(driveId, "/root:/").concat(encodedPath, ":/children") +
                            "?$select=id,name,folder,file&$top=200";
                        console.log("[VesselDMS] _getGraphChildren \u2192 ".concat(url));
                        _e.label = 1;
                    case 1:
                        _e.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, graphClient.api(url).get()];
                    case 2:
                        result = _e.sent();
                        if (signal.aborted)
                            return [2 /*return*/, []];
                        return [2 /*return*/, ((_a = result.value) !== null && _a !== void 0 ? _a : []).map(function (item) { return ({
                                id: item.id,
                                name: item.name,
                                isFolder: !!item.folder,
                                monthDriven: false,
                                upload: !!item.folder
                            }); })];
                    case 3:
                        err_1 = _e.sent();
                        status_1 = (_c = (_b = err_1 === null || err_1 === void 0 ? void 0 : err_1.statusCode) !== null && _b !== void 0 ? _b : err_1 === null || err_1 === void 0 ? void 0 : err_1.code) !== null && _c !== void 0 ? _c : 0;
                        if (status_1 === 404)
                            return [2 /*return*/, []];
                        throw err_1;
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Recursively walk a Graph folder path and emit FlatRow entries.
     * Mirrors the walkFolder() + leafToRows() pattern from the reference project's VesselListView.tsx.
     */
    VesselEmail.prototype._walkGraphFolder = function (folderPath, pathParts, vesselName, signal, srCounter, onRows) {
        var _a, _b;
        return __awaiter(this, void 0, void 0, function () {
            var kids, _c, files, subFolders, stripPrefix, group, category, subPath, groupKey, canUpload, baseSr_1, suffixes_1, leafRows;
            var _this = this;
            return __generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        if (signal.aborted)
                            return [2 /*return*/];
                        _d.label = 1;
                    case 1:
                        _d.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._getGraphChildren(folderPath, signal)];
                    case 2:
                        kids = _d.sent();
                        return [3 /*break*/, 4];
                    case 3:
                        _c = _d.sent();
                        return [2 /*return*/];
                    case 4:
                        if (signal.aborted)
                            return [2 /*return*/];
                        files = kids.filter(function (k) { return !k.isFolder; });
                        subFolders = kids.filter(function (k) { return k.isFolder; });
                        stripPrefix = function (s) { return s.replace(/^Folder-\d+\s+/i, ''); };
                        group = pathParts.length >= 2 ? stripPrefix(pathParts[1]) : stripPrefix((_b = (_a = pathParts[0]) !== null && _a !== void 0 ? _a : folderPath.split('/').pop()) !== null && _b !== void 0 ? _b : folderPath);
                        category = pathParts.length >= 3 ? stripPrefix(pathParts[pathParts.length - 1]) : group;
                        subPath = pathParts.length >= 2
                            ? __spreadArray([vesselName], pathParts.slice(1).map(stripPrefix), true).join(' > ')
                            : __spreadArray([vesselName], pathParts.map(stripPrefix), true).join(' > ');
                        groupKey = "".concat(vesselName, "||").concat(group, "||").concat(category, "||").concat(subPath);
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
                                        subFolderPath: subPath, fileName: null, fileId: null,
                                        canUpload: canUpload,
                                        groupKey: groupKey,
                                        uploadFolderId: folderPath,
                                        monthDriven: false
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
                                        subFolderPath: subPath,
                                        fileName: f.name, fileId: f.id,
                                        canUpload: canUpload,
                                        groupKey: groupKey,
                                        uploadFolderId: folderPath,
                                        monthDriven: false
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
                        _d.sent();
                        _d.label = 6;
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
            var allRows, srCounter;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        allRows = [];
                        srCounter = { value: 0 };
                        // Walk each main folder in parallel
                        return [4 /*yield*/, this._mapLimit(this.MAIN_FOLDER_NAMES, 3, function (mainFolderName) { return __awaiter(_this, void 0, void 0, function () {
                                var vesselPath, topCats, folders, files, displayMainFolder_1, leafRows;
                                var _this = this;
                                return __generator(this, function (_a) {
                                    switch (_a.label) {
                                        case 0:
                                            if (signal.aborted)
                                                return [2 /*return*/];
                                            vesselPath = "".concat(this.VESSEL_ROOT, "/").concat(mainFolderName, "/").concat(vesselName);
                                            return [4 /*yield*/, this._getGraphChildren(vesselPath, signal)["catch"](function () { return []; })];
                                        case 1:
                                            topCats = _a.sent();
                                            if (signal.aborted)
                                                return [2 /*return*/];
                                            folders = topCats.filter(function (c) { return c.isFolder; });
                                            files = topCats.filter(function (c) { return !c.isFolder; });
                                            if (folders.length === 0) {
                                                displayMainFolder_1 = mainFolderName.replace(/^Folder-\d+\s+/i, '');
                                                srCounter.value += 1;
                                                leafRows = files.map(function (f, idx) { return ({
                                                    srNo: idx === 0 ? String(srCounter.value) : "".concat(srCounter.value).concat(String.fromCharCode(97 + idx - 1)),
                                                    vesselName: vesselName,
                                                    group: displayMainFolder_1,
                                                    category: displayMainFolder_1,
                                                    subFolderPath: "".concat(vesselName, " > ").concat(displayMainFolder_1),
                                                    fileName: f.name,
                                                    fileId: f.id,
                                                    canUpload: true,
                                                    groupKey: "".concat(vesselName, "||").concat(displayMainFolder_1, "||").concat(displayMainFolder_1, "||").concat(vesselName, " > ").concat(displayMainFolder_1),
                                                    uploadFolderId: f.id,
                                                    monthDriven: false
                                                }); });
                                                if (leafRows.length === 0) {
                                                    leafRows.push({
                                                        srNo: String(srCounter.value),
                                                        vesselName: vesselName,
                                                        group: displayMainFolder_1,
                                                        category: displayMainFolder_1,
                                                        subFolderPath: "".concat(vesselName, " > ").concat(displayMainFolder_1),
                                                        fileName: null, fileId: null, canUpload: true,
                                                        groupKey: "".concat(vesselName, "||").concat(displayMainFolder_1, "||").concat(displayMainFolder_1),
                                                        uploadFolderId: vesselPath,
                                                        monthDriven: false
                                                    });
                                                }
                                                allRows.push.apply(allRows, leafRows);
                                                onChunk(leafRows);
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
                                                                return [4 /*yield*/, this._walkGraphFolder("".concat(vesselPath, "/").concat(cat.name), [mainFolderName, cat.name], vesselName, signal, srCounter, function (rows) {
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
                                        case 2:
                                            // Walk each category folder with bounded concurrency
                                            _a.sent();
                                            return [2 /*return*/];
                                    }
                                });
                            }); })];
                    case 1:
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
            var allRows, flushTimer, queueFlush;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        allRows = [];
                        flushTimer = null;
                        queueFlush = function () {
                            if (flushTimer !== null)
                                return;
                            flushTimer = setTimeout(function () {
                                flushTimer = null;
                                if (signal.aborted)
                                    return;
                                var snapshot = __spreadArray([], allRows, true);
                                _this.setState(function (prev) { return ({ rows: _this._normalize(snapshot), vessels: prev.vessels }); });
                            }, 120);
                        };
                        return [4 /*yield*/, this._mapLimit(vessels, 3, function (vessel) { return __awaiter(_this, void 0, void 0, function () {
                                return __generator(this, function (_a) {
                                    switch (_a.label) {
                                        case 0:
                                            if (signal.aborted)
                                                return [2 /*return*/];
                                            return [4 /*yield*/, this._flattenVesselViaGraph(vessel.name, signal, function (chunk) {
                                                    allRows.push.apply(allRows, chunk);
                                                    queueFlush();
                                                })];
                                        case 1:
                                            _a.sent();
                                            return [2 /*return*/];
                                    }
                                });
                            }); })];
                    case 1:
                        _a.sent();
                        if (flushTimer !== null) {
                            clearTimeout(flushTimer);
                            flushTimer = null;
                        }
                        return [2 /*return*/, allRows];
                }
            });
        });
    };
    VesselEmail.prototype._flattenAll = function (vessels, signal) {
        return __awaiter(this, void 0, void 0, function () {
            var base, out, sr, _i, vessels_1, v, mains, _a, mains_1, m, cats, _b, cats_1, cat, sp, _c, _d;
            return __generator(this, function (_e) {
                switch (_e.label) {
                    case 0:
                        base = this._base();
                        out = [];
                        sr = 0;
                        _i = 0, vessels_1 = vessels;
                        _e.label = 1;
                    case 1:
                        if (!(_i < vessels_1.length)) return [3 /*break*/, 12];
                        v = vessels_1[_i];
                        if (signal.aborted)
                            return [3 /*break*/, 12];
                        _e.label = 2;
                    case 2:
                        _e.trys.push([2, 10, , 11]);
                        return [4 /*yield*/, this._fetchJson("".concat(base, "/api/vessels/").concat(v.id, "/mains"), signal)];
                    case 3:
                        mains = _e.sent();
                        _a = 0, mains_1 = mains;
                        _e.label = 4;
                    case 4:
                        if (!(_a < mains_1.length)) return [3 /*break*/, 9];
                        m = mains_1[_a];
                        if (signal.aborted)
                            return [3 /*break*/, 9];
                        _e.label = 5;
                    case 5:
                        _e.trys.push([5, 7, , 8]);
                        return [4 /*yield*/, this._fetchJson("".concat(base, "/api/folders/").concat(m.id, "/children"), signal)];
                    case 6:
                        cats = _e.sent();
                        for (_b = 0, cats_1 = cats; _b < cats_1.length; _b++) {
                            cat = cats_1[_b];
                            if ((cat === null || cat === void 0 ? void 0 : cat.kind) === 'file')
                                continue;
                            sr++;
                            sp = "".concat(v.name, " > ").concat(m.name, " > ").concat(cat.name);
                            out.push({ srNo: String(sr), vesselName: v.name, group: m.name, category: cat.name, subFolderPath: sp, fileName: null, fileId: null, canUpload: true, groupKey: "".concat(v.id, "||").concat(m.name, "||").concat(cat.name, "||").concat(sp), uploadFolderId: cat.id, monthDriven: Boolean(cat.month_driven) });
                        }
                        return [3 /*break*/, 8];
                    case 7:
                        _c = _e.sent();
                        return [3 /*break*/, 8];
                    case 8:
                        _a++;
                        return [3 /*break*/, 4];
                    case 9: return [3 /*break*/, 11];
                    case 10:
                        _d = _e.sent();
                        return [3 /*break*/, 11];
                    case 11:
                        _i++;
                        return [3 /*break*/, 1];
                    case 12: return [2 /*return*/, out];
                }
            });
        });
    };
    // SharePoint Folder Provisioning ─────────────────────────────────────────
    VesselEmail.prototype._provisionVesselFolders = function (vesselName, vesselId) {
        var _a, _b;
        return __awaiter(this, void 0, void 0, function () {
            var _c, graphClient, siteId, driveId, result, err_2;
            return __generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        _c = this.props, graphClient = _c.graphClient, siteId = _c.siteId, driveId = _c.driveId;
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
                        this.setState({
                            folderCreationBusy: false,
                            folderProvisioningVesselId: null,
                            folderCreationResults: result.results,
                            folderCreationError: result.success ? null : 'Some folders could not be created. Check the creation log.'
                        });
                        // Trigger an immediate delta sync so the new folders appear in the tree
                        void ((_a = this._syncScheduler) === null || _a === void 0 ? void 0 : _a.triggerNow()["catch"](function () { return undefined; }));
                        return [3 /*break*/, 4];
                    case 3:
                        err_2 = _d.sent();
                        this.setState({
                            folderCreationBusy: false,
                            folderProvisioningVesselId: null,
                            folderCreationError: (_b = err_2 === null || err_2 === void 0 ? void 0 : err_2.message) !== null && _b !== void 0 ? _b : 'Folder provisioning failed.'
                        });
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    // ── Load Files for Vessel ─────────────────────────────────────────────────
    VesselEmail.prototype._loadFilesForVessel = function (vesselName) {
        return __awaiter(this, void 0, void 0, function () {
            var rows, vesselRows, uniqueFolderIds, BATCH, i, batch;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        rows = this.state.rows;
                        vesselRows = rows.filter(function (r) { return r.vesselName === vesselName && r.uploadFolderId; });
                        uniqueFolderIds = Array.from(new Set(vesselRows.map(function (r) { return r.uploadFolderId; })));
                        if (uniqueFolderIds.length === 0)
                            return [2 /*return*/];
                        BATCH = 5;
                        i = 0;
                        _a.label = 1;
                    case 1:
                        if (!(i < uniqueFolderIds.length)) return [3 /*break*/, 4];
                        batch = uniqueFolderIds.slice(i, i + BATCH);
                        return [4 /*yield*/, Promise.all(batch.map(function (fid) { return _this._refreshFolderFiles(fid, ''); }))];
                    case 2:
                        _a.sent();
                        _a.label = 3;
                    case 3:
                        i += BATCH;
                        return [3 /*break*/, 1];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    VesselEmail.prototype._refreshFolderFiles = function (uploadFolderId, groupKey, force) {
        if (force === void 0) { force = false; }
        return __awaiter(this, void 0, void 0, function () {
            var children, fileItems_1, parsedUploads_1, _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (!uploadFolderId)
                            return [2 /*return*/];
                        if (/^f\d+$/.test(uploadFolderId) || uploadFolderId.includes('/'))
                            return [2 /*return*/];
                        if (!force && this._filesLoadedForFolders.has(uploadFolderId))
                            return [2 /*return*/];
                        this._filesLoadedForFolders.add(uploadFolderId);
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._fetchJson("".concat(this._base(), "/api/folders/").concat(encodeURIComponent(uploadFolderId), "/children"))];
                    case 2:
                        children = _b.sent();
                        fileItems_1 = (children || []).filter(function (c) { return c.kind === 'file'; });
                        parsedUploads_1 = fileItems_1.map(function (f) { return ({
                            name: f.name || f.displayName,
                            size: f.size ? "".concat((f.size / 1024).toFixed(1), " KB") : '142 KB',
                            date: f.modified || 'Today',
                            pending: false,
                            id: f.id
                        }); });
                        this.setState(function (prev) {
                            var _a;
                            var updatedByFolder = __assign(__assign({}, prev.uploadedFilesByFolder), (_a = {}, _a[uploadFolderId] = parsedUploads_1, _a));
                            var existingBaseRows = prev.rows.filter(function (r) { return r.uploadFolderId === uploadFolderId; });
                            var newRows = prev.rows;
                            if (existingBaseRows.length > 0) {
                                var baseRow_1 = existingBaseRows[0];
                                updatedByFolder[baseRow_1.groupKey] = parsedUploads_1;
                                updatedByFolder[baseRow_1.category] = parsedUploads_1;
                                updatedByFolder[baseRow_1.group] = parsedUploads_1;
                                var otherFolderRows = prev.rows.filter(function (r) { return r.uploadFolderId !== uploadFolderId; });
                                if (fileItems_1.length === 0) {
                                    newRows = __spreadArray(__spreadArray([], otherFolderRows, true), [__assign(__assign({}, baseRow_1), { fileName: null, fileId: null, filePending: false })], false);
                                }
                                else {
                                    var mappedRows = fileItems_1.map(function (f, i) { return (__assign(__assign({}, baseRow_1), { srNo: i === 0 ? baseRow_1.srNo : "".concat(baseRow_1.srNo, ".").concat(i + 1), fileName: f.name || f.displayName || null, fileId: f.id || null, filePending: false })); });
                                    newRows = __spreadArray(__spreadArray([], otherFolderRows, true), mappedRows, true);
                                }
                            }
                            return { rows: newRows, uploadedFilesByFolder: updatedByFolder };
                        });
                        return [3 /*break*/, 4];
                    case 3:
                        _a = _b.sent();
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
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
                map.set(row.groupKey, { srNo: row.srNo, vesselName: row.vesselName, group: row.group, category: row.category, subFolderPath: row.subFolderPath, groupKey: row.groupKey, uploadFolderId: row.uploadFolderId, monthDriven: row.monthDriven, canUpload: row.canUpload, files: row.fileId && row.fileName ? [{ id: row.fileId, name: row.fileName }] : [] });
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
        var _this = this;
        var _a = this.state, view = _a.view, notificationsList = _a.notificationsList;
        var unreadNotifs = notificationsList.filter(function (n) { return !n.read; }).length;
        var navItems = [
            { id: 'dashboard', label: 'Home', icon: '🏠' },
            { id: 'list', label: 'Documents', icon: '📄' },
            { id: 'vessels', label: 'Vessels', icon: '🚢' },
            { id: 'templates', label: 'Templates', icon: '📑' },
            { id: 'approvals', label: 'Approvals', icon: '☑️' },
            { id: 'notifications', label: 'Notifications', icon: '🔔', badge: unreadNotifs > 0 ? unreadNotifs : undefined },
            { id: 'reports', label: 'Reports', icon: '📊' },
            { id: 'users', label: 'User Management', icon: '👥' },
            { id: 'settings', label: 'Settings', icon: '⚙️' },
        ];
        return (<div style={{
                width: 230,
                minWidth: 230,
                background: '#ffffff',
                borderRight: '1px solid #e0e0e0',
                display: 'flex',
                flexDirection: 'column',
                justifyContent: 'space-between',
                padding: '16px 0',
                boxSizing: 'border-box',
                userSelect: 'none'
            }}>
        {/* Brand Header */}
        <div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '0 20px 20px', borderBottom: '1px solid #f0f0f0' }}>
            <div style={{
                width: 36, height: 36, borderRadius: 8, background: '#0078d4',
                color: '#fff', fontWeight: 700, fontSize: 16, display: 'flex',
                alignItems: 'center', justifyContent: 'center', boxShadow: '0 2px 6px rgba(0,120,212,0.3)'
            }}>
              VD
            </div>
            <div>
              <div style={{ fontWeight: 700, fontSize: 14, color: '#111827', lineHeight: 1.2 }}>Vessel Documents</div>
              <div style={{ fontSize: 12, color: '#6b7280', fontWeight: 500 }}>Management</div>
            </div>
          </div>

          {/* Navigation Links */}
          <div style={{ padding: '12px 10px' }}>
            {navItems.map(function (item) {
                var active = view === item.id;
                return (<button key={item.id} onClick={function () { return _this._goToView(item.id); }} style={{
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'space-between',
                        width: '100%',
                        padding: '9px 14px',
                        marginBottom: 3,
                        borderRadius: 6,
                        border: 'none',
                        background: active ? '#eff6ff' : 'transparent',
                        color: active ? '#0078d4' : '#4b5563',
                        fontWeight: active ? 600 : 500,
                        fontSize: 13,
                        cursor: 'pointer',
                        textAlign: 'left',
                        transition: 'all 0.15s ease'
                    }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                    <span style={{ fontSize: 16 }}>{item.icon}</span>
                    <span>{item.label}</span>
                  </div>
                  {item.badge !== undefined && (<span style={{
                            background: '#ef4444', color: '#fff', borderRadius: 10,
                            padding: '1px 7px', fontSize: 10, fontWeight: 700
                        }}>
                      {item.badge}
                    </span>)}
                </button>);
            })}
          </div>
        </div>

        {/* Bottom Auxiliary Links */}
        <div style={{ padding: '12px 10px', borderTop: '1px solid #f0f0f0' }}>
          <button onClick={function () { return _this._goToView('bento_email'); }} style={{
                display: 'flex', alignItems: 'center', gap: 10, width: '100%', padding: '8px 14px',
                borderRadius: 6, border: 'none', background: view === 'bento_email' ? '#eff6ff' : 'transparent',
                color: view === 'bento_email' ? '#0078d4' : '#6b7280', fontSize: 12, fontWeight: 500, cursor: 'pointer'
            }}>
            <span>🤖</span>
            <span>AI Bento Email</span>
          </button>

          <button onClick={function () {
                void _this._goToView('bento_email');
                _this.setState({
                    bentoComposeOpen: true,
                    bentoComposeMsg: null,
                    bentoComposeErr: null,
                    bentoComposeForm: {
                        vessel_name: '',
                        datasource_tag: 'mail',
                        subject_text: _this._buildAutoSubject('', 'mail', ''),
                        body: '',
                        file: null,
                        existing_attachment: '',
                        recipient: ''
                    }
                });
            }} style={{
                display: 'flex', alignItems: 'center', gap: 10, width: '100%', padding: '8px 14px',
                borderRadius: 6, border: 'none', background: 'transparent',
                color: '#6b7280', fontSize: 12, fontWeight: 500, cursor: 'pointer'
            }}>
            <span>✉</span>
            <span>Send Email</span>
          </button>

          <button onClick={function () { return _this._goToView('recycle'); }} style={{
                display: 'flex', alignItems: 'center', gap: 10, width: '100%', padding: '8px 14px',
                borderRadius: 6, border: 'none', background: view === 'recycle' ? '#eff6ff' : 'transparent',
                color: view === 'recycle' ? '#0078d4' : '#6b7280', fontSize: 12, fontWeight: 500, cursor: 'pointer'
            }}>
            <span>🗑️</span>
            <span>Recycle bin</span>
          </button>

          <button onClick={function () { return _this._goToView('archive'); }} style={{
                display: 'flex', alignItems: 'center', gap: 10, width: '100%', padding: '8px 14px',
                borderRadius: 6, border: 'none', background: view === 'archive' ? '#eff6ff' : 'transparent',
                color: view === 'archive' ? '#0078d4' : '#6b7280', fontSize: 12, fontWeight: 500, cursor: 'pointer'
            }}>
            <span>📦</span>
            <span>Archive</span>
          </button>
        </div>
      </div>);
    };
    // ── Layout Wrapper ────────────────────────────────────────────────────────
    VesselEmail.prototype._renderLayout = function (content) {
        var userDisplayName = this.props.userDisplayName || 'Priya';
        return (<div style={{ display: 'flex', minHeight: '100vh', background: '#f8fafc', width: '100%', fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif" }}>
        {this._renderSidebar()}
        <div style={{ flex: 1, display: 'flex', flexDirection: 'column', minWidth: 0 }}>
          {/* Top Bar Header */}
          <div style={{
                height: 52, background: '#ffffff', borderBottom: '1px solid #e2e8f0',
                display: 'flex', alignItems: 'center', justifyContent: 'space-between',
                padding: '0 24px', flexShrink: 0
            }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
              <span style={{ color: '#64748b', fontSize: 14, fontWeight: 500 }}>Vessel DMS</span>
              <span style={{ color: '#cbd5e1' }}>/</span>
              <span style={{ color: '#0f172a', fontSize: 14, fontWeight: 600, textTransform: 'capitalize' }}>
                {this.state.view.replace('_', ' ')}
              </span>
            </div>

            <div style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
              <div style={{ position: 'relative', width: 220 }}>
                <input type="text" placeholder="Search this site..." style={{
                width: '100%', padding: '6px 12px 6px 30px', borderRadius: 16,
                border: '1px solid #cbd5e1', fontSize: 12, background: '#f8fafc',
                outline: 'none', boxSizing: 'border-box'
            }}/>
                <span style={{ position: 'absolute', left: 10, top: '50%', transform: 'translateY(-50%)', color: '#94a3b8', fontSize: 12 }}>🔍</span>
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <div style={{
                width: 28, height: 28, borderRadius: '50%', background: '#0078d4',
                color: '#fff', fontSize: 12, fontWeight: 600, display: 'flex',
                alignItems: 'center', justifyContent: 'center'
            }}>
                  {userDisplayName.charAt(0)}
                </div>
                <span style={{ fontSize: 13, fontWeight: 600, color: '#334155' }}>{userDisplayName}</span>
              </div>
            </div>
          </div>

          {/* Main Module Content */}
          <div style={{ flex: 1, padding: 24, overflowY: 'auto' }}>
            {content}
          </div>

          {/* Right-side Aside Document Preview Drawer */}
          {this._renderDocPreviewDrawer()}
        </div>
      </div>);
    };
    VesselEmail.prototype._renderDocPreviewDrawer = function () {
        var _this = this;
        var selectedDocPreview = this.state.selectedDocPreview;
        if (!selectedDocPreview)
            return null;
        var base = this._base();
        var downloadUrl = selectedDocPreview.fileId ? "".concat(base, "/api/files/").concat(selectedDocPreview.fileId, "/content") : null;
        return (<div style={{
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
                animation: 'slideIn 0.2s ease-out'
            }}>
        {/* Drawer Header */}
        <div style={{
                padding: '16px 20px', background: '#0f172a', color: '#fff',
                display: 'flex', alignItems: 'center', justifyContent: 'space-between'
            }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, fontWeight: 700, fontSize: 15 }}>
            <span>📄</span> Document Overview
          </div>
          <button onClick={function () { return _this.setState({ selectedDocPreview: null }); }} style={{ border: 'none', background: 'transparent', color: '#cbd5e1', fontSize: 18, cursor: 'pointer', fontWeight: 700 }} title="Close Drawer">
            ✕
          </button>
        </div>

        {/* Drawer Content */}
        <div style={{ flex: 1, padding: 20, overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: 16 }}>
          {/* File Icon Banner */}
          <div style={{
                background: '#f0f9ff', borderRadius: 12, padding: 20, border: '1px solid #bae6fd',
                display: 'flex', flexDirection: 'column', alignItems: 'center', textAlign: 'center', gap: 8
            }}>
            <div style={{ fontSize: 42 }}>📄</div>
            <a href={downloadUrl || '#'} target="_blank" rel="noopener noreferrer" onClick={function (e) {
                if (!downloadUrl) {
                    e.preventDefault();
                    alert("File \"".concat(selectedDocPreview.fileName, "\" stored locally."));
                }
            }} style={{
                fontSize: 15, fontWeight: 700, color: '#0284c7', textDecoration: 'underline',
                wordBreak: 'break-all', cursor: 'pointer'
            }}>
              {selectedDocPreview.fileName}
            </a>
            <span style={{ fontSize: 11, background: '#dbeafe', color: '#1e40af', padding: '2px 8px', borderRadius: 10, fontWeight: 600 }}>
              Uploaded Document
            </span>
          </div>

          {/* Document Properties */}
          <div style={{ background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', padding: 14, display: 'flex', flexDirection: 'column', gap: 10 }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: '#475569', textTransform: 'uppercase', letterSpacing: '0.05em', marginBottom: 2 }}>
              Metadata Details
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, borderBottom: '1px solid #f1f5f9', paddingBottom: 6 }}>
              <span style={{ color: '#64748b' }}>🚢 Vessel Name:</span>
              <strong style={{ color: '#0f172a' }}>{selectedDocPreview.vesselName}</strong>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, borderBottom: '1px solid #f1f5f9', paddingBottom: 6 }}>
              <span style={{ color: '#64748b' }}>📁 Group:</span>
              <span style={{ color: '#2563eb', fontWeight: 600 }}>{selectedDocPreview.group}</span>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, borderBottom: '1px solid #f1f5f9', paddingBottom: 6 }}>
              <span style={{ color: '#64748b' }}>🏷️ Category:</span>
              <strong style={{ color: '#1e293b' }}>{selectedDocPreview.category}</strong>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, borderBottom: '1px solid #f1f5f9', paddingBottom: 6 }}>
              <span style={{ color: '#64748b' }}>💾 Size:</span>
              <span style={{ color: '#475569' }}>{selectedDocPreview.size || '142 KB'}</span>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, borderBottom: '1px solid #f1f5f9', paddingBottom: 6 }}>
              <span style={{ color: '#64748b' }}>📅 Date:</span>
              <span style={{ color: '#475569' }}>{selectedDocPreview.date || 'Today'}</span>
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 2, fontSize: 12 }}>
              <span style={{ color: '#64748b' }}>🗺️ Folder Path:</span>
              <span style={{ color: '#334155', fontFamily: 'monospace', fontSize: 11, background: '#f8fafc', padding: '4px 6px', borderRadius: 4, wordBreak: 'break-all' }}>
                {selectedDocPreview.folderPath}
              </span>
            </div>
          </div>

          {/* Quick Preview Card */}
          <div style={{ background: '#f8fafc', borderRadius: 10, border: '1px dashed #cbd5e1', padding: 14, textAlign: 'center' }}>
            <div style={{ fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 }}>
              Document Status: <span style={{ color: '#16a34a', fontWeight: 700 }}>Active / Ready</span>
            </div>
            <p style={{ margin: 0, fontSize: 11, color: '#64748b' }}>
              Click open below to view or download the uploaded document directly.
            </p>
          </div>
        </div>

        {/* Drawer Actions Footer */}
        <div style={{ padding: 16, background: '#f8fafc', borderTop: '1px solid #e2e8f0', display: 'flex', flexDirection: 'column', gap: 8 }}>
          <a href={downloadUrl || '#'} target="_blank" rel="noopener noreferrer" onClick={function (e) {
                if (!downloadUrl) {
                    e.preventDefault();
                    alert("File \"".concat(selectedDocPreview.fileName, "\" stored locally."));
                }
            }} style={{
                display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6,
                background: '#0284c7', color: '#fff', padding: '9px 16px', borderRadius: 8,
                fontSize: 13, fontWeight: 600, textDecoration: 'none', textAlign: 'center'
            }}>
            🌐 Open / View Document
          </a>
          <button onClick={function () {
                _this.setState({
                    view: 'bento_email',
                    bentoUploadVessel: selectedDocPreview.vesselName,
                    selectedDocPreview: null
                });
            }} style={{
                display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6,
                background: '#0f172a', color: '#fff', border: 'none', padding: '9px 16px', borderRadius: 8,
                fontSize: 13, fontWeight: 600, cursor: 'pointer'
            }}>
            ✉ Send via Bento Email
          </button>
          <button onClick={function () { return _this.setState({ selectedDocPreview: null }); }} style={{
                background: '#fff', border: '1px solid #cbd5e1', color: '#475569', padding: '8px 16px',
                borderRadius: 8, fontSize: 12, fontWeight: 600, cursor: 'pointer'
            }}>
            Close Drawer
          </button>
        </div>
      </div>);
    };
    // ── Module 1: Home (Dashboard) ────────────────────────────────────────────
    VesselEmail.prototype._renderDashboard = function () {
        var _this = this;
        var _a = this.state, rows = _a.rows, vessels = _a.vessels, documentsList = _a.documentsList;
        var userDisplayName = this.props.userDisplayName || 'Priya';
        var totalDocs = rows.length > 0 ? rows.length : 156;
        var expiringSoon = 23;
        var pendingApprovals = 7;
        var totalVessels = vessels.length;
        return (<div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
        {/* Welcome Banner */}
        <div style={{
                background: 'linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%)',
                borderRadius: 12, padding: '24px 32px', display: 'flex', justifyContent: 'space-between',
                alignItems: 'center', border: '1px solid #bfdbfe', position: 'relative', overflow: 'hidden'
            }}>
          <div>
            <h1 style={{ margin: 0, fontSize: 22, fontWeight: 700, color: '#1e3a8a' }}>
              Welcome back, {userDisplayName}! 👋
            </h1>
            <p style={{ margin: '6px 0 0', fontSize: 13, color: '#3b82f6' }}>
              Here's what's happening with your vessels and documents.
            </p>
          </div>
          {/* Real Vessel Graphic Banner */}
          <div style={{ height: 85, width: 200, borderRadius: 8, overflow: 'hidden', flexShrink: 0, boxShadow: '0 4px 12px rgba(37,99,235,0.15)' }}>
            <img src={(0, vesselImagePool_1.resolveImgUrl)(null)} alt="Vessel Illustration" style={{ width: '100%', height: '100%', objectFit: 'cover' }}/>
          </div>
        </div>

        {/* 4 Stat Cards */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: 16 }}>
          {/* Stat 1 */}
          <div style={{ background: '#fff', borderRadius: 10, padding: 20, border: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <div>
              <div style={{ fontSize: 28, fontWeight: 800, color: '#0f172a' }}>{totalDocs}</div>
              <div style={{ fontSize: 12, color: '#64748b', fontWeight: 500, marginTop: 2 }}>Total Documents</div>
              <button onClick={function () { return _this._goToView('list'); }} style={{ background: 'none', border: 'none', color: '#2563eb', fontSize: 11, fontWeight: 600, padding: 0, marginTop: 8, cursor: 'pointer' }}>View all →</button>
            </div>
            <div style={{ width: 44, height: 44, borderRadius: '50%', background: '#f0f9ff', color: '#0284c7', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20 }}>📄</div>
          </div>

          {/* Stat 2 */}
          <div style={{ background: '#fff', borderRadius: 10, padding: 20, border: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <div>
              <div style={{ fontSize: 28, fontWeight: 800, color: '#0f172a' }}>{expiringSoon}</div>
              <div style={{ fontSize: 12, color: '#64748b', fontWeight: 500, marginTop: 2 }}>Expiring Soon</div>
              <button onClick={function () { return _this._goToView('list'); }} style={{ background: 'none', border: 'none', color: '#2563eb', fontSize: 11, fontWeight: 600, padding: 0, marginTop: 8, cursor: 'pointer' }}>View all →</button>
            </div>
            <div style={{ width: 44, height: 44, borderRadius: '50%', background: '#ecfdf5', color: '#10b981', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20 }}>✅</div>
          </div>

          {/* Stat 3 */}
          <div style={{ background: '#fff', borderRadius: 10, padding: 20, border: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <div>
              <div style={{ fontSize: 28, fontWeight: 800, color: '#0f172a' }}>0{pendingApprovals}</div>
              <div style={{ fontSize: 12, color: '#64748b', fontWeight: 500, marginTop: 2 }}>Pending Approvals</div>
              <button onClick={function () { return _this._goToView('approvals'); }} style={{ background: 'none', border: 'none', color: '#2563eb', fontSize: 11, fontWeight: 600, padding: 0, marginTop: 8, cursor: 'pointer' }}>View all →</button>
            </div>
            <div style={{ width: 44, height: 44, borderRadius: '50%', background: '#fffbeb', color: '#f59e0b', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20 }}>⏰</div>
          </div>

          {/* Stat 4 */}
          <div style={{ background: '#fff', borderRadius: 10, padding: 20, border: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <div>
              <div style={{ fontSize: 28, fontWeight: 800, color: '#0f172a' }}>{totalVessels}</div>
              <div style={{ fontSize: 12, color: '#64748b', fontWeight: 500, marginTop: 2 }}>Total Vessels</div>
              <button onClick={function () { return _this._goToView('vessels'); }} style={{ background: 'none', border: 'none', color: '#2563eb', fontSize: 11, fontWeight: 600, padding: 0, marginTop: 8, cursor: 'pointer' }}>View all →</button>
            </div>
            <div style={{ width: 44, height: 44, borderRadius: '50%', background: '#eff6ff', color: '#2563eb', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20 }}>🚢</div>
          </div>
        </div>

        {/* Main Dashboard Grid: Recent Docs + Expiry Overview */}
        <div style={{ display: 'grid', gridTemplateColumns: '2fr 1fr', gap: 20 }}>
          {/* Recent Documents Table */}
          <div style={{ background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', padding: 20 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
              <h3 style={{ margin: 0, fontSize: 15, fontWeight: 700, color: '#0f172a' }}>Recent Documents</h3>
              <button onClick={function () { return _this._goToView('list'); }} style={{ background: 'none', border: 'none', color: '#2563eb', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>View all</button>
            </div>

            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
              <thead>
                <tr style={{ borderBottom: '1px solid #e2e8f0', textTransform: 'uppercase', fontSize: 11, color: '#64748b', textAlign: 'left' }}>
                  <th style={{ padding: '8px 12px' }}>Name</th>
                  <th style={{ padding: '8px 12px' }}>Vessel</th>
                  <th style={{ padding: '8px 12px' }}>Type</th>
                  <th style={{ padding: '8px 12px' }}>Modified</th>
                </tr>
              </thead>
              <tbody>
                {documentsList.slice(3, 7).map(function (doc, idx) { return (<tr key={doc.id} style={{ borderBottom: '1px solid #f1f5f9' }}>
                    <td style={{ padding: '10px 12px', fontWeight: 600, color: '#1e293b', display: 'flex', alignItems: 'center', gap: 8 }}>
                      <span style={{ color: doc.name.endsWith('.pdf') ? '#ef4444' : doc.name.endsWith('.docx') ? '#2563eb' : '#10b981' }}>📄</span>
                      {doc.name}
                    </td>
                    <td style={{ padding: '10px 12px', color: '#475569' }}>{doc.vessel}</td>
                    <td style={{ padding: '10px 12px' }}>
                      {badge(doc.type === 'Insurance' ? 'purple' : doc.type === 'Crew' ? 'blue' : doc.type === 'Maintenance' ? 'green' : 'orange', doc.type)}
                    </td>
                    <td style={{ padding: '10px 12px', color: '#64748b', fontSize: 12 }}>{doc.modified}</td>
                  </tr>); })}
              </tbody>
            </table>
          </div>

          {/* Document Expiry Overview Chart */}
          <div style={{ background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', padding: 20 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
              <h3 style={{ margin: 0, fontSize: 15, fontWeight: 700, color: '#0f172a' }}>Document Expiry Overview</h3>
              <button onClick={function () { return _this._goToView('reports'); }} style={{ background: 'none', border: 'none', color: '#2563eb', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>View all</button>
            </div>

            {/* SVG Donut Chart */}
            <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 16 }}>
              <div style={{ position: 'relative', width: 140, height: 140 }}>
                <svg width="140" height="140" viewBox="0 0 42 42">
                  <circle cx="21" cy="21" r="15.91549430918954" fill="transparent" stroke="#e2e8f0" strokeWidth="5"/>
                  {/* Valid 67% (Green) */}
                  <circle cx="21" cy="21" r="15.91549430918954" fill="transparent" stroke="#10b981" strokeWidth="5" strokeDasharray="67 33" strokeDashoffset="25"/>
                  {/* Expiring in 30 days 24% (Orange) */}
                  <circle cx="21" cy="21" r="15.91549430918954" fill="transparent" stroke="#f59e0b" strokeWidth="5" strokeDasharray="24 76" strokeDashoffset="58"/>
                  {/* Expired 9% (Red) */}
                  <circle cx="21" cy="21" r="15.91549430918954" fill="transparent" stroke="#ef4444" strokeWidth="5" strokeDasharray="9 91" strokeDashoffset="34"/>
                </svg>
                <div style={{ position: 'absolute', inset: 0, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center' }}>
                  <span style={{ fontSize: 20, fontWeight: 800, color: '#0f172a' }}>96</span>
                  <span style={{ fontSize: 10, color: '#64748b', textTransform: 'uppercase' }}>Total</span>
                </div>
              </div>

              {/* Legend */}
              <div style={{ width: '100%', display: 'flex', flexDirection: 'column', gap: 8, fontSize: 12 }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    <span style={{ width: 10, height: 10, borderRadius: '50%', background: '#ef4444' }}/>
                    <span style={{ color: '#475569' }}>Expired</span>
                  </div>
                  <span style={{ fontWeight: 600, color: '#0f172a' }}>8 (9%)</span>
                </div>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    <span style={{ width: 10, height: 10, borderRadius: '50%', background: '#f59e0b' }}/>
                    <span style={{ color: '#475569' }}>Expiring in 30 days</span>
                  </div>
                  <span style={{ fontWeight: 600, color: '#0f172a' }}>23 (24%)</span>
                </div>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    <span style={{ width: 10, height: 10, borderRadius: '50%', background: '#10b981' }}/>
                    <span style={{ color: '#475569' }}>Valid</span>
                  </div>
                  <span style={{ fontWeight: 600, color: '#0f172a' }}>65 (67%)</span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>);
    };
    // ── Module 2: Documents ───────────────────────────────────────────────────
    VesselEmail.prototype._renderDocumentsPage = function () {
        var _this = this;
        var _a = this.state, textFilter = _a.textFilter, vesselFilter = _a.vesselFilter, catFilter = _a.catFilter, docViewMode = _a.docViewMode, showAllVesselsInFolderView = _a.showAllVesselsInFolderView, vessels = _a.vessels, rows = _a.rows, docListPage = _a.docListPage, docListSort = _a.docListSort, docGroupFilter = _a.docGroupFilter, docUploadRowKey = _a.docUploadRowKey, docUploadBusy = _a.docUploadBusy, docUploadMsg = _a.docUploadMsg, documentsList = _a.documentsList, folderPathStack = _a.folderPathStack, uploadedFilesByFolder = _a.uploadedFilesByFolder, docMainFolder = _a.docMainFolder;
        var PAGE_ROWS = 10;
        var MAIN_FOLDERS = [
            { key: 'Technical & Crewing', icon: '⚙️', emoji: '⚙️', color: '#dc2626', bg: '#fee2e2' },
            { key: 'Commercial & Chartering', icon: '💼', emoji: '💼', color: '#16a34a', bg: '#dcfce7' },
            { key: 'Insurance', icon: '🛡️', emoji: '🛡️', color: '#d97706', bg: '#fef3c7' },
            { key: 'Knowledge Bank', icon: '📚', emoji: '📚', color: '#7c3aed', bg: '#ede9fe' },
        ];
        // ── Per main-folder subfolder structure (aligned with vesselFolderTemplate.ts) ──
        var DEFAULT_VESSEL_MAINS = {
            'Technical & Crewing': ['Month End Reports', 'Service Agreements', 'Registration', 'Drawings and Manuals', 'PO & Invoice', 'Incidents', 'Crewing', 'To be Classified'],
            'Commercial & Chartering': ['Agreements', 'Invoices & Payments', 'Claims & Disputes', 'To be Classified'],
            'Insurance': ['P&I', 'H&M', 'War Risk', 'Flag and MPA'],
            'Knowledge Bank': ['Circulars', 'Manuals & Procedures', 'Regulations', 'Company Policies', 'Training Materials']
        };
        var SUBFOLDERS_MAP = {
            'Month End Reports': ['Month Wise Folder'],
            'Drawings and Manuals': [
                'Flag & MPA', 'Ship Builder', 'Rules & Telecom', 'Crewing & SMOU', 'Novation', 'Automation and To be Classified',
                'Drawings', 'Manuals',
            ],
            'Drawings': ['Archive', 'Basic', 'Electrical', 'Engine', 'Hull', 'Other Drawings', 'Safety'],
            'Manuals': ['Automation', 'Auxiliary Engine', 'Boiler', 'Cargo', 'Deck Machinery', 'Electrical', 'Main Engine', 'Other Manuals', 'Pollution', 'Propulsion', 'Refrigeration', 'Safety', 'Shafting', 'Steering Gear', 'Thrusters'],
            'Invoices & Payments': ['Month Wise Folder'],
            'Claims & Disputes': ['Month Wise Folder'],
            'Common for all ships': ['Vendor & Service Agreements', 'Vendor Management'],
            'Common Agreements (Not Ship Specific)': ['Agreements', 'To be Classified'],
            'Common (Not Ship Specific)': ['Agreements', 'Miscellaneous']
        };
        var currentFolderNode = folderPathStack.length > 0 ? folderPathStack[folderPathStack.length - 1] : null;
        var currentFolderName = currentFolderNode ? currentFolderNode.name : null;
        // Determine subfolders for current depth
        var subfolderNames = currentFolderName
            ? (SUBFOLDERS_MAP[currentFolderName] || (
            // If we are at vessel-level depth (stack[0] is a vessel), show main folder's default top-level folders
            folderPathStack.length === 1 && docMainFolder
                ? DEFAULT_VESSEL_MAINS[docMainFolder]
                : []))
            : [];
        var currentFolderFiles = currentFolderName ? (uploadedFilesByFolder[currentFolderName] || []).filter(function (f) { return !f.pending; }) : [];
        // Also include approved files from backend rows that match the current folder node
        var backendFolderFiles = [];
        if (currentFolderNode && !/^(sf_|common)/.test(currentFolderNode.id)) {
            rows.filter(function (r) { return r.uploadFolderId === currentFolderNode.id && r.fileName && !r.filePending; })
                .forEach(function (r) { return backendFolderFiles.push({ name: r.fileName, size: '—', date: '—' }); });
        }
        var allCurrentFolderFiles = __spreadArray(__spreadArray([], backendFolderFiles, true), currentFolderFiles.filter(function (f) { return !backendFolderFiles.some(function (b) { return b.name === f.name; }); }), true);
        // ── Build source rows from API only — no mock fallback ──
        var allRows = rows || [];
        // Subfolders lookup map for Folder View levels
        // (subfolder resolution already computed above via SUBFOLDERS_MAP / DEFAULT_VESSEL_MAINS)
        // Filtered list rows for List view
        var activeVesselName = vesselFilter !== 'all' ? vesselFilter : null;
        var allGroups = Array.from(new Set(allRows.map(function (r) { return r.group; }))).sort();
        var allCategories = Array.from(new Set(allRows.filter(function (r) { return docGroupFilter === 'all' || r.group === docGroupFilter; }).map(function (r) { return r.category; }))).sort();
        // ── mainFolderGroupMap: which groups belong to which main folder (for list view filtering) ──
        // These must match the actual template group names returned by the backend API.
        var mainFolderGroupMap = {
            'Technical & Crewing': [
                'Month End Reports', 'Service Agreements', 'Registration',
                'Drawings and Manuals', 'PO & Invoice', 'Incidents', 'Crewing', 'To be Classified',
            ],
            'Commercial & Chartering': [
                'Agreements', 'Invoices & Payments', 'Claims & Disputes', 'To be Classified',
            ],
            'Insurance': ['P&I', 'H&M', 'War Risk', 'Flag - MPA', 'USA Related', 'Flag and MPA'],
            'Kaizen - Knowledge Bank': [
                'Templates', 'Procedures and Work Instructions', 'Lessons Learned', 'Circulars and Guidance',
            ],
            'Knowledge Bank': [
                'Templates', 'Procedures and Work Instructions', 'Lessons Learned', 'Circulars and Guidance',
            ]
        };
        var filtered = allRows.filter(function (r) {
            // Filter by active main folder using template-based group membership
            if (docMainFolder) {
                var allowedGroups = mainFolderGroupMap[docMainFolder];
                // If allowedGroups is defined, filter to only matching groups.
                // If the group is not in the map (e.g. flat folder or unknown), still show all rows.
                if (allowedGroups && allowedGroups.length > 0) {
                    if (!allowedGroups.includes(r.group) && !r.subFolderPath.toLowerCase().includes(docMainFolder.toLowerCase().split(' ')[0].toLowerCase())) {
                        return false;
                    }
                }
            }
            if (activeVesselName && r.vesselName !== activeVesselName)
                return false;
            if (docGroupFilter !== 'all' && r.group !== docGroupFilter)
                return false;
            if (catFilter !== 'all' && r.category !== catFilter && r.group !== catFilter)
                return false;
            if (textFilter) {
                var q = textFilter.toLowerCase();
                return r.vesselName.toLowerCase().includes(q) || r.group.toLowerCase().includes(q) ||
                    r.category.toLowerCase().includes(q) || (r.subFolderPath || '').toLowerCase().includes(q) ||
                    (r.fileName || '').toLowerCase().includes(q);
            }
            return true;
        });
        // ── Group filtered rows by groupKey (category/folder level) ──
        var groupedMap = new Map();
        var _loop_1 = function (r) {
            var folderUploads = (uploadedFilesByFolder[r.groupKey] || [])
                .concat(uploadedFilesByFolder[r.uploadFolderId] || [])
                .concat(uploadedFilesByFolder[r.category] || []);
            var existing = groupedMap.get(r.groupKey);
            if (!existing) {
                var files = [];
                if (r.fileName) {
                    files.push({ id: r.fileId || r.fileName, name: r.fileName });
                }
                var _loop_2 = function (f) {
                    if (!files.some(function (ex) { return ex.name === f.name; })) {
                        files.push({ id: f.id || f.name, name: f.name });
                    }
                };
                for (var _b = 0, folderUploads_1 = folderUploads; _b < folderUploads_1.length; _b++) {
                    var f = folderUploads_1[_b];
                    _loop_2(f);
                }
                groupedMap.set(r.groupKey, {
                    srNo: r.srNo,
                    vesselName: r.vesselName,
                    group: r.group,
                    category: r.category,
                    subFolderPath: r.subFolderPath,
                    groupKey: r.groupKey,
                    uploadFolderId: r.uploadFolderId,
                    monthDriven: r.monthDriven,
                    canUpload: r.canUpload,
                    files: files
                });
            }
            else {
                if (r.fileName && !existing.files.some(function (f) { return f.name === r.fileName; })) {
                    existing.files.push({ id: r.fileId || r.fileName, name: r.fileName });
                }
                var _loop_3 = function (f) {
                    if (!existing.files.some(function (ex) { return ex.name === f.name; })) {
                        existing.files.push({ id: f.id || f.name, name: f.name });
                    }
                };
                for (var _c = 0, folderUploads_2 = folderUploads; _c < folderUploads_2.length; _c++) {
                    var f = folderUploads_2[_c];
                    _loop_3(f);
                }
            }
        };
        for (var _i = 0, filtered_1 = filtered; _i < filtered_1.length; _i++) {
            var r = filtered_1[_i];
            _loop_1(r);
        }
        var groupedList = Array.from(groupedMap.values());
        if (docListSort === 'name_az')
            groupedList.sort(function (a, b) { return a.category.localeCompare(b.category); });
        else if (docListSort === 'newest')
            groupedList.reverse();
        var totalPages = Math.max(1, Math.ceil(groupedList.length / PAGE_ROWS));
        var safePage = Math.min(docListPage, totalPages - 1);
        var pageGroupedRows = groupedList.slice(safePage * PAGE_ROWS, (safePage + 1) * PAGE_ROWS);
        var displayVessels = showAllVesselsInFolderView ? vessels : vessels.slice(0, 4);
        return (<div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        {/* Breadcrumb Navigation Trail */}
        <div style={{ fontSize: 12, color: '#64748b', display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
          <span style={{ cursor: 'pointer', color: '#0284c7' }} onClick={function () { return _this.setState({ docMainFolder: null, folderPathStack: [] }); }}>Home</span>
          <span>›</span>
          <span style={{ cursor: 'pointer', color: !docMainFolder ? '#0f172a' : '#0284c7', fontWeight: !docMainFolder ? 600 : 400 }} onClick={function () { return _this.setState({ docMainFolder: null, folderPathStack: [] }); }}>
            Documents
          </span>
          {docMainFolder && (<>
              <span>›</span>
              <span style={{ cursor: folderPathStack.length === 0 ? 'default' : 'pointer', color: folderPathStack.length === 0 ? '#0f172a' : '#0284c7', fontWeight: folderPathStack.length === 0 ? 600 : 400 }} onClick={function () { return _this.setState({ folderPathStack: [] }); }}>
                {docMainFolder}
              </span>
            </>)}
          {folderPathStack.map(function (item, idx) {
                var isLast = idx === folderPathStack.length - 1;
                return (<React.Fragment key={item.id + idx}>
                <span>›</span>
                <span onClick={function () { return _this.setState({ folderPathStack: folderPathStack.slice(0, idx + 1) }); }} style={{ cursor: isLast ? 'default' : 'pointer', color: isLast ? '#0f172a' : '#0284c7', fontWeight: isLast ? 600 : 400 }}>
                  {item.name}
                </span>
              </React.Fragment>);
            })}
        </div>

        {/* Module Header */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 12 }}>
          <div>
            <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: '#0f172a', display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ fontSize: 20 }}>📁</span>
              {currentFolderName || docMainFolder || 'Documents'}
            </h2>
            <p style={{ margin: '4px 0 0', fontSize: 13, color: '#64748b' }}>
              {docViewMode === 'list'
                ? "".concat(filtered.length, " rows \u00B7 flattened list view")
                : !docMainFolder
                    ? "".concat(MAIN_FOLDERS.length, " main folders")
                    : folderPathStack.length === 0
                        ? "".concat(vessels.length, " vessels \u00B7 folder view")
                        : "".concat(subfolderNames.length, " folders \u00B7 ").concat(allCurrentFolderFiles.length, " files")}
            </p>
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            {/* Folder view / List view Pill Toggle */}
            <div style={{ display: 'inline-flex', background: '#fff', border: '1px solid #cbd5e1', borderRadius: 8, padding: 3, gap: 2 }}>
              <button onClick={function () { return _this.setState({ docViewMode: 'folder' }); }} style={{
                padding: '5px 12px', borderRadius: 6, border: 'none', fontSize: 12, fontWeight: 600, cursor: 'pointer',
                background: docViewMode === 'folder' ? '#0f172a' : 'transparent',
                color: docViewMode === 'folder' ? '#fff' : '#64748b'
            }}>
                ⊞ Folder view
              </button>
              <button onClick={function () { return _this.setState({ docViewMode: 'list' }); }} style={{
                padding: '5px 12px', borderRadius: 6, border: 'none', fontSize: 12, fontWeight: 600, cursor: 'pointer',
                background: docViewMode === 'list' ? '#0f172a' : 'transparent',
                color: docViewMode === 'list' ? '#fff' : '#64748b'
            }}>
                ☰ List view
              </button>
            </div>

            {/* Archive Button */}
            <button onClick={function () { return _this._goToView('archive'); }} style={{
                background: '#d97706', color: '#fff', border: 'none', borderRadius: 8,
                padding: '7px 16px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
                display: 'inline-flex', alignItems: 'center', gap: 6
            }}>
              🗑 Archive
            </button>

            {/* Top-Right Upload Button */}
            <label style={{
                background: '#0284c7', color: '#fff', border: 'none', borderRadius: 8,
                padding: '7px 16px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
                display: 'inline-flex', alignItems: 'center', gap: 6, boxShadow: '0 2px 4px rgba(2,132,199,0.2)'
            }}>
              <input type="file" style={{ display: 'none' }} onChange={function (e) { return __awaiter(_this, void 0, void 0, function () {
                var file, folderKey, topFolderId, currentVessel, subFolderPath, mainFolderPathMap, spoMainFolder, matchingRow, resolvedFolderId, _a, fileId, statusPending, msg, newUpload, existing, err_3;
                var _b;
                var _c;
                return __generator(this, function (_d) {
                    switch (_d.label) {
                        case 0:
                            file = (_c = e.target.files) === null || _c === void 0 ? void 0 : _c[0];
                            if (!file)
                                return [2 /*return*/];
                            folderKey = currentFolderName || docMainFolder || 'Documents';
                            topFolderId = currentFolderNode && !/^(sf_|common)/.test(currentFolderNode.id) ? currentFolderNode.id : '';
                            currentVessel = folderPathStack.length > 0 ? folderPathStack[0].name : (vessels.length > 0 ? vessels[0].name : 'Bow Fighter');
                            subFolderPath = currentFolderNode ? "".concat(currentVessel, " > ").concat(docMainFolder || '', " > ").concat(currentFolderNode.name) : "".concat(currentVessel, " > ").concat(docMainFolder || '');
                            mainFolderPathMap = {
                                'Technical & Crewing': 'Technical & Crewing',
                                'Commercial & Chartering': 'Commercial & Chartering',
                                'Insurance': 'Insurance',
                                'Kaizen - Knowledge Bank': 'Kaizen - Knowledge Bank',
                                'Knowledge Bank': 'Kaizen - Knowledge Bank'
                            };
                            spoMainFolder = (docMainFolder && mainFolderPathMap[docMainFolder]) || 'Technical & Crewing';
                            matchingRow = this.state.rows.find(function (r) {
                                return r.vesselName === currentVessel &&
                                    r.uploadFolderId &&
                                    !r.uploadFolderId.includes('/') &&
                                    (docMainFolder ? r.group.toLowerCase().includes(docMainFolder.toLowerCase().split(' ')[0]) : true);
                            });
                            resolvedFolderId = topFolderId || (matchingRow === null || matchingRow === void 0 ? void 0 : matchingRow.uploadFolderId) || "".concat(this.VESSEL_ROOT, "/").concat(spoMainFolder, "/").concat(currentVessel);
                            this.setState({ docUploadMsg: null });
                            _d.label = 1;
                        case 1:
                            _d.trys.push([1, 3, , 4]);
                            return [4 /*yield*/, this._uploadFileToFolder(resolvedFolderId, subFolderPath, currentVessel, file)];
                        case 2:
                            _a = _d.sent(), fileId = _a.fileId, statusPending = _a.statusPending;
                            msg = statusPending
                                ? "\"".concat(file.name, "\" submitted for approval.")
                                : "\"".concat(file.name, "\" uploaded successfully to ").concat(folderKey, "!");
                            newUpload = { name: file.name, size: "".concat((file.size / 1024).toFixed(1), " KB"), date: 'Just now', pending: statusPending, id: fileId };
                            existing = uploadedFilesByFolder[folderKey] || [];
                            this.setState({
                                uploadedFilesByFolder: __assign(__assign({}, uploadedFilesByFolder), (_b = {}, _b[folderKey] = __spreadArray(__spreadArray([], existing.filter(function (f) { return f.name !== file.name; }), true), [newUpload], false), _b)),
                                docUploadMsg: msg
                            });
                            return [3 /*break*/, 4];
                        case 3:
                            err_3 = _d.sent();
                            this.setState({ docUploadMsg: "Upload failed: ".concat((err_3 === null || err_3 === void 0 ? void 0 : err_3.message) || 'Error') });
                            return [3 /*break*/, 4];
                        case 4: return [2 /*return*/];
                    }
                });
            }); }}/>
              <span>⬆</span> Upload
            </label>
          </div>
        </div>

        {/* Filter Toolbar */}
        <div style={{ background: '#fff', borderRadius: 10, padding: '10px 12px', border: '1px solid #e2e8f0', display: 'flex', flexWrap: 'wrap', gap: 8, alignItems: 'center' }}>
          <div style={{ position: 'relative', flex: '1 1 180px', minWidth: 160 }}>
            <span style={{ position: 'absolute', left: 9, top: '50%', transform: 'translateY(-50%)', color: '#94a3b8', fontSize: 12 }}>🔍</span>
            <input type="text" placeholder="Filter by vessel, group, category, path..." value={textFilter} onChange={function (e) { return _this.setState({ textFilter: e.target.value, docListPage: 0 }); }} style={{ width: '100%', padding: '6px 10px 6px 28px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 12, outline: 'none', boxSizing: 'border-box' }}/>
          </div>
          <select value={vesselFilter} onChange={function (e) { return _this.setState({ vesselFilter: e.target.value, docListPage: 0 }); }} style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 12, background: '#fff', outline: 'none', maxWidth: 140 }}>
            <option value="all">All vessels</option>
            {vessels.map(function (v) { return (<option key={v.id || v.name} value={v.name}>{v.name}</option>); })}
          </select>
          <select value={docGroupFilter} onChange={function (e) { return _this.setState({ docGroupFilter: e.target.value, catFilter: 'all', docListPage: 0 }); }} style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 12, background: '#fff', outline: 'none', maxWidth: 140 }}>
            <option value="all">All groups</option>
            {allGroups.map(function (g) { return (<option key={g} value={g}>{g}</option>); })}
          </select>
          <select value={catFilter} onChange={function (e) { return _this.setState({ catFilter: e.target.value, docListPage: 0 }); }} style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 12, background: '#fff', outline: 'none', maxWidth: 140 }}>
            <option value="all">All categories</option>
            {allCategories.map(function (c) { return (<option key={c} value={c}>{c}</option>); })}
          </select>
          <select value={docListSort} onChange={function (e) { return _this.setState({ docListSort: e.target.value, docListPage: 0 }); }} style={{ padding: '6px 10px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 12, background: '#fff', outline: 'none' }}>
            <option value="default">Default order</option>
            <option value="name_az">Name A–Z</option>
            <option value="newest">Newest</option>
          </select>
          <div style={{ display: 'flex', border: '1px solid #cbd5e1', borderRadius: 8, overflow: 'hidden', marginLeft: 'auto' }}>
            <button onClick={function () { return _this.setState({ docViewMode: 'folder' }); }} style={{ padding: '5px 10px', background: docViewMode === 'folder' ? '#e2e8f0' : '#fff', border: 'none', cursor: 'pointer', fontSize: 13 }} title="Folder view">::</button>
            <button onClick={function () { return _this.setState({ docViewMode: 'list' }); }} style={{ padding: '5px 10px', background: docViewMode === 'list' ? '#e2e8f0' : '#fff', border: 'none', borderLeft: '1px solid #cbd5e1', cursor: 'pointer', fontSize: 13 }} title="List view">☰</button>
          </div>
        </div>

        {/* Success Message Banner */}
        {docUploadMsg && (<div style={{ background: '#f0fdf4', border: '1px solid #86efac', borderRadius: 8, padding: '8px 14px', fontSize: 12, color: '#16a34a', display: 'flex', justifyContent: 'space-between' }}>
            <span>✓ {docUploadMsg}</span>
            <button onClick={function () { return _this.setState({ docUploadMsg: null }); }} style={{ border: 'none', background: 'none', cursor: 'pointer', color: '#16a34a', fontWeight: 700 }}>✕</button>
          </div>)}

        {/* View Mode Content */}
        {docViewMode === 'folder' ? (<div style={{ display: 'flex', flexDirection: 'column', gap: 24, marginTop: 4 }}>

            {/* ── Level 0: Four Main Folders ── */}
            {!docMainFolder ? (<div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: 16 }}>
                {MAIN_FOLDERS.map(function (mf) { return (<div key={mf.key} onClick={function () { return _this.setState({ docMainFolder: mf.key, folderPathStack: [], vesselFilter: 'all' }); }} style={{
                            background: '#fff', borderRadius: 14, border: '1px solid #e2e8f0', padding: 18,
                            display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer',
                            boxShadow: '0 1px 3px rgba(0,0,0,0.05)', transition: 'transform 0.15s, box-shadow 0.15s'
                        }}>
                    <div style={{ width: 44, height: 44, borderRadius: 10, background: mf.bg, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 22 }}>
                      {mf.emoji}
                    </div>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontWeight: 700, fontSize: 14, color: '#0f172a', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{mf.key}</div>
                      <div style={{ fontSize: 12, color: '#64748b', marginTop: 2 }}>{vessels.length} vessels</div>
                    </div>
                    <span style={{ color: '#94a3b8', fontSize: 16 }}>›</span>
                  </div>); })}
              </div>) : folderPathStack.length === 0 ? (
                /* ── Level 1: Vessel list inside the selected main folder ── */
                <>
                {/* Common Documents section */}
                <div>
                  <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: '0.05em', color: '#94a3b8', textTransform: 'uppercase', marginBottom: 12 }}>COMMON AGREEMENTS / DOCUMENTS</div>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 340px))', gap: 16 }}>
                    <div onClick={function () { return _this.setState({ folderPathStack: [{ id: 'common', name: 'Common for all ships' }] }); }} style={{ background: '#fff', borderRadius: 14, border: '1px solid #e2e8f0', padding: 18, display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer' }}>
                      <div style={{ width: 44, height: 44, borderRadius: 10, background: '#fef3c7', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20, color: '#d97706' }}>📁</div>
                      <div style={{ flex: 1 }}>
                        <div style={{ fontWeight: 700, fontSize: 14, color: '#0f172a' }}>Common for all ships</div>
                        <div style={{ fontSize: 12, color: '#64748b', marginTop: 2 }}>Common</div>
                      </div>
                      <span style={{ color: '#94a3b8', fontSize: 16 }}>›</span>
                    </div>
                  </div>
                </div>

                {/* Vessels section */}
                <div>
                  <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: '0.05em', color: '#94a3b8', textTransform: 'uppercase', marginBottom: 12 }}>VESSELS</div>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: 16 }}>
                    {displayVessels.map(function (v) { return (<div key={v.id} onClick={function () {
                            _this.setState({ vesselFilter: v.name, folderPathStack: [{ id: v.id, name: v.name }] });
                            // Pre-load files for this vessel so approved files show immediately
                            if (!_this._filesLoadedForVessels.has(v.name)) {
                                _this._filesLoadedForVessels.add(v.name);
                                _this._loadFilesForVessel(v.name)["catch"](function () { return undefined; });
                            }
                        }} style={{ background: '#fff', borderRadius: 14, border: '1px solid #e2e8f0', padding: 18, display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer' }}>
                        <div style={{ width: 44, height: 44, borderRadius: 10, background: '#e0f2fe', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20, color: '#0284c7' }}>🚢</div>
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ fontWeight: 700, fontSize: 14, color: '#0f172a', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{v.name}</div>
                          <div style={{ fontSize: 12, color: '#64748b', marginTop: 2 }}>Vessel</div>
                        </div>
                        <span style={{ color: '#94a3b8', fontSize: 16 }}>›</span>
                      </div>); })}
                  </div>
                  {vessels.length > 4 && (<div style={{ display: 'flex', justifyContent: 'center', marginTop: 20 }}>
                      <button onClick={function () { return _this.setState({ showAllVesselsInFolderView: !showAllVesselsInFolderView }); }} style={{ background: '#fff', border: '1px solid #cbd5e1', borderRadius: 20, padding: '8px 20px', fontSize: 13, fontWeight: 600, color: '#334155', cursor: 'pointer' }}>
                        {showAllVesselsInFolderView ? 'Show Less ⌃' : "More Vessels (".concat(vessels.length - 4, ") \u2304")}
                      </button>
                    </div>)}
                </div>
              </>) : subfolderNames.length > 0 ? (
                /* Subfolders Grid (Screenshot 1) */
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(260px, 1fr))', gap: 16 }}>
                {subfolderNames.map(function (sfName, idx) { return (<div key={sfName + idx} onClick={function () {
                            var newStack = __spreadArray(__spreadArray([], folderPathStack, true), [{ id: "sf_".concat(idx), name: sfName }], false);
                            _this.setState({ folderPathStack: newStack });
                            // If the current folder node has a real backend ID, refresh its files
                            if (currentFolderNode && !/^(sf_|common)/.test(currentFolderNode.id)) {
                                _this._refreshFolderFiles(currentFolderNode.id, currentFolderNode.id)["catch"](function () { return undefined; });
                            }
                        }} style={{
                            background: '#fff', borderRadius: 14, border: '1px solid #e2e8f0', padding: 18,
                            display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer',
                            boxShadow: '0 1px 3px rgba(0,0,0,0.05)', transition: 'transform 0.15s, box-shadow 0.15s'
                        }}>
                    <div style={{
                            width: 44, height: 44, borderRadius: 10, background: '#e0f2fe',
                            display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20, color: '#0284c7'
                        }}>
                      📁
                    </div>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontWeight: 700, fontSize: 14, color: '#0f172a', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{sfName}</div>
                      <div style={{ fontSize: 12, color: '#64748b', marginTop: 2 }}>Folder</div>
                    </div>
                    <span style={{ color: '#94a3b8', fontSize: 16 }}>›</span>
                  </div>); })}
              </div>) : allCurrentFolderFiles.length > 0 ? (
                /* Folder File Items List */
                <div style={{ background: '#fff', borderRadius: 12, border: '1px solid #e2e8f0', overflow: 'hidden' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                  <thead>
                    <tr style={{ background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 11, fontWeight: 700, textTransform: 'uppercase', textAlign: 'left' }}>
                      <th style={{ padding: '10px 16px' }}>FILE NAME</th>
                      <th style={{ padding: '10px 16px' }}>SIZE</th>
                      <th style={{ padding: '10px 16px' }}>DATE UPLOADED</th>
                      <th style={{ padding: '10px 16px', textAlign: 'right' }}>ACTION</th>
                    </tr>
                  </thead>
                  <tbody>
                    {allCurrentFolderFiles.map(function (file, idx) { return (<tr key={file.name + idx} style={{ borderBottom: '1px solid #f1f5f9' }}>
                        <td style={{ padding: '12px 16px', fontWeight: 600, color: '#0f172a', display: 'flex', alignItems: 'center', gap: 10 }}>
                          <span style={{ fontSize: 18 }}>📄</span> {file.name}
                        </td>
                        <td style={{ padding: '12px 16px', color: '#64748b' }}>{file.size}</td>
                        <td style={{ padding: '12px 16px', color: '#64748b' }}>{file.date}</td>
                        <td style={{ padding: '12px 16px', textAlign: 'right' }}>
                          <button style={{ border: '1px solid #cbd5e1', background: '#fff', borderRadius: 6, padding: '4px 10px', fontSize: 11, fontWeight: 600, cursor: 'pointer', color: '#0078d4' }}>
                            View / Download
                          </button>
                        </td>
                      </tr>); })}
                  </tbody>
                </table>
              </div>) : (
                /* Empty Folder View (Screenshot 2) */
                <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', padding: '60px 20px', width: '100%' }}>
                <div style={{
                        background: 'rgba(240, 249, 255, 0.6)', border: '1px solid #e0f2fe',
                        borderRadius: 24, padding: '48px 40px', maxWidth: 500, width: '100%',
                        textAlign: 'center', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 12
                    }}>
                  <div style={{
                        width: 56, height: 56, borderRadius: 14, background: '#e0f2fe',
                        display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 24, color: '#0284c7',
                        marginBottom: 4
                    }}>
                    📁
                  </div>
                  <h3 style={{ margin: 0, fontSize: 18, fontWeight: 700, color: '#0f172a' }}>This folder is empty</h3>
                  <p style={{ margin: 0, fontSize: 13, color: '#64748b', maxWidth: 320, lineHeight: 1.5 }}>
                    Use the Upload button in the top-right to add a document.
                  </p>
                </div>
              </div>)}
          </div>) : (
            /* ── LIST VIEW ── */
            <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
            {/* Table wrapper */}
            <div style={{ background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', overflowX: 'auto', width: '100%' }}>
              <table style={{ width: '100%', minWidth: 850, borderCollapse: 'collapse', fontSize: 12 }}>
                <thead>
                  <tr style={{ background: '#f8fafc', borderBottom: '2px solid #e2e8f0', color: '#64748b', fontSize: 11, fontWeight: 700, textTransform: 'uppercase', textAlign: 'left' }}>
                    <th style={{ padding: '10px 10px', width: 44, textAlign: 'center' }}>SR.</th>
                    <th style={{ padding: '10px 12px' }}>VESSEL NAME</th>
                    <th style={{ padding: '10px 12px' }}>GROUP</th>
                    <th style={{ padding: '10px 12px' }}>SUB-CATEGORY</th>
                    <th style={{ padding: '10px 12px' }}>FOLDER PATH</th>
                    <th style={{ padding: '10px 12px' }}>FILE NAME</th>
                    <th style={{ padding: '10px 12px', textAlign: 'right' }}>ATTACHMENT</th>
                  </tr>
                </thead>
                <tbody>
                  {pageGroupedRows.length === 0 ? (<tr>
                      <td colSpan={7} style={{ padding: '36px 16px', textAlign: 'center', color: '#94a3b8', fontSize: 13 }}>
                        No documents found. {textFilter || vesselFilter !== 'all' || docGroupFilter !== 'all' || catFilter !== 'all' ? 'Try clearing the filters.' : ''}
                      </td>
                    </tr>) : pageGroupedRows.map(function (r, idx) {
                    var globalIdx = safePage * PAGE_ROWS + idx + 1;
                    var isUploading = docUploadRowKey === r.groupKey && docUploadBusy;
                    var hasFiles = r.files.length > 0;
                    return (<tr key={"".concat(r.groupKey, "-").concat(idx)} style={{ borderBottom: '1px solid #f1f5f9', transition: 'background 0.1s' }} onMouseEnter={function (e) { return (e.currentTarget.style.background = '#f8fafc'); }} onMouseLeave={function (e) { return (e.currentTarget.style.background = ''); }}>
                        <td style={{ padding: '10px 10px', color: '#94a3b8', fontSize: 11, fontFamily: 'monospace', textAlign: 'center' }}>{globalIdx}</td>
                        <td style={{ padding: '10px 12px', fontWeight: 700, color: '#0f172a' }}>{r.vesselName}</td>
                        <td style={{ padding: '10px 12px' }}>
                          <span style={{ display: 'inline-block', borderRadius: 8, padding: '3px 8px', fontSize: 11, fontWeight: 600, background: '#eff6ff', color: '#2563eb' }}>
                            {r.group}
                          </span>
                        </td>
                        <td style={{ padding: '10px 12px', fontWeight: 600, color: '#1e293b' }}>{r.category}</td>
                        <td style={{ padding: '10px 12px', color: '#64748b', fontSize: 11 }} title={r.subFolderPath}>
                          {r.subFolderPath}
                        </td>
                        <td style={{ padding: '10px 12px' }}>
                          {hasFiles ? (<div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                              {r.files.map(function (file) { return (<div key={file.name} style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                                  <span style={{ fontSize: 14 }}>📄</span>
                                  <span onClick={function () {
                                    if (file.id && !file.id.startsWith('file_')) {
                                        // Numeric IDs are approval DB row IDs (pending files) — backend
                                        // now resolves them to the staged drive_item_id for preview.
                                        window.open("".concat(_this._base(), "/api/files/").concat(file.id, "/content"), '_blank');
                                    }
                                    else {
                                        alert("File \"".concat(file.name, "\" is pending \u2014 it will be available after approval."));
                                    }
                                }} style={{ color: '#0284c7', textDecoration: 'underline', fontWeight: 600, cursor: 'pointer' }} title={file.id && /^\d+$/.test(file.id) ? "".concat(file.name, " (pending approval - click to preview staged copy)") : "Click to open ".concat(file.name)}>
                                    {file.name}{file.id && /^\d+$/.test(file.id) ? ' ⏳' : ''}
                                  </span>
                                </div>); })}
                            </div>) : (<span style={{ color: '#94a3b8', fontStyle: 'italic', fontSize: 11 }}>—</span>)}
                        </td>
                        <td style={{ padding: '10px 12px', textAlign: 'right' }}>
                          <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: 4 }}>
                            {hasFiles ? (<button onClick={function () {
                                var firstFile = r.files[0];
                                if (firstFile.id && !firstFile.id.startsWith('file_')) {
                                    window.open("".concat(_this._base(), "/api/files/").concat(firstFile.id, "/content"), '_blank');
                                }
                                else {
                                    alert("File \"".concat(firstFile.name, "\" in folder: ").concat(r.subFolderPath));
                                }
                            }} style={{
                                background: '#eff6ff', color: '#1d4ed8', border: '1px solid #bfdbfe',
                                borderRadius: 6, padding: '3px 8px', fontSize: 11, fontWeight: 600, cursor: 'pointer',
                                display: 'inline-flex', alignItems: 'center', gap: 4
                            }}>
                                ↗ Open ({r.files.length})
                              </button>) : (<span style={{ fontSize: 11, color: '#94a3b8', fontStyle: 'italic' }}>No attachment</span>)}

                            <label style={{
                            background: isUploading ? '#f1f5f9' : '#fff', border: '1px solid #cbd5e1', borderRadius: 6,
                            padding: '3px 8px', fontSize: 11, fontWeight: 600,
                            color: isUploading ? '#94a3b8' : '#334155',
                            cursor: isUploading ? 'not-allowed' : 'pointer',
                            display: 'inline-flex', alignItems: 'center', gap: 4, whiteSpace: 'nowrap'
                        }}>
                              <input type="file" style={{ display: 'none' }} disabled={isUploading} onChange={function (e) { return __awaiter(_this, void 0, void 0, function () {
                            var file, _a, fileId_1, statusPending_1, msg, newApproval_1, newUpload, updatedByFolder, baseRows, updatedRows, err_4;
                            var _b;
                            var _this = this;
                            var _c;
                            return __generator(this, function (_d) {
                                switch (_d.label) {
                                    case 0:
                                        file = (_c = e.target.files) === null || _c === void 0 ? void 0 : _c[0];
                                        if (!file)
                                            return [2 /*return*/];
                                        this.setState({ docUploadRowKey: r.groupKey, docUploadBusy: true, docUploadMsg: null });
                                        _d.label = 1;
                                    case 1:
                                        _d.trys.push([1, 3, , 4]);
                                        return [4 /*yield*/, this._uploadFileToFolder(r.uploadFolderId, r.subFolderPath, r.vesselName, file, r.monthDriven)];
                                    case 2:
                                        _a = _d.sent(), fileId_1 = _a.fileId, statusPending_1 = _a.statusPending;
                                        msg = statusPending_1
                                            ? "\"".concat(file.name, "\" submitted for approval.")
                                            : "\"".concat(file.name, "\" uploaded successfully!");
                                        if (statusPending_1) {
                                            newApproval_1 = {
                                                id: fileId_1 || "a_".concat(Date.now()),
                                                documentName: file.name,
                                                vessel: r.vesselName,
                                                requestedBy: this.props.userDisplayName || this.props.userEmail || 'You',
                                                requestedOn: new Date().toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }),
                                                status: 'Pending'
                                            };
                                            this.setState(function (prev) { return ({ approvalsList: __spreadArray([newApproval_1], prev.approvalsList, true) }); });
                                        }
                                        newUpload = { name: file.name, size: "".concat((file.size / 1024).toFixed(1), " KB"), date: 'Just now', pending: statusPending_1, id: fileId_1 };
                                        updatedByFolder = __assign(__assign({}, uploadedFilesByFolder), (_b = {}, _b[r.groupKey] = __spreadArray(__spreadArray([], (uploadedFilesByFolder[r.groupKey] || []).filter(function (f) { return f.name !== file.name; }), true), [newUpload], false), _b[r.category] = __spreadArray(__spreadArray([], (uploadedFilesByFolder[r.category] || []).filter(function (f) { return f.name !== file.name; }), true), [newUpload], false), _b[r.group] = __spreadArray(__spreadArray([], (uploadedFilesByFolder[r.group] || []).filter(function (f) { return f.name !== file.name; }), true), [newUpload], false), _b[r.uploadFolderId] = __spreadArray(__spreadArray([], (uploadedFilesByFolder[r.uploadFolderId] || []).filter(function (f) { return f.name !== file.name; }), true), [newUpload], false), _b));
                                        baseRows = (rows && rows.length > 0) ? rows : allRows;
                                        updatedRows = baseRows.map(function (row) {
                                            return row.groupKey === r.groupKey
                                                ? __assign(__assign({}, row), { fileName: file.name, fileId: fileId_1 || row.fileId || "file_".concat(Date.now()), filePending: statusPending_1 }) : row;
                                        });
                                        this.setState({
                                            docUploadBusy: false,
                                            docUploadRowKey: null,
                                            docUploadMsg: msg,
                                            uploadedFilesByFolder: updatedByFolder,
                                            rows: updatedRows
                                        });
                                        if (statusPending_1 && r.uploadFolderId && !/^f\d+$/.test(r.uploadFolderId)) {
                                            setTimeout(function () { return _this._refreshFolderFiles(r.uploadFolderId, r.groupKey, true)["catch"](function () { return undefined; }); }, 1500);
                                        }
                                        return [3 /*break*/, 4];
                                    case 3:
                                        err_4 = _d.sent();
                                        this.setState({ docUploadBusy: false, docUploadRowKey: null, docUploadMsg: "Upload failed: ".concat((err_4 === null || err_4 === void 0 ? void 0 : err_4.message) || 'Error') });
                                        return [3 /*break*/, 4];
                                    case 4: return [2 /*return*/];
                                }
                            });
                        }); }}/>
                              {isUploading ? '⏳...' : '↑ Upload'}
                            </label>
                          </div>
                        </td>
                      </tr>);
                })}
                </tbody>
              </table>

              {/* Pagination footer */}
              <div style={{ padding: '10px 14px', color: '#64748b', fontSize: 11, display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: '#f8fafc', borderTop: '1px solid #e2e8f0' }}>
                <span>
                  Showing {filtered.length === 0 ? 0 : safePage * PAGE_ROWS + 1}–{Math.min((safePage + 1) * PAGE_ROWS, filtered.length)} of {filtered.length} rows
                </span>
                <div style={{ display: 'flex', gap: 3, alignItems: 'center' }}>
                  <button onClick={function () { return _this.setState({ docListPage: Math.max(0, safePage - 1) }); }} disabled={safePage === 0} style={{ border: '1px solid #cbd5e1', background: '#fff', borderRadius: 4, padding: '3px 8px', fontSize: 11, cursor: safePage === 0 ? 'not-allowed' : 'pointer', opacity: safePage === 0 ? 0.4 : 1 }}>‹</button>
                  {Array.from({ length: totalPages }, function (_, i) { return (<button key={i} onClick={function () { return _this.setState({ docListPage: i }); }} style={{
                        border: i === safePage ? 'none' : '1px solid #cbd5e1',
                        background: i === safePage ? '#0078d4' : '#fff',
                        color: i === safePage ? '#fff' : '#334155',
                        borderRadius: 4, padding: '3px 8px', fontSize: 11,
                        fontWeight: i === safePage ? 700 : 400,
                        cursor: 'pointer',
                        minWidth: 26
                    }}>{i + 1}</button>); })}
                  <button onClick={function () { return _this.setState({ docListPage: Math.min(totalPages - 1, safePage + 1) }); }} disabled={safePage >= totalPages - 1} style={{ border: '1px solid #cbd5e1', background: '#fff', borderRadius: 4, padding: '3px 8px', fontSize: 11, cursor: safePage >= totalPages - 1 ? 'not-allowed' : 'pointer', opacity: safePage >= totalPages - 1 ? 0.4 : 1 }}>›</button>
                </div>
              </div>
            </div>
          </div>)}
      </div>);
    };
    // ── Module 3: Vessels ─────────────────────────────────────────────────────
    VesselEmail.prototype._renderVesselsPage = function () {
        var _this = this;
        var _a = this.state, vessels = _a.vessels, vesselsSearch = _a.vesselsSearch, vesselStatusFilter = _a.vesselStatusFilter, vesselTypeFilter = _a.vesselTypeFilter, modal = _a.modal, selectedVessel = _a.selectedVessel, folderProvisioningVesselId = _a.folderProvisioningVesselId, folderCreationError = _a.folderCreationError, folderCreationResults = _a.folderCreationResults, panelLoading = _a.panelLoading, loading = _a.loading;
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
        return (<div style={{ display: 'flex', flexDirection: 'column', gap: 0 }}>

        {/* ── Header ── */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 20, flexWrap: 'wrap', gap: 12 }}>
          <div>
            <h2 style={{ margin: 0, fontSize: 22, fontWeight: 700, color: '#0f172a', display: 'flex', alignItems: 'center', gap: 10 }}>
              🚢 Vessels
              {vessels.length > 0 && (<span style={{ background: '#e0f2fe', color: '#0284c7', borderRadius: 20, padding: '2px 10px', fontSize: 13, fontWeight: 700 }}>
                  {vessels.length}
                </span>)}
            </h2>
            <p style={{ margin: '4px 0 0', fontSize: 13, color: '#64748b' }}>
              Manage fleet vessels, provision SharePoint folders, and view documents.
            </p>
          </div>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
            <button onClick={function () { return _this._goToView('vessels')["catch"](function () { return undefined; }); }} title="Reload vessel list from database" style={{ background: '#f1f5f9', color: '#475569', border: '1px solid #cbd5e1', borderRadius: 8, padding: '9px 14px', fontSize: 13, fontWeight: 600, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
              🔄 Refresh
            </button>
            <button onClick={function () { return selectedVessel ? _this._openDeleteVessel(selectedVessel) : alert('Please select a vessel first.'); }} style={{ background: 'linear-gradient(135deg, #f43f5e, #e11d48)', color: '#fff', border: 'none', borderRadius: 8, padding: '9px 18px', fontSize: 13, fontWeight: 600, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
              🗑 Delete
            </button>
            <button onClick={function () { return selectedVessel ? _this._openEditVessel(selectedVessel) : alert('Please select a vessel first.'); }} style={{ background: 'linear-gradient(135deg, #38bdf8, #0284c7)', color: '#fff', border: 'none', borderRadius: 8, padding: '9px 18px', fontSize: 13, fontWeight: 600, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
              ✏️ Edit
            </button>
            <button onClick={this._openCreate} style={{ background: 'linear-gradient(135deg, #10b981, #059669)', color: '#fff', border: 'none', borderRadius: 8, padding: '9px 18px', fontSize: 13, fontWeight: 600, cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6 }}>
              ＋ New Vessel
            </button>
          </div>
        </div>

        {/* ── Search & Filter Bar ── */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 20, flexWrap: 'wrap' }}>
          <div style={{ position: 'relative', flex: '1 1 220px', minWidth: 180 }}>
            <span style={{ position: 'absolute', left: 12, top: '50%', transform: 'translateY(-50%)', color: '#94a3b8', fontSize: 14 }}>🔍</span>
            <input type="text" placeholder="Search by name, IMO, type, shipyard…" value={vesselsSearch} onChange={function (e) { return _this.setState({ vesselsSearch: e.target.value }); }} style={{ width: '100%', padding: '10px 14px 10px 38px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 13, outline: 'none', boxSizing: 'border-box', background: '#fff' }}/>
          </div>
          <select value={vesselStatusFilter} onChange={function (e) { return _this.setState({ vesselStatusFilter: e.target.value }); }} style={{ padding: '10px 14px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 13, background: '#fff', outline: 'none', minWidth: 130 }}>
            <option value="all">All Status</option>
            <option value="Active">Active</option>
            <option value="In Maintenance">In Maintenance</option>
            <option value="Inactive">Inactive</option>
          </select>
          <select value={vesselTypeFilter || 'all'} onChange={function (e) { return _this.setState({ vesselTypeFilter: e.target.value }); }} style={{ padding: '10px 14px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 13, background: '#fff', outline: 'none', minWidth: 140 }}>
            <option value="all">All Types</option>
            {allTypes.map(function (t) { return <option key={t} value={t}>{t}</option>; })}
          </select>
          {(vesselsSearch || vesselStatusFilter !== 'all' || (vesselTypeFilter && vesselTypeFilter !== 'all')) && (<button onClick={function () { return _this.setState({ vesselsSearch: '', vesselStatusFilter: 'all', vesselTypeFilter: 'all' }); }} style={{ background: 'transparent', color: '#64748b', border: '1px solid #cbd5e1', borderRadius: 8, padding: '9px 14px', fontSize: 13, cursor: 'pointer' }}>
              ✕ Clear
            </button>)}
          <span style={{ marginLeft: 'auto', fontSize: 12, color: '#94a3b8', fontWeight: 500 }}>
            {filtered.length} vessel{filtered.length !== 1 ? 's' : ''}
          </span>
        </div>

        {/* ── Folder creation status banners ── */}
        {folderCreationError && (<div style={{ marginBottom: 12, background: '#fde7e9', border: '1px solid #fca5a5', borderRadius: 8, padding: '10px 14px', fontSize: 12, color: '#a4262c', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <span>⚠ {folderCreationError}</span>
            <button onClick={function () { return _this.setState({ folderCreationError: null }); }} style={{ border: 'none', background: 'none', cursor: 'pointer', color: '#a4262c', fontWeight: 700 }}>✕</button>
          </div>)}
        {folderCreationResults && !folderCreationError && (<div style={{ marginBottom: 12, background: '#dff6dd', border: '1px solid #86efac', borderRadius: 8, padding: '10px 14px', fontSize: 12, color: '#107c10', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <span>✅ SharePoint folders provisioned — {folderCreationResults.filter(function (r) { return r.status === 'created'; }).length} created, {folderCreationResults.filter(function (r) { return r.status === 'existed'; }).length} already existed.</span>
            <button onClick={function () { return _this.setState({ folderCreationResults: null }); }} style={{ border: 'none', background: 'none', cursor: 'pointer', color: '#107c10', fontWeight: 700 }}>✕</button>
          </div>)}

        {/* ── Loading State ── */}
        {isLoading ? (<div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', minHeight: 280, gap: 16, background: '#fff', borderRadius: 12, border: '1px solid #e2e8f0' }}>
            <div style={{
                    width: 40, height: 40, border: '3px solid #e2e8f0',
                    borderTopColor: '#0078d4', borderRadius: '50%',
                    animation: 'spin 0.8s linear infinite'
                }}/>
            <p style={{ margin: 0, fontSize: 14, color: '#64748b', fontWeight: 500 }}>Loading vessels from database…</p>
            <p style={{ margin: 0, fontSize: 12, color: '#94a3b8' }}>Fetching vessel records and folder structure</p>
          </div>) : filtered.length === 0 ? (
            /* ── Empty State ── */
            <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', minHeight: 280, gap: 16, background: '#fff', borderRadius: 12, border: '2px dashed #e2e8f0' }}>
            <span style={{ fontSize: 48 }}>🚢</span>
            <div style={{ textAlign: 'center' }}>
              <p style={{ margin: 0, fontSize: 16, fontWeight: 700, color: '#0f172a' }}>
                {vessels.length === 0 ? 'No vessels yet' : 'No vessels match your filter'}
              </p>
              <p style={{ margin: '6px 0 0', fontSize: 13, color: '#64748b' }}>
                {vessels.length === 0
                    ? 'Create your first vessel to provision its SharePoint folder structure.'
                    : 'Try clearing the search or filters.'}
              </p>
            </div>
            {vessels.length === 0 && (<button onClick={this._openCreate} style={{ background: '#10b981', color: '#fff', border: 'none', borderRadius: 8, padding: '10px 24px', fontSize: 14, fontWeight: 600, cursor: 'pointer' }}>
                ＋ Create First Vessel
              </button>)}
          </div>) : (
            /* ── Vessel Cards Grid ── */
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(320px, 1fr))', gap: 16 }}>
            {filtered.map(function (vessel) {
                    var isSelected = (selectedVessel === null || selectedVessel === void 0 ? void 0 : selectedVessel.id) === vessel.id;
                    var status = vessel.status || 'Active';
                    var statusColor = status === 'Active' ? '#10b981' : status === 'In Maintenance' ? '#f59e0b' : '#ef4444';
                    var statusBg = status === 'Active' ? '#f0fdf4' : status === 'In Maintenance' ? '#fffbeb' : '#fef2f2';
                    var isProvisioning = folderProvisioningVesselId === vessel.id;
                    // Use deterministic vessel image from pool
                    var imgSrc = (0, vesselImagePool_1.getVesselImageForId)(vessel.id);
                    return (<div key={vessel.id} onClick={function () { return _this.setState({ selectedVessel: isSelected ? null : vessel }); }} style={{
                            background: '#fff',
                            borderRadius: 14,
                            border: isSelected ? '2px solid #0078d4' : '1px solid #e2e8f0',
                            boxShadow: isSelected ? '0 0 0 3px rgba(0,120,212,0.15)' : '0 1px 4px rgba(0,0,0,0.06)',
                            overflow: 'hidden',
                            cursor: 'pointer',
                            transition: 'all 0.18s ease',
                            display: 'flex',
                            flexDirection: 'column'
                        }} onMouseEnter={function (e) { if (!isSelected)
                        e.currentTarget.style.boxShadow = '0 4px 16px rgba(0,0,0,0.10)'; }} onMouseLeave={function (e) { if (!isSelected)
                        e.currentTarget.style.boxShadow = '0 1px 4px rgba(0,0,0,0.06)'; }}>
                  {/* Card image header */}
                  <div style={{ position: 'relative', height: 120, overflow: 'hidden', background: '#1e3a5f' }}>
                    <img src={(0, vesselImagePool_1.resolveImgUrl)(imgSrc)} alt={vessel.name} style={{ width: '100%', height: '100%', objectFit: 'cover', opacity: 0.85 }}/>
                    <div style={{
                            position: 'absolute', inset: 0,
                            background: 'linear-gradient(to bottom, transparent 30%, rgba(15,23,42,0.75) 100%)'
                        }}/>
                    {/* Status badge */}
                    <span style={{
                            position: 'absolute', top: 10, right: 10,
                            background: statusBg, color: statusColor,
                            borderRadius: 20, padding: '2px 10px', fontSize: 11, fontWeight: 700,
                            border: "1px solid ".concat(statusColor, "40")
                        }}>
                      {status}
                    </span>
                    {/* Selection indicator */}
                    {isSelected && (<span style={{
                                position: 'absolute', top: 10, left: 10,
                                background: '#0078d4', color: '#fff',
                                borderRadius: '50%', width: 22, height: 22,
                                display: 'flex', alignItems: 'center', justifyContent: 'center',
                                fontSize: 13, fontWeight: 700, boxShadow: '0 2px 6px rgba(0,0,0,0.25)'
                            }}>
                        ✓
                      </span>)}
                    {/* Vessel name over image */}
                    <div style={{ position: 'absolute', bottom: 10, left: 14, right: 14 }}>
                      <p style={{ margin: 0, fontSize: 15, fontWeight: 700, color: '#fff', textShadow: '0 1px 3px rgba(0,0,0,0.5)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                        🚢 {vessel.name}
                      </p>
                    </div>
                  </div>

                  {/* Card body */}
                  <div style={{ padding: '14px 16px', flex: 1, display: 'flex', flexDirection: 'column', gap: 8 }}>
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '6px 12px' }}>
                      <div>
                        <span style={{ fontSize: 10, color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase' }}>IMO</span>
                        <p style={{ margin: 0, fontSize: 13, color: '#1e293b', fontWeight: 600 }}>{vessel.imo || '—'}</p>
                      </div>
                      <div>
                        <span style={{ fontSize: 10, color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase' }}>Type</span>
                        <p style={{ margin: 0, fontSize: 13, color: '#1e293b', fontWeight: 600, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{vessel.vessel_type || '—'}</p>
                      </div>
                      <div>
                        <span style={{ fontSize: 10, color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase' }}>Shipyard</span>
                        <p style={{ margin: 0, fontSize: 12, color: '#475569', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{vessel.shipyard || '—'}</p>
                      </div>
                      <div>
                        <span style={{ fontSize: 10, color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase' }}>Hull No.</span>
                        <p style={{ margin: 0, fontSize: 12, color: '#475569' }}>{vessel.hull_number || '—'}</p>
                      </div>
                    </div>

                    {/* Action buttons */}
                    <div style={{ display: 'flex', gap: 8, marginTop: 4 }}>
                      {/* View Documents — navigate to Documents filtered for this vessel */}
                      <button onClick={function (e) {
                            e.stopPropagation();
                            _this.setState({ vesselFilter: vessel.name, docMainFolder: null, folderPathStack: [] });
                            void _this._goToView('list');
                        }} title="View documents for this vessel" style={{
                            flex: 1, background: '#eff6ff', color: '#1d4ed8',
                            border: '1px solid #bfdbfe', borderRadius: 7, padding: '7px 10px',
                            fontSize: 11, fontWeight: 700, cursor: 'pointer',
                            display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 4
                        }}>
                        📄 View Documents
                      </button>
                      {/* Provision SPO folders */}
                      <button disabled={!!folderProvisioningVesselId} onClick={function (e) {
                            e.stopPropagation();
                            _this._provisionVesselFolders(vessel.name, vessel.id)["catch"](function () { return undefined; });
                        }} title="Create SharePoint folder structure for this vessel" style={{
                            flex: 1,
                            border: '1px solid #cbd5e1', borderRadius: 7, padding: '7px 10px',
                            fontSize: 11, fontWeight: 700, cursor: folderProvisioningVesselId ? 'not-allowed' : 'pointer',
                            background: isProvisioning ? '#f0f9ff' : '#f8fafc',
                            color: isProvisioning ? '#0284c7' : '#334155',
                            display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 4
                        }}>
                        {isProvisioning ? (<><span style={{ display: 'inline-block', width: 11, height: 11, border: '2px solid #0284c7', borderTopColor: 'transparent', borderRadius: '50%', animation: 'spin 0.7s linear infinite' }}/> Creating…</>) : '📁 Provision'}
                      </button>
                    </div>
                  </div>
                </div>);
                })}
          </div>)}

        {/* Keyframe animation for spinners */}
        <style>{"\n          @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }\n        "}</style>

        {/* Modals */}
        {modal === 'create' && this._renderVesselForm('create')}
        {modal === 'edit' && this._renderVesselForm('edit')}
        {modal === 'delete' && this._renderDeleteModal()}
      </div>);
    };
    // ── Module 4: Templates ───────────────────────────────────────────────────
    VesselEmail.prototype._renderTemplatesPage = function () {
        var templatesList = this.state.templatesList;
        return (<div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 12 }}>
          <div>
            <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: '#0f172a' }}>Templates</h2>
            <p style={{ margin: '4px 0 0', fontSize: 13, color: '#64748b' }}>Create and manage document templates.</p>
          </div>
          <button onClick={function () { return alert('New Template dialog'); }} style={{
                background: '#0078d4', color: '#fff', border: 'none', borderRadius: 6,
                padding: '8px 18px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
                display: 'inline-flex', alignItems: 'center', gap: 6
            }}>
            ＋ New Template
          </button>
        </div>

        <div style={{ background: '#fff', borderRadius: 8, border: '1px solid #e2e8f0', overflow: 'hidden' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
            <thead>
              <tr style={{ background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 11, textTransform: 'uppercase', textAlign: 'left' }}>
                <th style={{ padding: '10px 16px' }}>Name</th>
                <th style={{ padding: '10px 16px' }}>Type</th>
                <th style={{ padding: '10px 16px' }}>Description</th>
                <th style={{ padding: '10px 16px' }}>Modified</th>
              </tr>
            </thead>
            <tbody>
              {templatesList.map(function (tpl) { return (<tr key={tpl.id} style={{ borderBottom: '1px solid #f1f5f9' }}>
                  <td style={{ padding: '12px 16px', fontWeight: 600, color: '#1e293b', display: 'flex', alignItems: 'center', gap: 8 }}>
                    <span style={{ color: '#2563eb' }}>📝</span>
                    {tpl.name}
                  </td>
                  <td style={{ padding: '12px 16px' }}>{badge('blue', tpl.type)}</td>
                  <td style={{ padding: '12px 16px', color: '#475569' }}>{tpl.description}</td>
                  <td style={{ padding: '12px 16px', color: '#64748b' }}>{tpl.modified}</td>
                </tr>); })}
            </tbody>
          </table>
        </div>
      </div>);
    };
    // ── Module 5: Approvals ───────────────────────────────────────────────────
    VesselEmail.prototype._renderApprovalsPage = function () {
        var _this = this;
        var _a = this.state, approvalsList = _a.approvalsList, approvalTab = _a.approvalTab, panelLoading = _a.panelLoading;
        var filtered = approvalsList.filter(function (a) { return a.status === approvalTab; });
        return (<div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div>
          <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: '#0f172a' }}>Approvals</h2>
          <p style={{ margin: '4px 0 0', fontSize: 13, color: '#64748b' }}>Review and take action on pending approvals.</p>
        </div>

        {/* Tabs */}
        <div style={{ display: 'flex', gap: 8, borderBottom: '1px solid #e2e8f0', paddingBottom: 8 }}>
          {['Pending', 'Approved', 'Rejected'].map(function (tab) { return (<button key={tab} onClick={function () { return _this.setState({ approvalTab: tab }); }} style={{
                    border: 'none', background: approvalTab === tab ? '#eff6ff' : 'transparent',
                    color: approvalTab === tab ? '#0078d4' : '#64748b', fontWeight: approvalTab === tab ? 700 : 500,
                    fontSize: 13, padding: '6px 14px', borderRadius: 6, cursor: 'pointer'
                }}>
              {tab} ({approvalsList.filter(function (a) { return a.status === tab; }).length})
            </button>); })}
        </div>

        {/* Table */}
        {panelLoading ? (<div style={{ padding: 32, textAlign: 'center', color: '#64748b', fontSize: 13 }}>⏳ Loading approvals...</div>) : (<div style={{ background: '#fff', borderRadius: 8, border: '1px solid #e2e8f0', overflowX: 'auto' }}>
          <table style={{ width: '100%', minWidth: 700, borderCollapse: 'collapse', fontSize: 13 }}>
            <thead>
              <tr style={{ background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 11, textTransform: 'uppercase', textAlign: 'left' }}>
                <th style={{ padding: '10px 16px', whiteSpace: 'nowrap' }}>Document Name</th>
                <th style={{ padding: '10px 16px', whiteSpace: 'nowrap' }}>Vessel</th>
                <th style={{ padding: '10px 16px', whiteSpace: 'nowrap' }}>Requested By</th>
                <th style={{ padding: '10px 16px', whiteSpace: 'nowrap' }}>Requested On</th>
                <th style={{ padding: '10px 16px', textAlign: 'right', whiteSpace: 'nowrap' }}>Action</th>
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 ? (<tr><td colSpan={5} style={{ padding: 24, textAlign: 'center', color: '#94a3b8', fontSize: 13 }}>No {approvalTab.toLowerCase()} approvals.</td></tr>) : filtered.map(function (app) { return (<tr key={app.id} style={{ borderBottom: '1px solid #f1f5f9' }}>
                  <td style={{ padding: '12px 16px', fontWeight: 600, color: '#1e293b', maxWidth: 200, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={app.documentName}>{app.documentName}</td>
                  <td style={{ padding: '12px 16px', color: '#475569', whiteSpace: 'nowrap' }}>{app.vessel}</td>
                  <td style={{ padding: '12px 16px', color: '#475569', maxWidth: 180, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={app.requestedBy}>{app.requestedBy}</td>
                  <td style={{ padding: '12px 16px', color: '#64748b', whiteSpace: 'nowrap' }}>{app.requestedOn}</td>
                  <td style={{ padding: '12px 16px', textAlign: 'right' }}>
                    {app.status === 'Pending' ? (<div style={{ display: 'flex', gap: 6, justifyContent: 'flex-end' }}>
                        <button onClick={function () { return __awaiter(_this, void 0, void 0, function () {
                            var userEmail, adminParam, _a, updated, vesselName_1;
                            var _this = this;
                            return __generator(this, function (_b) {
                                switch (_b.label) {
                                    case 0:
                                        userEmail = this.props.userEmail || '';
                                        adminParam = userEmail ? "?admin=".concat(encodeURIComponent(userEmail)) : '';
                                        _b.label = 1;
                                    case 1:
                                        _b.trys.push([1, 3, , 4]);
                                        return [4 /*yield*/, fetch("".concat(this._base(), "/api/approvals/").concat(app.id, "/approve").concat(adminParam), { method: 'POST', headers: this._headers() })];
                                    case 2:
                                        _b.sent();
                                        return [3 /*break*/, 4];
                                    case 3:
                                        _a = _b.sent();
                                        return [3 /*break*/, 4];
                                    case 4:
                                        updated = approvalsList.map(function (item) { return item.id === app.id ? __assign(__assign({}, item), { status: 'Approved' }) : item; });
                                        this.setState({ approvalsList: updated });
                                        this._filesLoadedForFolders.clear();
                                        this._filesLoadedForVessels.clear();
                                        this._loadData();
                                        // Re-fetch files for the approved item's vessel so list view updates
                                        if (app.vessel && app.vessel !== '—') {
                                            vesselName_1 = app.vessel;
                                            setTimeout(function () {
                                                _this._filesLoadedForVessels.add(vesselName_1);
                                                _this._loadFilesForVessel(vesselName_1)["catch"](function () { return undefined; });
                                            }, 2000);
                                        }
                                        return [2 /*return*/];
                                }
                            });
                        }); }} style={{ border: 'none', background: '#dff6dd', color: '#107c10', width: 28, height: 28, borderRadius: 4, cursor: 'pointer', fontWeight: 700 }} title="Approve">
                          ✓
                        </button>
                        <button onClick={function () { return __awaiter(_this, void 0, void 0, function () {
                            var userEmail, adminParam, _a, updated;
                            return __generator(this, function (_b) {
                                switch (_b.label) {
                                    case 0:
                                        userEmail = this.props.userEmail || '';
                                        adminParam = userEmail ? "?admin=".concat(encodeURIComponent(userEmail)) : '';
                                        _b.label = 1;
                                    case 1:
                                        _b.trys.push([1, 3, , 4]);
                                        return [4 /*yield*/, fetch("".concat(this._base(), "/api/approvals/").concat(app.id, "/reject").concat(adminParam), {
                                                method: 'POST',
                                                headers: __assign(__assign({}, this._headers()), { 'Content-Type': 'application/json' }),
                                                body: JSON.stringify({ reason: null })
                                            })];
                                    case 2:
                                        _b.sent();
                                        return [3 /*break*/, 4];
                                    case 3:
                                        _a = _b.sent();
                                        return [3 /*break*/, 4];
                                    case 4:
                                        updated = approvalsList.map(function (item) { return item.id === app.id ? __assign(__assign({}, item), { status: 'Rejected' }) : item; });
                                        this.setState({ approvalsList: updated });
                                        return [2 /*return*/];
                                }
                            });
                        }); }} style={{ border: 'none', background: '#fde7e9', color: '#a4262c', width: 28, height: 28, borderRadius: 4, cursor: 'pointer', fontWeight: 700 }} title="Reject">
                          ✕
                        </button>
                      </div>) : (badge(app.status === 'Approved' ? 'green' : 'red', app.status))}
                  </td>
                </tr>); })}
            </tbody>
          </table>
        </div>)}
      </div>);
    };
    // ── Module 6: Notifications ───────────────────────────────────────────────
    VesselEmail.prototype._renderNotificationsPage = function () {
        var _this = this;
        var _a = this.state, notificationsList = _a.notificationsList, notificationFilter = _a.notificationFilter;
        var filtered = notificationsList.filter(function (n) { return notificationFilter === 'all' || !n.read; });
        return (<div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 12 }}>
          <div>
            <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: '#0f172a' }}>Notifications</h2>
            <p style={{ margin: '4px 0 0', fontSize: 13, color: '#64748b' }}>Stay updated with important alerts and reminders.</p>
          </div>
          <button onClick={function () {
                var updated = notificationsList.map(function (n) { return (__assign(__assign({}, n), { read: true })); });
                _this.setState({ notificationsList: updated });
            }} style={{ background: 'none', border: 'none', color: '#2563eb', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>
            Mark all as read
          </button>
        </div>

        {/* Filter buttons */}
        <div style={{ display: 'flex', gap: 8 }}>
          <button onClick={function () { return _this.setState({ notificationFilter: 'all' }); }} style={{
                border: 'none', background: notificationFilter === 'all' ? '#0078d4' : '#f1f5f9',
                color: notificationFilter === 'all' ? '#fff' : '#475569', borderRadius: 16,
                padding: '4px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer'
            }}>
            All ({notificationsList.length})
          </button>
          <button onClick={function () { return _this.setState({ notificationFilter: 'unread' }); }} style={{
                border: 'none', background: notificationFilter === 'unread' ? '#0078d4' : '#f1f5f9',
                color: notificationFilter === 'unread' ? '#fff' : '#475569', borderRadius: 16,
                padding: '4px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer'
            }}>
            Unread ({notificationsList.filter(function (n) { return !n.read; }).length})
          </button>
        </div>

        {/* Notification list */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
          {filtered.map(function (notif) { return (<div key={notif.id} style={{
                    background: notif.read ? '#fff' : '#f0f9ff',
                    borderRadius: 8, border: '1px solid #e2e8f0', padding: '14px 18px',
                    display: 'flex', justifyContent: 'space-between', alignItems: 'center'
                }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                <div style={{
                    width: 32, height: 32, borderRadius: '50%',
                    background: notif.type === 'alert' ? '#fde7e9' : notif.type === 'warning' ? '#fff4ce' : notif.type === 'success' ? '#dff6dd' : '#e1efff',
                    color: notif.type === 'alert' ? '#a4262c' : notif.type === 'warning' ? '#8a5700' : notif.type === 'success' ? '#107c10' : '#0078d4',
                    display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 700
                }}>
                  {notif.type === 'alert' ? '⚠' : notif.type === 'warning' ? '🔔' : notif.type === 'success' ? '✓' : 'ℹ'}
                </div>
                <div>
                  <div style={{ fontSize: 13, color: '#1e293b', fontWeight: notif.read ? 500 : 700 }}>{notif.message}</div>
                  <div style={{ fontSize: 11, color: '#64748b', marginTop: 2 }}>{notif.timestamp}</div>
                </div>
              </div>

              <div>
                {badge(notif.priority === 'High' ? 'red' : notif.priority === 'Medium' ? 'orange' : 'blue', notif.priority)}
              </div>
            </div>); })}
        </div>
      </div>);
    };
    // ── Module 7: Reports ─────────────────────────────────────────────────────
    VesselEmail.prototype._renderReportsPage = function () {
        return (<div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
        {/* Header */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 12 }}>
          <div>
            <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: '#0f172a' }}>Reports</h2>
            <p style={{ margin: '4px 0 0', fontSize: 13, color: '#64748b' }}>Generate and view analytics reports.</p>
          </div>

          <button onClick={function () { return alert('Downloading report PDF...'); }} style={{
                background: '#0078d4', color: '#fff', border: 'none', borderRadius: 6,
                padding: '8px 18px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
                display: 'inline-flex', alignItems: 'center', gap: 6
            }}>
            ⬇ Download Report
          </button>
        </div>

        {/* Date / Filter bar */}
        <div style={{ background: '#fff', borderRadius: 8, padding: 12, border: '1px solid #e2e8f0', display: 'flex', gap: 12, alignItems: 'center' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13, color: '#334155' }}>
            <span>📅</span>
            <input type="text" defaultValue="May 1, 2024 - May 31, 2024" style={{ border: '1px solid #cbd5e1', borderRadius: 6, padding: '6px 12px', fontSize: 13, outline: 'none' }}/>
          </div>
          <select style={{ border: '1px solid #cbd5e1', borderRadius: 6, padding: '6px 12px', fontSize: 13, outline: 'none', background: '#fff' }}>
            <option>All Vessels</option>
            <option>Ocean Star</option>
            <option>Sea Breeze</option>
          </select>
        </div>

        {/* 4 Stat Cards */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: 16 }}>
          <div style={{ background: '#fff', borderRadius: 8, padding: 18, border: '1px solid #e2e8f0' }}>
            <div style={{ fontSize: 12, color: '#64748b', fontWeight: 600, textTransform: 'uppercase' }}>Total Documents</div>
            <div style={{ fontSize: 26, fontWeight: 800, color: '#0078d4', marginTop: 4 }}>156</div>
          </div>
          <div style={{ background: '#fff', borderRadius: 8, padding: 18, border: '1px solid #e2e8f0' }}>
            <div style={{ fontSize: 12, color: '#64748b', fontWeight: 600, textTransform: 'uppercase' }}>Expired Documents</div>
            <div style={{ fontSize: 26, fontWeight: 800, color: '#ef4444', marginTop: 4 }}>8</div>
          </div>
          <div style={{ background: '#fff', borderRadius: 8, padding: 18, border: '1px solid #e2e8f0' }}>
            <div style={{ fontSize: 12, color: '#64748b', fontWeight: 600, textTransform: 'uppercase' }}>Expiring in 30 days</div>
            <div style={{ fontSize: 26, fontWeight: 800, color: '#f59e0b', marginTop: 4 }}>23</div>
          </div>
          <div style={{ background: '#fff', borderRadius: 8, padding: 18, border: '1px solid #e2e8f0' }}>
            <div style={{ fontSize: 12, color: '#64748b', fontWeight: 600, textTransform: 'uppercase' }}>Valid Documents</div>
            <div style={{ fontSize: 26, fontWeight: 800, color: '#10b981', marginTop: 4 }}>125</div>
          </div>
        </div>

        {/* Charts Grid */}
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 20 }}>
          {/* Documents by Type (Donut Chart) */}
          <div style={{ background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', padding: 20 }}>
            <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 700, color: '#0f172a' }}>Documents by Type</h3>
            <div style={{ display: 'flex', alignItems: 'center', gap: 24 }}>
              <div style={{ position: 'relative', width: 140, height: 140 }}>
                <svg width="140" height="140" viewBox="0 0 42 42">
                  <circle cx="21" cy="21" r="15.91549430918954" fill="transparent" stroke="#e2e8f0" strokeWidth="5"/>
                  <circle cx="21" cy="21" r="15.91549430918954" fill="transparent" stroke="#2563eb" strokeWidth="5" strokeDasharray="29 71" strokeDashoffset="25"/>
                  <circle cx="21" cy="21" r="15.91549430918954" fill="transparent" stroke="#f59e0b" strokeWidth="5" strokeDasharray="26 74" strokeDashoffset="96"/>
                  <circle cx="21" cy="21" r="15.91549430918954" fill="transparent" stroke="#10b981" strokeWidth="5" strokeDasharray="22 78" strokeDashoffset="70"/>
                  <circle cx="21" cy="21" r="15.91549430918954" fill="transparent" stroke="#8b5cf6" strokeWidth="5" strokeDasharray="13 87" strokeDashoffset="48"/>
                  <circle cx="21" cy="21" r="15.91549430918954" fill="transparent" stroke="#06b6d4" strokeWidth="5" strokeDasharray="10 90" strokeDashoffset="35"/>
                </svg>
              </div>

              <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 6, fontSize: 12 }}>
                <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                  <span style={{ color: '#475569' }}>● Insurance</span>
                  <span style={{ fontWeight: 600 }}>45 (29%)</span>
                </div>
                <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                  <span style={{ color: '#475569' }}>● Crew</span>
                  <span style={{ fontWeight: 600 }}>40 (26%)</span>
                </div>
                <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                  <span style={{ color: '#475569' }}>● Maintenance</span>
                  <span style={{ fontWeight: 600 }}>35 (22%)</span>
                </div>
                <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                  <span style={{ color: '#475569' }}>● Certificate</span>
                  <span style={{ fontWeight: 600 }}>20 (13%)</span>
                </div>
                <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                  <span style={{ color: '#475569' }}>● Survey</span>
                  <span style={{ fontWeight: 600 }}>16 (10%)</span>
                </div>
                <div style={{ borderTop: '1px solid #e2e8f0', paddingTop: 6, marginTop: 4, display: 'flex', justifyContent: 'space-between', fontWeight: 700 }}>
                  <span>Total</span>
                  <span>156</span>
                </div>
              </div>
            </div>
          </div>

          {/* Documents by Vessel (Bar Chart) */}
          <div style={{ background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', padding: 20 }}>
            <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 700, color: '#0f172a' }}>Documents by Vessel</h3>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
              {[
                { name: 'Ocean Star', count: 45, max: 50 },
                { name: 'Sea Breeze', count: 32, max: 50 },
                { name: 'Blue Horizon', count: 28, max: 50 },
                { name: 'Pacific Dawn', count: 26, max: 50 },
                { name: 'Atlantic Wave', count: 25, max: 50 },
            ].map(function (v) { return (<div key={v.name} style={{ display: 'flex', alignItems: 'center', gap: 12, fontSize: 12 }}>
                  <span style={{ width: 90, color: '#475569', fontWeight: 500 }}>{v.name}</span>
                  <div style={{ flex: 1, background: '#f1f5f9', borderRadius: 4, height: 16, overflow: 'hidden' }}>
                    <div style={{ width: "".concat((v.count / v.max) * 100, "%"), background: '#2563eb', height: '100%', borderRadius: 4 }}/>
                  </div>
                  <span style={{ width: 24, fontWeight: 600, textAlign: 'right' }}>{v.count}</span>
                </div>); })}
              <div style={{ fontSize: 10, color: '#94a3b8', textAlign: 'center', marginTop: 8 }}>No. of Documents</div>
            </div>
          </div>
        </div>
      </div>);
    };
    // ── Module 8: User Management ─────────────────────────────────────────────
    VesselEmail.prototype._renderUsersPage = function () {
        var _this = this;
        var _a = this.state, usersList = _a.usersList, userSearch = _a.userSearch;
        var filtered = usersList.filter(function (u) {
            return u.name.toLowerCase().indexOf(userSearch.toLowerCase()) !== -1 ||
                u.email.toLowerCase().indexOf(userSearch.toLowerCase()) !== -1;
        });
        return (<div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 12 }}>
          <div>
            <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: '#0f172a' }}>User Management</h2>
            <p style={{ margin: '4px 0 0', fontSize: 13, color: '#64748b' }}>Manage users and their roles.</p>
          </div>
          <button onClick={function () { return alert('Add User Modal'); }} style={{
                background: '#0078d4', color: '#fff', border: 'none', borderRadius: 6,
                padding: '8px 18px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
                display: 'inline-flex', alignItems: 'center', gap: 6
            }}>
            ＋ Add User
          </button>
        </div>

        <div style={{ background: '#fff', borderRadius: 8, padding: 12, border: '1px solid #e2e8f0', display: 'flex', gap: 12 }}>
          <input type="text" placeholder="Search users" value={userSearch} onChange={function (e) { return _this.setState({ userSearch: e.target.value }); }} style={{ padding: '6px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, outline: 'none', width: 240 }}/>
        </div>

        <div style={{ background: '#fff', borderRadius: 8, border: '1px solid #e2e8f0', overflow: 'hidden' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
            <thead>
              <tr style={{ background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 11, textTransform: 'uppercase', textAlign: 'left' }}>
                <th style={{ padding: '10px 16px' }}>Name</th>
                <th style={{ padding: '10px 16px' }}>Email</th>
                <th style={{ padding: '10px 16px' }}>Role</th>
                <th style={{ padding: '10px 16px' }}>Status</th>
                <th style={{ padding: '10px 16px' }}>Last Login</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map(function (usr) { return (<tr key={usr.id} style={{ borderBottom: '1px solid #f1f5f9' }}>
                  <td style={{ padding: '12px 16px', fontWeight: 600, color: '#1e293b', display: 'flex', alignItems: 'center', gap: 10 }}>
                    <div style={{ width: 28, height: 28, borderRadius: '50%', background: '#e2e8f0', color: '#475569', fontSize: 12, fontWeight: 700, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                      {usr.name.charAt(0)}
                    </div>
                    {usr.name}
                  </td>
                  <td style={{ padding: '12px 16px', color: '#475569' }}>{usr.email}</td>
                  <td style={{ padding: '12px 16px', color: '#334155' }}>{usr.role}</td>
                  <td style={{ padding: '12px 16px' }}>
                    {badge(usr.status === 'Active' ? 'green' : 'red', usr.status)}
                  </td>
                  <td style={{ padding: '12px 16px', color: '#64748b' }}>{usr.lastLogin}</td>
                </tr>); })}
            </tbody>
          </table>
        </div>
      </div>);
    };
    // ── Module 9: Settings ────────────────────────────────────────────────────
    VesselEmail.prototype._renderSettingsPage = function () {
        var _this = this;
        var _a = this.state, settingsTab = _a.settingsTab, settingsForm = _a.settingsForm, settingsSavedMsg = _a.settingsSavedMsg;
        return (<div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div>
          <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: '#0f172a' }}>Settings</h2>
          <p style={{ margin: '4px 0 0', fontSize: 13, color: '#64748b' }}>Configure application settings and preferences.</p>
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: '200px 1fr', gap: 24, background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', padding: 24 }}>
          {/* Settings Left Nav */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: 4, borderRight: '1px solid #f1f5f9', paddingRight: 16 }}>
            {[
                'General', 'Document Settings', 'Notification Settings',
                'Permission Settings', 'Integration', 'Audit Logs'
            ].map(function (tab) { return (<button key={tab} onClick={function () { return _this.setState({ settingsTab: tab }); }} style={{
                    border: 'none', background: settingsTab === tab ? '#eff6ff' : 'transparent',
                    color: settingsTab === tab ? '#0078d4' : '#475569', fontWeight: settingsTab === tab ? 700 : 500,
                    fontSize: 13, padding: '8px 12px', borderRadius: 6, textAlign: 'left', cursor: 'pointer'
                }}>
                {tab}
              </button>); })}
          </div>

          {/* Settings Content Area */}
          <div style={{ maxWidth: 500 }}>
            <h3 style={{ margin: '0 0 16px', fontSize: 15, fontWeight: 700, color: '#0f172a' }}>{settingsTab}</h3>

            {settingsSavedMsg && (<div style={{ background: '#dff6dd', color: '#107c10', padding: '8px 12px', borderRadius: 6, fontSize: 13, marginBottom: 16 }}>
                ✓ Settings saved successfully.
              </div>)}

            {settingsTab === 'General' && (<div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
                <div>
                  <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 }}>Site Title</label>
                  <input type="text" value={settingsForm.siteTitle} onChange={function (e) { return _this.setState({ settingsForm: __assign(__assign({}, settingsForm), { siteTitle: e.target.value }) }); }} style={{ width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, outline: 'none', boxSizing: 'border-box' }}/>
                </div>

                <div>
                  <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 }}>Site Description</label>
                  <textarea value={settingsForm.siteDescription} onChange={function (e) { return _this.setState({ settingsForm: __assign(__assign({}, settingsForm), { siteDescription: e.target.value }) }); }} style={{ width: '100%', height: 60, padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, outline: 'none', boxSizing: 'border-box', resize: 'vertical' }}/>
                </div>

                <div>
                  <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 }}>Date Format</label>
                  <select value={settingsForm.dateFormat} onChange={function (e) { return _this.setState({ settingsForm: __assign(__assign({}, settingsForm), { dateFormat: e.target.value }) }); }} style={{ width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, outline: 'none', background: '#fff' }}>
                    <option value="MM/DD/YYYY">MM/DD/YYYY</option>
                    <option value="DD/MM/YYYY">DD/MM/YYYY</option>
                    <option value="YYYY-MM-DD">YYYY-MM-DD</option>
                  </select>
                </div>

                <div>
                  <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 }}>Time Zone</label>
                  <select value={settingsForm.timeZone} onChange={function (e) { return _this.setState({ settingsForm: __assign(__assign({}, settingsForm), { timeZone: e.target.value }) }); }} style={{ width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, outline: 'none', background: '#fff' }}>
                    <option value="(UTC+05:30) Chennai, Kolkata, Mumbai, New Delhi">(UTC+05:30) Chennai, Kolkata, Mumbai, New Delhi</option>
                    <option value="(UTC+00:00) UTC">(UTC+00:00) UTC</option>
                  </select>
                </div>

                <div style={{ marginTop: 12 }}>
                  <button onClick={function () {
                    _this.setState({ settingsSavedMsg: true });
                    setTimeout(function () { return _this.setState({ settingsSavedMsg: false }); }, 2000);
                }} style={{
                    background: '#0078d4', color: '#fff', border: 'none', borderRadius: 6,
                    padding: '8px 20px', fontSize: 13, fontWeight: 600, cursor: 'pointer'
                }}>
                    Save Changes
                  </button>
                </div>
              </div>)}

            {settingsTab !== 'General' && (<div style={{ color: '#64748b', fontSize: 13 }}>
                Configuration settings for {settingsTab} are active with system default policies.
              </div>)}
          </div>
        </div>
      </div>);
    };
    // ── Existing Views: Bento Email Dashboard ──────────────────────────────────
    VesselEmail.prototype._renderBentoEmailDashboardPage = function () {
        var _this = this;
        var _a = this.state, bentoLogs = _a.bentoLogs, panelLoading = _a.panelLoading, bentoStatusFilter = _a.bentoStatusFilter, bentoSearch = _a.bentoSearch, bentoComposeOpen = _a.bentoComposeOpen, bentoComposeForm = _a.bentoComposeForm, bentoComposeBusy = _a.bentoComposeBusy, bentoComposeMsg = _a.bentoComposeMsg, bentoComposeErr = _a.bentoComposeErr, bentoDetailLog = _a.bentoDetailLog, vessels = _a.vessels, bentoUploadFile = _a.bentoUploadFile, bentoUploadVessel = _a.bentoUploadVessel, bentoUploadTag = _a.bentoUploadTag, bentoUploadBusy = _a.bentoUploadBusy, bentoUploadMsg = _a.bentoUploadMsg, bentoUploadErr = _a.bentoUploadErr;
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
        return (<div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div style={{ background: 'linear-gradient(135deg, #0f172a 0%, #1e293b 100%)', borderRadius: 12, padding: '24px 28px', color: '#fff', boxShadow: '0 4px 16px rgba(0,0,0,0.15)' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 16 }}>
            <div>
              <div style={{ display: 'inline-flex', alignItems: 'center', gap: 8, background: 'rgba(255,255,255,0.1)', padding: '4px 12px', borderRadius: 20, fontSize: 12, fontWeight: 600, color: '#38bdf8', marginBottom: 8 }}>
                🤖 AI BENTO AUTOMATION ENGINE
              </div>
              <h1 style={{ margin: 0, fontSize: 22, fontWeight: 800 }}>AI Bento Email Dashboard</h1>
              <p style={{ margin: '6px 0 0', fontSize: 13, color: '#94a3b8', maxWidth: 650 }}>
                Automated document tagging, status tracking, and Graph email dispatching.
              </p>
            </div>
            <button style={{ background: 'linear-gradient(135deg, #38bdf8, #0284c7)', color: '#fff', border: 'none', borderRadius: 8, padding: '10px 22px', fontSize: 13, fontWeight: 700, cursor: 'pointer' }} onClick={function () { return _this.setState({
                bentoComposeOpen: true,
                bentoComposeMsg: null,
                bentoComposeErr: null,
                bentoComposeForm: {
                    vessel_name: '',
                    datasource_tag: 'mail',
                    subject_text: _this._buildAutoSubject('', 'mail', ''),
                    body: '',
                    file: null,
                    existing_attachment: '',
                    recipient: ''
                }
            }); }}>
              ✉ Compose & Dispatch Email
            </button>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: 14, marginTop: 22 }}>
            <div style={{ background: 'rgba(255,255,255,0.06)', borderRadius: 10, padding: '14px 18px', border: '1px solid rgba(255,255,255,0.1)' }}>
              <div style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', color: '#94a3b8' }}>Total Processed</div>
              <div style={{ fontSize: 26, fontWeight: 800, color: '#fff', marginTop: 4 }}>{totalCount}</div>
            </div>
            <div style={{ background: 'rgba(16, 185, 129, 0.12)', borderRadius: 10, padding: '14px 18px', border: '1px solid rgba(16, 185, 129, 0.25)' }}>
              <div style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', color: '#6ee7b7' }}>Completed</div>
              <div style={{ fontSize: 26, fontWeight: 800, color: '#34d399', marginTop: 4 }}>{completedCount}</div>
            </div>
            <div style={{ background: 'rgba(245, 158, 11, 0.12)', borderRadius: 10, padding: '14px 18px', border: '1px solid rgba(245, 158, 11, 0.25)' }}>
              <div style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', color: '#fcd34d' }}>Pending</div>
              <div style={{ fontSize: 26, fontWeight: 800, color: '#fbbf24', marginTop: 4 }}>{pendingCount}</div>
            </div>
            <div style={{ background: 'rgba(239, 68, 68, 0.12)', borderRadius: 10, padding: '14px 18px', border: '1px solid rgba(239, 68, 68, 0.25)' }}>
              <div style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', color: '#fca5a5' }}>Failed</div>
              <div style={{ fontSize: 26, fontWeight: 800, color: '#f87171', marginTop: 4 }}>{failedCount}</div>
            </div>
          </div>
        </div>

        {/* Logs Table */}
        <div style={{ background: '#fff', borderRadius: 8, border: '1px solid #e2e8f0', overflowX: 'auto' }}>
          <table style={{ width: '100%', minWidth: 900, borderCollapse: 'collapse', fontSize: 13 }}>
            <thead>
              <tr style={{ background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 11, textTransform: 'uppercase', textAlign: 'left', whiteSpace: 'nowrap' }}>
                <th style={{ padding: '10px 16px' }}>#</th>
                <th style={{ padding: '10px 16px' }}>Vessel</th>
                <th style={{ padding: '10px 16px' }}>Tag</th>
                <th style={{ padding: '10px 16px' }}>Subject</th>
                <th style={{ padding: '10px 16px' }}>Recipient</th>
                <th style={{ padding: '10px 16px' }}>Status</th>
                <th style={{ padding: '10px 16px' }}>Attached File</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map(function (log) { return (<tr key={log.id} style={{ borderBottom: '1px solid #f1f5f9' }}>
                  <td style={{ padding: '10px 16px', color: '#64748b', whiteSpace: 'nowrap' }}>#{log.id}</td>
                  <td style={{ padding: '10px 16px', fontWeight: 600, whiteSpace: 'nowrap' }}>{log.vessel_name || '—'}</td>
                  <td style={{ padding: '10px 16px', whiteSpace: 'nowrap' }}>{badge('blue', log.datasource_tag_used)}</td>
                  <td style={{ padding: '10px 16px', color: '#1e293b', minWidth: 200 }}>{log.subject}</td>
                  <td style={{ padding: '10px 16px', color: '#475569', whiteSpace: 'nowrap' }}>{log.recipient}</td>
                  <td style={{ padding: '10px 16px', whiteSpace: 'nowrap' }}>
                    {badge(log.status === 'completed' || log.status === 'success' ? 'green' : log.status === 'pending' ? 'orange' : 'red', log.status)}
                  </td>
                  <td style={{ padding: '10px 16px', whiteSpace: 'nowrap' }}>
                    {log.attachment_names && log.attachment_names.length > 0
                    ? log.attachment_names.map(function (name, i) { return (<span key={i} style={{ display: 'inline-flex', alignItems: 'center', gap: 4, background: '#f0f9ff', border: '1px solid #bae6fd', borderRadius: 6, padding: '2px 8px', fontSize: 11, color: '#0284c7', fontWeight: 600, marginRight: 4 }}>
                            📎 {name}
                          </span>); })
                    : <span style={{ color: '#94a3b8', fontSize: 11 }}>—</span>}
                  </td>
                </tr>); })}
            </tbody>
          </table>
        </div>
      </div>);
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
                        return [4 /*yield*/, this._loadFilesForVessel(vesselName)["catch"](function () { return undefined; })];
                    case 2:
                        _c.sent();
                        _c.label = 3;
                    case 3: return [4 /*yield*/, this._fetchJson("".concat(base_1, "/api/my-approvals?status=approved"))["catch"](function () {
                            return _this._fetchJson("".concat(base_1, "/api/approvals?status=approved"))["catch"](function () { return []; });
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
                                bentoApprovedFileIds: __assign(__assign({}, (prev.bentoApprovedFileIds || {})), fileIdUpdates_1)
                            });
                        });
                        return [3 /*break*/, 6];
                    case 5:
                        _b = _c.sent();
                        this.setState(function (prev) {
                            var _a;
                            return ({
                                bentoApprovedFiles: __assign(__assign({}, (prev.bentoApprovedFiles || {})), (_a = {}, _a[vLower] = [], _a))
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
        var _this = this;
        var _a = this.state, bentoComposeOpen = _a.bentoComposeOpen, bentoComposeForm = _a.bentoComposeForm, bentoComposeBusy = _a.bentoComposeBusy, bentoComposeMsg = _a.bentoComposeMsg, bentoComposeErr = _a.bentoComposeErr, vessels = _a.vessels, bentoLogs = _a.bentoLogs, bentoApprovedFiles = _a.bentoApprovedFiles;
        if (!bentoComposeOpen)
            return null;
        var FIXED_RECIPIENT = this.state.bentoConfigRecipient || this.props.userEmail || '';
        var setForm = function (patch, autoSubject) {
            var merged = __assign(__assign({}, bentoComposeForm), patch);
            // Auto-fill subject whenever vessel or tag changes
            if (autoSubject) {
                merged.subject_text = _this._buildAutoSubject(merged.vessel_name, merged.datasource_tag, merged.subject_suffix || 'Subject');
            }
            _this.setState({ bentoComposeForm: merged });
        };
        // Approved files available for the selected vessel
        var approvedFiles = this._getApprovedFilesForVesselTag(bentoComposeForm.vessel_name, bentoComposeForm.datasource_tag);
        var vLowerKey = bentoComposeForm.vessel_name.trim().toLowerCase();
        var isLoadingFiles = !!bentoComposeForm.vessel_name && !(vLowerKey in (bentoApprovedFiles || {}));
        var handleSend = function () { return __awaiter(_this, void 0, void 0, function () {
            var tempId, pendingLog, fd, attachFileId, fileRes, blob, _a, res, data_1, e_4;
            var _this = this;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (!bentoComposeForm.subject_text.trim()) {
                            this.setState({ bentoComposeErr: 'Subject is required.' });
                            return [2 /*return*/];
                        }
                        this.setState({ bentoComposeBusy: true, bentoComposeErr: null, bentoComposeMsg: null });
                        tempId = Date.now();
                        pendingLog = {
                            id: tempId,
                            datasource_tag_used: bentoComposeForm.datasource_tag,
                            tag_label: DATASOURCE_TAGS_MAP[bentoComposeForm.datasource_tag],
                            tag_was_valid: true,
                            vessel_name: bentoComposeForm.vessel_name || undefined,
                            subject: bentoComposeForm.subject_text,
                            body: bentoComposeForm.body,
                            recipient: FIXED_RECIPIENT,
                            status: 'pending',
                            attachments_count: bentoComposeForm.existing_attachment ? 1 : 0,
                            attachment_names: bentoComposeForm.existing_attachment ? [bentoComposeForm.existing_attachment] : [],
                            created_at: new Date().toISOString()
                        };
                        this.setState({ bentoLogs: __spreadArray([pendingLog], bentoLogs, true) });
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
                        attachFileId = this._getFileIdForAttachment(bentoComposeForm.vessel_name, bentoComposeForm.existing_attachment);
                        if (!attachFileId) {
                            this.setState(function (prev) { return ({
                                bentoComposeBusy: false,
                                bentoComposeErr: "Couldn't locate \"".concat(bentoComposeForm.existing_attachment, "\" as a real file \u2014 please reselect it from the list."),
                                bentoLogs: prev.bentoLogs.filter(function (l) { return l.id !== tempId; })
                            }); });
                            return [2 /*return*/];
                        }
                        fd.append('file_id', attachFileId);
                        _b.label = 2;
                    case 2:
                        _b.trys.push([2, 7, , 8]);
                        return [4 /*yield*/, fetch("".concat(this._base(), "/api/files/").concat(encodeURIComponent(attachFileId), "/content"), {
                                headers: this._uploadHeaders()
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
                    case 8: return [4 /*yield*/, fetch("".concat(this._base(), "/api/bento/dispatch"), {
                            method: 'POST', headers: this._uploadHeaders(), body: fd
                        })];
                    case 9:
                        res = _b.sent();
                        if (!res.ok)
                            throw new Error("HTTP ".concat(res.status));
                        return [4 /*yield*/, res.json()["catch"](function () { return ({}); })];
                    case 10:
                        data_1 = _b.sent();
                        // Update the pending log to completed
                        this.setState(function (prev) { return ({
                            bentoLogs: prev.bentoLogs.map(function (l) {
                                return l.id === tempId
                                    ? __assign(__assign({}, l), { id: (data_1 === null || data_1 === void 0 ? void 0 : data_1.id) || tempId, status: 'completed', sent_at: new Date().toISOString(), display_status: 'Sent' }) : l;
                            }),
                            bentoComposeBusy: false,
                            bentoComposeMsg: "\u2705 Email dispatched to ".concat(FIXED_RECIPIENT),
                            bentoComposeForm: { vessel_name: '', datasource_tag: 'mail', subject_text: '[DataSource:mail]', body: '', file: null, existing_attachment: '', recipient: FIXED_RECIPIENT }
                        }); });
                        setTimeout(function () { return _this.setState({ bentoComposeOpen: false, bentoComposeMsg: null }); }, 2000);
                        return [3 /*break*/, 12];
                    case 11:
                        e_4 = _b.sent();
                        // Mark log as failed
                        this.setState(function (prev) { return ({
                            bentoLogs: prev.bentoLogs.map(function (l) {
                                return l.id === tempId ? __assign(__assign({}, l), { status: 'failed', error_message: e_4 === null || e_4 === void 0 ? void 0 : e_4.message }) : l;
                            }),
                            bentoComposeBusy: false,
                            bentoComposeErr: "Failed: ".concat((e_4 === null || e_4 === void 0 ? void 0 : e_4.message) || 'Unknown error')
                        }); });
                        return [3 /*break*/, 12];
                    case 12: return [2 /*return*/];
                }
            });
        }); };
        return (<div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999 }} onClick={function () { return !bentoComposeBusy && _this.setState({ bentoComposeOpen: false }); }}>
        <div style={{ background: '#fff', borderRadius: 12, padding: '28px 32px', width: 540, maxWidth: '95vw', maxHeight: '90vh', overflowY: 'auto', boxShadow: '0 12px 40px rgba(0,0,0,0.25)' }} onClick={function (e) { return e.stopPropagation(); }}>
          <div style={{ fontSize: 17, fontWeight: 800, color: '#0f172a', marginBottom: 20, display: 'flex', alignItems: 'center', gap: 8 }}>
            ✉ Compose & Dispatch Email
          </div>

          {bentoComposeMsg && <div style={{ background: '#dff6dd', color: '#107c10', padding: '8px 12px', borderRadius: 6, fontSize: 12, marginBottom: 12 }}>{bentoComposeMsg}</div>}
          {bentoComposeErr && <div style={{ background: '#fde7e9', color: '#a4262c', padding: '8px 12px', borderRadius: 6, fontSize: 12, marginBottom: 12 }}>{bentoComposeErr}</div>}

          <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>

            {/* Recipient – fixed, read-only */}
            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 }}>Recipient (Fixed)</label>
              <input type="text" value={FIXED_RECIPIENT} readOnly style={{ width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, background: '#f8fafc', color: '#64748b', boxSizing: 'border-box', cursor: 'not-allowed' }}/>
            </div>

            {/* Vessel Dropdown */}
            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 }}>Vessel</label>
              <select value={bentoComposeForm.vessel_name} onChange={function (e) {
                var selectedVesselName = e.target.value;
                setForm({ vessel_name: selectedVesselName, existing_attachment: '' }, true);
                if (selectedVesselName) {
                    // Fetch approved files directly from backend
                    _this._fetchApprovedFilesForVessel(selectedVesselName)["catch"](function () { return undefined; });
                    // Also load via existing mechanism as fallback
                    if (!_this._filesLoadedForVessels.has(selectedVesselName)) {
                        _this._filesLoadedForVessels.add(selectedVesselName);
                        _this._loadFilesForVessel(selectedVesselName)["catch"](function () { return undefined; });
                    }
                }
            }} disabled={bentoComposeBusy} style={{ width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, background: '#fff', outline: 'none' }}>
                <option value="">— Select vessel —</option>
                {vessels.map(function (v) { return <option key={v.id} value={v.name}>{v.name}</option>; })}
              </select>
            </div>

            {/* Document Tag Dropdown */}
            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 }}>Document Tag</label>
              <select value={bentoComposeForm.datasource_tag} onChange={function (e) { return setForm({ datasource_tag: e.target.value }, true); }} disabled={bentoComposeBusy} style={{ width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, background: '#fff', outline: 'none' }}>
                {Object.keys(DATASOURCE_TAGS_MAP).map(function (k) { return (<option key={k} value={k}>{DATASOURCE_TAGS_MAP[k]} ({k})</option>); })}
              </select>
            </div>

            {/* Subject – auto-filled, editable [DataSource:TAG] Vessel Name / Subject */}
            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 }}>
                Subject * <span style={{ fontWeight: 400, color: '#94a3b8', fontSize: 11 }}>(auto-filled, editable)</span>
              </label>
              <input type="text" value={bentoComposeForm.subject_text} onChange={function (e) {
                var val = e.target.value;
                var suffix = val;
                var slashIdx = val.indexOf(' / ');
                if (slashIdx !== -1) {
                    suffix = val.substring(slashIdx + 3);
                }
                setForm({ subject_text: val, subject_suffix: suffix });
            }} disabled={bentoComposeBusy} placeholder="[DataSource:TAG] Vessel Name / Subject..." style={{ width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #0284c7', fontSize: 13, outline: 'none', boxSizing: 'border-box', background: '#f0f9ff' }}/>
            </div>

            {/* Body */}
            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 }}>Body</label>
              <textarea value={bentoComposeForm.body} onChange={function (e) { return setForm({ body: e.target.value }); }} disabled={bentoComposeBusy} placeholder="Email body..." rows={4} style={{ width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, outline: 'none', boxSizing: 'border-box', resize: 'vertical' }}/>
            </div>

            {/* Attach Approved File Dropdown */}
            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4 }}>
                Attach Approved File
                <span style={{ fontWeight: 400, color: '#94a3b8', fontSize: 11, marginLeft: 6 }}>(from selected vessel documents)</span>
              </label>
              <select value={bentoComposeForm.existing_attachment || ''} onChange={function (e) { return setForm({ existing_attachment: e.target.value }); }} disabled={bentoComposeBusy || !bentoComposeForm.vessel_name || isLoadingFiles} style={{ width: '100%', padding: '8px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, background: bentoComposeForm.vessel_name ? '#fff' : '#f8fafc', outline: 'none' }}>
                {!bentoComposeForm.vessel_name ? (<option value="">— Select Vessel First —</option>) : isLoadingFiles ? (<option value="">⏳ Loading files... —</option>) : approvedFiles.length > 0 ? (<>
                    <option value="">— Select Attached File —</option>
                    {approvedFiles.map(function (f) { return <option key={f} value={f}>📎 {f}</option>; })}
                  </>) : (<option value="">— No approved files found for this vessel —</option>)}
              </select>
            </div>

            {/* Status indicator */}
            <div style={{ background: '#f8fafc', borderRadius: 6, padding: '8px 12px', fontSize: 11, color: '#64748b', border: '1px solid #e2e8f0' }}>
              📊 Status flow: <strong>Pending</strong> (on upload) → <strong>Pending</strong> (awaiting send) → <strong>Completed</strong> (after dispatch)
            </div>
          </div>

          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 20 }}>
            <button onClick={function () { return _this.setState({ bentoComposeOpen: false, bentoComposeErr: null, bentoComposeMsg: null }); }} disabled={bentoComposeBusy} style={{ background: '#f1f5f9', border: '1px solid #cbd5e1', borderRadius: 6, padding: '8px 16px', fontSize: 13, cursor: 'pointer' }}>Cancel</button>
            <button onClick={handleSend} disabled={bentoComposeBusy} style={{ background: '#0284c7', color: '#fff', border: 'none', borderRadius: 6, padding: '8px 20px', fontSize: 13, fontWeight: 700, cursor: 'pointer' }}>
              {bentoComposeBusy ? '⏳ Sending...' : '✉ Send Email'}
            </button>
          </div>
        </div>
      </div>);
    };
    VesselEmail.prototype._renderRecycleBinPage = function () {
        var _this = this;
        var _a = this.state, recycleBin = _a.recycleBin, panelLoading = _a.panelLoading;
        var vesselItems = recycleBin.filter(function (r) { return r.kind === 'vessel' || r.item_type === 'vessel'; });
        var otherItems = recycleBin.filter(function (r) { return r.kind !== 'vessel' && r.item_type !== 'vessel'; });
        var renderTable = function (items, emptyMsg) { return (<div style={{ background: '#fff', borderRadius: 8, border: '1px solid #e2e8f0', overflow: 'hidden' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
          <thead>
            <tr style={{ background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 11, textTransform: 'uppercase', textAlign: 'left' }}>
              <th style={{ padding: '10px 16px' }}>Name</th>
              <th style={{ padding: '10px 16px' }}>Type</th>
              <th style={{ padding: '10px 16px' }}>Original Path</th>
              <th style={{ padding: '10px 16px' }}>Deleted At</th>
              <th style={{ padding: '10px 16px', textAlign: 'right' }}>Actions</th>
            </tr>
          </thead>
          <tbody>
            {items.length === 0 ? (<tr><td colSpan={5} style={{ padding: 20, textAlign: 'center', color: '#94a3b8', fontSize: 12 }}>{emptyMsg}</td></tr>) : items.map(function (item) { return (<tr key={item.id} style={{ borderBottom: '1px solid #f1f5f9' }} onMouseEnter={function (e) { return (e.currentTarget.style.background = '#f8fafc'); }} onMouseLeave={function (e) { return (e.currentTarget.style.background = ''); }}>
                <td style={{ padding: '10px 16px', fontWeight: 600, color: '#0f172a' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                    <span>{item.kind === 'vessel' || item.item_type === 'vessel' ? 'ðŸš¢' : item.kind === 'folder' ? 'ðŸ“' : 'ðŸ“„'}</span>
                    {item.name}
                  </div>
                </td>
                <td style={{ padding: '10px 16px' }}>
                  {badge(item.kind === 'vessel' || item.item_type === 'vessel' ? 'blue' : item.kind === 'folder' ? 'orange' : 'default', item.kind || item.item_type || 'â€”')}
                </td>
                <td style={{ padding: '10px 16px', color: '#64748b', fontSize: 11, maxWidth: 200, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={item.original_path || 'â€”'}>
                  {item.original_path || 'â€”'}
                </td>
                <td style={{ padding: '10px 16px', color: '#64748b', fontSize: 12, whiteSpace: 'nowrap' }}>
                  {item.deleted_at ? new Date(item.deleted_at).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }) : 'â€”'}
                </td>
                <td style={{ padding: '10px 16px', textAlign: 'right' }}>
                  <div style={{ display: 'flex', gap: 6, justifyContent: 'flex-end' }}>
                    <button onClick={function () { return _this._restoreFromRecycleBin(item); }} style={{ background: '#dff6dd', color: '#107c10', border: 'none', borderRadius: 6, padding: '5px 12px', fontSize: 12, fontWeight: 600, cursor: 'pointer', whiteSpace: 'nowrap' }}>
                      â†© Restore
                    </button>
                    <button onClick={function () { return _this._permanentDeleteFromRecycleBin(item); }} style={{ background: '#fde7e9', color: '#a4262c', border: 'none', borderRadius: 6, padding: '5px 12px', fontSize: 12, fontWeight: 600, cursor: 'pointer', whiteSpace: 'nowrap' }}>
                      ðŸ—‘ Delete
                    </button>
                  </div>
                </td>
              </tr>); })}
          </tbody>
        </table>
      </div>); };
        return (<div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div style={{ background: '#fff4ce', border: '1px solid #f7cf72', borderRadius: 8, padding: '16px 20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div>
            <h2 style={{ margin: 0, fontSize: 18, fontWeight: 700, color: '#8a5700' }}>ðŸ—‘ï¸ Recycle Bin</h2>
            <p style={{ margin: '4px 0 0', fontSize: 12, color: '#8a5700', opacity: 0.9 }}>
              {recycleBin.length} item{recycleBin.length !== 1 ? 's' : ''} â€” Restore to reactivate or permanently delete.
            </p>
          </div>
          <button onClick={function () { return _this._goToView('recycle'); }} style={{ background: '#fff', border: '1px solid #f7cf72', borderRadius: 6, padding: '6px 14px', fontSize: 12, fontWeight: 600, color: '#8a5700', cursor: 'pointer' }}>
            ðŸ”„ Refresh
          </button>
        </div>

        {panelLoading ? (<div style={{ padding: 32, textAlign: 'center', color: '#64748b', fontSize: 13 }}>â³ Loading recycle bin...</div>) : (<>
            {vesselItems.length > 0 && (<div>
                <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: '0.05em', color: '#94a3b8', textTransform: 'uppercase', marginBottom: 8 }}>Deleted Vessels</div>
                {renderTable(vesselItems, 'No deleted vessels.')}
              </div>)}
            <div>
              {vesselItems.length > 0 && (<div style={{ fontSize: 11, fontWeight: 700, letterSpacing: '0.05em', color: '#94a3b8', textTransform: 'uppercase', marginBottom: 8 }}>Other Items</div>)}
              {renderTable(otherItems, recycleBin.length === 0 ? 'Recycle bin is empty.' : 'No other items.')}
            </div>
          </>)}
      </div>);
    };
    VesselEmail.prototype._renderArchivePage = function () {
        var archiveList = this.state.archiveList;
        return (<div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div style={{ background: '#ede8f5', border: '1px solid #b4a0d4', borderRadius: 8, padding: '16px 20px' }}>
          <h2 style={{ margin: 0, fontSize: 18, fontWeight: 700, color: '#5c2d91' }}>📦 Archive</h2>
          <p style={{ margin: '4px 0 0', fontSize: 12, color: '#5c2d91', opacity: 0.9 }}>
            Archived items stored for historical record-keeping.
          </p>
        </div>

        <div style={{ background: '#fff', borderRadius: 8, border: '1px solid #e2e8f0', overflow: 'hidden' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
            <thead>
              <tr style={{ background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 11, textTransform: 'uppercase', textAlign: 'left' }}>
                <th style={{ padding: '10px 16px' }}>Name</th>
                <th style={{ padding: '10px 16px' }}>Type</th>
                <th style={{ padding: '10px 16px' }}>Archived At</th>
              </tr>
            </thead>
            <tbody>
              {archiveList.length === 0 ? (<tr>
                  <td colSpan={3} style={{ padding: 20, textAlign: 'center', color: '#64748b' }}>No items archived.</td>
                </tr>) : (archiveList.map(function (item) { return (<tr key={item.id} style={{ borderBottom: '1px solid #f1f5f9' }}>
                    <td style={{ padding: '10px 16px', fontWeight: 600 }}>{item.name}</td>
                    <td style={{ padding: '10px 16px' }}>{item.kind}</td>
                    <td style={{ padding: '10px 16px' }}>{item.archived_at || '—'}</td>
                  </tr>); }))}
            </tbody>
          </table>
        </div>
      </div>);
    };
    // ── Render Vessel Modals ──────────────────────────────────────────────────
    VesselEmail.prototype._renderVesselForm = function (mode) {
        var _this = this;
        var _a = this.state, form = _a.form, modalBusy = _a.modalBusy, modalMsg = _a.modalMsg, modalError = _a.modalError;
        var set = function (k) { return function (e) {
            var _a;
            return _this.setState({ form: __assign(__assign({}, form), (_a = {}, _a[k] = e.target.value, _a)), modalError: null });
        }; };
        var isCreate = mode === 'create';
        return (<div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999 }} onClick={this._closeModal}>
        <div style={{ background: '#fff', borderRadius: 16, padding: '32px 36px', width: 480, maxWidth: '92vw', boxShadow: '0 12px 40px rgba(0,0,0,0.18)' }} onClick={function (e) { return e.stopPropagation(); }}>

          {/* Modal Header */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 24 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 14 }}>
              <div style={{ width: 44, height: 44, borderRadius: 10, background: '#e0f2fe', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 22 }}>
                🚢
              </div>
              <div>
                <div style={{ fontSize: 18, fontWeight: 700, color: '#0f172a' }}>{isCreate ? 'New Vessel' : 'Update Vessel'}</div>
                <div style={{ fontSize: 12, color: '#64748b', marginTop: 2, maxWidth: 280 }}>
                  {isCreate ? 'Provisions the full folder structure across all 3 main folders.' : 'Update the vessel details below.'}
                </div>
              </div>
            </div>
            <button onClick={this._closeModal} disabled={modalBusy} style={{ background: 'none', border: 'none', fontSize: 18, color: '#94a3b8', cursor: 'pointer', lineHeight: 1, padding: 4 }}>✕</button>
          </div>

          {modalMsg && <div style={{ background: '#dff6dd', color: '#107c10', padding: '8px 12px', borderRadius: 6, fontSize: 12, marginBottom: 14 }}>{modalMsg}</div>}
          {modalError && <div style={{ background: '#fde7e9', color: '#a4262c', padding: '8px 12px', borderRadius: 6, fontSize: 12, marginBottom: 14 }}>{modalError}</div>}

          {/* Fields */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
            <div>
              <label style={{ display: 'block', fontSize: 13, fontWeight: 600, color: '#334155', marginBottom: 6 }}>
                Vessel name <span style={{ color: '#ef4444' }}>*</span>
              </label>
              <input style={{ width: '100%', padding: '10px 14px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 14, outline: 'none', boxSizing: 'border-box', color: '#0f172a' }} value={form.name} onChange={set('name')} placeholder="e.g. MV Pacific Trader" disabled={modalBusy}/>
            </div>

            <div>
              <label style={{ display: 'block', fontSize: 13, fontWeight: 600, color: '#334155', marginBottom: 6 }}>
                IMO number <span style={{ color: '#ef4444' }}>*</span>
              </label>
              <input style={{ width: '100%', padding: '10px 14px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 14, outline: 'none', boxSizing: 'border-box', color: '#0f172a' }} value={form.imo} onChange={function (e) {
                var val = e.target.value.replace(/\D/g, '').slice(0, 7);
                _this.setState({ form: __assign(__assign({}, form), { imo: val }), modalError: null });
            }} placeholder="7 digits, e.g. 9074729" maxLength={7} inputMode="numeric" disabled={modalBusy}/>
            </div>

            <div>
              <label style={{ display: 'block', fontSize: 13, fontWeight: 600, color: '#334155', marginBottom: 6 }}>Ship yard name</label>
              <input style={{ width: '100%', padding: '10px 14px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 14, outline: 'none', boxSizing: 'border-box', color: '#0f172a' }} value={form.shipyard} onChange={set('shipyard')} placeholder="e.g. Hyundai Heavy Industries" disabled={modalBusy}/>
            </div>

            <div>
              <label style={{ display: 'block', fontSize: 13, fontWeight: 600, color: '#334155', marginBottom: 6 }}>Hull number</label>
              <input style={{ width: '100%', padding: '10px 14px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 14, outline: 'none', boxSizing: 'border-box', color: '#0f172a' }} value={form.hull_number} onChange={set('hull_number')} placeholder="e.g. H2456" disabled={modalBusy}/>
            </div>

            <div>
              <label style={{ display: 'block', fontSize: 13, fontWeight: 600, color: '#334155', marginBottom: 6 }}>Vessel type</label>
              <select style={{ width: '100%', padding: '10px 14px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 14, outline: 'none', background: '#fff', color: form.vessel_type ? '#0f172a' : '#94a3b8', appearance: 'auto' }} value={form.vessel_type} onChange={set('vessel_type')} disabled={modalBusy}>
                <option value="">Select a type...</option>
                {VESSEL_TYPES.map(function (t) { return <option key={t} value={t}>{t}</option>; })}
              </select>
            </div>
          </div>

          {/* Footer */}
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 28 }}>
            <button style={{ background: 'transparent', border: 'none', borderRadius: 8, padding: '10px 20px', fontSize: 14, fontWeight: 500, color: '#475569', cursor: 'pointer' }} onClick={this._closeModal} disabled={modalBusy}>Cancel</button>
            <button style={{ background: 'linear-gradient(135deg, #0d9488, #0f766e)', color: '#fff', border: 'none', borderRadius: 8, padding: '10px 24px', fontSize: 14, fontWeight: 600, cursor: 'pointer' }} onClick={isCreate ? this._submitCreate : this._submitEdit} disabled={modalBusy}>
              {modalBusy ? 'Saving...' : isCreate ? 'Create Vessel' : 'Update Vessel'}
            </button>
          </div>
        </div>
      </div>);
    };
    VesselEmail.prototype._renderDeleteModal = function () {
        var _a = this.state, selectedVessel = _a.selectedVessel, modalBusy = _a.modalBusy, modalMsg = _a.modalMsg, modalError = _a.modalError;
        return (<div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999 }} onClick={this._closeModal}>
        <div style={{ background: '#fff', borderRadius: 10, padding: '24px 28px', width: 400, maxWidth: '90vw', boxShadow: '0 8px 32px rgba(0,0,0,0.2)' }} onClick={function (e) { return e.stopPropagation(); }}>
          <div style={{ fontSize: 16, fontWeight: 700, marginBottom: 12 }}>🗑 Delete Vessel</div>
          <p style={{ fontSize: 13, color: '#475569', marginBottom: 16 }}>
            Are you sure you want to delete <strong>"{selectedVessel === null || selectedVessel === void 0 ? void 0 : selectedVessel.name}"</strong>?
          </p>
          {modalMsg && <div style={{ background: '#dff6dd', color: '#107c10', padding: 8, borderRadius: 6, fontSize: 12, marginBottom: 12 }}>{modalMsg}</div>}
          {modalError && <div style={{ background: '#fde7e9', color: '#a4262c', padding: 8, borderRadius: 6, fontSize: 12, marginBottom: 12 }}>{modalError}</div>}
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
            <button style={{ background: '#f1f5f9', border: '1px solid #cbd5e1', borderRadius: 6, padding: '8px 16px', fontSize: 13, cursor: 'pointer' }} onClick={this._closeModal} disabled={modalBusy}>Cancel</button>
            <button style={{ background: '#ef4444', color: '#fff', border: 'none', borderRadius: 6, padding: '8px 16px', fontSize: 13, fontWeight: 600, cursor: 'pointer' }} onClick={this._submitDelete} disabled={modalBusy}>
              {modalBusy ? 'Deleting...' : 'Delete'}
            </button>
          </div>
        </div>
      </div>);
    };
    // ── Main Render Router ────────────────────────────────────────────────────
    VesselEmail.prototype.render = function () {
        var _a = this.state, sessionExpired = _a.sessionExpired, view = _a.view;
        if (sessionExpired) {
            return (<div style={{
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                    minHeight: '100vh', background: '#f8fafc', fontFamily: "'Segoe UI', sans-serif"
                }}>
          <div style={{
                    background: '#fff', borderRadius: 12, padding: '40px 48px', maxWidth: 420,
                    textAlign: 'center', boxShadow: '0 4px 24px rgba(0,0,0,0.10)', border: '1px solid #e2e8f0'
                }}>
            <div style={{ fontSize: 48, marginBottom: 16 }}>🔒</div>
            <h2 style={{ margin: '0 0 8px', fontSize: 20, fontWeight: 700, color: '#0f172a' }}>Session Expired</h2>
            <p style={{ margin: '0 0 24px', fontSize: 14, color: '#64748b', lineHeight: 1.6 }}>
              Your session has expired or is no longer valid. Please sign out and sign back in to continue.
            </p>
            <button onClick={this._handleSignOut} style={{
                    background: '#0078d4', color: '#fff', border: 'none', borderRadius: 8,
                    padding: '10px 28px', fontSize: 14, fontWeight: 600, cursor: 'pointer'
                }}>
              Sign Out & Reload
            </button>
          </div>
        </div>);
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
        return (<>
        {this._renderLayout(content)}
        {this._renderBentoComposeModal()}
      </>);
    };
    return VesselEmail;
}(React.Component));
exports["default"] = VesselEmail;
