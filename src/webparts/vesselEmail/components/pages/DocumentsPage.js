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
exports.renderDocumentsPage = renderDocumentsPage;
var React = __importStar(require("react"));
var VesselsPage_1 = require("./VesselsPage");
function renderDocumentsPage(host) {
    var _this = this;
    var _a = host.state, textFilter = _a.textFilter, vesselFilter = _a.vesselFilter, catFilter = _a.catFilter, docViewMode = _a.docViewMode, showAllVesselsInFolderView = _a.showAllVesselsInFolderView, vessels = _a.vessels, rows = _a.rows, docListPage = _a.docListPage, docListSort = _a.docListSort, docGroupFilter = _a.docGroupFilter, documentVesselCount = _a.documentVesselCount, documentVesselsLoadingMore = _a.documentVesselsLoadingMore, docUploadRowKey = _a.docUploadRowKey, docUploadBusy = _a.docUploadBusy, docUploadMsg = _a.docUploadMsg, documentsList = _a.documentsList, folderPathStack = _a.folderPathStack, uploadedFilesByFolder = _a.uploadedFilesByFolder, docMainFolder = _a.docMainFolder, folderNavHistory = _a.folderNavHistory, folderNavIndex = _a.folderNavIndex;
    var canGoBack = folderNavIndex > 0;
    var canGoForward = folderNavIndex < folderNavHistory.length - 1;
    var goBack = function () {
        if (!canGoBack) return;
        var prev = folderNavHistory[folderNavIndex - 1];
        host.setState({ folderNavIndex: folderNavIndex - 1, folderPathStack: prev.folderPathStack, docMainFolder: prev.docMainFolder });
    };
    var goForward = function () {
        if (!canGoForward) return;
        var next = folderNavHistory[folderNavIndex + 1];
        host.setState({ folderNavIndex: folderNavIndex + 1, folderPathStack: next.folderPathStack, docMainFolder: next.docMainFolder });
    };
    var MAIN_FOLDERS = [
        { key: 'Technical & Crewing', icon: '⚙️', emoji: '⚙️', color: '#dc2626', bg: '#fee2e2' },
        { key: 'Commercial & Chartering', icon: '💼', emoji: '💼', color: '#16a34a', bg: '#dcfce7' },
        { key: 'Insurance', icon: '🛡️', emoji: '🛡️', color: '#d97706', bg: '#fef3c7' },
        { key: 'Kaizen - Knowledge Bank', icon: '📚', emoji: '📚', color: '#7c3aed', bg: '#ede9fe' },
    ];
    // ── Per main-folder subfolder structure (aligned with template.py DMS specification) ──
    var DEFAULT_VESSEL_MAINS = {
        'Technical & Crewing': ['Month End Reports', 'Service Agreements', 'Registration', 'Drawings and Manuals', 'PO & Invoice', 'Incidents', 'Crewing', 'To be Classified'],
        'Commercial & Chartering': ['Agreements', 'Invoices & Payments', 'Claims & Disputes', 'To be Classified'],
        'Insurance': ['P&I', 'H&M', 'War Risk', 'Flag & MPA', 'USA Related'],
        'Kaizen - Knowledge Bank': ['Templates', 'Procedures and Work Instructions', 'Lessons Learned', 'Circulars and Guidance'],
        'Knowledge Bank': ['Templates', 'Procedures and Work Instructions', 'Lessons Learned', 'Circulars and Guidance'],
    };
    var SUBFOLDERS_MAP = {
        'Month End Reports': ['Main Engine', 'Aux Engine', 'Cooling Water', 'Inspection Reports', 'Defect Reports', 'Guarantee Claims', 'To be Classified'],
        'Service Agreements': ['Technical Management', 'Crew Management', 'Vendor & Service Provider', 'To be Classified'],
        'Registration': ['Flag & MPA', 'Ship Builder', 'Radio & Telecom', 'Crewing & SMOU', 'Novation', 'To be Classified'],
        'Drawings and Manuals': ['Drawing', 'Manual', 'To be Classified'],
        'PO & Invoice': ['Purchase Order and Vendor Invoice'],
        'Agreements': ['Charter party', 'Pool Agreement', 'Commission Agreement', 'To be Classified'],
        'Invoices & Payments': ['Invoice', 'Payments', 'To be Classified'],
        'Claims & Disputes': ['Disputes', 'Claims', 'To be Classified'],
        'Circulars and Guidance': ['Equipment Maker', 'Class', 'Flag / Port State', 'SIRE/OCIMF/RightShip', 'Shipyard'],
        'Common for all ships': ['Vendor & Service Agreements', 'Vendor Management'],
    };
    var stackLevel = folderPathStack.length;
    var atVesselsRoot = stackLevel === 1 && folderPathStack[0].id === 'vessels_root';
    var atSpecificVessels = stackLevel === 2 && folderPathStack[1].id === 'specific_vessels';
    var KNOWN_NAV_IDS = new Set(['vessels_root', 'specific_vessels', 'common', 'kaizen_root']);
    var VESSEL_MAIN_FOLDER_KEYS = ['Technical & Crewing', 'Commercial & Chartering', 'Insurance'];
    var vesselStackIdx = (function () {
        // Standard path: vessels_root > specific_vessels > {vessel}
        if (stackLevel >= 3 && folderPathStack[0].id === 'vessels_root' && folderPathStack[1].id === 'specific_vessels') {
            return 2;
        }
        // Fallback: find first node that is not a known nav node and not a main folder name
        return folderPathStack.findIndex(function (n) {
            return !KNOWN_NAV_IDS.has(n.id) && !VESSEL_MAIN_FOLDER_KEYS.includes(n.name);
        });
    })();
    var vesselNodeInStack = vesselStackIdx !== -1 ? folderPathStack[vesselStackIdx] : null;
    var atVesselMainFolderSelect = vesselNodeInStack !== null && !docMainFolder && stackLevel === vesselStackIdx + 1;
    var currentVesselNameFromStack = vesselNodeInStack ? vesselNodeInStack.name : null;
    var currentFolderNode = folderPathStack.length > 0 ? folderPathStack[folderPathStack.length - 1] : null;
    var currentFolderName = currentFolderNode ? currentFolderNode.name : null;
    if (currentVesselNameFromStack && !host._filesLoadedForVessels.has(currentVesselNameFromStack)) {
        host._filesLoadedForVessels.add(currentVesselNameFromStack);
        setTimeout(function () { return host._loadFilesForVessel(currentVesselNameFromStack).catch(function () { return undefined; }); }, 0);
    }
    // Resolve real SPO folder ID for the current folder node from rows
    var resolvedCurrentFolderId = (function () {
        if (!currentFolderNode)
            return null;
        if (!/^(sf_|common|vessels_root|specific_vessels|kaizen_root)/.test(currentFolderNode.id))
            return currentFolderNode.id;
        var vesselName = currentVesselNameFromStack || (folderPathStack.length > 0 ? folderPathStack[0].name : null);
        if (!vesselName)
            return null;
        var folderName = currentFolderNode.name;
        var matchRow = host.state.rows.find(function (r) {
            return r.vesselName === vesselName &&
                r.uploadFolderId &&
                !r.uploadFolderId.includes('/') &&
                (r.subCategory === folderName || r.category === folderName);
        });
        return (matchRow === null || matchRow === void 0 ? void 0 : matchRow.uploadFolderId) || null;
    })();
    // Trigger live SPO file refresh when we have a real folder ID
    if (resolvedCurrentFolderId && !/^(sf_|common)/.test(resolvedCurrentFolderId) && !/^\d+$/.test(resolvedCurrentFolderId)) {
        var alreadyLoaded = host._filesLoadedForFolders.has(resolvedCurrentFolderId);
        if (!alreadyLoaded) {
            setTimeout(function () { return host._refreshFolderFiles(resolvedCurrentFolderId, resolvedCurrentFolderId, true).catch(function () { return undefined; }); }, 0);
        }
    }
    // Determine subfolders for current depth
    var subfolderNames = [];
    if (currentFolderName && SUBFOLDERS_MAP[currentFolderName]) {
        subfolderNames = SUBFOLDERS_MAP[currentFolderName];
    } else if (docMainFolder && DEFAULT_VESSEL_MAINS[docMainFolder]) {
        if (!currentFolderName || currentFolderName === docMainFolder ||
            (vesselNodeInStack && currentFolderNode && currentFolderNode.id === vesselNodeInStack.id) ||
            atVesselsRoot || atSpecificVessels || folderPathStack.length <= 3) {
            subfolderNames = DEFAULT_VESSEL_MAINS[docMainFolder];
        }
    }
    var currentFolderFiles = (resolvedCurrentFolderId ? (uploadedFilesByFolder[resolvedCurrentFolderId] || []) : [])
        .filter(function (f, i, arr) { return arr.findIndex(function (x) { return x.name === f.name; }) === i; });
    // Also include approved files from backend rows that match the current folder node
    var backendFolderFiles = [];
    if (currentFolderNode) {
        var matchIds_1 = new Set();
        if (!/^(sf_|common)/.test(currentFolderNode.id))
            matchIds_1.add(currentFolderNode.id);
        if (resolvedCurrentFolderId)
            matchIds_1.add(resolvedCurrentFolderId);
        rows.filter(function (r) { return matchIds_1.has(r.uploadFolderId) && r.fileName && !r.filePending; })
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
        ],
    };
    // Vessels are returned newest first. Start with four and extend in pages
    // of eight when the user chooses "More vessels".
    var visibleVesselNames = new Set(vessels.slice(0, documentVesselCount).map(function (v) { return v.name; }));
    var filtered = allRows.filter(function (r) {
        // Keep the list aligned with the vessels currently loaded for Documents.
        if (!activeVesselName) {
            if (!visibleVesselNames.has(r.vesselName))
                return false;
        }
        // Filter by active main folder
        if (docMainFolder) {
            var allowedCats = mainFolderGroupMap[docMainFolder] || [];
            var mainMatch = r.group === docMainFolder ||
                r.group.toLowerCase().includes(docMainFolder.toLowerCase().split(' ')[0].toLowerCase()) ||
                allowedCats.includes(r.group) ||
                allowedCats.includes(r.category) ||
                (r.subFolderPath || '').toLowerCase().includes(docMainFolder.toLowerCase().split(' ')[0].toLowerCase());
            if (!mainMatch)
                return false;
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
                r.category.toLowerCase().includes(q) || (r.subCategory || '').toLowerCase().includes(q) ||
                (r.subFolderPath || '').toLowerCase().includes(q) ||
                (r.fileName || '').toLowerCase().includes(q);
        }
        return true;
    });
    // ── Group filtered rows by groupKey (category/folder level) ──
    var groupedMap = new Map();
    var _loop_1 = function (r) {
        var folderUploads = (uploadedFilesByFolder[r.groupKey] || [])
            .concat(uploadedFilesByFolder[r.uploadFolderId] || []);
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
                subCategory: r.subCategory || r.category,
                subFolderPath: r.subFolderPath,
                groupKey: r.groupKey,
                uploadFolderId: r.uploadFolderId,
                monthDriven: r.monthDriven,
                canUpload: r.canUpload,
                files: files,
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
    // Navigation folder names that should never appear as vessel cards
    var NAV_FOLDER_NAMES = new Set(['specific vessels', 'common for all ships', 'vessels', 'kaizen - knowledge bank', 'knowledge bank', 'common']);
    // Deduplicate by name and exclude any vessel whose name matches a navigation folder
    var displayVessels = vessels.slice(0, documentVesselCount).filter(function (v, i, arr) {
        return arr.findIndex(function (x) { return x.name.trim().toLowerCase() === v.name.trim().toLowerCase(); }) === i &&
            !NAV_FOLDER_NAMES.has(v.name.trim().toLowerCase());
    });
    return (React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 16 } },
        React.createElement("div", { style: { fontSize: 12, color: '#64748b', display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' } },
            React.createElement("button", { onClick: goBack, disabled: !canGoBack, style: { width: 28, height: 28, borderRadius: 6, border: '1px solid #cbd5e1', background: canGoBack ? '#fff' : '#f1f5f9', color: canGoBack ? '#334155' : '#cbd5e1', cursor: canGoBack ? 'pointer' : 'not-allowed', fontSize: 14, display: 'inline-flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 } }, '\u2039'),
            React.createElement("button", { onClick: goForward, disabled: !canGoForward, style: { width: 28, height: 28, borderRadius: 6, border: '1px solid #cbd5e1', background: canGoForward ? '#fff' : '#f1f5f9', color: canGoForward ? '#334155' : '#cbd5e1', cursor: canGoForward ? 'pointer' : 'not-allowed', fontSize: 14, display: 'inline-flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 } }, '\u203A'),
            React.createElement("span", { style: { cursor: 'pointer', color: '#0284c7' }, onClick: function () { return host._pushFolderNav([], null); } }, "Home"),
            React.createElement("span", null, "\u203A"),
            React.createElement("span", { style: { cursor: stackLevel === 0 && !docMainFolder ? 'default' : 'pointer', color: stackLevel === 0 && !docMainFolder ? '#0f172a' : '#0284c7', fontWeight: stackLevel === 0 && !docMainFolder ? 600 : 400 }, onClick: function () { return host._pushFolderNav([], null); } }, "Documents"),
            folderPathStack.map(function (item, idx) {
                var isLast = idx === folderPathStack.length - 1 && !docMainFolder;
                return (React.createElement(React.Fragment, { key: item.id + idx },
                    React.createElement("span", null, "\u203A"),
                    React.createElement("span", { onClick: function () { return host._pushFolderNav(folderPathStack.slice(0, idx + 1), null); }, style: { cursor: isLast ? 'default' : 'pointer', color: isLast ? '#0f172a' : '#0284c7', fontWeight: isLast ? 600 : 400 } }, item.name)));
            }),
            docMainFolder && (React.createElement(React.Fragment, null,
                React.createElement("span", null, "\u203A"),
                React.createElement("span", { style: { color: '#0f172a', fontWeight: 600 } }, docMainFolder)))),
        React.createElement("div", { style: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 12 } },
            React.createElement("div", null,
                React.createElement("h2", { style: { margin: 0, fontSize: 20, fontWeight: 700, color: '#0f172a', display: 'flex', alignItems: 'center', gap: 8 } },
                    React.createElement("span", { style: { fontSize: 20 } }, "\uD83D\uDCC1"),
                    currentFolderName || docMainFolder || 'Documents'),
                React.createElement("p", { style: { margin: '4px 0 0', fontSize: 13, color: '#64748b' } }, docViewMode === 'list'
                    ? "".concat(filtered.length, " rows \u00B7 flattened list view")
                    : !docMainFolder
                        ? "".concat(MAIN_FOLDERS.length, " main folders")
                        : folderPathStack.length === 0
                            ? "".concat(vessels.length, " vessels \u00B7 folder view")
                            : "".concat(subfolderNames.length, " folders \u00B7 ").concat(allCurrentFolderFiles.length, " files"))),
            React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 10 } },
                React.createElement("div", { style: { display: 'inline-flex', background: '#fff', border: '1px solid #cbd5e1', borderRadius: 8, padding: 3, gap: 2 } },
                    React.createElement("button", { onClick: function () { return host.setState({ docViewMode: 'folder' }); }, style: {
                            padding: '5px 12px', borderRadius: 6, border: 'none', fontSize: 12, fontWeight: 600, cursor: 'pointer',
                            background: docViewMode === 'folder' ? '#0f172a' : 'transparent',
                            color: docViewMode === 'folder' ? '#fff' : '#64748b',
                        } }, "\u229E Folder view"),
                    React.createElement("button", { onClick: function () { host.setState({ docViewMode: 'list' }); host._pushFolderNav([], null); }, style: {
                            padding: '5px 12px', borderRadius: 6, border: 'none', fontSize: 12, fontWeight: 600, cursor: 'pointer',
                            background: docViewMode === 'list' ? '#0f172a' : 'transparent',
                            color: docViewMode === 'list' ? '#fff' : '#64748b',
                        } }, "\u2630 List view")),
                React.createElement("button", { onClick: function () { return host._goToView('archive'); }, style: {
                        background: '#d97706', color: '#fff', border: 'none', borderRadius: 8,
                        padding: '7px 16px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
                        display: 'inline-flex', alignItems: 'center', gap: 6,
                    } }, "\uD83D\uDDD1 Archive"),
                React.createElement("label", { style: {
                        background: '#0284c7', color: '#fff', border: 'none', borderRadius: 8,
                        padding: '7px 16px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
                        display: 'inline-flex', alignItems: 'center', gap: 6, boxShadow: '0 2px 4px rgba(2,132,199,0.2)',
                    } },
                    React.createElement("input", { type: "file", style: { display: 'none' }, onChange: function (e) { return __awaiter(_this, void 0, void 0, function () {
                            var file, folderKey, topFolderId, currentVessel, subFolderPath, mainFolderPathMap, spoMainFolder, matchingRow, resolvedFolderId, _a, fileId, statusPending, folderId, msg, newUpload, existingUploads, updatedByFolder, liveFolderId, refreshId_1, err_1;
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
                                            'Knowledge Bank': 'Kaizen - Knowledge Bank',
                                        };
                                        spoMainFolder = (docMainFolder && mainFolderPathMap[docMainFolder]) || 'Technical & Crewing';
                                        matchingRow = host.state.rows.find(function (r) {
                                            return r.vesselName === currentVessel &&
                                                r.uploadFolderId &&
                                                !r.uploadFolderId.includes('/') &&
                                                (docMainFolder ? r.group.toLowerCase().includes(docMainFolder.toLowerCase().split(' ')[0]) : true);
                                        });
                                        resolvedFolderId = topFolderId || (matchingRow === null || matchingRow === void 0 ? void 0 : matchingRow.uploadFolderId) || "".concat(host.VESSEL_ROOT, "/").concat(spoMainFolder, "/").concat(currentVessel);
                                        host.setState({ docUploadMsg: null });
                                        _d.label = 1;
                                    case 1:
                                        _d.trys.push([1, 3, , 4]);
                                        return [4 /*yield*/, host._uploadFileToFolder(resolvedFolderId, subFolderPath, currentVessel, file)];
                                    case 2:
                                        _a = _d.sent(), fileId = _a.fileId, statusPending = _a.statusPending, folderId = _a.folderId;
                                        msg = statusPending
                                            ? "\"".concat(file.name, "\" submitted for approval.")
                                            : "\"".concat(file.name, "\" uploaded successfully to ").concat(folderKey, "!");
                                        newUpload = { name: file.name, size: "".concat((file.size / 1024).toFixed(1), " KB"), date: 'Just now', pending: statusPending, id: fileId };
                                        existingUploads = uploadedFilesByFolder[folderKey] || [];
                                        updatedByFolder = __assign(__assign({}, uploadedFilesByFolder), (_b = {}, _b[folderKey] = __spreadArray(__spreadArray([], existingUploads.filter(function (f) { return f.name !== file.name; }), true), [newUpload], false), _b));
                                        liveFolderId = folderId || topFolderId;
                                        if (liveFolderId) {
                                            updatedByFolder[liveFolderId] = __spreadArray(__spreadArray([], (uploadedFilesByFolder[liveFolderId] || []).filter(function (f) { return f.name !== file.name; }), true), [newUpload], false);
                                        }
                                        if (resolvedFolderId && resolvedFolderId !== topFolderId) {
                                            updatedByFolder[resolvedFolderId] = __spreadArray(__spreadArray([], (uploadedFilesByFolder[resolvedFolderId] || []).filter(function (f) { return f.name !== file.name; }), true), [newUpload], false);
                                        }
                                        host.setState({ uploadedFilesByFolder: updatedByFolder, docUploadMsg: msg });
                                        refreshId_1 = folderId || topFolderId || resolvedFolderId;
                                        if (refreshId_1 && !/^(sf_|common)/.test(refreshId_1)) {
                                            void host._refreshFolderFiles(refreshId_1, refreshId_1, true).catch(function () { return undefined; });
                                            // Graph can be briefly eventually consistent after an upload.
                                            setTimeout(function () { return host._refreshFolderFiles(refreshId_1, refreshId_1, true).catch(function () { return undefined; }); }, 1500);
                                        }
                                        return [3 /*break*/, 4];
                                    case 3:
                                        err_1 = _d.sent();
                                        host.setState({ docUploadMsg: "Upload failed: ".concat((err_1 === null || err_1 === void 0 ? void 0 : err_1.message) || 'Error') });
                                        return [3 /*break*/, 4];
                                    case 4: return [2 /*return*/];
                                }
                            });
                        }); } }),
                    React.createElement("span", null, "\u2B06"),
                    " Upload"))),
        React.createElement("div", { style: { background: '#fff', borderRadius: 10, padding: '10px 12px', border: '1px solid #e2e8f0', display: 'flex', flexWrap: 'wrap', gap: 8, alignItems: 'center' } },
            React.createElement("div", { style: { position: 'relative', flex: '1 1 180px', minWidth: 160 } },
                React.createElement("span", { style: { position: 'absolute', left: 9, top: '50%', transform: 'translateY(-50%)', color: '#94a3b8', fontSize: 12 } }, "\uD83D\uDD0D"),
                React.createElement("input", { type: "text", placeholder: "Filter by vessel, group, category, path...", value: textFilter, onChange: function (e) { return host.setState({ textFilter: e.target.value, docListPage: 0 }); }, style: { width: '100%', padding: '6px 10px 6px 28px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 12, outline: 'none', boxSizing: 'border-box' } })),
            React.createElement("select", { value: vesselFilter, onChange: function (e) { return host.setState({ vesselFilter: e.target.value, docListPage: 0 }); }, style: { padding: '6px 10px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 12, background: '#fff', outline: 'none', maxWidth: 140 } },
                React.createElement("option", { value: "all" }, "All vessels"),
                vessels.map(function (v) { return (React.createElement("option", { key: v.id || v.name, value: v.name }, v.name)); })),
            React.createElement("select", { value: docGroupFilter, onChange: function (e) { return host.setState({ docGroupFilter: e.target.value, catFilter: 'all', docListPage: 0 }); }, style: { padding: '6px 10px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 12, background: '#fff', outline: 'none', maxWidth: 140 } },
                React.createElement("option", { value: "all" }, "All groups"),
                allGroups.map(function (g) { return (React.createElement("option", { key: g, value: g }, g)); })),
            React.createElement("select", { value: catFilter, onChange: function (e) { return host.setState({ catFilter: e.target.value, docListPage: 0 }); }, style: { padding: '6px 10px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 12, background: '#fff', outline: 'none', maxWidth: 140 } },
                React.createElement("option", { value: "all" }, "All categories"),
                allCategories.map(function (c) { return (React.createElement("option", { key: c, value: c }, c)); })),
            React.createElement("select", { value: docListSort, onChange: function (e) { return host.setState({ docListSort: e.target.value, docListPage: 0 }); }, style: { padding: '6px 10px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 12, background: '#fff', outline: 'none' } },
                React.createElement("option", { value: "default" }, "Default order"),
                React.createElement("option", { value: "name_az" }, "Name A\u2013Z"),
                React.createElement("option", { value: "newest" }, "Newest")),
            false && vessels.length > 4 && (React.createElement("button", { onClick: function () { return host.setState({ showAllVesselsInFolderView: !showAllVesselsInFolderView }); }, style: {
                    padding: '6px 12px', borderRadius: 8, border: '1px solid #cbd5e1', fontSize: 12,
                    background: showAllVesselsInFolderView ? '#0f172a' : '#f1f5f9',
                    color: showAllVesselsInFolderView ? '#fff' : '#334155',
                    fontWeight: 600, cursor: 'pointer', outline: 'none', display: 'inline-flex', alignItems: 'center', gap: 4
                } }, showAllVesselsInFolderView ? 'Showing All Vessels ⌃' : "Show All Vessels (".concat(vessels.length, ") \u2304"))),
            documentVesselCount < vessels.length && (React.createElement("button", { onClick: function () { return void host._loadMoreDocumentVessels(); }, disabled: documentVesselsLoadingMore, style: {
                    padding: '6px 12px', borderRadius: 8, border: '1px solid #0284c7', fontSize: 12,
                    background: documentVesselsLoadingMore ? '#e2e8f0' : '#0284c7', color: documentVesselsLoadingMore ? '#64748b' : '#fff',
                    fontWeight: 600, cursor: documentVesselsLoadingMore ? 'wait' : 'pointer', outline: 'none',
                } }, documentVesselsLoadingMore ? 'Loading vessels...' : "More vessels (+".concat(Math.min(8, vessels.length - documentVesselCount), ")"))),
            React.createElement("div", { style: { display: 'flex', border: '1px solid #cbd5e1', borderRadius: 8, overflow: 'hidden', marginLeft: 'auto' } },
                React.createElement("button", { onClick: function () { return host.setState({ docViewMode: 'folder' }); }, style: { padding: '5px 10px', background: docViewMode === 'folder' ? '#e2e8f0' : '#fff', border: 'none', cursor: 'pointer', fontSize: 13 }, title: "Folder view" }, "::"),
                React.createElement("button", { onClick: function () { return host.setState({ docViewMode: 'list', docMainFolder: null, folderPathStack: [] }); }, style: { padding: '5px 10px', background: docViewMode === 'list' ? '#e2e8f0' : '#fff', border: 'none', borderLeft: '1px solid #cbd5e1', cursor: 'pointer', fontSize: 13 }, title: "List view" }, "\u2630"))),
        docUploadMsg && (React.createElement("div", { style: { background: '#f0fdf4', border: '1px solid #86efac', borderRadius: 8, padding: '8px 14px', fontSize: 12, color: '#16a34a', display: 'flex', justifyContent: 'space-between' } },
            React.createElement("span", null,
                "\u2713 ",
                docUploadMsg),
            React.createElement("button", { onClick: function () { return host.setState({ docUploadMsg: null }); }, style: { border: 'none', background: 'none', cursor: 'pointer', color: '#16a34a', fontWeight: 700 } }, "\u2715"))),
        docViewMode === 'folder' ? (React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 24, marginTop: 4 } },
            stackLevel === 0 ? (
            /* Level 0: Root — Vessels + Kaizen */
            React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 24 } },
                React.createElement("div", { style: { display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: 16 } },
                    [
                        { id: 'vessels_root', name: 'Vessels', emoji: '\uD83D\uDEA2', bg: '#e0f2fe', sub: vessels.length + ' vessels' },
                        { id: 'kaizen_root', name: 'Kaizen - Knowledge Bank', emoji: '\uD83D\uDCDA', bg: '#ede9fe', sub: 'Knowledge base' },
                    ].map(function (item) { return (React.createElement("div", { key: item.id, onClick: function () {
                            if (item.id === 'kaizen_root') {
                                host._pushFolderNav([{ id: 'kaizen_root', name: 'Kaizen - Knowledge Bank' }], 'Kaizen - Knowledge Bank');
                                host.setState({ vesselFilter: 'all' });
                            } else {
                                host._pushFolderNav([{ id: 'vessels_root', name: 'Vessels' }], null);
                            }
                        }, style: { background: '#fff', borderRadius: 14, border: '1px solid #e2e8f0', padding: 18, display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' } },
                        React.createElement("div", { style: { width: 44, height: 44, borderRadius: 10, background: item.bg, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 22 } }, item.emoji),
                        React.createElement("div", { style: { flex: 1, minWidth: 0 } },
                            React.createElement("div", { style: { fontWeight: 700, fontSize: 14, color: '#0f172a', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' } }, item.name),
                            React.createElement("div", { style: { fontSize: 12, color: '#64748b', marginTop: 2 } }, item.sub)),
                        React.createElement("span", { style: { color: '#94a3b8', fontSize: 16 } }, '\u203A'))); })))) : atVesselsRoot ? (
            /* Level 1: Specific Vessels + Common */
            React.createElement("div", { style: { display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: 16 } },
                [
                    { id: 'specific_vessels', name: 'Specific Vessels', emoji: '\uD83D\uDEA2', bg: '#e0f2fe', sub: vessels.length + ' vessels' },
                    { id: 'common', name: 'Common for all ships', emoji: '\uD83D\uDCC1', bg: '#fef3c7', sub: 'Shared documents' },
                ].map(function (item) { return (React.createElement("div", { key: item.id, onClick: function () { return host._pushFolderNav(__spreadArray(__spreadArray([], folderPathStack, true), [{ id: item.id, name: item.name }], false), null); }, style: { background: '#fff', borderRadius: 14, border: '1px solid #e2e8f0', padding: 18, display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer' } },
                    React.createElement("div", { style: { width: 44, height: 44, borderRadius: 10, background: item.bg, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20 } }, item.emoji),
                    React.createElement("div", { style: { flex: 1, minWidth: 0 } },
                        React.createElement("div", { style: { fontWeight: 700, fontSize: 14, color: '#0f172a' } }, item.name),
                        React.createElement("div", { style: { fontSize: 12, color: '#64748b', marginTop: 2 } }, item.sub)),
                    React.createElement("span", { style: { color: '#94a3b8', fontSize: 16 } }, '\u203A'))); }))) : atSpecificVessels ? (
            /* Level 2: Vessel list */
            React.createElement("div", { style: { display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: 16 } },
                displayVessels.map(function (v) { return (React.createElement("div", { key: v.id, onClick: function () {
                        host._pushFolderNav(__spreadArray(__spreadArray([], folderPathStack, true), [{ id: v.id, name: v.name }], false), null);
                        host.setState({ vesselFilter: v.name });
                        host._loadFilesForVessel(v.name).catch(function () { return undefined; });
                    }, style: { background: '#fff', borderRadius: 14, border: '1px solid #e2e8f0', padding: 18, display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer' } },
                    React.createElement("div", { style: { width: 44, height: 44, borderRadius: 10, background: '#e0f2fe', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20, color: '#0284c7' } }, '\uD83D\uDEA2'),
                    React.createElement("div", { style: { flex: 1, minWidth: 0 } },
                        React.createElement("div", { style: { fontWeight: 700, fontSize: 14, color: '#0f172a', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' } }, v.name),
                        React.createElement("div", { style: { fontSize: 12, color: '#64748b', marginTop: 2 } }, 'Vessel')),
                    React.createElement("span", { style: { color: '#94a3b8', fontSize: 16 } }, '\u203A'))); }))) : atVesselMainFolderSelect ? (
            /* Level 3: Main folder selection inside a vessel */
            React.createElement("div", { style: { display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: 16 } },
                [
                    { key: 'Technical & Crewing', emoji: '\u2699\uFE0F', bg: '#fee2e2' },
                    { key: 'Commercial & Chartering', emoji: '\uD83D\uDCBC', bg: '#dcfce7' },
                    { key: 'Insurance', emoji: '\uD83D\uDEE1\uFE0F', bg: '#fef3c7' },
                ].map(function (mf) { return (React.createElement("div", { key: mf.key, onClick: function () {
                        host._pushFolderNav(folderPathStack, mf.key);
                        host.setState({ vesselFilter: currentVesselNameFromStack || 'all' });
                    }, style: { background: '#fff', borderRadius: 14, border: '1px solid #e2e8f0', padding: 18, display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' } },
                    React.createElement("div", { style: { width: 44, height: 44, borderRadius: 10, background: mf.bg, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 22 } }, mf.emoji),
                    React.createElement("div", { style: { flex: 1, minWidth: 0 } },
                        React.createElement("div", { style: { fontWeight: 700, fontSize: 14, color: '#0f172a', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' } }, mf.key),
                        React.createElement("div", { style: { fontSize: 12, color: '#64748b', marginTop: 2 } }, 'Main folder')),
                    React.createElement("span", { style: { color: '#94a3b8', fontSize: 16 } }, '\u203A'))); }))) : subfolderNames.length > 0 ? (
            /* Subfolders Grid */
            React.createElement("div", null,
                React.createElement("div", { style: { display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(260px, 1fr))', gap: 16 } }, subfolderNames.map(function (sfName, idx) { return (React.createElement("div", { key: sfName + idx, onClick: function () {
                        var vesselName = currentVesselNameFromStack || (folderPathStack.length > 0 ? folderPathStack[0].name : null);
                        var candidateSubPath = vesselName ? vesselName + ' > ' + (docMainFolder || '') + ' > ' + sfName : '';
                        var liveFolderId = candidateSubPath ? host._getLiveSharePointFolderId(candidateSubPath) : null;
                        var realFolderId = liveFolderId || (vesselName
                            ? ((host.state.rows.find(function (r) {
                                return r.vesselName === vesselName &&
                                    r.uploadFolderId &&
                                    !r.uploadFolderId.includes('/') &&
                                    (r.subCategory === sfName || r.category === sfName);
                            }) || {}).uploadFolderId || 'sf_' + idx)
                            : 'sf_' + idx);
                        var newStack = __spreadArray(__spreadArray([], folderPathStack, true), [{ id: realFolderId, name: sfName }], false);
                        host._pushFolderNav(newStack, docMainFolder);
                        if (!/^sf_/.test(realFolderId)) {
                            void host._refreshFolderFiles(realFolderId, realFolderId, true).catch(function () { return undefined; });
                        } else if (currentFolderNode && !/^(sf_|common)/.test(currentFolderNode.id)) {
                            host._refreshFolderFiles(currentFolderNode.id, currentFolderNode.id).catch(function () { return undefined; });
                        }
                    }, style: { background: '#fff', borderRadius: 14, border: '1px solid #e2e8f0', padding: 18, display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' } },
                    React.createElement("div", { style: { width: 44, height: 44, borderRadius: 10, background: '#e0f2fe', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20, color: '#0284c7' } }, '\uD83D\uDCC1'),
                    React.createElement("div", { style: { flex: 1, minWidth: 0 } },
                        React.createElement("div", { style: { fontWeight: 700, fontSize: 14, color: '#0f172a', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' } }, sfName),
                        React.createElement("div", { style: { fontSize: 12, color: '#64748b', marginTop: 2 } }, 'Folder')),
                    React.createElement("span", { style: { color: '#94a3b8', fontSize: 16 } }, '\u203A'))); })),
                allCurrentFolderFiles.length > 0 && (React.createElement("div", { style: { marginTop: 24, background: '#fff', borderRadius: 12, border: '1px solid #e2e8f0', overflow: 'hidden' } },
                    React.createElement("table", { style: { width: '100%', borderCollapse: 'collapse', fontSize: 13 } },
                        React.createElement("thead", null, React.createElement("tr", { style: { background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 11, fontWeight: 700, textTransform: 'uppercase', textAlign: 'left' } },
                            React.createElement("th", { style: { padding: '10px 16px' } }, 'FILE NAME'),
                            React.createElement("th", { style: { padding: '10px 16px' } }, 'SIZE'),
                            React.createElement("th", { style: { padding: '10px 16px' } }, 'DATE UPLOADED'),
                            React.createElement("th", { style: { padding: '10px 16px', textAlign: 'right' } }, 'ACTION'))),
                        React.createElement("tbody", null, allCurrentFolderFiles.map(function (file, idx) { return (React.createElement("tr", { key: file.name + idx, style: { borderBottom: '1px solid #f1f5f9' } },
                            React.createElement("td", { style: { padding: '12px 16px', fontWeight: 600, color: '#0f172a', display: 'flex', alignItems: 'center', gap: 10 } },
                                React.createElement("span", { style: { fontSize: 18 } }, file.pending ? '\u23F3' : '\uD83D\uDCC4'), file.name,
                                file.pending && React.createElement("span", { style: { fontSize: 11, color: '#d97706', fontWeight: 600, background: '#fef3c7', borderRadius: 4, padding: '1px 6px' } }, 'Pending Approval')),
                            React.createElement("td", { style: { padding: '12px 16px', color: '#64748b' } }, file.size),
                            React.createElement("td", { style: { padding: '12px 16px', color: '#64748b' } }, file.date),
                            React.createElement("td", { style: { padding: '12px 16px', textAlign: 'right' } },
                                React.createElement("div", { style: { display: 'inline-flex', gap: 6, alignItems: 'center' } },
                                    React.createElement("button", { style: { border: '1px solid #cbd5e1', background: '#fff', borderRadius: 6, padding: '4px 10px', fontSize: 11, fontWeight: 600, cursor: 'pointer', color: '#0078d4' } }, 'View / Download'),
                                    React.createElement("button", { onClick: function () { return host._openFileDeleteDialog([{ id: file.id || file.name, name: file.name, folderId: currentFolderNode ? currentFolderNode.id : '', folderPath: currentFolderNode ? currentFolderNode.name : '' }]); }, style: { border: '1px solid #fca5a5', background: '#fff5f5', borderRadius: 6, padding: '4px 10px', fontSize: 11, fontWeight: 600, cursor: 'pointer', color: '#ef4444' }, title: 'Delete file' }, '\uD83D\uDDD1 Delete'))))); }))))))) :
                (function () {
                    var mainAnomalies = (host.state.folderAnomalies || []).filter(function (a) { return a.anomaly_type === 'main_folder_unmatched'; });
                    if (mainAnomalies.length === 0)
                        return null;
                    var mainFolders = mainAnomalies.filter(function (a) { return a.item_type === 'folder'; });
                    var mainFiles = mainAnomalies.filter(function (a) { return a.item_type === 'file'; });
                    return (React.createElement("div", { style: { marginTop: 24, display: 'flex', flexDirection: 'column', gap: 16 } },
                        mainFolders.length > 0 && (React.createElement("div", { style: { background: 'linear-gradient(135deg, #fffbeb 0%, #fff9e6 100%)', border: '2px solid #f59e0b', borderRadius: 14, padding: 20 } },
                            React.createElement("h4", { style: { margin: '0 0 6px', fontSize: 15, fontWeight: 700, color: '#92400e', display: 'flex', alignItems: 'center', gap: 8 } },
                                "\u26A0\uFE0F Folders Created Outside Standard Main Folders (",
                                mainFolders.length,
                                ")"),
                            React.createElement("p", { style: { margin: '0 0 12px', fontSize: 12, color: '#a16207' } }, "These folders were created directly in SharePoint Online at the main folder root level (`Vessel Management`) outside standard category structures."),
                            React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 10 } }, mainFolders.map(function (item) { return (React.createElement("div", { key: item.id, style: {
                                    background: '#fff', borderRadius: 10, border: '1px solid #fde68a', padding: '12px 16px',
                                    display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap',
                                } },
                                React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 12, flex: 1, minWidth: 200 } },
                                    React.createElement("span", { style: { fontSize: 24 } }, "\uD83D\uDCC1"),
                                    React.createElement("div", null,
                                        React.createElement("div", { style: { fontWeight: 700, fontSize: 13, color: '#1f1f1f' } }, item.name),
                                        React.createElement("div", { style: { fontSize: 11, color: '#78716c', fontFamily: 'monospace' } }, item.spo_path))),
                                React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 8 } },
                                    React.createElement("button", { onClick: function () { return host.setState({ spoClassifyDialog: { anomaly: item, provisioning: false, done: false, error: null } }); }, style: {
                                            background: 'linear-gradient(135deg, #f59e0b, #d97706)',
                                            color: '#fff', border: 'none', borderRadius: 6, padding: '6px 12px',
                                            fontSize: 12, fontWeight: 700, cursor: 'pointer',
                                            display: 'inline-flex', alignItems: 'center', gap: 4,
                                        } }, "\uD83D\uDD0D Classify"),
                                    React.createElement("button", { onClick: function () { return host._dismissAnomaly(item.id); }, style: { background: '#fff', color: '#78716c', border: '1px solid #d6d3d1', borderRadius: 6, padding: '6px 10px', fontSize: 11, cursor: 'pointer' } }, "\u2715 Dismiss")))); })))),
                        mainFiles.length > 0 && (React.createElement("div", { style: { background: 'linear-gradient(135deg, #faf5ff 0%, #f3e8ff 100%)', border: '2px solid #a855f7', borderRadius: 14, padding: 20 } },
                            React.createElement("h4", { style: { margin: '0 0 6px', fontSize: 15, fontWeight: 700, color: '#6b21a8', display: 'flex', alignItems: 'center', gap: 8 } },
                                "\uD83D\uDCC4 Files Uploaded Outside Main Category Structure (",
                                mainFiles.length,
                                ")"),
                            React.createElement("p", { style: { margin: '0 0 12px', fontSize: 12, color: '#7c3aed' } }, "These files were uploaded directly to the main folder root in SharePoint Online."),
                            React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 8 } }, mainFiles.map(function (item) {
                                var _a;
                                var ext = ((_a = item.name.split('.').pop()) === null || _a === void 0 ? void 0 : _a.toUpperCase()) || 'FILE';
                                return (React.createElement("div", { key: item.id, style: {
                                        background: '#fff', borderRadius: 10, border: '1px solid #e9d5ff', padding: '10px 14px',
                                        display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap',
                                    } },
                                    React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 10 } },
                                        React.createElement("span", { style: { background: '#f3e8ff', color: '#6b21a8', borderRadius: 4, padding: '2px 6px', fontSize: 10, fontWeight: 700 } }, ext),
                                        React.createElement("div", null,
                                            React.createElement("span", { style: { fontWeight: 600, fontSize: 13, color: '#1e293b' } }, item.name),
                                            React.createElement("span", { style: { fontSize: 11, color: '#64748b', marginLeft: 8, fontFamily: 'monospace' } }, item.spo_path))),
                                    React.createElement("button", { onClick: function () { return host._dismissAnomaly(item.id); }, style: { background: '#fff', color: '#6b21a8', border: '1px solid #e9d5ff', borderRadius: 6, padding: '4px 10px', fontSize: 11, cursor: 'pointer' } }, "\u2715 Dismiss")));
                            }))))));
                })())) : folderPathStack.length === 0 ? (
            /* ── Level 1: Vessel list inside the selected main folder ── */
            React.createElement(React.Fragment, null,
                React.createElement("div", null,
                    React.createElement("div", { style: { fontSize: 11, fontWeight: 700, letterSpacing: '0.05em', color: '#94a3b8', textTransform: 'uppercase', marginBottom: 12 } }, "COMMON AGREEMENTS / DOCUMENTS"),
                    React.createElement("div", { style: { display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 340px))', gap: 16 } },
                        React.createElement("div", { onClick: function () { return host.setState({ folderPathStack: [{ id: 'common', name: 'Common for all ships' }] }); }, style: { background: '#fff', borderRadius: 14, border: '1px solid #e2e8f0', padding: 18, display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer' } },
                            React.createElement("div", { style: { width: 44, height: 44, borderRadius: 10, background: '#fef3c7', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20, color: '#d97706' } }, "\uD83D\uDCC1"),
                            React.createElement("div", { style: { flex: 1 } },
                                React.createElement("div", { style: { fontWeight: 700, fontSize: 14, color: '#0f172a' } }, "Common for all ships"),
                                React.createElement("div", { style: { fontSize: 12, color: '#64748b', marginTop: 2 } }, "Common")),
                            React.createElement("span", { style: { color: '#94a3b8', fontSize: 16 } }, "\u203A")))),
                React.createElement("div", null,
                    React.createElement("div", { style: { fontSize: 11, fontWeight: 700, letterSpacing: '0.05em', color: '#94a3b8', textTransform: 'uppercase', marginBottom: 12 } }, "VESSELS"),
                    React.createElement("div", { style: { display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: 16 } }, displayVessels.map(function (v) { return (React.createElement("div", { key: v.id, onClick: function () {
                            host.setState({ vesselFilter: v.name, folderPathStack: [{ id: v.id, name: v.name }] });
                            // Pre-load files & scan live SPO structure for host vessel so new SPO folders show immediately
                            host._loadFilesForVessel(v.name).catch(function () { return undefined; });
                        }, style: { background: '#fff', borderRadius: 14, border: '1px solid #e2e8f0', padding: 18, display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer' } },
                        React.createElement("div", { style: { width: 44, height: 44, borderRadius: 10, background: '#e0f2fe', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20, color: '#0284c7' } }, "\uD83D\uDEA2"),
                        React.createElement("div", { style: { flex: 1, minWidth: 0 } },
                            React.createElement("div", { style: { fontWeight: 700, fontSize: 14, color: '#0f172a', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' } }, v.name),
                            React.createElement("div", { style: { fontSize: 12, color: '#64748b', marginTop: 2 } }, "Vessel")),
                        React.createElement("span", { style: { color: '#94a3b8', fontSize: 16 } }, "\u203A"))); })),
                    false && vessels.length > 4 && (React.createElement("div", { style: { display: 'flex', justifyContent: 'center', marginTop: 20 } },
                        React.createElement("button", { onClick: function () { return host.setState({ showAllVesselsInFolderView: !showAllVesselsInFolderView }); }, style: { background: '#fff', border: '1px solid #cbd5e1', borderRadius: 20, padding: '8px 20px', fontSize: 13, fontWeight: 600, color: '#334155', cursor: 'pointer' } }, showAllVesselsInFolderView ? 'Show Less ⌃' : "More Vessels (".concat(vessels.length - 4, ") \u2304"))))))) : subfolderNames.length > 0 && allCurrentFolderFiles.length === 0 ? (
            /* Subfolders Grid (Screenshot 1) */
            React.createElement("div", { style: { display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(260px, 1fr))', gap: 16 } }, subfolderNames.map(function (sfName, idx) { return (React.createElement("div", { key: sfName + idx, onClick: function () {
                    var _a;
                    // Try to resolve a real SPO folder ID for this subfolder from rows
                    var vesselName = folderPathStack.length > 0 ? folderPathStack[0].name : null;
                    var realFolderId = vesselName
                        ? (((_a = host.state.rows.find(function (r) {
                            return r.vesselName === vesselName &&
                                r.uploadFolderId &&
                                !r.uploadFolderId.includes('/') &&
                                (r.subCategory === sfName || r.category === sfName);
                        })) === null || _a === void 0 ? void 0 : _a.uploadFolderId) || "sf_".concat(idx))
                        : "sf_".concat(idx);
                    var newStack = __spreadArray(__spreadArray([], folderPathStack, true), [{ id: realFolderId, name: sfName }], false);
                    host.setState({ folderPathStack: newStack });
                    // Immediately refresh files for this folder from live SPO
                    if (!/^sf_/.test(realFolderId)) {
                        void host._refreshFolderFiles(realFolderId, realFolderId, true).catch(function () { return undefined; });
                    }
                    else if (currentFolderNode && !/^(sf_|common)/.test(currentFolderNode.id)) {
                        host._refreshFolderFiles(currentFolderNode.id, currentFolderNode.id).catch(function () { return undefined; });
                    }
                }, style: {
                    background: '#fff', borderRadius: 14, border: '1px solid #e2e8f0', padding: 18,
                    display: 'flex', alignItems: 'center', gap: 14, cursor: 'pointer',
                    boxShadow: '0 1px 3px rgba(0,0,0,0.05)', transition: 'transform 0.15s, box-shadow 0.15s',
                } },
                React.createElement("div", { style: {
                        width: 44, height: 44, borderRadius: 10, background: '#e0f2fe',
                        display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 20, color: '#0284c7',
                    } }, "\uD83D\uDCC1"),
                React.createElement("div", { style: { flex: 1, minWidth: 0 } },
                    React.createElement("div", { style: { fontWeight: 700, fontSize: 14, color: '#0f172a', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' } }, sfName),
                    React.createElement("div", { style: { fontSize: 12, color: '#64748b', marginTop: 2 } }, "Folder")),
                React.createElement("span", { style: { color: '#94a3b8', fontSize: 16 } }, "\u203A"))); }))) : allCurrentFolderFiles.length > 0 ? (
            /* Folder File Items List */
            React.createElement("div", { style: { background: '#fff', borderRadius: 12, border: '1px solid #e2e8f0', overflow: 'hidden' } },
                React.createElement("table", { style: { width: '100%', borderCollapse: 'collapse', fontSize: 13 } },
                    React.createElement("thead", null,
                        React.createElement("tr", { style: { background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 11, fontWeight: 700, textTransform: 'uppercase', textAlign: 'left' } },
                            React.createElement("th", { style: { padding: '10px 16px' } }, "FILE NAME"),
                            React.createElement("th", { style: { padding: '10px 16px' } }, "SIZE"),
                            React.createElement("th", { style: { padding: '10px 16px' } }, "DATE UPLOADED"),
                            React.createElement("th", { style: { padding: '10px 16px', textAlign: 'right' } }, "ACTION"))),
                    React.createElement("tbody", null, allCurrentFolderFiles.map(function (file, idx) { return (React.createElement("tr", { key: file.name + idx, style: { borderBottom: '1px solid #f1f5f9' } },
                        React.createElement("td", { style: { padding: '12px 16px', fontWeight: 600, color: '#0f172a', display: 'flex', alignItems: 'center', gap: 10 } },
                            React.createElement("span", { style: { fontSize: 18 } }, file.pending ? '⏳' : '📄'),
                            file.name,
                            file.pending && React.createElement("span", { style: { fontSize: 11, color: '#d97706', fontWeight: 600, background: '#fef3c7', borderRadius: 4, padding: '1px 6px' } }, "Pending Approval")),
                        React.createElement("td", { style: { padding: '12px 16px', color: '#64748b' } }, file.size),
                        React.createElement("td", { style: { padding: '12px 16px', color: '#64748b' } }, file.date),
                        React.createElement("td", { style: { padding: '12px 16px', textAlign: 'right' } },
                            React.createElement("div", { style: { display: 'inline-flex', gap: 6, alignItems: 'center' } },
                                React.createElement("button", { style: { border: '1px solid #cbd5e1', background: '#fff', borderRadius: 6, padding: '4px 10px', fontSize: 11, fontWeight: 600, cursor: 'pointer', color: '#0078d4' } }, "View / Download"),
                                React.createElement("button", { onClick: function () { return host._openFileDeleteDialog([{ id: file.id || file.name, name: file.name, folderId: currentFolderNode ? currentFolderNode.id : '', folderPath: currentFolderNode ? currentFolderNode.name : '' }]); }, style: { border: '1px solid #fca5a5', background: '#fff5f5', borderRadius: 6, padding: '4px 10px', fontSize: 11, fontWeight: 600, cursor: 'pointer', color: '#ef4444' }, title: 'Delete file' }, '\uD83D\uDDD1 Delete'))))); })))))) : (
            /* Empty Folder View (Screenshot 2) */
            React.createElement("div", { style: { display: 'flex', justifyContent: 'center', alignItems: 'center', padding: '60px 20px', width: '100%' } },
                React.createElement("div", { style: {
                        background: 'rgba(240, 249, 255, 0.6)', border: '1px solid #e0f2fe',
                        borderRadius: 24, padding: '48px 40px', maxWidth: 500, width: '100%',
                        textAlign: 'center', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 12,
                    } },
                    React.createElement("div", { style: {
                            width: 56, height: 56, borderRadius: 14, background: '#e0f2fe',
                            display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 24, color: '#0284c7',
                            marginBottom: 4,
                        } }, "\uD83D\uDCC1"),
                    React.createElement("h3", { style: { margin: 0, fontSize: 18, fontWeight: 700, color: '#0f172a' } }, "This folder is empty"),
                    React.createElement("p", { style: { margin: 0, fontSize: 13, color: '#64748b', maxWidth: 320, lineHeight: 1.5 } }, "Use the Upload button in the top-right to add a document.")))),
            (function () {
                var currentVessel = folderPathStack.length > 0 ? folderPathStack[0].name : null;
                var subAnomalies = (host.state.folderAnomalies || []).filter(function (a) {
                    return a.anomaly_type === 'subfolder_unmatched' && (!currentVessel || a.vessel_name === currentVessel);
                });
                if (subAnomalies.length === 0)
                    return null;
                return (React.createElement("div", { style: { background: '#fff7ed', border: '1px solid #ffedd5', borderRadius: 14, padding: 18, marginTop: 16 } },
                    React.createElement("h4", { style: { margin: '0 0 8px', fontSize: 14, fontWeight: 700, color: '#c2410c', display: 'flex', alignItems: 'center', gap: 8 } },
                        "\u26A0\uFE0F Other / Unclassified Items Inside Vessel Tree (",
                        subAnomalies.length,
                        ")"),
                    React.createElement("p", { style: { margin: '0 0 12px', fontSize: 12, color: '#9a3412' } }, "These items were added inside the vessel folder in SharePoint but are not part of the standard template structure."),
                    React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 8 } }, subAnomalies.map(function (item) { return (React.createElement("div", { key: item.id, style: {
                            background: '#fff', borderRadius: 8, border: '1px solid #fed7aa', padding: '10px 14px',
                            display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap',
                        } },
                        React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 10 } },
                            React.createElement("span", { style: { fontSize: 18 } }, item.item_type === 'folder' ? '📁' : '📄'),
                            React.createElement("div", null,
                                React.createElement("span", { style: { fontWeight: 600, fontSize: 13, color: '#1e293b' } }, item.name),
                                React.createElement("span", { style: { fontSize: 11, color: '#64748b', marginLeft: 8 } },
                                    "Path: ",
                                    item.spo_path))),
                        React.createElement("button", { onClick: function () { return host._dismissAnomaly(item.id); }, style: { background: '#fff', color: '#c2410c', border: '1px solid #fed7aa', borderRadius: 6, padding: '4px 10px', fontSize: 11, cursor: 'pointer' } }, "\u2715 Dismiss"))); }))));
            })())) : (
        /* ── LIST VIEW ── */
        React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 12 } },
            React.createElement("div", { style: { background: '#fff', borderRadius: 10, border: '1px solid #e2e8f0', overflowX: 'auto', width: '100%' } },
                React.createElement("table", { style: { width: '100%', minWidth: 850, borderCollapse: 'collapse', fontSize: 12 } },
                    React.createElement("thead", null,
                        React.createElement("tr", { style: { background: '#f8fafc', borderBottom: '2px solid #e2e8f0', color: '#64748b', fontSize: 11, fontWeight: 700, textTransform: 'uppercase', textAlign: 'left' } },
                            React.createElement("th", { style: { padding: '10px 10px', width: 44, textAlign: 'center' } }, "SR."),
                            React.createElement("th", { style: { padding: '10px 12px' } }, "VESSEL NAME"),
                            React.createElement("th", { style: { padding: '10px 12px' } }, "GROUP"),
                            React.createElement("th", { style: { padding: '10px 12px' } }, "CATEGORY"),
                            React.createElement("th", { style: { padding: '10px 12px' } }, "SUB-CATEGORY"),
                            React.createElement("th", { style: { padding: '10px 12px' } }, "FOLDER PATH"),
                            React.createElement("th", { style: { padding: '10px 12px' } }, "FILE NAME"),
                            React.createElement("th", { style: { padding: '10px 12px', textAlign: 'right' } }, "ATTACHMENT"))),
                    React.createElement("tbody", null, pageGroupedRows.length === 0 ? (React.createElement("tr", null,
                        React.createElement("td", { colSpan: 8, style: { padding: '36px 16px', textAlign: 'center', color: '#94a3b8', fontSize: 13 } },
                            "No documents found. ",
                            textFilter || vesselFilter !== 'all' || docGroupFilter !== 'all' || catFilter !== 'all' ? 'Try clearing the filters.' : ''))) : pageGroupedRows.map(function (r, idx) {
                        var globalIdx = safePage * PAGE_ROWS + idx + 1;
                        var isUploading = docUploadRowKey === r.groupKey && docUploadBusy;
                        var hasFiles = r.files.length > 0;
                        return (React.createElement("tr", { key: "".concat(r.groupKey, "-").concat(idx), style: { borderBottom: '1px solid #f1f5f9', transition: 'background 0.1s' }, onMouseEnter: function (e) { return (e.currentTarget.style.background = '#f8fafc'); }, onMouseLeave: function (e) { return (e.currentTarget.style.background = ''); } },
                            React.createElement("td", { style: { padding: '10px 10px', color: '#94a3b8', fontSize: 11, fontFamily: 'monospace', textAlign: 'center' } }, globalIdx),
                            React.createElement("td", { style: { padding: '10px 12px', fontWeight: 700, color: '#0f172a' } }, r.vesselName),
                            React.createElement("td", { style: { padding: '10px 12px' } },
                                React.createElement("span", { style: { display: 'inline-block', borderRadius: 8, padding: '3px 8px', fontSize: 11, fontWeight: 600, background: '#eff6ff', color: '#2563eb' } }, r.group)),
                            React.createElement("td", { style: { padding: '10px 12px', fontWeight: 600, color: '#334155' } }, r.category),
                            React.createElement("td", { style: { padding: '10px 12px', fontWeight: 600, color: '#1e293b' } }, r.subCategory || r.category),
                            React.createElement("td", { style: { padding: '10px 12px', color: '#64748b', fontSize: 11 }, title: r.subFolderPath }, r.subFolderPath),
                            React.createElement("td", { style: { padding: '10px 12px' } }, hasFiles ? (React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 4 } }, r.files.map(function (file) { return (React.createElement("div", { key: file.name, style: { display: 'flex', alignItems: 'center', gap: 6 } },
                                React.createElement("input", { type: "checkbox", checked: host.state.listViewSelectedFiles.has(file.id), onChange: function () { host.setState(function (prev) { var next = new Set(prev.listViewSelectedFiles); if (next.has(file.id)) next.delete(file.id); else next.add(file.id); return { listViewSelectedFiles: next }; }); }, style: { width: 14, height: 14, accentColor: '#ef4444', cursor: 'pointer', flexShrink: 0 } }),
                                React.createElement("span", { style: { fontSize: 14 } }, "\uD83D\uDCC4"),
                                React.createElement("span", { onClick: function () {
                                        if (file.id && !file.id.startsWith('file_')) {
                                            window.open("".concat(host._base(), "/api/files/").concat(file.id, "/content"), '_blank');
                                        }
                                        else {
                                            alert("File \"".concat(file.name, "\" is pending \u2014 it will be available after approval."));
                                        }
                                    }, style: { color: '#0284c7', textDecoration: 'underline', fontWeight: 600, cursor: 'pointer' }, title: file.id && /^\d+$/.test(file.id) ? "".concat(file.name, " (pending approval - click to preview staged copy)") : "Click to open ".concat(file.name) },
                                    file.name,
                                    file.id && /^\d+$/.test(file.id) ? ' ⏳' : ''))); }))) : (React.createElement("span", { style: { color: '#94a3b8', fontStyle: 'italic', fontSize: 11 } }, "\u2014"))),
                            React.createElement("td", { style: { padding: '10px 12px', textAlign: 'right' } },
                                React.createElement("div", { style: { display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: 4 } },
                                    React.createElement("div", { style: { display: 'inline-flex', gap: 4, alignItems: 'center' } },
                                    false && hasFiles ? (React.createElement("button", { onClick: function () {
                                            var firstFile = r.files[0];
                                            if (firstFile.id && !firstFile.id.startsWith('file_')) {
                                                window.open("".concat(host._base(), "/api/files/").concat(firstFile.id, "/content"), '_blank');
                                            }
                                            else {
                                                alert("File \"".concat(firstFile.name, "\" in folder: ").concat(r.subFolderPath));
                                            }
                                        }, style: {
                                            background: '#eff6ff', color: '#1d4ed8', border: '1px solid #bfdbfe',
                                            borderRadius: 6, padding: '3px 8px', fontSize: 11, fontWeight: 600, cursor: 'pointer',
                                            display: 'inline-flex', alignItems: 'center', gap: 4,
                                        } },
                                        "\u2197 Open (",
                                        r.files.length,
                                        ")")) : null,
                                    React.createElement("label", { style: {
                                            background: isUploading ? '#f1f5f9' : '#fff', border: '1px solid #cbd5e1', borderRadius: 6,
                                            padding: '3px 8px', fontSize: 11, fontWeight: 600,
                                            color: isUploading ? '#94a3b8' : '#334155',
                                            cursor: isUploading ? 'not-allowed' : 'pointer',
                                            display: 'inline-flex', alignItems: 'center', gap: 4, whiteSpace: 'nowrap',
                                        } },
                                        React.createElement("input", { type: "file", style: { display: 'none' }, disabled: isUploading, onChange: function (e) { return __awaiter(_this, void 0, void 0, function () {
                                                var file, _a, fileId_1, statusPending_1, folderId, msg, newApproval_1, newUpload, liveFolderId_1, updatedByFolder, baseRows, existingRowIdx_1, updatedRows, refRow, newRow, err_2;
                                                var _b;
                                                var _c;
                                                return __generator(this, function (_d) {
                                                    switch (_d.label) {
                                                        case 0:
                                                            file = (_c = e.target.files) === null || _c === void 0 ? void 0 : _c[0];
                                                            if (!file)
                                                                return [2 /*return*/];
                                                            host.setState({ docUploadRowKey: r.groupKey, docUploadBusy: true, docUploadMsg: null });
                                                            _d.label = 1;
                                                        case 1:
                                                            _d.trys.push([1, 3, , 4]);
                                                            return [4 /*yield*/, host._uploadFileToFolder(r.uploadFolderId, r.subFolderPath, r.vesselName, file, r.monthDriven)];
                                                        case 2:
                                                            _a = _d.sent(), fileId_1 = _a.fileId, statusPending_1 = _a.statusPending, folderId = _a.folderId;
                                                            msg = statusPending_1
                                                                ? "\"".concat(file.name, "\" submitted for approval.")
                                                                : "\"".concat(file.name, "\" uploaded successfully!");
                                                            if (statusPending_1) {
                                                                newApproval_1 = {
                                                                    id: fileId_1 || "a_".concat(Date.now()),
                                                                    documentName: file.name,
                                                                    vessel: r.vesselName,
                                                                    requestedBy: host.props.userDisplayName || host.props.userEmail || 'You',
                                                                    requestedOn: new Date().toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }),
                                                                    status: 'Pending',
                                                                };
                                                                host.setState(function (prev) { return ({ approvalsList: __spreadArray([newApproval_1], prev.approvalsList, true) }); });
                                                            }
                                                            newUpload = { name: file.name, size: "".concat((file.size / 1024).toFixed(1), " KB"), date: 'Just now', pending: statusPending_1, id: fileId_1 };
                                                            liveFolderId_1 = folderId || r.uploadFolderId;
                                                            updatedByFolder = __assign(__assign({}, uploadedFilesByFolder), (_b = {}, _b[r.groupKey] = __spreadArray(__spreadArray([], (uploadedFilesByFolder[r.groupKey] || []).filter(function (f) { return f.name !== file.name; }), true), [newUpload], false), _b));
                                                            if (liveFolderId_1) {
                                                                updatedByFolder[liveFolderId_1] = __spreadArray(__spreadArray([], (uploadedFilesByFolder[liveFolderId_1] || []).filter(function (f) { return f.name !== file.name; }), true), [newUpload], false);
                                                            }
                                                            baseRows = (rows && rows.length > 0) ? rows : allRows;
                                                            existingRowIdx_1 = baseRows.findIndex(function (row) { return row.groupKey === r.groupKey && !row.fileName; });
                                                            updatedRows = void 0;
                                                            if (existingRowIdx_1 !== -1) {
                                                                updatedRows = baseRows.map(function (row, i) {
                                                                    return i === existingRowIdx_1
                                                                        ? __assign(__assign({}, row), { fileName: file.name, fileId: fileId_1 || "file_".concat(Date.now()), filePending: statusPending_1 }) : row;
                                                                });
                                                            }
                                                            else {
                                                                refRow = baseRows.find(function (row) { return row.groupKey === r.groupKey; }) || r;
                                                                newRow = __assign(__assign({}, refRow), { fileName: file.name, fileId: fileId_1 || "file_".concat(Date.now()), filePending: statusPending_1 });
                                                                updatedRows = __spreadArray(__spreadArray([], baseRows, true), [newRow], false);
                                                            }
                                                            host.setState({
                                                                docUploadBusy: false,
                                                                docUploadRowKey: null,
                                                                docUploadMsg: msg,
                                                                uploadedFilesByFolder: updatedByFolder,
                                                                rows: updatedRows,
                                                            });
                                                            // Refresh the live folder, rather than replacing this
                                                            // vessel's rows with the DB-only response. The latter can
                                                            // temporarily hide files that were already in SharePoint.
                                                            if (liveFolderId_1 && !/^f\d+$/.test(liveFolderId_1)) {
                                                                void host._refreshFolderFiles(liveFolderId_1, r.groupKey, true).catch(function () { return undefined; });
                                                                // Retry once for Graph's post-upload consistency window.
                                                                setTimeout(function () { return host._refreshFolderFiles(liveFolderId_1, r.groupKey, true).catch(function () { return undefined; }); }, 1500);
                                                            }
                                                            return [3 /*break*/, 4];
                                                        case 3:
                                                            err_2 = _d.sent();
                                                            host.setState({ docUploadBusy: false, docUploadRowKey: null, docUploadMsg: "Upload failed: ".concat((err_2 === null || err_2 === void 0 ? void 0 : err_2.message) || 'Error') });
                                                            return [3 /*break*/, 4];
                                                        case 4: return [2 /*return*/];
                                                    }
                                                });
                                            }); } }),
                                        isUploading ? '⏳...' : '↑ Upload'),
                                    React.createElement("button", { type: "button", onClick: function () { return void host._openSharePointFolder(r); }, title: "Open this folder in SharePoint", "aria-label": "Open ".concat(r.subFolderPath, " in SharePoint"), style: {
                                            width: 28, height: 27, padding: 0, borderRadius: 6,
                                            border: '1px solid #bfdbfe', background: '#eff6ff', color: '#1d4ed8',
                                            cursor: 'pointer', fontSize: 16, fontWeight: 700, lineHeight: 1,
                                            display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                                        } }, "\u2197")))));
                    }))),
                React.createElement("div", { style: { padding: '10px 14px', color: '#64748b', fontSize: 11, display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: '#f8fafc', borderTop: '1px solid #e2e8f0' } },
                    React.createElement("span", null,
                        "Showing ",
                        filtered.length === 0 ? 0 : safePage * PAGE_ROWS + 1,
                        "\u2013",
                        Math.min((safePage + 1) * PAGE_ROWS, filtered.length),
                        " of ",
                        filtered.length,
                        " rows"),
                    React.createElement("div", { style: { display: 'flex', gap: 3, alignItems: 'center' } },
                        React.createElement("button", { onClick: function () { return host.setState({ docListPage: Math.max(0, safePage - 1) }); }, disabled: safePage === 0, style: { border: '1px solid #cbd5e1', background: '#fff', borderRadius: 4, padding: '3px 8px', fontSize: 11, cursor: safePage === 0 ? 'not-allowed' : 'pointer', opacity: safePage === 0 ? 0.4 : 1 } }, "\u2039"),
                        Array.from({ length: totalPages }, function (_, i) { return (React.createElement("button", { key: i, onClick: function () { return host.setState({ docListPage: i }); }, style: {
                                border: i === safePage ? 'none' : '1px solid #cbd5e1',
                                background: i === safePage ? '#0078d4' : '#fff',
                                color: i === safePage ? '#fff' : '#334155',
                                borderRadius: 4, padding: '3px 8px', fontSize: 11,
                                fontWeight: i === safePage ? 700 : 400,
                                cursor: 'pointer',
                                minWidth: 26,
                            } }, i + 1)); }),
                        React.createElement("button", { onClick: function () { return host.setState({ docListPage: Math.min(totalPages - 1, safePage + 1) }); }, disabled: safePage >= totalPages - 1, style: { border: '1px solid #cbd5e1', background: '#fff', borderRadius: 4, padding: '3px 8px', fontSize: 11, cursor: safePage >= totalPages - 1 ? 'not-allowed' : 'pointer', opacity: safePage >= totalPages - 1 ? 0.4 : 1 } }, "\u203A")))))),
        (0, VesselsPage_1.renderClassifyDialog)(host)));
}
