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
exports.renderRecycleBinPage = renderRecycleBinPage;
var React = __importStar(require("react"));
var MAIN_FOLDERS = [
    'Technical & Crewing', 'Commercial & Chartering', 'Insurance',
    'Kaizen - Knowledge Bank', 'Knowledge Bank',
];
function getLocationDetails(item) {
    var parts = (item.original_path || '')
        .split('/')
        .map(function (part) { return part.trim(); })
        .filter(Boolean);
    if (parts[parts.length - 1] === item.name)
        parts.pop();
    var mainIndex = parts.findIndex(function (part) { return MAIN_FOLDERS.some(function (main) { return main.toLowerCase() === part.toLowerCase(); }); });
    var vessel = item.vessel_name || '';
    var category = item.category || '';
    var subCategory = item.sub_category || '';
    var remainder = [];
    if (mainIndex >= 0) {
        var beforeMain = parts.slice(0, mainIndex);
        var afterMain = parts.slice(mainIndex + 1);
        // Supports both legacy "Main folder/Vessel/..." and current "Vessels/Vessel/Main folder/..." paths.
        if (beforeMain[0] === 'Vessels') {
            vessel = vessel || beforeMain[1] || '';
            remainder = afterMain;
        }
        else {
            vessel = vessel || afterMain[0] || beforeMain[0] || '';
            remainder = afterMain.slice(1);
        }
        category = category || remainder[0] || parts[mainIndex] || '';
    }
    else {
        var vesselsIndex = parts.findIndex(function (part) { return part.toLowerCase() === 'vessels'; });
        if (vesselsIndex >= 0) {
            vessel = vessel || parts[vesselsIndex + 1] || '';
            remainder = parts.slice(vesselsIndex + 2);
        }
        else {
            vessel = vessel || parts[0] || '';
            remainder = parts.slice(1);
        }
        category = category || remainder[0] || '';
    }
    subCategory = subCategory || (remainder.length > 1 ? remainder[remainder.length - 1] : '');
    return {
        folderPath: parts.join(' / ') || '—',
        vessel: vessel || '—',
        category: category || item.main_folder || '—',
        subCategory: subCategory || '—',
    };
}
function isVessel(item) {
    return item.kind === 'vessel' || item.item_type === 'vessel';
}
function isFile(item) {
    return item.kind === 'file' || / file$/i.test(item.item_type || '');
}
function renderRecycleBinPage(host) {
    return React.createElement(RecycleBinContent, { host: host });
}
function RecycleBinContent(_a) {
    var _this = this;
    var host = _a.host;
    var _b = host.state, recycleBin = _b.recycleBin, panelLoading = _b.panelLoading;
    var _c = React.useState(new Set()), selectedIds = _c[0], setSelectedIds = _c[1];
    var _d = React.useState(null), permanentDeleteItems = _d[0], setPermanentDeleteItems = _d[1];
    var _e = React.useState(false), permanentDeleteBusy = _e[0], setPermanentDeleteBusy = _e[1];
    var _f = React.useState([]), permanentDeleteProgress = _f[0], setPermanentDeleteProgress = _f[1];
    var _g = React.useState(0), deleteElapsedSeconds = _g[0], setDeleteElapsedSeconds = _g[1];
    var _h = React.useState(null), autoCloseSeconds = _h[0], setAutoCloseSeconds = _h[1];
    var vesselItems = recycleBin.filter(isVessel);
    var fileItems = recycleBin.filter(function (item) { return !isVessel(item) && isFile(item); });
    var folderItems = recycleBin.filter(function (item) { return !isVessel(item) && !isFile(item); });
    var toggleOne = function (id) {
        setSelectedIds(function (previous) {
            var next = new Set(previous);
            if (next.has(id))
                next.delete(id);
            else
                next.add(id);
            return next;
        });
    };
    var toggleAllInSection = function (items) {
        var ids = items.map(function (item) { return item.id; });
        var selected = ids.length > 0 && ids.every(function (id) { return selectedIds.has(id); });
        setSelectedIds(function (previous) {
            var next = new Set(previous);
            ids.forEach(function (id) { return selected ? next.delete(id) : next.add(id); });
            return next;
        });
    };
    var restoreSelected = function () { return __awaiter(_this, void 0, void 0, function () {
        var items, _i, items_1, item;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    items = recycleBin.filter(function (item) { return selectedIds.has(item.id); });
                    _i = 0, items_1 = items;
                    _a.label = 1;
                case 1:
                    if (!(_i < items_1.length)) return [3 /*break*/, 4];
                    item = items_1[_i];
                    return [4 /*yield*/, host._restoreFromRecycleBin(item)];
                case 2:
                    _a.sent();
                    _a.label = 3;
                case 3:
                    _i++;
                    return [3 /*break*/, 1];
                case 4:
                    setSelectedIds(new Set());
                    return [2 /*return*/];
            }
        });
    }); };
    var permanentlyDeleteSelected = function () { return __awaiter(_this, void 0, void 0, function () {
        var items;
        return __generator(this, function (_a) {
            items = recycleBin.filter(function (item) { return selectedIds.has(item.id); });
            if (!items.length)
                return [2 /*return*/];
            openPermanentDeleteModal(items);
            return [2 /*return*/];
        });
    }); };
    var confirmPermanentDelete = function () { return __awaiter(_this, void 0, void 0, function () {
        var startedAt, _loop_1, _i, permanentDeleteItems_1, item;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!(permanentDeleteItems === null || permanentDeleteItems === void 0 ? void 0 : permanentDeleteItems.length))
                        return [2 /*return*/];
                    setPermanentDeleteBusy(true);
                    setDeleteElapsedSeconds(0);
                    setAutoCloseSeconds(null);
                    setPermanentDeleteProgress(permanentDeleteItems.map(function (item) { return ({ id: item.id, name: item.name, status: 'waiting' }); }));
                    startedAt = Date.now();
                    _loop_1 = function (item) {
                        var result;
                        return __generator(this, function (_b) {
                            switch (_b.label) {
                                case 0:
                                    setPermanentDeleteProgress(function (previous) { return previous.map(function (progress) { return progress.id === item.id ? __assign(__assign({}, progress), { status: 'deleting' }) : progress; }); });
                                    return [4 /*yield*/, host._permanentDeleteFromRecycleBin(item, true)];
                                case 1:
                                    result = _b.sent();
                                    setPermanentDeleteProgress(function (previous) { return previous.map(function (progress) { return progress.id === item.id ? __assign(__assign({}, progress), { status: result.ok ? 'success' : 'failed', message: result.message }) : progress; }); });
                                    setDeleteElapsedSeconds(Math.max(1, Math.round((Date.now() - startedAt) / 1000)));
                                    return [2 /*return*/];
                            }
                        });
                    };
                    _i = 0, permanentDeleteItems_1 = permanentDeleteItems;
                    _a.label = 1;
                case 1:
                    if (!(_i < permanentDeleteItems_1.length)) return [3 /*break*/, 4];
                    item = permanentDeleteItems_1[_i];
                    return [5 /*yield**/, _loop_1(item)];
                case 2:
                    _a.sent();
                    _a.label = 3;
                case 3:
                    _i++;
                    return [3 /*break*/, 1];
                case 4:
                    setPermanentDeleteBusy(false);
                    setSelectedIds(new Set());
                    setAutoCloseSeconds(10);
                    void host._goToView('recycle').catch(function () { return undefined; });
                    return [2 /*return*/];
            }
        });
    }); };
    var cancelPermanentDelete = function () {
        if (!permanentDeleteBusy) {
            setPermanentDeleteItems(null);
            setPermanentDeleteProgress([]);
            setAutoCloseSeconds(null);
        }
    };
    var openPermanentDeleteModal = function (items) {
        if (items.length) {
            setPermanentDeleteItems(items);
            setPermanentDeleteProgress([]);
            setDeleteElapsedSeconds(0);
            setAutoCloseSeconds(null);
        }
    };
    React.useEffect(function () {
        if (!permanentDeleteBusy)
            return undefined;
        var startedAt = Date.now() - deleteElapsedSeconds * 1000;
        var timer = window.setInterval(function () { return setDeleteElapsedSeconds(Math.floor((Date.now() - startedAt) / 1000)); }, 250);
        return function () { return window.clearInterval(timer); };
    }, [permanentDeleteBusy]);
    React.useEffect(function () {
        if (autoCloseSeconds === null)
            return undefined;
        if (autoCloseSeconds <= 0) {
            setPermanentDeleteItems(null);
            setPermanentDeleteProgress([]);
            setAutoCloseSeconds(null);
            return undefined;
        }
        var timer = window.setTimeout(function () { return setAutoCloseSeconds(function (value) { return value === null ? null : value - 1; }); }, 1000);
        return function () { return window.clearTimeout(timer); };
    }, [autoCloseSeconds]);
    var renderPermanentDeleteModal = function () {
        if (!(permanentDeleteItems === null || permanentDeleteItems === void 0 ? void 0 : permanentDeleteItems.length))
            return null;
        var itemCount = permanentDeleteItems.length;
        var itemNames = permanentDeleteItems.slice(0, 3).map(function (item) { return item.name; });
        var extraCount = itemCount - itemNames.length;
        var completed = !permanentDeleteBusy && permanentDeleteProgress.length > 0;
        var successCount = permanentDeleteProgress.filter(function (item) { return item.status === 'success'; }).length;
        var failedCount = permanentDeleteProgress.filter(function (item) { return item.status === 'failed'; }).length;
        return (React.createElement("div", { role: "presentation", onClick: cancelPermanentDelete, style: { position: 'fixed', inset: 0, zIndex: 10000, background: 'rgba(15, 23, 42, 0.52)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20 } },
            React.createElement("div", { role: "dialog", "aria-modal": "true", "aria-labelledby": "permanent-delete-title", onClick: function (event) { return event.stopPropagation(); }, style: { width: 500, maxWidth: '100%', background: '#fff', borderRadius: 12, boxShadow: '0 20px 50px rgba(15, 23, 42, 0.3)', overflow: 'hidden' } },
                React.createElement("div", { style: { padding: '20px 22px 14px', borderBottom: '1px solid #fee2e2' } },
                    React.createElement("div", { style: { fontSize: 17, fontWeight: 700, color: completed ? '#107c10' : '#991b1b' }, id: "permanent-delete-title" }, completed ? 'Permanent deletion completed' : permanentDeleteBusy ? 'Permanently deleting items' : "Permanently delete ".concat(itemCount, " item").concat(itemCount === 1 ? '' : 's', "?")),
                    React.createElement("p", { style: { margin: '10px 0 0', color: '#475569', fontSize: 13, lineHeight: 1.5 } }, completed ? "".concat(successCount, " item").concat(successCount === 1 ? '' : 's', " permanently deleted").concat(failedCount ? "; ".concat(failedCount, " could not be deleted.") : '.') : permanentDeleteBusy ? "Deleting one item at a time. Elapsed time: ".concat(deleteElapsedSeconds, "s.") : "This action cannot be undone. The selected item".concat(itemCount === 1 ? '' : 's', " will be permanently removed from the Recycle Bin."))),
                permanentDeleteProgress.length ? (React.createElement("div", { style: { margin: '14px 22px', maxHeight: 190, overflowY: 'auto', border: '1px solid #e2e8f0', borderRadius: 7 } }, permanentDeleteProgress.map(function (progress) { return React.createElement("div", { key: progress.id, style: { display: 'flex', gap: 8, alignItems: 'center', padding: '9px 10px', borderBottom: '1px solid #f1f5f9', fontSize: 12 } },
                    React.createElement("span", { style: { width: 18, textAlign: 'center' } }, progress.status === 'success' ? '✓' : progress.status === 'failed' ? '!' : progress.status === 'deleting' ? '…' : '○'),
                    React.createElement("span", { style: { flex: 1, color: '#334155' } }, progress.name),
                    React.createElement("span", { style: { color: progress.status === 'success' ? '#107c10' : progress.status === 'failed' ? '#a4262c' : '#64748b' } }, progress.status === 'success' ? 'Deleted' : progress.status === 'failed' ? progress.message || 'Failed' : progress.status === 'deleting' ? 'Deleting...' : 'Waiting')); }))) : React.createElement("div", { style: { margin: '14px 22px', padding: '10px 12px', background: '#f8fafc', borderRadius: 7, color: '#475569', fontSize: 12 } },
                    React.createElement("strong", { style: { color: '#334155' } }, itemNames.join(', ')),
                    extraCount > 0 ? " and ".concat(extraCount, " more") : ''),
                React.createElement("div", { style: { display: 'flex', justifyContent: 'flex-end', gap: 8, padding: '0 22px 20px' } }, completed ? React.createElement("button", { onClick: cancelPermanentDelete, style: buttonStyle('#fff', '#475569', '#cbd5e1') },
                    "Close now",
                    autoCloseSeconds !== null ? " (".concat(autoCloseSeconds, "s)") : '') : React.createElement(React.Fragment, null,
                    React.createElement("button", { onClick: cancelPermanentDelete, disabled: permanentDeleteBusy, style: buttonStyle('#fff', '#475569', '#cbd5e1') }, "Cancel"),
                    React.createElement("button", { onClick: confirmPermanentDelete, disabled: permanentDeleteBusy, style: __assign(__assign({}, buttonStyle('#dc2626', '#fff', '#dc2626')), { opacity: permanentDeleteBusy ? 0.65 : 1 }) }, permanentDeleteBusy ? 'Deleting...' : 'Delete permanently'))))));
    };
    var actionCell = function (item) { return (React.createElement("div", { style: { display: 'flex', justifyContent: 'flex-end', gap: 8 } },
        React.createElement("button", { onClick: function () { return host._restoreFromRecycleBin(item); }, style: buttonStyle('#dff6dd', '#107c10', '#86efac') }, "Restore"),
        React.createElement("button", { onClick: function () { return openPermanentDeleteModal([item]); }, style: buttonStyle('#fde7e9', '#a4262c', '#fca5a5') }, "Delete permanently"))); };
    var renderTable = function (title, items, columns, cells, empty) {
        var sectionAllSelected = items.length > 0 && items.every(function (item) { return selectedIds.has(item.id); });
        return (React.createElement("section", null,
            React.createElement("h3", { style: { margin: '0 0 10px', fontSize: 14, color: '#334155' } },
                title,
                " (",
                items.length,
                ")"),
            React.createElement("div", { style: { background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflowX: 'auto' } },
                React.createElement("table", { style: { width: '100%', minWidth: 760, borderCollapse: 'collapse', fontSize: 13 } },
                    React.createElement("thead", null,
                        React.createElement("tr", { style: { background: '#f8fafc', color: '#64748b', fontSize: 11, textTransform: 'uppercase', textAlign: 'left' } },
                            React.createElement("th", { style: { padding: 12, width: 34, textAlign: 'center' } },
                                React.createElement("input", { type: "checkbox", checked: sectionAllSelected, onChange: function () { return toggleAllInSection(items); }, disabled: !items.length })),
                            columns.map(function (column) { return React.createElement("th", { key: column, style: { padding: 12 } }, column); }),
                            React.createElement("th", { style: { padding: 12, textAlign: 'right' } }, "Actions"))),
                    React.createElement("tbody", null, !items.length ? React.createElement("tr", null,
                        React.createElement("td", { colSpan: columns.length + 2, style: { padding: 22, color: '#94a3b8', textAlign: 'center' } }, empty)) : items.map(function (item) { return (React.createElement("tr", { key: item.id, style: { borderTop: '1px solid #f1f5f9', background: selectedIds.has(item.id) ? '#f0f9ff' : '#fff' } },
                        React.createElement("td", { style: { padding: 12, textAlign: 'center' } },
                            React.createElement("input", { type: "checkbox", checked: selectedIds.has(item.id), onChange: function () { return toggleOne(item.id); } })),
                        cells(item).map(function (cell, index) { return React.createElement("td", { key: index, style: { padding: 12, color: index === 0 ? '#0f172a' : '#475569', fontWeight: index === 0 ? 600 : 400 } }, cell); }),
                        React.createElement("td", { style: { padding: 12 } }, actionCell(item)))); }))))));
    };
    return (React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 20 } },
        React.createElement("div", { style: { background: '#fffbe6', border: '1px solid #f59e0b', borderRadius: 12, padding: '18px 22px', display: 'flex', justifyContent: 'space-between', gap: 12, alignItems: 'center', flexWrap: 'wrap' } },
            React.createElement("div", null,
                React.createElement("h2", { style: { margin: 0, fontSize: 20, color: '#92400e' } }, "Recycle Bin"),
                React.createElement("p", { style: { margin: '4px 0 0', color: '#92400e', fontSize: 13 } }, "Deleted vessels, folders and files remain here until restored or permanently deleted.")),
            React.createElement("button", { onClick: function () { return host._goToView('recycle').catch(function () { return undefined; }); }, style: buttonStyle('#fff', '#475569', '#cbd5e1') }, "Refresh")),
        selectedIds.size > 0 && React.createElement("div", { style: { background: '#0f172a', color: '#fff', borderRadius: 10, padding: '12px 16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 10, flexWrap: 'wrap' } },
            React.createElement("span", null,
                selectedIds.size,
                " selected"),
            React.createElement("div", { style: { display: 'flex', gap: 8 } },
                React.createElement("button", { onClick: restoreSelected, style: buttonStyle('#10b981', '#fff', '#10b981') }, "Restore selected"),
                React.createElement("button", { onClick: permanentlyDeleteSelected, style: buttonStyle('#ef4444', '#fff', '#ef4444') }, "Delete permanently"))),
        panelLoading ? React.createElement("div", { style: { padding: 48, textAlign: 'center', color: '#64748b' } }, "Loading Recycle Bin...") : React.createElement(React.Fragment, null,
            renderTable('Deleted vessels', vesselItems, ['Vessel', 'Original path', 'Deleted'], function (item) { return [item.name, item.original_path || '—', item.deleted_at ? new Date(item.deleted_at).toLocaleString() : '—']; }, 'No deleted vessels.'),
            renderTable('Deleted normal folders', folderItems, ['Folder', 'Original path', 'Vessel', 'Category', 'Sub-category'], function (item) { var location = getLocationDetails(item); return [item.name, location.folderPath, location.vessel, location.category, location.subCategory]; }, 'No deleted normal folders.'),
            renderTable('Deleted individual files', fileItems, ['File', 'Folder path', 'Vessel', 'Category', 'Sub-category'], function (item) { var location = getLocationDetails(item); return [item.name, location.folderPath, location.vessel, location.category, location.subCategory]; }, 'No deleted individual files.')),
        renderPermanentDeleteModal()));
}
function buttonStyle(background, color, border) {
    return { background: background, color: color, border: "1px solid ".concat(border), borderRadius: 6, padding: '6px 10px', cursor: 'pointer', fontSize: 12, fontWeight: 600, whiteSpace: 'nowrap' };
}
