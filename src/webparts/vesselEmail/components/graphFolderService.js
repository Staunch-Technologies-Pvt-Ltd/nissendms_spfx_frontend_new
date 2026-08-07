"use strict";
/**
 * graphFolderService.ts
 *
 * Idempotent, recursive Graph API folder-creation service.
 *
 * Responsibilities:
 *  - Ensure the top-level "Vessel Management" root exists.
 *  - For each MainFolder (Folder-1/2/3):
 *      • Create the MainFolder itself if absent.
 *      • Create all "common" sub-trees once (skip if already present).
 *      • Create the {VesselName} sub-folder and its full per-vessel tree.
 *  - Return a per-path success/failure log so the caller can retry only
 *    the paths that failed (fully idempotent re-runs).
 *
 * Graph endpoint used:
 *   POST /sites/{siteId}/drives/{driveId}/root:/{parentPath}:/children
 *   { name, folder: {}, "@microsoft.graph.conflictBehavior": "fail" }
 *
 * A 409 Conflict response means the folder already exists → treated as success.
 */
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
exports.retryFailedFolders = exports.createVesselFolders = exports.ensureRootStructure = void 0;
var vesselFolderTemplate_1 = require("./vesselFolderTemplate");
// ── Internal helpers ──────────────────────────────────────────────────────────
/**
 * Encode each path segment individually (preserves "/" separators).
 * Graph requires each segment to be URL-encoded but the "/" must stay literal.
 */
function encodePath(path) {
    return path
        .split('/')
        .map(function (seg) { return encodeURIComponent(seg); })
        .join('/');
}
/**
 * Attempt to create a single folder under parentPath.
 * Returns the folder item on success/already-exists, throws on real errors.
 */
function createFolder(client, siteId, driveId, parentPath, folderName) {
    var _a, _b, _c, _d, _e, _f;
    return __awaiter(this, void 0, void 0, function () {
        var url, response, err_1, status_1, isConflict, fullPath, existing, _g;
        return __generator(this, function (_h) {
            switch (_h.label) {
                case 0:
                    url = parentPath
                        ? "/sites/".concat(siteId, "/drives/").concat(driveId, "/root:/").concat(encodePath(parentPath), ":/children")
                        : "/sites/".concat(siteId, "/drives/").concat(driveId, "/root/children");
                    console.log("[VesselDMS] createFolder \u2192 POST ".concat(url, " name=\"").concat(folderName, "\""));
                    _h.label = 1;
                case 1:
                    _h.trys.push([1, 3, , 8]);
                    return [4 /*yield*/, client
                            .api(url)
                            .post({
                            name: folderName,
                            folder: {},
                            '@microsoft.graph.conflictBehavior': 'fail'
                        })];
                case 2:
                    response = _h.sent();
                    console.log("[VesselDMS] createFolder \u2713 created id=".concat(response.id, " \"").concat(folderName, "\""));
                    return [2 /*return*/, { id: response.id, existed: false }];
                case 3:
                    err_1 = _h.sent();
                    status_1 = (_d = (_c = (_a = err_1 === null || err_1 === void 0 ? void 0 : err_1.statusCode) !== null && _a !== void 0 ? _a : (_b = err_1 === null || err_1 === void 0 ? void 0 : err_1.response) === null || _b === void 0 ? void 0 : _b.status) !== null && _c !== void 0 ? _c : err_1 === null || err_1 === void 0 ? void 0 : err_1.code) !== null && _d !== void 0 ? _d : 0;
                    isConflict = status_1 === 409 ||
                        ((_e = err_1 === null || err_1 === void 0 ? void 0 : err_1.message) !== null && _e !== void 0 ? _e : '').toLowerCase().includes('namealreadyexists');
                    if (!isConflict) return [3 /*break*/, 7];
                    console.log("[VesselDMS] createFolder \u21A9 existed \"".concat(folderName, "\""));
                    _h.label = 4;
                case 4:
                    _h.trys.push([4, 6, , 7]);
                    fullPath = parentPath ? "".concat(parentPath, "/").concat(folderName) : folderName;
                    return [4 /*yield*/, client
                            .api("/sites/".concat(siteId, "/drives/").concat(driveId, "/root:/").concat(encodePath(fullPath)))
                            .get()];
                case 5:
                    existing = _h.sent();
                    return [2 /*return*/, { id: existing.id, existed: true }];
                case 6:
                    _g = _h.sent();
                    return [2 /*return*/, { id: '', existed: true }];
                case 7:
                    console.error("[VesselDMS] createFolder \u2717 FAILED \"".concat(folderName, "\" status=").concat(status_1), (_f = err_1 === null || err_1 === void 0 ? void 0 : err_1.message) !== null && _f !== void 0 ? _f : err_1);
                    throw err_1;
                case 8: return [2 /*return*/];
            }
        });
    });
}
/**
 * Recursively walk a FolderNode tree, creating each node under its parent path.
 * Pushes a FolderResult entry for every node attempted.
 */
function createTree(client, siteId, driveId, parentPath, nodes, log) {
    var _a;
    return __awaiter(this, void 0, void 0, function () {
        var _i, nodes_1, node, fullPath, _b, id, existed, err_2;
        return __generator(this, function (_c) {
            switch (_c.label) {
                case 0:
                    _i = 0, nodes_1 = nodes;
                    _c.label = 1;
                case 1:
                    if (!(_i < nodes_1.length)) return [3 /*break*/, 8];
                    node = nodes_1[_i];
                    fullPath = "".concat(parentPath, "/").concat(node.name);
                    _c.label = 2;
                case 2:
                    _c.trys.push([2, 6, , 7]);
                    return [4 /*yield*/, createFolder(client, siteId, driveId, parentPath, node.name)];
                case 3:
                    _b = _c.sent(), id = _b.id, existed = _b.existed;
                    log.push({ path: fullPath, id: id, status: existed ? 'existed' : 'created' });
                    if (!(node.children && node.children.length > 0)) return [3 /*break*/, 5];
                    return [4 /*yield*/, createTree(client, siteId, driveId, fullPath, node.children, log)];
                case 4:
                    _c.sent();
                    _c.label = 5;
                case 5: return [3 /*break*/, 7];
                case 6:
                    err_2 = _c.sent();
                    log.push({
                        path: fullPath,
                        status: 'failed',
                        error: (_a = err_2 === null || err_2 === void 0 ? void 0 : err_2.message) !== null && _a !== void 0 ? _a : String(err_2)
                    });
                    return [3 /*break*/, 7];
                case 7:
                    _i++;
                    return [3 /*break*/, 1];
                case 8: return [2 /*return*/];
            }
        });
    });
}
// ── Public API ────────────────────────────────────────────────────────────────
/**
 * Ensure the "Vessel Management" root and all three MainFolder roots exist.
 * Safe to call on every app load — all operations are idempotent.
 */
function ensureRootStructure(client, siteId, driveId) {
    var _a, _b;
    return __awaiter(this, void 0, void 0, function () {
        var log, _c, id, existed, err_3, _i, MAIN_FOLDERS_1, mf, mfPath, _d, id, existed, err_4;
        return __generator(this, function (_e) {
            switch (_e.label) {
                case 0:
                    log = [];
                    console.log("[VesselDMS] ensureRootStructure siteId=".concat(siteId, " driveId=").concat(driveId));
                    if (!siteId || !driveId) {
                        console.error('[VesselDMS] ensureRootStructure aborted — siteId or driveId is empty!');
                        return [2 /*return*/, log];
                    }
                    _e.label = 1;
                case 1:
                    _e.trys.push([1, 3, , 4]);
                    return [4 /*yield*/, createFolder(client, siteId, driveId, '', vesselFolderTemplate_1.VESSEL_MANAGEMENT_ROOT)];
                case 2:
                    _c = _e.sent(), id = _c.id, existed = _c.existed;
                    log.push({
                        path: vesselFolderTemplate_1.VESSEL_MANAGEMENT_ROOT,
                        id: id,
                        status: existed ? 'existed' : 'created'
                    });
                    return [3 /*break*/, 4];
                case 3:
                    err_3 = _e.sent();
                    log.push({
                        path: vesselFolderTemplate_1.VESSEL_MANAGEMENT_ROOT,
                        status: 'failed',
                        error: (_a = err_3 === null || err_3 === void 0 ? void 0 : err_3.message) !== null && _a !== void 0 ? _a : String(err_3)
                    });
                    return [2 /*return*/, log]; // Can't continue without the root
                case 4:
                    _i = 0, MAIN_FOLDERS_1 = vesselFolderTemplate_1.MAIN_FOLDERS;
                    _e.label = 5;
                case 5:
                    if (!(_i < MAIN_FOLDERS_1.length)) return [3 /*break*/, 12];
                    mf = MAIN_FOLDERS_1[_i];
                    mfPath = "".concat(vesselFolderTemplate_1.VESSEL_MANAGEMENT_ROOT, "/").concat(mf.name);
                    _e.label = 6;
                case 6:
                    _e.trys.push([6, 8, , 9]);
                    return [4 /*yield*/, createFolder(client, siteId, driveId, vesselFolderTemplate_1.VESSEL_MANAGEMENT_ROOT, mf.name)];
                case 7:
                    _d = _e.sent(), id = _d.id, existed = _d.existed;
                    log.push({ path: mfPath, id: id, status: existed ? 'existed' : 'created' });
                    return [3 /*break*/, 9];
                case 8:
                    err_4 = _e.sent();
                    log.push({ path: mfPath, status: 'failed', error: (_b = err_4 === null || err_4 === void 0 ? void 0 : err_4.message) !== null && _b !== void 0 ? _b : String(err_4) });
                    return [3 /*break*/, 11];
                case 9:
                    if (!(mf.commonTree.length > 0)) return [3 /*break*/, 11];
                    return [4 /*yield*/, createTree(client, siteId, driveId, mfPath, mf.commonTree, log)];
                case 10:
                    _e.sent();
                    _e.label = 11;
                case 11:
                    _i++;
                    return [3 /*break*/, 5];
                case 12: return [2 /*return*/, log];
            }
        });
    });
}
exports.ensureRootStructure = ensureRootStructure;
/**
 * Create the full per-vessel folder tree for a new vessel.
 *
 * @param client      MSGraphClientV3 from SPFx context
 * @param siteId      SharePoint site ID
 * @param driveId     Document library drive ID
 * @param vesselName  Exact vessel name (used as the folder name)
 *
 * Returns a VesselFolderCreationResult with per-path success/failure entries.
 * Re-running for an existing vessel is safe — already-present folders are skipped.
 */
function createVesselFolders(client, siteId, driveId, vesselName) {
    var _a;
    return __awaiter(this, void 0, void 0, function () {
        var log, rootLog, rootFailed, _loop_1, _i, MAIN_FOLDERS_2, mf, success;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    log = [];
                    console.log("[VesselDMS] createVesselFolders START vessel=\"".concat(vesselName, "\" siteId=").concat(siteId, " driveId=").concat(driveId));
                    if (!siteId || !driveId) {
                        console.error('[VesselDMS] createVesselFolders aborted — siteId or driveId is empty!');
                        return [2 /*return*/, { vesselName: vesselName, siteId: siteId, driveId: driveId, results: log, success: false }];
                    }
                    return [4 /*yield*/, ensureRootStructure(client, siteId, driveId)];
                case 1:
                    rootLog = _b.sent();
                    log.push.apply(log, rootLog);
                    rootFailed = rootLog.some(function (r) { return r.path === vesselFolderTemplate_1.VESSEL_MANAGEMENT_ROOT && r.status === 'failed'; });
                    if (rootFailed) {
                        return [2 /*return*/, { vesselName: vesselName, siteId: siteId, driveId: driveId, results: log, success: false }];
                    }
                    _loop_1 = function (mf) {
                        var mfPath, mfFailed, vesselFolderPath, _c, id, existed, err_5;
                        return __generator(this, function (_d) {
                            switch (_d.label) {
                                case 0:
                                    mfPath = "".concat(vesselFolderTemplate_1.VESSEL_MANAGEMENT_ROOT, "/").concat(mf.name);
                                    mfFailed = log.some(function (r) { return r.path === mfPath && r.status === 'failed'; });
                                    if (mfFailed)
                                        return [2 /*return*/, "continue"];
                                    if (mf.perVesselTree.length === 0)
                                        return [2 /*return*/, "continue"];
                                    vesselFolderPath = "".concat(mfPath, "/").concat(vesselName);
                                    _d.label = 1;
                                case 1:
                                    _d.trys.push([1, 3, , 4]);
                                    return [4 /*yield*/, createFolder(client, siteId, driveId, mfPath, vesselName)];
                                case 2:
                                    _c = _d.sent(), id = _c.id, existed = _c.existed;
                                    log.push({
                                        path: vesselFolderPath,
                                        id: id,
                                        status: existed ? 'existed' : 'created'
                                    });
                                    return [3 /*break*/, 4];
                                case 3:
                                    err_5 = _d.sent();
                                    log.push({
                                        path: vesselFolderPath,
                                        status: 'failed',
                                        error: (_a = err_5 === null || err_5 === void 0 ? void 0 : err_5.message) !== null && _a !== void 0 ? _a : String(err_5)
                                    });
                                    return [2 /*return*/, "continue"];
                                case 4: 
                                // Recursively create the per-vessel sub-tree
                                return [4 /*yield*/, createTree(client, siteId, driveId, vesselFolderPath, mf.perVesselTree, log)];
                                case 5:
                                    // Recursively create the per-vessel sub-tree
                                    _d.sent();
                                    return [2 /*return*/];
                            }
                        });
                    };
                    _i = 0, MAIN_FOLDERS_2 = vesselFolderTemplate_1.MAIN_FOLDERS;
                    _b.label = 2;
                case 2:
                    if (!(_i < MAIN_FOLDERS_2.length)) return [3 /*break*/, 5];
                    mf = MAIN_FOLDERS_2[_i];
                    return [5 /*yield**/, _loop_1(mf)];
                case 3:
                    _b.sent();
                    _b.label = 4;
                case 4:
                    _i++;
                    return [3 /*break*/, 2];
                case 5:
                    success = log.every(function (r) { return r.status !== 'failed'; });
                    console.log("[VesselDMS] createVesselFolders DONE vessel=\"".concat(vesselName, "\" success=").concat(success), log);
                    return [2 /*return*/, { vesselName: vesselName, siteId: siteId, driveId: driveId, results: log, success: success }];
            }
        });
    });
}
exports.createVesselFolders = createVesselFolders;
/**
 * Retry only the paths that previously failed.
 * Pass the results array from a previous createVesselFolders call.
 * Returns a merged result with updated statuses.
 */
function retryFailedFolders(client, siteId, driveId, previousResults) {
    var _a;
    return __awaiter(this, void 0, void 0, function () {
        var failed, updated, _loop_2, _i, failed_1, entry;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    failed = previousResults.filter(function (r) { return r.status === 'failed'; });
                    updated = __spreadArray([], previousResults, true);
                    _loop_2 = function (entry) {
                        var lastSlash, parentPath, folderName, _c, id, existed, idx, err_6, idx;
                        return __generator(this, function (_d) {
                            switch (_d.label) {
                                case 0:
                                    lastSlash = entry.path.lastIndexOf('/');
                                    parentPath = lastSlash > 0 ? entry.path.substring(0, lastSlash) : '';
                                    folderName = entry.path.substring(lastSlash + 1);
                                    _d.label = 1;
                                case 1:
                                    _d.trys.push([1, 3, , 4]);
                                    return [4 /*yield*/, createFolder(client, siteId, driveId, parentPath, folderName)];
                                case 2:
                                    _c = _d.sent(), id = _c.id, existed = _c.existed;
                                    idx = updated.findIndex(function (r) { return r.path === entry.path; });
                                    if (idx !== -1) {
                                        updated[idx] = { path: entry.path, id: id, status: existed ? 'existed' : 'created' };
                                    }
                                    return [3 /*break*/, 4];
                                case 3:
                                    err_6 = _d.sent();
                                    idx = updated.findIndex(function (r) { return r.path === entry.path; });
                                    if (idx !== -1) {
                                        updated[idx] = __assign(__assign({}, updated[idx]), { error: (_a = err_6 === null || err_6 === void 0 ? void 0 : err_6.message) !== null && _a !== void 0 ? _a : String(err_6) });
                                    }
                                    return [3 /*break*/, 4];
                                case 4: return [2 /*return*/];
                            }
                        });
                    };
                    _i = 0, failed_1 = failed;
                    _b.label = 1;
                case 1:
                    if (!(_i < failed_1.length)) return [3 /*break*/, 4];
                    entry = failed_1[_i];
                    return [5 /*yield**/, _loop_2(entry)];
                case 2:
                    _b.sent();
                    _b.label = 3;
                case 3:
                    _i++;
                    return [3 /*break*/, 1];
                case 4: return [2 /*return*/, updated];
            }
        });
    });
}
exports.retryFailedFolders = retryFailedFolders;
