"use strict";
/**
 * deltaSync.ts
 *
 * Graph delta-query sync service.
 *
 * Detects folders created, renamed, moved, or deleted directly in SharePoint
 * Online (outside the SPFx app) and merges them into the correct position in
 * the app's in-memory folder tree — treating SPO as the source of truth.
 *
 * Key design decisions:
 *  - Uses /drives/{driveId}/root/delta so changes anywhere in the drive are
 *    captured, not just under a specific known parent.
 *  - Resolves each changed item's true parent via parentReference.path.
 *  - If a changed item's parent branch hasn't been loaded yet, triggers a
 *    targeted re-fetch of that path rather than guessing its position.
 *  - Persists the deltaLink in localStorage (keyed by driveId) so each poll
 *    only fetches what changed since the last sync.
 *  - Designed to be called on a timer (every 30–60 s) and/or on demand.
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
exports.__esModule = true;
exports.createSyncScheduler = exports.removeNodeFromMap = exports.mergeNodeIntoMap = exports.pollDelta = exports.fetchFolderChildren = void 0;
// ── Delta link persistence (localStorage) ────────────────────────────────────
var DELTA_LINK_KEY_PREFIX = 'vesselDMS_deltaLink_';
function loadDeltaLink(driveId) {
    try {
        return localStorage.getItem("".concat(DELTA_LINK_KEY_PREFIX).concat(driveId));
    }
    catch (_a) {
        return null;
    }
}
function saveDeltaLink(driveId, link) {
    try {
        localStorage.setItem("".concat(DELTA_LINK_KEY_PREFIX).concat(driveId), link);
    }
    catch ( /* storage unavailable — non-fatal */_a) { /* storage unavailable — non-fatal */ }
}
// ── Graph helpers ─────────────────────────────────────────────────────────────
/**
 * Walk all pages of a delta/list response, collecting every item.
 * Returns { items, nextDeltaLink }.
 */
function drainPages(client, firstUrl) {
    var _a;
    return __awaiter(this, void 0, void 0, function () {
        var items, url, deltaLink, relativeUrl, page;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    items = [];
                    url = firstUrl;
                    deltaLink = '';
                    _b.label = 1;
                case 1:
                    if (!url) return [3 /*break*/, 3];
                    relativeUrl = url.startsWith('https://')
                        ? url.replace('https://graph.microsoft.com/v1.0', '')
                        : url;
                    return [4 /*yield*/, client.api(relativeUrl).get()];
                case 2:
                    page = _b.sent();
                    if (page.value)
                        items.push.apply(items, page.value);
                    if (page['@odata.deltaLink']) {
                        deltaLink = page['@odata.deltaLink'];
                        return [3 /*break*/, 3];
                    }
                    url = (_a = page['@odata.nextLink']) !== null && _a !== void 0 ? _a : '';
                    return [3 /*break*/, 1];
                case 3: return [2 /*return*/, { items: items, deltaLink: deltaLink }];
            }
        });
    });
}
/**
 * Fetch children of a specific path to rebuild an unloaded branch.
 */
function fetchFolderChildren(client, siteId, driveId, folderPath) {
    var _a;
    return __awaiter(this, void 0, void 0, function () {
        var encoded, url, result, _b;
        return __generator(this, function (_c) {
            switch (_c.label) {
                case 0:
                    encoded = folderPath
                        .split('/')
                        .map(function (s) { return encodeURIComponent(s); })
                        .join('/');
                    _c.label = 1;
                case 1:
                    _c.trys.push([1, 3, , 4]);
                    url = "/sites/".concat(siteId, "/drives/").concat(driveId, "/root:/").concat(encoded, ":/children?$filter=folder ne null&$select=id,name,parentReference,folder");
                    return [4 /*yield*/, client.api(url).get()];
                case 2:
                    result = _c.sent();
                    return [2 /*return*/, ((_a = result.value) !== null && _a !== void 0 ? _a : []).map(function (item) { return graphItemToNode(item); })];
                case 3:
                    _b = _c.sent();
                    return [2 /*return*/, []];
                case 4: return [2 /*return*/];
            }
        });
    });
}
exports.fetchFolderChildren = fetchFolderChildren;
// ── Item → SpoFolderNode mapping ──────────────────────────────────────────────
/**
 * Extract the server-relative path from a Graph item's parentReference.
 * parentReference.path looks like:
 *   "/drives/{driveId}/root:/Vessel Management/Folder-1 Technical & Crewing"
 * We strip the "/drives/{driveId}/root:" prefix to get the clean path.
 */
function resolveServerPath(item) {
    var _a, _b;
    var raw = (_b = (_a = item.parentReference) === null || _a === void 0 ? void 0 : _a.path) !== null && _b !== void 0 ? _b : '';
    var rootMarker = '/root:';
    var idx = raw.indexOf(rootMarker);
    var parentPath = idx !== -1 ? raw.substring(idx + rootMarker.length) : '';
    // Decode URI components (Graph encodes spaces as %20 etc.)
    var decodedParent = decodeURIComponent(parentPath);
    return decodedParent ? "".concat(decodedParent, "/").concat(item.name) : "/".concat(item.name);
}
function graphItemToNode(item) {
    var _a, _b;
    return {
        id: item.id,
        name: item.name,
        parentId: (_b = (_a = item.parentReference) === null || _a === void 0 ? void 0 : _a.id) !== null && _b !== void 0 ? _b : null,
        serverRelativePath: resolveServerPath(item),
        children: [],
        deleted: !!item.deleted
    };
}
// ── Main sync function ────────────────────────────────────────────────────────
/**
 * Poll Graph for changes since the last sync.
 *
 * On first call (no stored deltaLink) it performs a full drive scan to build
 * the initial deltaLink baseline — this may be slow for large drives but only
 * happens once per browser session.
 *
 * @param client   MSGraphClientV3
 * @param siteId   SharePoint site ID
 * @param driveId  Document library drive ID
 * @returns DeltaSyncResult with added/updated/deleted items and the new deltaLink
 */
function pollDelta(client, siteId, driveId) {
    return __awaiter(this, void 0, void 0, function () {
        var storedLink, startUrl, _a, items, deltaLink, added, updated, deleted, _i, items_1, item, node;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    storedLink = loadDeltaLink(driveId);
                    startUrl = storedLink
                        ? storedLink
                        : "/sites/".concat(siteId, "/drives/").concat(driveId, "/root/delta?$select=id,name,parentReference,folder,deleted");
                    return [4 /*yield*/, drainPages(client, startUrl)];
                case 1:
                    _a = _b.sent(), items = _a.items, deltaLink = _a.deltaLink;
                    if (deltaLink)
                        saveDeltaLink(driveId, deltaLink);
                    added = [];
                    updated = [];
                    deleted = [];
                    for (_i = 0, items_1 = items; _i < items_1.length; _i++) {
                        item = items_1[_i];
                        // Only care about folders (items without a 'folder' facet are files)
                        if (!item.folder && !item.deleted)
                            continue;
                        if (item.deleted) {
                            deleted.push(item.id);
                        }
                        else {
                            node = graphItemToNode(item);
                            // Distinguish add vs update: callers can compare against their existing tree
                            // We emit both as separate arrays; the merger decides based on its own state.
                            added.push(node); // caller deduplicates against existing tree
                        }
                    }
                    return [2 /*return*/, { added: added, updated: updated, deleted: deleted, newDeltaLink: deltaLink }];
            }
        });
    });
}
exports.pollDelta = pollDelta;
// ── Tree merge helpers (used by the React component) ─────────────────────────
/**
 * Insert or update a node at the correct position in a flat id→node map.
 * The map is keyed by folder ID for O(1) lookups.
 *
 * If the node's parent is not yet in the map (branch not loaded), returns
 * the parentId so the caller can trigger a targeted re-fetch.
 */
function mergeNodeIntoMap(map, node) {
    // Update or insert the node itself
    var existing = map.get(node.id);
    if (existing) {
        // Update name/path in place, preserve children
        existing.name = node.name;
        existing.serverRelativePath = node.serverRelativePath;
        existing.parentId = node.parentId;
    }
    else {
        map.set(node.id, __assign(__assign({}, node), { children: [] }));
    }
    // Wire into parent's children array
    if (node.parentId) {
        var parent_1 = map.get(node.parentId);
        if (!parent_1) {
            // Parent branch not loaded yet — caller must re-fetch
            return { missingParentId: node.parentId };
        }
        var alreadyChild = parent_1.children.some(function (c) { return c.id === node.id; });
        if (!alreadyChild) {
            parent_1.children.push(map.get(node.id));
        }
    }
    return { missingParentId: null };
}
exports.mergeNodeIntoMap = mergeNodeIntoMap;
/**
 * Remove a deleted node (and all its descendants) from the map.
 */
function removeNodeFromMap(map, deletedId) {
    var node = map.get(deletedId);
    if (!node)
        return;
    // Recursively remove children first
    for (var _i = 0, _a = node.children; _i < _a.length; _i++) {
        var child = _a[_i];
        removeNodeFromMap(map, child.id);
    }
    // Detach from parent
    if (node.parentId) {
        var parent_2 = map.get(node.parentId);
        if (parent_2) {
            parent_2.children = parent_2.children.filter(function (c) { return c.id !== deletedId; });
        }
    }
    map["delete"](deletedId);
}
exports.removeNodeFromMap = removeNodeFromMap;
/**
 * Create a scheduler that calls pollDelta on a fixed interval.
 *
 * @param client        MSGraphClientV3
 * @param siteId        SharePoint site ID
 * @param driveId       Document library drive ID
 * @param intervalMs    Poll interval in milliseconds (default 45 000 = 45 s)
 * @param onResult      Callback invoked with each DeltaSyncResult
 * @param onError       Optional error callback
 */
function createSyncScheduler(client, siteId, driveId, onResult, intervalMs, onError) {
    var _this = this;
    if (intervalMs === void 0) { intervalMs = 45000; }
    var timerId = null;
    var run = function () { return __awaiter(_this, void 0, void 0, function () {
        var result, err_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    return [4 /*yield*/, pollDelta(client, siteId, driveId)];
                case 1:
                    result = _a.sent();
                    onResult(result);
                    return [2 /*return*/, result];
                case 2:
                    err_1 = _a.sent();
                    onError === null || onError === void 0 ? void 0 : onError(err_1);
                    return [2 /*return*/, null];
                case 3: return [2 /*return*/];
            }
        });
    }); };
    return {
        start: function () {
            if (timerId !== null)
                return;
            timerId = setInterval(function () { run()["catch"](function () { return undefined; }); }, intervalMs);
        },
        stop: function () {
            if (timerId !== null) {
                clearInterval(timerId);
                timerId = null;
            }
        },
        triggerNow: run
    };
}
exports.createSyncScheduler = createSyncScheduler;
