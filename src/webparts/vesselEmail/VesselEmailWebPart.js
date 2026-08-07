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
var React = require("react");
var ReactDom = require("react-dom");
var sp_core_library_1 = require("@microsoft/sp-core-library");
var sp_property_pane_1 = require("@microsoft/sp-property-pane");
var sp_webpart_base_1 = require("@microsoft/sp-webpart-base");
var strings = require("VesselEmailWebPartStrings");
var VesselEmail_1 = require("./components/VesselEmail");
var VesselEmailWebPart = /** @class */ (function (_super) {
    __extends(VesselEmailWebPart, _super);
    function VesselEmailWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._isDarkTheme = false;
        _this._environmentMessage = '';
        _this._userEmail = '';
        _this._sessionId = '';
        // True once the bypass-login call has SETTLED (success, no-session-stub-mode,
        // failure, or network error) — NOT the same as having a session id. The React
        // component uses this to know it's safe to load data even when the backend
        // is running without a database (where session_id is legitimately null).
        _this._sessionInitialized = false;
        return _this;
    }
    VesselEmailWebPart.prototype.render = function () {
        var element = React.createElement(VesselEmail_1["default"], {
            apiBaseUrl: this.properties.apiBaseUrl,
            apiToken: this.properties.apiToken,
            isDarkTheme: this._isDarkTheme,
            userDisplayName: this.context.pageContext.user.displayName,
            userEmail: this._userEmail || this.context.pageContext.user.email,
            graphClient: this._graphClient,
            siteId: this._siteId,
            driveId: this._driveId,
            sessionId: this._sessionId,
            sessionInitialized: this._sessionInitialized
        });
        ReactDom.render(element, this.domElement);
    };
    VesselEmailWebPart.prototype.onInit = function () {
        var _this = this;
        return this._getEnvironmentMessage().then(function (message) {
            _this._environmentMessage = message;
        }).then(function () { return _this._initGraphContextAndSession(); });
    };
    VesselEmailWebPart.prototype._initGraphContextAndSession = function () {
        return __awaiter(this, void 0, void 0, function () {
            var _a, me, siteUrl, siteRes, drivesRes, docLib, e_1;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 5, , 6]);
                        _a = this;
                        return [4 /*yield*/, this.context.msGraphClientFactory.getClient('3')];
                    case 1:
                        _a._graphClient = _b.sent();
                        return [4 /*yield*/, this._graphClient.api('/me').select('mail,userPrincipalName').get()];
                    case 2:
                        me = _b.sent();
                        this._userEmail = me.mail || me.userPrincipalName || this.context.pageContext.user.email;
                        siteUrl = this.context.pageContext.site.absoluteUrl;
                        return [4 /*yield*/, this._graphClient
                                .api("/sites/".concat(new URL(siteUrl).hostname, ":").concat(new URL(siteUrl).pathname))
                                .get()];
                    case 3:
                        siteRes = _b.sent();
                        this._siteId = siteRes.id;
                        return [4 /*yield*/, this._graphClient.api("/sites/".concat(this._siteId, "/drives")).get()];
                    case 4:
                        drivesRes = _b.sent();
                        docLib = drivesRes.value.find(function (d) { return d.driveType === 'documentLibrary' && d.name === 'Documents'; }) || drivesRes.value[0];
                        this._driveId = docLib === null || docLib === void 0 ? void 0 : docLib.id;
                        return [3 /*break*/, 6];
                    case 5:
                        e_1 = _b.sent();
                        console.warn('[VesselDMS] Graph context init failed:', e_1);
                        return [3 /*break*/, 6];
                    case 6: 
                    // Step 2: create a backend session BEFORE first render so _loadData has a valid session_id
                    return [4 /*yield*/, this._ensureSession()];
                    case 7:
                        // Step 2: create a backend session BEFORE first render so _loadData has a valid session_id
                        _b.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    VesselEmailWebPart.prototype._ensureSession = function () {
        var _a, _b;
        return __awaiter(this, void 0, void 0, function () {
            var base, res, data, _c, _d, _e, e_2;
            return __generator(this, function (_f) {
                switch (_f.label) {
                    case 0:
                        if (!this._userEmail) {
                            // Fallback: use SharePoint page context email directly
                            this._userEmail = this.context.pageContext.user.email || '';
                        }
                        if (!this._userEmail) {
                            // No email could be resolved at all — still mark init as settled and
                            // render so the UI (and vessel list) isn't stuck waiting forever.
                            this._sessionInitialized = true;
                            this.render();
                            return [2 /*return*/];
                        }
                        base = (this.properties.apiBaseUrl || 'http://localhost:8000').replace(/\/$/, '');
                        _f.label = 1;
                    case 1:
                        _f.trys.push([1, 7, 8, 9]);
                        return [4 /*yield*/, fetch("".concat(base, "/api/auth/bypass-login"), {
                                method: 'POST',
                                headers: { 'Content-Type': 'application/json' },
                                body: JSON.stringify({
                                    email: this._userEmail,
                                    display_name: this.context.pageContext.user.displayName,
                                    tenant_id: ((_b = (_a = this.context.pageContext.aadInfo) === null || _a === void 0 ? void 0 : _a.tenantId) === null || _b === void 0 ? void 0 : _b.toString()) || ''
                                })
                            })];
                    case 2:
                        res = _f.sent();
                        if (!res.ok) return [3 /*break*/, 4];
                        return [4 /*yield*/, res.json()];
                    case 3:
                        data = _f.sent();
                        if (data.session_id) {
                            this._sessionId = data.session_id;
                        }
                        else {
                            // Backend is running in stub/no-DB mode (session_id intentionally null —
                            // require_session() allows unauthenticated calls there). This is NOT a
                            // failure: fall through and still mark init as done so the UI loads data.
                            console.warn('[VesselDMS] bypass-login succeeded without a session_id (stub/no-DB mode).');
                        }
                        return [3 /*break*/, 6];
                    case 4:
                        _d = (_c = console).warn;
                        _e = ['[VesselDMS] bypass-login failed:', res.status];
                        return [4 /*yield*/, res.text()["catch"](function () { return ''; })];
                    case 5:
                        _d.apply(_c, _e.concat([_f.sent()]));
                        _f.label = 6;
                    case 6: return [3 /*break*/, 9];
                    case 7:
                        e_2 = _f.sent();
                        console.warn('[VesselDMS] Session init failed:', e_2);
                        return [3 /*break*/, 9];
                    case 8:
                        // ALWAYS mark session init as settled and re-render — regardless of whether
                        // a session_id was actually obtained. Previously this.render() was only
                        // called inside the `if (data.session_id)` branch, so when the backend had
                        // no database configured (session_id === null) this method returned
                        // silently: the React component's props.sessionId stayed '' forever, its
                        // _loadData() guard ("if (!this.props.sessionId) return;") never passed,
                        // and the vessel list stayed blank with no error shown.
                        this._sessionInitialized = true;
                        this.render();
                        return [7 /*endfinally*/];
                    case 9: return [2 /*return*/];
                }
            });
        });
    };
    VesselEmailWebPart.prototype._getEnvironmentMessage = function () {
        var _this = this;
        if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
            return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
                .then(function (context) {
                var environmentMessage = '';
                switch (context.app.host.name) {
                    case 'Office': // running in Office
                        environmentMessage = _this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
                        break;
                    case 'Outlook': // running in Outlook
                        environmentMessage = _this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
                        break;
                    case 'Teams': // running in Teams
                    case 'TeamsModern':
                        environmentMessage = _this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
                        break;
                    default:
                        environmentMessage = strings.UnknownEnvironment;
                }
                return environmentMessage;
            });
        }
        return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
    };
    VesselEmailWebPart.prototype.onThemeChanged = function (currentTheme) {
        if (!currentTheme) {
            return;
        }
        this._isDarkTheme = !!currentTheme.isInverted;
        var semanticColors = currentTheme.semanticColors;
        if (semanticColors) {
            this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
            this.domElement.style.setProperty('--link', semanticColors.link || null);
            this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
        }
    };
    VesselEmailWebPart.prototype.onDispose = function () {
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(VesselEmailWebPart.prototype, "dataVersion", {
        get: function () {
            return sp_core_library_1.Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    VesselEmailWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                (0, sp_property_pane_1.PropertyPaneTextField)('apiBaseUrl', {
                                    label: 'API Base URL'
                                }),
                                (0, sp_property_pane_1.PropertyPaneTextField)('apiToken', {
                                    label: 'API Token'
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return VesselEmailWebPart;
}(sp_webpart_base_1.BaseClientSideWebPart));
exports["default"] = VesselEmailWebPart;
