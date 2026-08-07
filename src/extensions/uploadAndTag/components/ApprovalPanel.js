"use strict";
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
var react_1 = require("@fluentui/react");
var sp_http_1 = require("@microsoft/sp-http");
var ApprovalPanel = function (_a) {
    var isOpen = _a.isOpen, onDismiss = _a.onDismiss, context = _a.context, itemId = _a.itemId, itemTitle = _a.itemTitle, listId = _a.listId, onComplete = _a.onComplete;
    var _b = React.useState(''), comments = _b[0], setComments = _b[1];
    var _c = React.useState(false), submitting = _c[0], setSubmitting = _c[1];
    var _d = React.useState(null), statusMessage = _d[0], setStatusMessage = _d[1];
    var handleDecision = function (approve) { return __awaiter(void 0, void 0, void 0, function () {
        var webUrl, metaResp, metaJson, itemType, updateResp, errJson, err_1;
        var _a, _b, _c, _d;
        return __generator(this, function (_e) {
            switch (_e.label) {
                case 0:
                    setSubmitting(true);
                    setStatusMessage(null);
                    _e.label = 1;
                case 1:
                    _e.trys.push([1, 7, 8, 9]);
                    webUrl = context.pageContext.web.absoluteUrl;
                    return [4 /*yield*/, context.spHttpClient.get("".concat(webUrl, "/_api/web/lists(guid'").concat(listId, "')/items(").concat(itemId, ")?$select=ID"), sp_http_1.SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=verbose' } })];
                case 2:
                    metaResp = _e.sent();
                    return [4 /*yield*/, metaResp.json()];
                case 3:
                    metaJson = _e.sent();
                    itemType = ((_b = (_a = metaJson === null || metaJson === void 0 ? void 0 : metaJson.d) === null || _a === void 0 ? void 0 : _a.__metadata) === null || _b === void 0 ? void 0 : _b.type) || 'SP.ListItem';
                    return [4 /*yield*/, context.spHttpClient.post("".concat(webUrl, "/_api/web/lists(guid'").concat(listId, "')/items(").concat(itemId, ")"), sp_http_1.SPHttpClient.configurations.v1, {
                            headers: {
                                'Accept': 'application/json;odata=verbose',
                                'Content-Type': 'application/json;odata=verbose',
                                'X-HTTP-Method': 'MERGE',
                                'IF-MATCH': '*'
                            },
                            body: JSON.stringify({
                                '__metadata': { type: itemType },
                                '_ModerationStatus': approve ? 0 : 1,
                                '_ModerationComments': comments || (approve ? 'Approved' : 'Rejected')
                            })
                        })];
                case 4:
                    updateResp = _e.sent();
                    if (!!updateResp.ok) return [3 /*break*/, 6];
                    return [4 /*yield*/, updateResp.json()["catch"](function () { return ({}); })];
                case 5:
                    errJson = _e.sent();
                    throw new Error(((_d = (_c = errJson === null || errJson === void 0 ? void 0 : errJson.error) === null || _c === void 0 ? void 0 : _c.message) === null || _d === void 0 ? void 0 : _d.value) || "Update failed with status ".concat(updateResp.status));
                case 6:
                    setStatusMessage({
                        type: approve ? react_1.MessageBarType.success : react_1.MessageBarType.warning,
                        text: "Document \"".concat(itemTitle, "\" has been ").concat(approve ? 'approved' : 'rejected', ".")
                    });
                    setComments('');
                    if (onComplete)
                        onComplete();
                    return [3 /*break*/, 9];
                case 7:
                    err_1 = _e.sent();
                    setStatusMessage({ type: react_1.MessageBarType.error, text: err_1.message || 'Approval action failed.' });
                    return [3 /*break*/, 9];
                case 8:
                    setSubmitting(false);
                    return [7 /*endfinally*/];
                case 9: return [2 /*return*/];
            }
        });
    }); };
    return (<react_1.Panel isOpen={isOpen} onDismiss={onDismiss} type={react_1.PanelType.medium} headerText="Document Approval" closeButtonAriaLabel="Close">
      <react_1.Stack tokens={{ childrenGap: 16 }} style={{ marginTop: 10 }}>
        {statusMessage && (<react_1.MessageBar messageBarType={statusMessage.type} isMultiline>
            {statusMessage.text}
          </react_1.MessageBar>)}

        <div>
          <react_1.Label>Document</react_1.Label>
          <span style={{ fontSize: 14 }}>{itemTitle}</span>
        </div>

        <react_1.TextField label="Comments" multiline rows={4} value={comments} onChange={function (_, v) { return setComments(v || ''); }} placeholder="Optional approval/rejection comments..." disabled={submitting}/>

        <react_1.Stack horizontal tokens={{ childrenGap: 10 }} style={{ marginTop: 20 }}>
          <react_1.PrimaryButton text={submitting ? '' : 'Approve'} onClick={function () { return handleDecision(true); }} disabled={submitting} styles={{ root: { backgroundColor: '#107c10', borderColor: '#107c10' } }}>
            {submitting && <react_1.Spinner size={react_1.SpinnerSize.small}/>}
          </react_1.PrimaryButton>
          <react_1.DefaultButton text={submitting ? '' : 'Reject'} onClick={function () { return handleDecision(false); }} disabled={submitting} styles={{ root: { color: '#a4262c', borderColor: '#a4262c' } }}>
            {submitting && <react_1.Spinner size={react_1.SpinnerSize.small}/>}
          </react_1.DefaultButton>
          <react_1.DefaultButton onClick={onDismiss} disabled={submitting}>Cancel</react_1.DefaultButton>
        </react_1.Stack>
      </react_1.Stack>
    </react_1.Panel>);
};
exports["default"] = ApprovalPanel;
