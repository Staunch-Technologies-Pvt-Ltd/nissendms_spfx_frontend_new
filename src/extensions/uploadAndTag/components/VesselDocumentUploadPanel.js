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
exports.VesselDocumentUploadPanel = exports.DATA_SOURCE_OPTIONS = void 0;
var React = require("react");
var react_1 = require("@fluentui/react");
var sp_http_1 = require("@microsoft/sp-http");
exports.DATA_SOURCE_OPTIONS = [
    { key: 'contract', text: 'contract — CP (Charter Party)' },
    { key: 'other_contract', text: 'other_contract — Other Contracts' },
    { key: 'vessels_certificate', text: 'vessels_certificate — Certificate' },
    { key: 'vessels_drawing', text: 'vessels_drawing — Drawing' },
    { key: 'mail', text: 'mail — Mail (default fallback)' },
    { key: 'imo', text: 'imo — IMO' },
    { key: 'uscg', text: 'uscg — USCG' },
    { key: 'msib', text: 'msib — USCG MSIB' },
    { key: 'imo_flag_country_others', text: 'imo_flag_country_others — IMO Flag Country Others' },
    { key: 'panama_flag_circular', text: 'panama_flag_circular — Panama Flag Circular' },
    { key: 'imo_flag_country_flag', text: 'imo_flag_country_flag — Flag' },
    { key: 'nk', text: 'nk — NK' },
    { key: 'japan_p_and_i', text: 'japan_p_and_i — Japan P&I' },
    { key: 'ukpandi', text: 'ukpandi — UK P&I' },
    { key: 'gard', text: 'gard — GARD' },
    { key: 'scmg', text: 'scmg — Standard Club' },
    { key: 'britannia_p_and_i', text: 'britannia_p_and_i — Britannia P&I' },
    { key: 'bimco', text: 'bimco — BIMCO' },
    { key: 'security_information', text: 'security_information — Security Information' },
    { key: 'omc_kaikoumu', text: 'omc_kaikoumu — OMC Marine & Tech. Support Center' },
    { key: 'ice_information', text: 'ice_information — Ice Information' },
    { key: 'right_ship', text: 'right_ship — RightShip' },
    { key: 'others', text: 'others — Others' },
];
var VesselDocumentUploadPanel = function (props) {
    var isOpen = props.isOpen, onDismiss = props.onDismiss, context = props.context, onUploadSuccess = props.onUploadSuccess;
    var _a = React.useState(null), selectedFile = _a[0], setSelectedFile = _a[1];
    var _b = React.useState(''), vesselName = _b[0], setVesselName = _b[1];
    var _c = React.useState('contract'), dataSource = _c[0], setDataSource = _c[1];
    var _d = React.useState(props.apiBaseUrl || 'http://localhost:8000'), apiBaseUrl = _d[0], setApiBaseUrl = _d[1];
    var _e = React.useState(false), submitting = _e[0], setSubmitting = _e[1];
    var _f = React.useState(null), statusMessage = _f[0], setStatusMessage = _f[1];
    var handleFileChange = function (e) {
        if (e.target.files && e.target.files.length > 0) {
            setSelectedFile(e.target.files[0]);
        }
    };
    var handleSubmit = function () { return __awaiter(void 0, void 0, void 0, function () {
        var webUrl, listId, arrayBuffer, fileName, uploadUrl, uploadResp, errJson, uploadData, itemUri, itemType, updateResp, backendUrl, formData, backendSuccess, backendResp, backendErr_1, err_1, errorObj;
        var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k;
        return __generator(this, function (_l) {
            switch (_l.label) {
                case 0:
                    if (!selectedFile) {
                        setStatusMessage({ type: react_1.MessageBarType.error, text: 'Please select a document file to upload.' });
                        return [2 /*return*/];
                    }
                    if (!vesselName.trim()) {
                        setStatusMessage({ type: react_1.MessageBarType.error, text: 'Please enter a Vessel Name.' });
                        return [2 /*return*/];
                    }
                    setSubmitting(true);
                    setStatusMessage(null);
                    _l.label = 1;
                case 1:
                    _l.trys.push([1, 13, 14, 15]);
                    webUrl = context.pageContext.web.absoluteUrl;
                    listId = ((_b = (_a = context.pageContext.list) === null || _a === void 0 ? void 0 : _a.id) === null || _b === void 0 ? void 0 : _b.toString()) || '';
                    return [4 /*yield*/, selectedFile.arrayBuffer()];
                case 2:
                    arrayBuffer = _l.sent();
                    fileName = selectedFile.name;
                    uploadUrl = "".concat(webUrl, "/_api/web/lists(guid'").concat(listId, "')/RootFolder/Files/add(url='").concat(encodeURIComponent(fileName), "',overwrite=true)");
                    return [4 /*yield*/, context.spHttpClient.post(uploadUrl, sp_http_1.SPHttpClient.configurations.v1, {
                            headers: {
                                'Accept': 'application/json;odata=verbose'
                            },
                            body: arrayBuffer
                        })];
                case 3:
                    uploadResp = _l.sent();
                    if (!!uploadResp.ok) return [3 /*break*/, 5];
                    return [4 /*yield*/, uploadResp.json()["catch"](function () { return ({}); })];
                case 4:
                    errJson = _l.sent();
                    throw new Error(((_d = (_c = errJson === null || errJson === void 0 ? void 0 : errJson.error) === null || _c === void 0 ? void 0 : _c.message) === null || _d === void 0 ? void 0 : _d.value) || "SharePoint upload failed with status ".concat(uploadResp.status));
                case 5: return [4 /*yield*/, uploadResp.json()];
                case 6:
                    uploadData = _l.sent();
                    itemUri = (_g = (_f = (_e = uploadData.d) === null || _e === void 0 ? void 0 : _e.ListItemAllFields) === null || _f === void 0 ? void 0 : _f.__deferred) === null || _g === void 0 ? void 0 : _g.uri;
                    if (!itemUri) return [3 /*break*/, 8];
                    itemType = ((_k = (_j = (_h = uploadData.d) === null || _h === void 0 ? void 0 : _h.ListItemAllFields) === null || _j === void 0 ? void 0 : _j.__metadata) === null || _k === void 0 ? void 0 : _k.type) || 'SP.Data.Shared_x0020_DocumentsItem';
                    return [4 /*yield*/, context.spHttpClient.post(itemUri, sp_http_1.SPHttpClient.configurations.v1, {
                            headers: {
                                'Accept': 'application/json;odata=verbose',
                                'Content-Type': 'application/json;odata=verbose',
                                'X-HTTP-Method': 'MERGE',
                                'IF-MATCH': '*'
                            },
                            body: JSON.stringify({
                                '__metadata': { 'type': itemType },
                                'VesselName': vesselName,
                                'DataSource': dataSource,
                                '_ModerationStatus': 2,
                                '_ModerationComments': 'Pending approval'
                            })
                        })];
                case 7:
                    updateResp = _l.sent();
                    if (!updateResp.ok) {
                        console.warn('SharePoint metadata update returned status', updateResp.status);
                    }
                    _l.label = 8;
                case 8:
                    backendUrl = apiBaseUrl.replace(/\/$/, '');
                    formData = new FormData();
                    formData.append('file', selectedFile);
                    formData.append('vessel_name', vesselName);
                    formData.append('data_source', dataSource);
                    formData.append('sp_web_url', webUrl);
                    backendSuccess = false;
                    _l.label = 9;
                case 9:
                    _l.trys.push([9, 11, , 12]);
                    return [4 /*yield*/, fetch("".concat(backendUrl, "/api/vessels/upload-and-tag"), {
                            method: 'POST',
                            body: formData
                        })];
                case 10:
                    backendResp = _l.sent();
                    if (backendResp.ok) {
                        backendSuccess = true;
                    }
                    else {
                        console.warn('FastAPI upload-and-tag status', backendResp.status);
                    }
                    return [3 /*break*/, 12];
                case 11:
                    backendErr_1 = _l.sent();
                    console.warn('FastAPI connection error', backendErr_1);
                    return [3 /*break*/, 12];
                case 12:
                    // 4. Show confirmation MessageBar
                    if (backendSuccess) {
                        setStatusMessage({
                            type: react_1.MessageBarType.success,
                            text: "Document \"".concat(fileName, "\" uploaded successfully and is pending approval.")
                        });
                    }
                    else {
                        setStatusMessage({
                            type: react_1.MessageBarType.warning,
                            text: "Document \"".concat(fileName, "\" uploaded to SharePoint (pending approval). Note: FastAPI backend at ").concat(backendUrl, " was unreachable.")
                        });
                    }
                    // Reset form
                    setSelectedFile(null);
                    setVesselName('');
                    if (onUploadSuccess)
                        onUploadSuccess();
                    return [3 /*break*/, 15];
                case 13:
                    err_1 = _l.sent();
                    errorObj = err_1;
                    setStatusMessage({
                        type: react_1.MessageBarType.error,
                        text: (errorObj === null || errorObj === void 0 ? void 0 : errorObj.message) || 'An error occurred during document upload.'
                    });
                    return [3 /*break*/, 15];
                case 14:
                    setSubmitting(false);
                    return [7 /*endfinally*/];
                case 15: return [2 /*return*/];
            }
        });
    }); };
    return (<react_1.Panel isOpen={isOpen} onDismiss={onDismiss} type={react_1.PanelType.medium} headerText="Vessel Document Upload & Tagging" closeButtonAriaLabel="Close">
      <react_1.Stack tokens={{ childrenGap: 16 }} style={{ marginTop: 10 }}>
        {statusMessage && (<react_1.MessageBar messageBarType={statusMessage.type} isMultiline={true}>
            {statusMessage.text}
          </react_1.MessageBar>)}

        <div>
          <label style={{ fontWeight: 600, display: 'block', marginBottom: 6 }}>Select Document File *</label>
          <input type="file" onChange={handleFileChange} disabled={submitting} style={{ display: 'block', width: '100%', padding: '6px', border: '1px solid #c8c6c4', borderRadius: '4px' }}/>
          {selectedFile && (<span style={{ fontSize: 12, color: '#605e5c', marginTop: 4, display: 'block' }}>
              Selected: {selectedFile.name} ({(selectedFile.size / 1024).toFixed(1)} KB)
            </span>)}
        </div>

        <react_1.TextField label="Vessel Name *" value={vesselName} onChange={function (_, newVal) { return setVesselName(newVal || ''); }} placeholder="e.g. M/V Alpine Hero" disabled={submitting} required/>

        <react_1.Dropdown label="DataSource Tag (Category) *" selectedKey={dataSource} options={exports.DATA_SOURCE_OPTIONS} onChange={function (_, option) { return option && setDataSource(option.key); }} disabled={submitting} required/>

        <react_1.TextField label="FastAPI Backend URL" value={apiBaseUrl} onChange={function (_, newVal) { return setApiBaseUrl(newVal || ''); }} placeholder="http://localhost:8000" disabled={submitting}/>

        <react_1.Stack horizontal tokens={{ childrenGap: 10 }} style={{ marginTop: 20 }}>
          <react_1.PrimaryButton onClick={handleSubmit} disabled={submitting || !selectedFile || !vesselName}>
            {submitting ? <react_1.Spinner size={react_1.SpinnerSize.small}/> : 'Submit & Tag Document'}
          </react_1.PrimaryButton>
          <react_1.DefaultButton onClick={onDismiss} disabled={submitting}>
            Close
          </react_1.DefaultButton>
        </react_1.Stack>
      </react_1.Stack>
    </react_1.Panel>);
};
exports.VesselDocumentUploadPanel = VesselDocumentUploadPanel;
exports["default"] = exports.VesselDocumentUploadPanel;
