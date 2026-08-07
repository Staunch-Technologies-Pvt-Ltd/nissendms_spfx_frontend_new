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
exports.__esModule = true;
var React = require("react");
var ReactDom = require("react-dom");
var sp_core_library_1 = require("@microsoft/sp-core-library");
var sp_listview_extensibility_1 = require("@microsoft/sp-listview-extensibility");
var VesselDocumentUploadPanel_1 = require("./components/VesselDocumentUploadPanel");
var ApprovalPanel_1 = require("./components/ApprovalPanel");
var LOG_SOURCE = 'UploadAndTagCommandSet';
var UploadAndTagCommandSet = /** @class */ (function (_super) {
    __extends(UploadAndTagCommandSet, _super);
    function UploadAndTagCommandSet() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._panelContainer = null;
        _this._dismissPanel = function () {
            if (_this._panelContainer) {
                ReactDom.unmountComponentAtNode(_this._panelContainer);
            }
        };
        _this._refreshListView = function () {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            var lv = _this.context.listView;
            if (typeof lv.refresh === 'function')
                lv.refresh();
        };
        _this._onListViewStateChanged = function (_args) {
            sp_core_library_1.Log.info(LOG_SOURCE, 'List view state changed');
            var uploadCommand = _this.tryGetCommand('UPLOAD_AND_TAG_DOC');
            if (uploadCommand)
                uploadCommand.visible = true;
            var approveCommand = _this.tryGetCommand('APPROVE_DOC');
            if (approveCommand) {
                var selectedRows = _this.context.listView.selectedRows;
                approveCommand.visible = !!(selectedRows && selectedRows.length === 1);
            }
            _this.raiseOnChange();
        };
        return _this;
    }
    UploadAndTagCommandSet.prototype.onInit = function () {
        sp_core_library_1.Log.info(LOG_SOURCE, 'Initialized UploadAndTagCommandSet');
        var uploadCommand = this.tryGetCommand('UPLOAD_AND_TAG_DOC');
        if (uploadCommand)
            uploadCommand.visible = true;
        var approveCommand = this.tryGetCommand('APPROVE_DOC');
        if (approveCommand)
            approveCommand.visible = false;
        this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);
        return Promise.resolve();
    };
    UploadAndTagCommandSet.prototype.onExecute = function (event) {
        switch (event.itemId) {
            case 'UPLOAD_AND_TAG_DOC':
            case 'COMMAND_1':
                this._showUploadPanel();
                break;
            case 'APPROVE_DOC':
                this._showApprovalPanel();
                break;
            default:
                throw new Error("Unknown command ".concat(event.itemId));
        }
    };
    UploadAndTagCommandSet.prototype._getContainer = function () {
        if (!this._panelContainer) {
            this._panelContainer = document.createElement('div');
            document.body.appendChild(this._panelContainer);
        }
        return this._panelContainer;
    };
    UploadAndTagCommandSet.prototype._showUploadPanel = function () {
        var element = React.createElement(VesselDocumentUploadPanel_1["default"], {
            isOpen: true,
            onDismiss: this._dismissPanel,
            context: this.context,
            apiBaseUrl: this.properties.apiBaseUrl || 'http://localhost:8000',
            onUploadSuccess: this._refreshListView
        });
        ReactDom.render(element, this._getContainer());
    };
    UploadAndTagCommandSet.prototype._showApprovalPanel = function () {
        var _a, _b;
        var rows = this.context.listView.selectedRows;
        if (!rows || rows.length === 0)
            return;
        var row = rows[0];
        var itemId = row.getValueByName('ID');
        var itemTitle = (row.getValueByName('FileLeafRef') || row.getValueByName('Title') || 'Selected Document');
        var listId = ((_b = (_a = this.context.pageContext.list) === null || _a === void 0 ? void 0 : _a.id) === null || _b === void 0 ? void 0 : _b.toString()) || '';
        var element = React.createElement(ApprovalPanel_1["default"], {
            isOpen: true,
            onDismiss: this._dismissPanel,
            context: this.context,
            itemId: itemId,
            itemTitle: itemTitle,
            listId: listId,
            onComplete: this._refreshListView
        });
        ReactDom.render(element, this._getContainer());
    };
    return UploadAndTagCommandSet;
}(sp_listview_extensibility_1.BaseListViewCommandSet));
exports["default"] = UploadAndTagCommandSet;
