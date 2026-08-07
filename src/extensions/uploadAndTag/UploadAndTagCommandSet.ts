import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';

import VesselDocumentUploadPanel from './components/VesselDocumentUploadPanel';
import ApprovalPanel from './components/ApprovalPanel';
import type { IVesselDocumentUploadPanelProps, IApprovalPanelProps } from './components/IVesselDocumentUploadPanelProps';

export interface IUploadAndTagCommandSetProperties {
  apiBaseUrl?: string;
}

const LOG_SOURCE: string = 'UploadAndTagCommandSet';

export default class UploadAndTagCommandSet extends BaseListViewCommandSet<IUploadAndTagCommandSetProperties> {
  private _panelContainer: HTMLDivElement | null = null;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized UploadAndTagCommandSet');

    const uploadCommand: Command = this.tryGetCommand('UPLOAD_AND_TAG_DOC');
    if (uploadCommand) uploadCommand.visible = true;

    const approveCommand: Command = this.tryGetCommand('APPROVE_DOC');
    if (approveCommand) approveCommand.visible = false;

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);
    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'UPLOAD_AND_TAG_DOC':
      case 'COMMAND_1':
        this._showUploadPanel();
        break;
      case 'APPROVE_DOC':
        this._showApprovalPanel();
        break;
      default:
        throw new Error(`Unknown command ${event.itemId}`);
    }
  }

  private _getContainer(): HTMLDivElement {
    if (!this._panelContainer) {
      this._panelContainer = document.createElement('div');
      document.body.appendChild(this._panelContainer);
    }
    return this._panelContainer;
  }

  private _dismissPanel = (): void => {
    if (this._panelContainer) {
      ReactDom.unmountComponentAtNode(this._panelContainer);
    }
  };

  private _refreshListView = (): void => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const lv = this.context.listView as any;
    if (typeof lv.refresh === 'function') lv.refresh();
  };

  private _showUploadPanel(): void {
    const element: React.ReactElement<IVesselDocumentUploadPanelProps> = React.createElement(
      VesselDocumentUploadPanel,
      {
        isOpen: true,
        onDismiss: this._dismissPanel,
        context: this.context,
        apiBaseUrl: this.properties.apiBaseUrl || 'http://localhost:8000',
        onUploadSuccess: this._refreshListView,
      }
    );
    ReactDom.render(element, this._getContainer());
  }

  private _showApprovalPanel(): void {
    const rows = this.context.listView.selectedRows;
    if (!rows || rows.length === 0) return;

    const row = rows[0];
    const itemId = row.getValueByName('ID') as number;
    const itemTitle = (row.getValueByName('FileLeafRef') || row.getValueByName('Title') || 'Selected Document') as string;
    const listId = this.context.pageContext.list?.id?.toString() || '';

    const element: React.ReactElement<IApprovalPanelProps> = React.createElement(
      ApprovalPanel,
      {
        isOpen: true,
        onDismiss: this._dismissPanel,
        context: this.context,
        itemId,
        itemTitle,
        listId,
        onComplete: this._refreshListView,
      }
    );
    ReactDom.render(element, this._getContainer());
  }

  private _onListViewStateChanged = (_args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');

    const uploadCommand: Command = this.tryGetCommand('UPLOAD_AND_TAG_DOC');
    if (uploadCommand) uploadCommand.visible = true;

    const approveCommand: Command = this.tryGetCommand('APPROVE_DOC');
    if (approveCommand) {
      const selectedRows = this.context.listView.selectedRows;
      approveCommand.visible = !!(selectedRows && selectedRows.length === 1);
    }

    this.raiseOnChange();
  };
}
