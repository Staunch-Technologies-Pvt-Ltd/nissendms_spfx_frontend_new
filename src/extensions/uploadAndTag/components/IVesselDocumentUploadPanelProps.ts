import { ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';

export interface IVesselDocumentUploadPanelProps {
  isOpen: boolean;
  onDismiss: () => void;
  context: ListViewCommandSetContext;
  apiBaseUrl?: string;
  onUploadSuccess?: () => void;
}

export interface IApprovalPanelProps {
  isOpen: boolean;
  onDismiss: () => void;
  context: ListViewCommandSetContext;
  itemId: number;
  itemTitle: string;
  listId: string;
  onComplete?: () => void;
}
