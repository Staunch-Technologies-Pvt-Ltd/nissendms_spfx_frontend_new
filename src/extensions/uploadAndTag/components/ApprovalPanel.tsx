import * as React from 'react';
import {
  Panel,
  PanelType,
  PrimaryButton,
  DefaultButton,
  TextField,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Stack,
  Label,
} from '@fluentui/react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import type { IApprovalPanelProps } from './IVesselDocumentUploadPanelProps';

const ApprovalPanel: React.FC<IApprovalPanelProps> = ({ isOpen, onDismiss, context, itemId, itemTitle, listId, onComplete }) => {
  const [comments, setComments] = React.useState<string>('');
  const [submitting, setSubmitting] = React.useState<boolean>(false);
  const [statusMessage, setStatusMessage] = React.useState<{ type: MessageBarType; text: string } | null>(null);

  const handleDecision = async (approve: boolean): Promise<void> => {
    setSubmitting(true);
    setStatusMessage(null);
    try {
      const webUrl = context.pageContext.web.absoluteUrl;

      // Get item entity type first
      const metaResp: SPHttpClientResponse = await context.spHttpClient.get(
        `${webUrl}/_api/web/lists(guid'${listId}')/items(${itemId})?$select=ID`,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=verbose' } }
      );
      const metaJson = await metaResp.json();
      const itemType = metaJson?.d?.__metadata?.type || 'SP.ListItem';

      const updateResp: SPHttpClientResponse = await context.spHttpClient.post(
        `${webUrl}/_api/web/lists(guid'${listId}')/items(${itemId})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': '*',
          },
          body: JSON.stringify({
            '__metadata': { type: itemType },
            '_ModerationStatus': approve ? 0 : 1,
            '_ModerationComments': comments || (approve ? 'Approved' : 'Rejected'),
          }),
        }
      );

      if (!updateResp.ok) {
        const errJson = await updateResp.json().catch(() => ({}));
        throw new Error(errJson?.error?.message?.value || `Update failed with status ${updateResp.status}`);
      }

      setStatusMessage({
        type: approve ? MessageBarType.success : MessageBarType.warning,
        text: `Document "${itemTitle}" has been ${approve ? 'approved' : 'rejected'}.`,
      });
      setComments('');
      if (onComplete) onComplete();
    } catch (err) {
      setStatusMessage({ type: MessageBarType.error, text: (err as Error).message || 'Approval action failed.' });
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.medium}
      headerText="Document Approval"
      closeButtonAriaLabel="Close"
    >
      <Stack tokens={{ childrenGap: 16 }} style={{ marginTop: 10 }}>
        {statusMessage && (
          <MessageBar messageBarType={statusMessage.type} isMultiline>
            {statusMessage.text}
          </MessageBar>
        )}

        <div>
          <Label>Document</Label>
          <span style={{ fontSize: 14 }}>{itemTitle}</span>
        </div>

        <TextField
          label="Comments"
          multiline
          rows={4}
          value={comments}
          onChange={(_, v) => setComments(v || '')}
          placeholder="Optional approval/rejection comments..."
          disabled={submitting}
        />

        <Stack horizontal tokens={{ childrenGap: 10 }} style={{ marginTop: 20 }}>
          <PrimaryButton
            text={submitting ? '' : 'Approve'}
            onClick={() => handleDecision(true)}
            disabled={submitting}
            styles={{ root: { backgroundColor: '#107c10', borderColor: '#107c10' } }}
          >
            {submitting && <Spinner size={SpinnerSize.small} />}
          </PrimaryButton>
          <DefaultButton
            text={submitting ? '' : 'Reject'}
            onClick={() => handleDecision(false)}
            disabled={submitting}
            styles={{ root: { color: '#a4262c', borderColor: '#a4262c' } }}
          >
            {submitting && <Spinner size={SpinnerSize.small} />}
          </DefaultButton>
          <DefaultButton onClick={onDismiss} disabled={submitting}>Cancel</DefaultButton>
        </Stack>
      </Stack>
    </Panel>
  );
};

export default ApprovalPanel;
