import * as React from 'react';
import {
  Panel,
  PanelType,
  PrimaryButton,
  DefaultButton,
  TextField,
  Dropdown,
  IDropdownOption,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Stack
} from '@fluentui/react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import type { IVesselDocumentUploadPanelProps } from './IVesselDocumentUploadPanelProps';

export const DATA_SOURCE_OPTIONS: IDropdownOption[] = [
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

export const VesselDocumentUploadPanel: React.FC<IVesselDocumentUploadPanelProps> = (props) => {
  const { isOpen, onDismiss, context, onUploadSuccess } = props;

  const [selectedFile, setSelectedFile] = React.useState<File | null>(null);
  const [vesselName, setVesselName] = React.useState<string>('');
  const [dataSource, setDataSource] = React.useState<string>('contract');
  const [apiBaseUrl, setApiBaseUrl] = React.useState<string>(props.apiBaseUrl || 'http://localhost:8000');
  
  const [submitting, setSubmitting] = React.useState<boolean>(false);
  const [statusMessage, setStatusMessage] = React.useState<{ type: MessageBarType; text: string } | null>(null);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    if (e.target.files && e.target.files.length > 0) {
      setSelectedFile(e.target.files[0]);
    }
  };

  const handleSubmit = async (): Promise<void> => {
    if (!selectedFile) {
      setStatusMessage({ type: MessageBarType.error, text: 'Please select a document file to upload.' });
      return;
    }
    if (!vesselName.trim()) {
      setStatusMessage({ type: MessageBarType.error, text: 'Please enter a Vessel Name.' });
      return;
    }

    setSubmitting(true);
    setStatusMessage(null);

    try {
      const webUrl = context.pageContext.web.absoluteUrl;
      const listId = context.pageContext.list?.id?.toString() || '';

      // 1. Upload File to SharePoint Document Library
      const arrayBuffer = await selectedFile.arrayBuffer();
      const fileName = selectedFile.name;

      const uploadUrl = `${webUrl}/_api/web/lists(guid'${listId}')/RootFolder/Files/add(url='${encodeURIComponent(fileName)}',overwrite=true)`;
      
      const uploadResp: SPHttpClientResponse = await context.spHttpClient.post(
        uploadUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
          },
          body: arrayBuffer,
        }
      );

      if (!uploadResp.ok) {
        const errJson = await uploadResp.json().catch(() => ({}));
        throw new Error(errJson?.error?.message?.value || `SharePoint upload failed with status ${uploadResp.status}`);
      }

      const uploadData = await uploadResp.json();
      const itemUri = uploadData.d?.ListItemAllFields?.__deferred?.uri;

      // 2. Write VesselName, DataSource metadata + set approval status to SharePoint item
      if (itemUri) {
        const itemType = uploadData.d?.ListItemAllFields?.__metadata?.type || 'SP.Data.Shared_x0020_DocumentsItem';
        const updateResp: SPHttpClientResponse = await context.spHttpClient.post(
          itemUri,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=verbose',
              'Content-Type': 'application/json;odata=verbose',
              'X-HTTP-Method': 'MERGE',
              'IF-MATCH': '*',
            },
            body: JSON.stringify({
              '__metadata': { 'type': itemType },
              'VesselName': vesselName,
              'DataSource': dataSource,
              '_ModerationStatus': 2,       // 0 = Approved, 1 = Rejected, 2 = Pending
              '_ModerationComments': 'Pending approval',
            }),
          }
        );

        if (!updateResp.ok) {
          console.warn('SharePoint metadata update returned status', updateResp.status);
        }
      }

      // 3. Send file + metadata to FastAPI Backend
      const backendUrl = apiBaseUrl.replace(/\/$/, '');
      const formData = new FormData();
      formData.append('file', selectedFile);
      formData.append('vessel_name', vesselName);
      formData.append('data_source', dataSource);
      formData.append('sp_web_url', webUrl);

      let backendSuccess = false;
      try {
        const backendResp = await fetch(`${backendUrl}/api/vessels/upload-and-tag`, {
          method: 'POST',
          body: formData,
        });

        if (backendResp.ok) {
          backendSuccess = true;
        } else {
          console.warn('FastAPI upload-and-tag status', backendResp.status);
        }
      } catch (backendErr) {
        console.warn('FastAPI connection error', backendErr);
      }

      // 4. Show confirmation MessageBar
      if (backendSuccess) {
        setStatusMessage({
          type: MessageBarType.success,
          text: `Document "${fileName}" uploaded successfully and is pending approval.`,
        });
      } else {
        setStatusMessage({
          type: MessageBarType.warning,
          text: `Document "${fileName}" uploaded to SharePoint (pending approval). Note: FastAPI backend at ${backendUrl} was unreachable.`,
        });
      }

      // Reset form
      setSelectedFile(null);
      setVesselName('');
      if (onUploadSuccess) onUploadSuccess();
    } catch (err) {
      const errorObj = err as Error;
      setStatusMessage({
        type: MessageBarType.error,
        text: errorObj?.message || 'An error occurred during document upload.',
      });
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.medium}
      headerText="Vessel Document Upload & Tagging"
      closeButtonAriaLabel="Close"
    >
      <Stack tokens={{ childrenGap: 16 }} style={{ marginTop: 10 }}>
        {statusMessage && (
          <MessageBar messageBarType={statusMessage.type} isMultiline={true}>
            {statusMessage.text}
          </MessageBar>
        )}

        <div>
          <label style={{ fontWeight: 600, display: 'block', marginBottom: 6 }}>Select Document File *</label>
          <input
            type="file"
            onChange={handleFileChange}
            disabled={submitting}
            style={{ display: 'block', width: '100%', padding: '6px', border: '1px solid #c8c6c4', borderRadius: '4px' }}
          />
          {selectedFile && (
            <span style={{ fontSize: 12, color: '#605e5c', marginTop: 4, display: 'block' }}>
              Selected: {selectedFile.name} ({(selectedFile.size / 1024).toFixed(1)} KB)
            </span>
          )}
        </div>

        <TextField
          label="Vessel Name *"
          value={vesselName}
          onChange={(_, newVal) => setVesselName(newVal || '')}
          placeholder="e.g. M/V Alpine Hero"
          disabled={submitting}
          required
        />

        <Dropdown
          label="DataSource Tag (Category) *"
          selectedKey={dataSource}
          options={DATA_SOURCE_OPTIONS}
          onChange={(_, option) => option && setDataSource(option.key as string)}
          disabled={submitting}
          required
        />

        <TextField
          label="FastAPI Backend URL"
          value={apiBaseUrl}
          onChange={(_, newVal) => setApiBaseUrl(newVal || '')}
          placeholder="http://localhost:8000"
          disabled={submitting}
        />

        <Stack horizontal tokens={{ childrenGap: 10 }} style={{ marginTop: 20 }}>
          <PrimaryButton onClick={handleSubmit} disabled={submitting || !selectedFile || !vesselName}>
            {submitting ? <Spinner size={SpinnerSize.small} /> : 'Submit & Tag Document'}
          </PrimaryButton>
          <DefaultButton onClick={onDismiss} disabled={submitting}>
            Close
          </DefaultButton>
        </Stack>
      </Stack>
    </Panel>
  );
};

export default VesselDocumentUploadPanel;
