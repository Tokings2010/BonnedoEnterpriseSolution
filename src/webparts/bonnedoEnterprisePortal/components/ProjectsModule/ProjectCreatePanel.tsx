import * as React from 'react';
import {
  Panel,
  PanelType,
  PrimaryButton,
  DefaultButton,
  TextField,
  DatePicker,
  Dropdown,
  IDropdownOption,
  Stack,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Label,
  Separator,
  Text,
  Icon,
} from '@fluentui/react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import PeoplePicker from '../PeoplePicker';
import {
  getProjectDocumentDisplayPath,
  getProjectDocumentStorage,
  IProjectDocumentStorage,
  sanitizeProjectCodeForStorage,
  uploadProjectDocument,
} from '../../services/ProjectDocumentStorageService';

// ─── Interfaces ──────────────────────────────────────────────────────────────

export interface IProjectCreatePanelProps {
  isOpen: boolean;
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  webPartContext?: WebPartContext;
  onDismiss: () => void;
  editMode?: boolean;
  editProject?: {
    ID: number;
    Project_Code: string;
    Project_Name: string;
    Client_Name: string;
    Project_ManagerId?: string | number;
    Project_Manager?: string | { Title?: string; Email?: string; LoginName?: string; ID?: number };
    Contract_Value: number;
    Start_Date: string;
    End_Date: string;
    Project_Status?: string;
  };
  onProjectCreated?: (projectDetails: {
    projectId: number;
    projectName: string;
    projectCode: string;
    projectManagerId?: string | number;
    projectManagerEmail?: string;
  }) => void | Promise<void>;
  onProjectUpdated?: () => void | Promise<void>;
}

interface IProjectCreateProjectDetails {
  projectId: number;
  projectName: string;
  projectCode: string;
  projectManagerId?: string | number;
  projectManagerEmail?: string;
}

// ─── Component ───────────────────────────────────────────────────────────────

const ProjectCreatePanel: React.FC<IProjectCreatePanelProps> = ({
  isOpen,
  spHttpClient,
  pageContext,
  webPartContext,
  onDismiss,
  editMode = false,
  editProject,
  onProjectCreated,
  onProjectUpdated,
}) => {
  // Form state
  const [projectCode, setProjectCode] = React.useState('');
  const [projectName, setProjectName] = React.useState('');
  const [clientName, setClientName] = React.useState('');
  const [contractValue, setContractValue] = React.useState('');
  const [startDate, setStartDate] = React.useState<Date | undefined>(undefined);
  const [endDate, setEndDate] = React.useState<Date | undefined>(undefined);
  const [projectStatus, setProjectStatus] = React.useState('Active');
  const [projectManagerId, setProjectManagerId] = React.useState<string | number | undefined>(undefined);
  const [projectManagerEmail, setProjectManagerEmail] = React.useState('');
  const [projectManagerLoginName, setProjectManagerLoginName] = React.useState('');

  const normalizeCodePart = (value: string): string => value.trim().toUpperCase().replace(/[^A-Z0-9]/g, '');

  const buildProjectCode = (name: string, client: string): string => {
    const projectPart = normalizeCodePart(name).substring(0, 3);
    const clientFirstName = client.trim().split(/\s+/)[0] || '';
    const clientPart = normalizeCodePart(clientFirstName).substring(0, 4);

    if (!projectPart || !clientPart) {
      return '';
    }

    return `PRJ-${projectPart}-${clientPart}-${new Date().getFullYear()}`;
  };

  const updateGeneratedProjectCode = (name: string, client: string): void => {
    if (editMode) {
      return;
    }

    const generatedCode = buildProjectCode(name, client);
    if (generatedCode) {
      setProjectCode(generatedCode);
    }
  };

  // File upload state
  const [cbsFile, setCbsFile] = React.useState<File | null>(null);
  const [wbsFile, setWbsFile] = React.useState<File | null>(null);

  // UI state
  const [isSubmitting, setIsSubmitting] = React.useState(false);
  const [message, setMessage] = React.useState<{ type: MessageBarType; text: string } | null>(null);
  const [uploadProgress, setUploadProgress] = React.useState('');

  // Refs for file inputs
  const cbsInputRef = React.useRef<HTMLInputElement>(null);
  const wbsInputRef = React.useRef<HTMLInputElement>(null);

  React.useEffect(() => {
    if (!isOpen || !editMode || !editProject) {
      return;
    }

    const manager = typeof editProject.Project_Manager === 'object' ? editProject.Project_Manager : undefined;
    setProjectCode(editProject.Project_Code || '');
    setProjectName(editProject.Project_Name || '');
    setClientName(editProject.Client_Name || '');
    setContractValue(editProject.Contract_Value?.toString() || '');
    setStartDate(editProject.Start_Date ? new Date(editProject.Start_Date) : undefined);
    setEndDate(editProject.End_Date ? new Date(editProject.End_Date) : undefined);
    setProjectStatus(editProject.Project_Status || 'Active');
    setProjectManagerId(editProject.Project_ManagerId ?? manager?.ID);
    setProjectManagerEmail(manager?.Email || '');
    setProjectManagerLoginName(manager?.LoginName || '');
  }, [isOpen, editMode, editProject]);

  // ─── Reset Form ──────────────────────────────────────────────────────────

  const resetForm = (): void => {
    setProjectCode('');
    setProjectName('');
    setClientName('');
    setContractValue('');
    setStartDate(undefined);
    setEndDate(undefined);
    setProjectStatus('Active');
    setProjectManagerId(undefined);
    setProjectManagerEmail('');
    setProjectManagerLoginName('');
    setCbsFile(null);
    setWbsFile(null);
    setMessage(null);
    setUploadProgress('');
  };

  // ─── File Handlers ───────────────────────────────────────────────────────

  const handleCbsFileChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    if (e.target.files && e.target.files.length > 0) {
      setCbsFile(e.target.files[0]);
    }
  };

  const handleWbsFileChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    if (e.target.files && e.target.files.length > 0) {
      setWbsFile(e.target.files[0]);
    }
  };

  const handleProjectManagerChange = (users: Array<{ id?: string | number; Email?: string; LoginName?: string }>): void => {
    const selectedUser = users && users.length > 0 ? users[0] : undefined;
    setProjectManagerId(selectedUser?.id);
    setProjectManagerEmail(selectedUser?.Email || '');
    setProjectManagerLoginName(selectedUser?.LoginName || '');
  };

  const validateForm = (): boolean => {
    if (!projectCode.trim()) {
      setMessage({ type: MessageBarType.error, text: 'Project Code is required' });
      return false;
    }
    if (!projectName.trim()) {
      setMessage({ type: MessageBarType.error, text: 'Project Name is required' });
      return false;
    }
    if (!clientName.trim()) {
      setMessage({ type: MessageBarType.error, text: 'Client Name is required' });
      return false;
    }
    if (!projectManagerId && !projectManagerEmail && !projectManagerLoginName) {
      setMessage({ type: MessageBarType.error, text: 'Project Manager is required' });
      return false;
    }
    if (!contractValue.trim()) {
      setMessage({ type: MessageBarType.error, text: 'Contract Value is required' });
      return false;
    }
    if (!startDate) {
      setMessage({ type: MessageBarType.error, text: 'Start Date is required' });
      return false;
    }
    if (!endDate) {
      setMessage({ type: MessageBarType.error, text: 'End Date is required' });
      return false;
    }
    if (endDate <= startDate) {
      setMessage({ type: MessageBarType.error, text: 'End Date must be after Start Date' });
      return false;
    }

    return true;
  };

  const resolveProjectManagerId = async (): Promise<number | null> => {
    if (!projectManagerId) {
      return null;
    }

    const numericId = parseInt(String(projectManagerId), 10);
    if (!isNaN(numericId)) {
      return numericId;
    }

    const isGraphGuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(String(projectManagerId));
    const loginNameCandidates = [
      projectManagerEmail,
      isGraphGuid ? '' : projectManagerLoginName,
    ].filter((value): value is string => !!value);

    if (loginNameCandidates.length === 0) {
      return null;
    }

    const ensureUserUrl = `${pageContext.web.absoluteUrl}/_api/web/ensureuser`;
    let lastError = '';

    for (const loginName of loginNameCandidates) {
      const ensureUserResponse: SPHttpClientResponse = await spHttpClient.post(
        ensureUserUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
          },
          body: JSON.stringify({ logonName: loginName }),
        }
      );

      if (ensureUserResponse.ok) {
        const ensureUserData = await ensureUserResponse.json();
        return ensureUserData.Data?.Id || ensureUserData.Id || null;
      }

      lastError = (await ensureUserResponse.text().catch(() => '')).trim();
    }

    throw new Error(`Unable to resolve Project Manager: ${lastError || 'No valid SharePoint user identifier was provided'}`);
  };

  // ─── Submit ──────────────────────────────────────────────────────────────

  const handleSubmit = async (): Promise<void> => {
    if (!validateForm()) {
      return;
    }

    setIsSubmitting(true);
    setMessage(null);

    try {
      const resolvedProjectManagerId = await resolveProjectManagerId();
      setUploadProgress(editMode ? 'Updating project...' : 'Creating project...');
      const projectUrl = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ENT_Project_Master')/items`;
      const projectBody: Record<string, string | number | null> = {
        Title: projectCode,
        Project_Code: projectCode,
        Project_Name: projectName,
        Client_Name: clientName || null,
        Contract_Value: contractValue ? parseFloat(contractValue) : null,
        Project_Status: projectStatus,
        Project_ManagerId: resolvedProjectManagerId ?? null,
      };

      if (startDate) projectBody.Start_Date = startDate.toISOString();
      if (endDate) projectBody.End_Date = endDate.toISOString();

      const itemUrl = editMode && editProject ? `${projectUrl}(${editProject.ID})` : projectUrl;
      const projectResponse: SPHttpClientResponse = await spHttpClient.post(
        itemUrl,
        SPHttpClient.configurations.v1,
        {
          headers: editMode
            ? {
              'Accept': 'application/json',
              'Content-Type': 'application/json',
              'X-HTTP-Method': 'MERGE',
              'If-Match': '*',
            }
            : {
              'Accept': 'application/json;odata=nometadata',
              'Content-Type': 'application/json;odata=nometadata',
            },
          body: JSON.stringify(projectBody),
        }
      );

      if (!projectResponse.ok) {
        const errData = await projectResponse.text().catch(() => '');
        throw new Error(`${editMode ? 'Failed to update project' : 'Failed to create project'}: ${projectResponse.status}${errData ? ' - ' + errData : ''}`);
      }

      if (editMode) {
        setUploadProgress('');
        setMessage({ type: MessageBarType.success, text: 'Project updated successfully!' });
        if (onProjectUpdated) {
          const projectUpdateResult = onProjectUpdated();
          if (projectUpdateResult) {
            (projectUpdateResult as Promise<void>).catch((error) => {
              console.error('Error notifying project update:', error);
            });
          }
        }
        setTimeout(() => {
          resetForm();
        }, 2000);
        return;
      }

      const projectData = await projectResponse.json();
      const projectId = projectData.ID || projectData.Id;
      const projectDetails: IProjectCreateProjectDetails = {
        projectId,
        projectName,
        projectCode,
        projectManagerId,
        projectManagerEmail: projectManagerEmail || projectManagerLoginName,
      };

      if (onProjectCreated) {
        const projectCreationResult = onProjectCreated(projectDetails);
        if (projectCreationResult) {
          (projectCreationResult as Promise<void>).catch((error) => {
            console.error('Error notifying project creation:', error);
          });
        }
      }

      let documentStorage: IProjectDocumentStorage | null = null;

      if (cbsFile || wbsFile) {
        try {
          setUploadProgress('Preparing project document storage...');
          documentStorage = await getProjectDocumentStorage(spHttpClient, pageContext, projectCode);
        } catch (err) {
          console.error('Error preparing project document storage:', err);
          setMessage({
            type: MessageBarType.warning,
            text: `${editMode ? 'Project updated' : 'Project created and workspace provisioning started'}. Document storage setup failed. Please upload the CBS/WBS files manually to the project document library.`,
          });
          setUploadProgress('');
          return;
        }
      }

      if (cbsFile && documentStorage) {
        setUploadProgress('Uploading budget (CBS) file...');
        try {
          await uploadProjectDocument(
            spHttpClient,
            pageContext,
            cbsFile,
            documentStorage.budgetFileName,
            documentStorage.budgetFolderUrl
          );
        } catch (err) {
          console.error('Error uploading CBS file:', err);
          setMessage({
            type: MessageBarType.warning,
            text: `${editMode ? 'Project updated' : 'Project created and workspace provisioning started'}. CBS file upload failed. Please upload manually to ${getProjectDocumentDisplayPath('Budgets', sanitizeProjectCodeForStorage(projectCode), documentStorage.budgetFileName)}.`,
          });
          setUploadProgress('');
          return;
        }
      }

      if (wbsFile && documentStorage) {
        setUploadProgress('Uploading schedule (WBS) file...');
        try {
          await uploadProjectDocument(
            spHttpClient,
            pageContext,
            wbsFile,
            documentStorage.wbsFileName,
            documentStorage.wbsFolderUrl
          );
        } catch (err) {
          console.error('Error uploading WBS file:', err);
          setMessage({
            type: MessageBarType.warning,
            text: `${editMode ? 'Project updated' : 'Project created'}. CBS file uploaded. WBS file upload failed. Please upload manually to ${getProjectDocumentDisplayPath('WBS', sanitizeProjectCodeForStorage(projectCode), documentStorage.wbsFileName)}.`,
          });
          setUploadProgress('');
          return;
        }
      }

      setUploadProgress('');
      const successPrefix = editMode ? 'Project updated successfully!' : 'Project created successfully!';
      setMessage({
        type: MessageBarType.success,
        text: cbsFile && wbsFile
          ? `${successPrefix} Budget and schedule files are stored in the project document library and will be processed automatically.`
          : cbsFile
            ? `${successPrefix} Budget file is stored in the project document library and will be processed automatically.`
            : wbsFile
              ? `${successPrefix} Schedule file is stored in the project document library and will be processed automatically.`
              : successPrefix
      });
      setTimeout(() => {
        resetForm();
      }, 2000);
    } catch (err) {
      console.error('Error creating project:', err);
      setMessage({ type: MessageBarType.error, text: err instanceof Error ? err.message : 'Failed to create project' });
      setUploadProgress('');
    } finally {
      setIsSubmitting(false);
    }
  };

  // ─── Status Options ──────────────────────────────────────────────────────

  const statusOptions: IDropdownOption[] = [
    { key: 'Active', text: 'Active' },
    { key: 'Planning', text: 'Planning' },
    { key: 'On Hold', text: 'On Hold' },
  ];

  const handleDismiss = (): void => {
    resetForm();
    onDismiss();
  };

  // ─── Render ──────────────────────────────────────────────────────────────

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={handleDismiss}
      type={PanelType.medium}
      headerText={editMode ? 'Edit Project' : 'Create New Project'}
      closeButtonAriaLabel="Close"
      isLightDismiss={!isSubmitting}
      onRenderFooterContent={() => (
        <Stack horizontal tokens={{ childrenGap: 8 }} style={{ padding: '16px 0' }}>
          <PrimaryButton
            text={isSubmitting ? (editMode ? 'Updating...' : 'Creating...') : editMode ? 'Update Project' : 'Create Project'}
            onClick={handleSubmit}
            disabled={isSubmitting}
            iconProps={{ iconName: 'Add' }}
          />
          <DefaultButton text="Cancel" onClick={handleDismiss} disabled={isSubmitting} />
        </Stack>
      )}
      isFooterAtBottom={true}
    >
      <Stack tokens={{ childrenGap: 16 }} style={{ padding: '16px 0' }}>
        {message && (
          <MessageBar messageBarType={message.type} onDismiss={() => setMessage(null)}>
            {message.text}
          </MessageBar>
        )}

        {isSubmitting && uploadProgress && (
          <Stack tokens={{ childrenGap: 8 }}>
            <Spinner size={SpinnerSize.small} label={uploadProgress} />
          </Stack>
        )}

        {/* Project Details */}
        <TextField
          label="Project Code"
          required
          placeholder="e.g. PRJ-ISU-HIRS-2026"
          description="Auto-generated from Project Name and Client Name"
          value={projectCode}
          onChange={(_, v) => setProjectCode(v || '')}
          disabled={isSubmitting || (!editMode && projectCode.trim().length > 0)}
        />

        <TextField
          label="Project Name"
          required
          placeholder="e.g. ISU Flowline Construction"
          value={projectName}
          onChange={(_, v) => {
            const nextProjectName = v || '';
            setProjectName(nextProjectName);
            updateGeneratedProjectCode(nextProjectName, clientName);
          }}
          disabled={isSubmitting}
        />

        <TextField
          label="Client Name"
          placeholder="e.g. Heirs Energies"
          value={clientName}
          onChange={(_, v) => {
            const nextClientName = v || '';
            setClientName(nextClientName);
            updateGeneratedProjectCode(projectName, nextClientName);
          }}
          disabled={isSubmitting}
        />

        <PeoplePicker
          titleText="Project Manager"
          selectedUsers={projectManagerId ? [{ id: projectManagerId, Title: '', Email: '', LoginName: '' }] : []}
          onChange={handleProjectManagerChange}
          personSelectionLimit={1}
          required={true}
          disabled={isSubmitting}
          spHttpClient={spHttpClient}
          pageContext={pageContext}
          webPartContext={webPartContext}
        />

        <TextField
          label="Contract Value (NGN)"
          placeholder="e.g. 500000000"
          value={contractValue}
          onChange={(_, v) => setContractValue(v || '')}
          type="number"
          disabled={isSubmitting}
        />

        <Stack horizontal tokens={{ childrenGap: 16 }}>
          <DatePicker
            label="Start Date"
            value={startDate}
            onSelectDate={(date) => setStartDate(date || undefined)}
            disabled={isSubmitting}
            style={{ flex: 1 }}
          />
          <DatePicker
            label="End Date"
            value={endDate}
            onSelectDate={(date) => setEndDate(date || undefined)}
            disabled={isSubmitting}
            style={{ flex: 1 }}
          />
        </Stack>

        <Dropdown
          label="Project Status"
          selectedKey={projectStatus}
          options={statusOptions}
          onChange={(_, option) => setProjectStatus(option?.key as string || 'Active')}
          disabled={isSubmitting}
        />

        <Separator />

        {/* File Upload Section */}
        <Text variant="mediumPlus" styles={{ root: { fontWeight: 600, color: '#1E2532' } }}>
          Budget & Schedule Upload
        </Text>
        <Text variant="small" styles={{ root: { color: '#6B7280' } }}>
          Upload the CBS (budget) and WBS (schedule) Excel files. Files are stored under Documents/Budgets/{'ProjectCode'}/, then renamed to {'ProjectCode'}_CBS.xlsx and {'ProjectCode'}_WBS.xlsx. Power Automate will parse them into the project tracking lists.
        </Text>

        {/* CBS Upload */}
        <div style={{ backgroundColor: '#F5F6FA', borderRadius: '8px', padding: '16px' }}>
          <Stack tokens={{ childrenGap: 8 }}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
              <Icon iconName="ExcelDocument" styles={{ root: { fontSize: 20, color: '#10B981' } }} />
              <Label>Budget (CBS) File</Label>
            </Stack>
            <input
              ref={cbsInputRef}
              type="file"
              accept=".xlsx,.xls"
              onChange={handleCbsFileChange}
              style={{ display: 'none' }}
            />
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
              <DefaultButton
                text={cbsFile ? 'Change File' : 'Select CBS File'}
                iconProps={{ iconName: 'Upload' }}
                onClick={() => cbsInputRef.current?.click()}
                disabled={isSubmitting}
              />
              {cbsFile && (
                <Text variant="small" styles={{ root: { color: '#10B981', fontWeight: 500 } }}>
                  {cbsFile.name} ({(cbsFile.size / 1024).toFixed(1)} KB)
                </Text>
              )}
            </Stack>
          </Stack>
        </div>

        {/* WBS Upload */}
        <div style={{ backgroundColor: '#F5F6FA', borderRadius: '8px', padding: '16px' }}>
          <Stack tokens={{ childrenGap: 8 }}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
              <Icon iconName="ExcelDocument" styles={{ root: { fontSize: 20, color: '#2563EB' } }} />
              <Label>Schedule (WBS) File</Label>
            </Stack>
            <input
              ref={wbsInputRef}
              type="file"
              accept=".xlsx,.xls"
              onChange={handleWbsFileChange}
              style={{ display: 'none' }}
            />
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
              <DefaultButton
                text={wbsFile ? 'Change File' : 'Select WBS File'}
                iconProps={{ iconName: 'Upload' }}
                onClick={() => wbsInputRef.current?.click()}
                disabled={isSubmitting}
              />
              {wbsFile && (
                <Text variant="small" styles={{ root: { color: '#2563EB', fontWeight: 500 } }}>
                  {wbsFile.name} ({(wbsFile.size / 1024).toFixed(1)} KB)
                </Text>
              )}
            </Stack>
          </Stack>
        </div>

        <MessageBar messageBarType={MessageBarType.info} styles={{ root: { borderRadius: '4px' } }}>
          Files must use the Bonnedo Budget & Schedule Template (v2) with named Excel tables (CBS_Table, WBS_Table).
        </MessageBar>
      </Stack>
    </Panel>
  );
};

export default ProjectCreatePanel;
