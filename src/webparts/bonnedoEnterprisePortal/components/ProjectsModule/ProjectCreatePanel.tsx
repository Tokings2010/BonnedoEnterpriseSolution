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
import { SharePointService } from '../../services/SharePointService';
import { ProjectProvisioningService } from '../../services/ProjectProvisioningService';
import {
  getProjectDocumentDisplayPath,
  getProjectDocumentStorage,
  IProjectDocumentStorage,
  sanitizeProjectCodeForStorage,
  uploadProjectDocument,
} from '../../services/ProjectDocumentStorageService';
import { createCbsBudgetImportService, ICbsImportResult } from '../../services/CbsBudgetImportService';
import { createWbsScheduleImportService } from '../../services/WbsScheduleImportService';

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

  // ─── CBS Import State ──────────────────────────────────────────────────────

  const [cbsImportResult, setCbsImportResult] = React.useState<ICbsImportResult | null>(null);

  const cbsImportService = React.useMemo(
    () => createCbsBudgetImportService(spHttpClient, pageContext),
    [spHttpClient, pageContext]
  );

  const wbsImportService = React.useMemo(
    () => createWbsScheduleImportService(spHttpClient, pageContext),
    [spHttpClient, pageContext]
  );

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
    // Enforce CBS file for new projects (edit mode can skip if already uploaded)
    if (!editMode && !cbsFile) {
      setMessage({ type: MessageBarType.error, text: 'Budget (CBS) file is required. Please select a CBS Excel file before creating the project.' });
      return false;
    }

    return true;
  };

  const resolveProjectManagerId = async (): Promise<number | null> => {
    if (!projectManagerId && !projectManagerEmail) {
      return null;
    }

    // If projectManagerId is already a number, use it directly
    const numericId = parseInt(String(projectManagerId), 10);
    if (!isNaN(numericId)) {
      return numericId;
    }

    const isGraphGuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(String(projectManagerId));

    // Build login name candidates in priority order
    const loginNameCandidates: string[] = [];

    // 1. SharePoint claims format: i:0#.f|membership|{email}
    if (projectManagerEmail) {
      loginNameCandidates.push(`i:0#.f|membership|${projectManagerEmail}`);
      loginNameCandidates.push(projectManagerEmail);
    }

    // 2. LoginName from PeoplePicker (if not a Graph GUID)
    if (projectManagerLoginName && !isGraphGuid) {
      loginNameCandidates.push(projectManagerLoginName);
    }

    // 3. For Graph GUIDs, try resolving via the email first (already added above)
    //    Then try the classic "domain\username" format if email has an @
    if (projectManagerEmail && projectManagerEmail.includes('@')) {
      const domain = projectManagerEmail.split('@')[1]?.split('.')[0] || '';
      const username = projectManagerEmail.split('@')[0] || '';
      if (domain && username) {
        loginNameCandidates.push(`${domain}\\${username}`);
      }
    }

    // Filter out empty strings
    const candidates = loginNameCandidates.filter((v): v is string => !!v);

    if (candidates.length === 0) {
      return null;
    }

    const ensureUserUrl = `${pageContext.web.absoluteUrl}/_api/web/ensureuser`;
    let lastError = '';

    for (const loginName of candidates) {
      try {
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
          const resolvedId = ensureUserData.Data?.Id || ensureUserData.Id;
          if (resolvedId) {
            return resolvedId;
          }
        }

        // Collect error for last resort
        lastError = (await ensureUserResponse.text().catch(() => '')).trim();
      } catch (e) {
        lastError = e instanceof Error ? e.message : String(e);
      }
    }

    // If all SharePoint methods failed, try direct Graph API to get the user
    if (webPartContext?.msGraphClientFactory && projectManagerEmail) {
      try {
        const graphClient = await webPartContext.msGraphClientFactory.getClient('3');
        const graphUser = await graphClient
          .api(`/users/${projectManagerEmail}`)
          .select('id,userPrincipalName,mail')
          .get();

        if (graphUser?.userPrincipalName) {
          // Try again with the userPrincipalName from Graph
          const upnCandidates = [
            `i:0#.f|membership|${graphUser.userPrincipalName}`,
            graphUser.userPrincipalName,
          ];
          for (const upn of upnCandidates) {
            try {
              const upnResp = await spHttpClient.post(ensureUserUrl, SPHttpClient.configurations.v1, {
                headers: { 'Accept': 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' },
                body: JSON.stringify({ logonName: upn }),
              });
              if (upnResp.ok) {
                const upnData = await upnResp.json();
                const resolvedId = upnData.Data?.Id || upnData.Id;
                if (resolvedId) return resolvedId;
              }
            } catch { /* skip */ }
          }
        }
      } catch { /* Graph resolve failed */ }
    }

    throw new Error(
      `Unable to resolve Project Manager for "${projectManagerEmail || projectManagerLoginName}". ` +
      `The user may not have access to this SharePoint site. ` +
      `Please ensure the user is a member of this site (Site Permissions) and try again.`
    );
  };

  // ─── File Deletion Helper (rollback) ───────────────────────────────────────

  const deleteUploadedFile = async (folderServerRelativeUrl: string, fileName: string): Promise<void> => {
    try {
      const normalizedFolderUrl = folderServerRelativeUrl.startsWith('/') ? folderServerRelativeUrl : `/${folderServerRelativeUrl.replace(/^\/+/, '')}`;
      const encodedFolder = encodeURIComponent(normalizedFolderUrl);
      const encodedFile = encodeURIComponent(fileName);
      const fileUrl = `${pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodedFolder}')/Files('${encodedFile}')`;
      await spHttpClient.post(fileUrl, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=verbose',
          'X-HTTP-Method': 'DELETE',
          'IF-MATCH': '*',
        },
        body: '',
      });
    } catch (e) {
      console.warn('[Rollback] Failed to delete uploaded file:', fileName, e);
    }
  };

  // ─── Check if budget or schedule data already exists (for edit mode) ──────

  const checkBudgetDataExists = async (): Promise<boolean> => {
    try {
      const webUrl = pageContext.web.absoluteUrl;
      const escaped = projectCode.replace(/'/g, "''");
      const url = `${webUrl}/_api/web/lists/getbytitle('Project_Budget_Items')/items?$filter=Project_Code eq '${escaped}'&$top=1&$select=ID`;
      const resp = await spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!resp.ok) return false;
      const data = await resp.json();
      return (data.value || []).length > 0;
    } catch { return false; }
  };

  const checkScheduleDataExists = async (): Promise<boolean> => {
    try {
      const webUrl = pageContext.web.absoluteUrl;
      const escaped = projectCode.replace(/'/g, "''");
      const url = `${webUrl}/_api/web/lists/getbytitle('Project_Schedule')/items?$filter=Project_Code eq '${escaped}'&$top=1&$select=ID`;
      const resp = await spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!resp.ok) return false;
      const data = await resp.json();
      return (data.value || []).length > 0;
    } catch { return false; }
  };

  // ─── Local upload helper (uses pre-read buffer, avoids double arrayBuffer) ──

  const uploadBufferToSharePoint = async (
    buffer: ArrayBuffer,
    fileName: string,
    folderServerRelativeUrl: string
  ): Promise<void> => {
    const normalizedFolderUrl = folderServerRelativeUrl.startsWith('/')
      ? folderServerRelativeUrl
      : `/${folderServerRelativeUrl.replace(/^\/+/, '')}`;
    const encodedFolder = encodeURIComponent(normalizedFolderUrl);
    const encodedFile = encodeURIComponent(fileName);
    const uploadUrl = `${pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodedFolder}')/Files/add(url='${encodedFile}',overwrite=true)`;
    const resp = await spHttpClient.post(uploadUrl, SPHttpClient.configurations.v1, {
      headers: { 'Accept': 'application/json;odata=verbose', 'Content-Type': 'application/octet-stream' },
      body: buffer,
    });
    if (!resp.ok) {
      const errText = await resp.text().catch(() => '');
      throw new Error(`Upload failed for '${fileName}': ${resp.status} ${errText.substring(0, 200)}`);
    }
  };

  // ─── Submit ──────────────────────────────────────────────────────────────

  const handleSubmit = async (): Promise<void> => {
    if (!validateForm()) return;

    setIsSubmitting(true);
    setMessage(null);
    let documentStorage: IProjectDocumentStorage | null = null;
    let cbsParseMessage = '';
    let wbsParseMessage = '';
    let cbsUploaded = false;
    let wbsUploaded = false;

    try {
      // ── Step 1: Save/update project record ──
      const resolvedProjectManagerId = await resolveProjectManagerId();
      setUploadProgress(editMode ? 'Updating project...' : 'Creating project...');

      const projectBaseUrl = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ENT_Project_Master')/items`;

      const projectBody: Record<string, string | number> = {};

      // Only include non-empty fields — SharePoint REST API returns 500
      // when null is sent for Text, Currency, or User fields.
      projectBody.Title = projectCode;
      projectBody.Project_Code = projectCode;
      projectBody.Project_Name = projectName;

      if (clientName) projectBody.Client_Name = clientName;
      if (contractValue) projectBody.Contract_Value = parseFloat(contractValue);
      projectBody.Project_Status = projectStatus;
      if (resolvedProjectManagerId != null) projectBody.Project_ManagerId = resolvedProjectManagerId;
      if (startDate) projectBody.Start_Date = startDate.toISOString();
      if (endDate) projectBody.End_Date = endDate.toISOString();

      const sharePointService = new SharePointService(spHttpClient, pageContext);

      // Proactively check for duplicate project code — SharePoint's
      // SPDuplicateValuesFoundException returns a hard 500, so we
      // catch it earlier with a friendly message.
      if (!editMode) {
        try {
          const existing = await sharePointService.getListData(
            'ENT_Project_Master',
            `Project_Code eq '${projectCode.replace(/'/g, "''")}'`,
            1
          );
          if (existing && existing.length > 0) {
            setMessage({
              type: MessageBarType.error,
              text: `A project with code "${projectCode}" already exists. Each project must have a unique code. Please use a different Project Code.`,
            });
            setUploadProgress('');
            setIsSubmitting(false);
            return;
          }
        } catch {
          // Silently continue — best-effort check; SharePoint will reject
          // duplicates anyway.
        }
      }

      let projectId: number;
      if (editMode && editProject) {
        await sharePointService.updateListItem('ENT_Project_Master', editProject.ID, projectBody);
        projectId = editProject.ID;
      } else {
        const createdItem = await sharePointService.createListItem('ENT_Project_Master', projectBody);
        projectId = createdItem.ID;
      }

      // ── Step 3: Prepare document storage if files need uploading ──
      if (cbsFile || wbsFile) {
        setUploadProgress('Preparing project document storage...');
        documentStorage = await getProjectDocumentStorage(spHttpClient, pageContext, projectCode);
      }

      // ── Step 4: Process CBS file (upload + parse) ──
      if (cbsFile && documentStorage) {
        // Read buffer ONCE — reused for both upload and parsing (avoid double-read issue)
        let cbsBuffer: ArrayBuffer;
        try {
          cbsBuffer = await cbsFile.arrayBuffer();
        } catch {
          setMessage({ type: MessageBarType.error, text: 'Failed to read CBS file. The file may be corrupted.' });
          setUploadProgress('');
          setIsSubmitting(false);
          return;
        }

        // Check for duplicates in edit mode
        if (editMode) {
          const budgetExists = await checkBudgetDataExists();
          if (budgetExists) {
            // Duplicate check: warn user but still process as variation
            console.log('[ProjectCreate] Budget data exists — processing as variation.');
          }
        }

        setUploadProgress('Uploading budget (CBS) file...');
        try {
          // Upload to: Documents/Budgets/{projectCode}/{projectCode}_CBS.xlsx
          await uploadBufferToSharePoint(cbsBuffer, documentStorage.budgetFileName, documentStorage.budgetFolderUrl);
          cbsUploaded = true;
        } catch (err) {
          const errorDetail = err instanceof Error ? err.message : 'Unknown error';
          setMessage({
            type: MessageBarType.error,
            text: `Budget file upload failed: ${errorDetail}. Please ensure the Documents library and Budgets folder exist, then try again.`,
          });
          setUploadProgress('');
          setIsSubmitting(false);
          return;
        }

        // Parse immediately (using same buffer read before upload)
        setUploadProgress('Parsing CBS budget data...');
        try {
          const importResult = await cbsImportService.importCbsFile({
            fileName: documentStorage.budgetFileName,
            projectCode,
            fileBuffer: cbsBuffer,
          });
          setCbsImportResult(importResult);

          if (importResult.success) {
            cbsParseMessage = importResult.isVariation
              ? `Budget variation applied: ${importResult.rowsCreated} created, ${importResult.rowsUpdated} updated`
              : `Budget imported: ${importResult.rowsCreated} line items created`;
          } else {
            // Rollback: delete the uploaded file since parsing failed
            await deleteUploadedFile(documentStorage.budgetFolderUrl, documentStorage.budgetFileName);
            cbsUploaded = false;
            const userMsg = `Budget file was uploaded but could not be parsed: ${importResult.summary}. The file has been removed. Please fix the file and try again.`;
            setMessage({ type: MessageBarType.error, text: userMsg });
            setUploadProgress('');
            setIsSubmitting(false);
            return;
          }
        } catch (err) {
          // Parse threw an exception — rollback file
          await deleteUploadedFile(documentStorage.budgetFolderUrl, documentStorage.budgetFileName);
          cbsUploaded = false;
          const errorDetail = err instanceof Error ? err.message : 'Unknown error';
          setMessage({
            type: MessageBarType.error,
            text: `Budget parsing failed: ${errorDetail}. The file has been removed. Please fix the issue and try again.`,
          });
          setUploadProgress('');
          setIsSubmitting(false);
          return;
        }
      }

      // ── Step 5: Process WBS file (upload + parse) ──
      if (wbsFile && documentStorage) {
        // Read buffer ONCE — reused for both upload and parsing
        let wbsBuffer: ArrayBuffer;
        try {
          wbsBuffer = await wbsFile.arrayBuffer();
        } catch {
          setMessage({
            type: MessageBarType.warning,
            text: `${cbsParseMessage || 'Project created.'} Failed to read WBS file. The file may be corrupted. Go to Schedule tab to re-upload.`,
          });
          setUploadProgress('');
          setIsSubmitting(false);
          return;
        }

        // Check for duplicates in edit mode
        if (editMode) {
          const scheduleExists = await checkScheduleDataExists();
          if (scheduleExists) {
            console.log('[ProjectCreate] Schedule data exists — processing as update.');
          }
        }

        setUploadProgress('Uploading schedule (WBS) file...');
        try {
          // Upload to: Documents/Budgets/{projectCode}/{projectCode}_WBS.xlsx
          await uploadBufferToSharePoint(wbsBuffer, documentStorage.wbsFileName, documentStorage.wbsFolderUrl);
          wbsUploaded = true;
        } catch (err) {
          const errorDetail = err instanceof Error ? err.message : 'Unknown error';
          setMessage({
            type: MessageBarType.warning,
            text: `${cbsParseMessage || 'Budget imported.'} Schedule file upload failed: ${errorDetail}. Go to Schedule tab to re-upload.`,
          });
          setUploadProgress('');
          setIsSubmitting(false);
          return;
        }

        // Parse immediately (using same buffer)
        setUploadProgress('Parsing WBS schedule data...');
        try {
          const wbsResult = await wbsImportService.importWbsFile({
            fileName: documentStorage.wbsFileName,
            projectCode,
            fileBuffer: wbsBuffer,
          });

          if (wbsResult.success) {
            wbsParseMessage = wbsResult.isUpdate
              ? `Schedule update applied: ${wbsResult.rowsCreated} created, ${wbsResult.rowsUpdated} updated`
              : `Schedule imported: ${wbsResult.rowsCreated} activities created`;
          } else {
            // Rollback WBS file only — CBS already succeeded
            await deleteUploadedFile(documentStorage.wbsFolderUrl, documentStorage.wbsFileName);
            wbsUploaded = false;
            const wbsFailMsg = `Schedule file was uploaded but could not be parsed: ${wbsResult.summary}. The file has been removed. Go to Schedule tab to re-upload.`;
            setMessage({ type: MessageBarType.warning, text: `${cbsParseMessage || 'Project updated.'} ${wbsFailMsg}` });
            setUploadProgress('');
            setIsSubmitting(false);
            return;
          }
        } catch (err) {
          if (wbsUploaded) {
            await deleteUploadedFile(documentStorage.wbsFolderUrl, documentStorage.wbsFileName);
            wbsUploaded = false;
          }
          const errorDetail = err instanceof Error ? err.message : 'Unknown error';
          setMessage({
            type: MessageBarType.warning,
            text: `${cbsParseMessage || 'Budget imported.'} WBS parsing failed: ${errorDetail}. Go to Schedule tab to re-upload.`,
          });
          setUploadProgress('');
          setIsSubmitting(false);
          return;
        }
      }

      // ── Step 6: Success & provisioning ──
      setUploadProgress('');

      // Fire-and-forget provisioning after both CBS and WBS succeed
      if (webPartContext && cbsFile && wbsFile && cbsParseMessage && wbsParseMessage) {
        // Don't await — let the UI respond immediately
        const provisioningService = new ProjectProvisioningService(webPartContext);
        provisioningService.provisionProjectWorkspace({
          projectId,
          projectName,
          projectCode,
          projectManagerEmail: projectManagerEmail || undefined,
        }).catch((err: Error) => {
          console.warn('[ProjectCreate] Provisioning background task:', err.message);
        });
      }

      const successPrefix = editMode ? 'Project updated successfully!' : 'Project created successfully!';
      let successDetail = '';
      if (cbsParseMessage) successDetail = cbsParseMessage;
      if (wbsParseMessage) successDetail = successDetail ? `${successDetail} ${wbsParseMessage}` : wbsParseMessage;

      // If edit mode with no file changes
      if (!cbsFile && !wbsFile) {
        setMessage({ type: MessageBarType.success, text: successPrefix });
      } else {
        setMessage({
          type: MessageBarType.success,
          text: successDetail ? `${successPrefix} ${successDetail}` : successPrefix,
        });
      }

      setTimeout(() => { resetForm(); }, 4000);
    } catch (err) {
      console.error('Error in project submission:', err);
      // Rollback any uploaded files
      if (documentStorage) {
        if (cbsUploaded) await deleteUploadedFile(documentStorage.budgetFolderUrl, documentStorage.budgetFileName);
        if (wbsUploaded) await deleteUploadedFile(documentStorage.wbsFolderUrl, documentStorage.wbsFileName);
      }
      setMessage({ type: MessageBarType.error, text: err instanceof Error ? err.message : 'Failed to process project' });
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
          Upload the CBS (budget) and WBS (schedule) Excel files. Files are stored under Documents/Budgets/{'ProjectCode'}/, then renamed to {'ProjectCode'}_CBS.xlsx and {'ProjectCode'}_WBS.xlsx. The CBS file is parsed immediately — no Power Automate needed.
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
          Files must use the Bonnedo Budget &amp; Schedule Template (v2) with named Excel tables (CBS_Table, WBS_Table). The CBS file is parsed immediately upon upload — no Power Automate required.
        </MessageBar>
      </Stack>
    </Panel>
  );
};

export default ProjectCreatePanel;
