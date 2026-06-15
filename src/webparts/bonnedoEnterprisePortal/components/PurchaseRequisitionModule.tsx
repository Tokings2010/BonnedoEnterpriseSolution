import * as React from 'react';
import {
    Panel,
    PanelType,
    Stack,
    TextField,
    PrimaryButton,
    DefaultButton,
    MessageBar,
    MessageBarType,
    Spinner,
    SpinnerSize,
    mergeStyleSets,
    Dropdown,
    IDropdownOption,
    getTheme,
} from '@fluentui/react';
import PeoplePicker from './PeoplePicker';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointService, IListItem } from '../services/SharePointService';
import { NotificationService } from '../services/NotificationService';
import { MSGraphClientV3 } from '@microsoft/sp-http';

export interface IPurchaseRequisitionFormProps {
    isOpen: boolean;
    onDismiss: () => void;
    onSubmitSuccess?: () => void;
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    webPartContext?: WebPartContext;
}

interface IFormData {
    Title: string;
    Project_CodeId: number | undefined;
    Material_RequestId: number | undefined;
    Requested_ById: number | undefined;
    Project_ManagerId: number | undefined;
    Description: string;
    Quantity: string;
    Estimated_Cost: number | undefined;
}

interface IFormState {
    formData: IFormData;
    isSubmitting: boolean;
    error: string | undefined;
    successMessage: string | undefined;
    projects: IListItem[];
    materialRequests: IListItem[];
    isLoadingLookups: boolean;
    currentUserId: number | undefined;
}

const PurchaseRequisitionForm: React.FC<IPurchaseRequisitionFormProps> = ({
    isOpen,
    onDismiss,
    onSubmitSuccess,
    spHttpClient,
    pageContext,
    webPartContext,
}) => {
    const theme = getTheme();
    const [state, setState] = React.useState<IFormState>({
            formData: {
            Title: '',
            Project_CodeId: undefined,
            Material_RequestId: undefined,
            Requested_ById: undefined,
            Project_ManagerId: undefined,
            Description: '',
            Quantity: '',
            Estimated_Cost: undefined,
        },
        isSubmitting: false,
        error: undefined,
        successMessage: undefined,
        projects: [],
        materialRequests: [],
        isLoadingLookups: true,
        currentUserId: undefined,
    });

    const sharePointService = React.useMemo(
        () => new SharePointService(spHttpClient, pageContext),
        [spHttpClient, pageContext]
    );

    // Load lookup data on mount (projects, approved non-converted MRs for duplicate prevention, current user)
    React.useEffect(() => {
        const loadLookups = async (): Promise<void> => {
            try {
                const [projects, allMaterialRequests, currentUser, allPRs] = await Promise.all([
                    sharePointService.getProjects(),
                    sharePointService.getListData('PRC_Material_Request_Register', undefined, 100, 0, 'Project_Code,Material'),
                    sharePointService.getCurrentUser(),
                    sharePointService.getListData('PRC_Purchase_Requisition_Register', undefined, 100),
                ]);

                const convertedMrIds = new Set<number>();
                allPRs.forEach((pr: any) => {
                    const mrRef = pr.Material_RequestId || (pr.Material_Request && (pr.Material_Request.ID || pr.Material_Request.Id || pr.Material_Request.id));
                    const mrId = typeof mrRef === 'number' ? mrRef : (typeof mrRef === 'string' ? parseInt(mrRef, 10) : undefined);
                    if (mrId) convertedMrIds.add(mrId);
                });

                const availableMRs = allMaterialRequests.filter((mr: any) => {
                    const mrId = mr.Id || mr.ID;
                    return mrId != null && !convertedMrIds.has(mrId);
                });

                setState((prev) => ({
                    ...prev,
                    projects,
                    materialRequests: availableMRs,
                    currentUserId: currentUser?.ID,
                    formData: {
                        ...prev.formData,
                        Requested_ById: currentUser?.ID,
                    },
                    isLoadingLookups: false,
                }));
            } catch (error) {
                console.error('Error loading lookups for PR form:', error);
                setState((prev) => ({
                    ...prev,
                    isLoadingLookups: false,
                    error: 'Failed to load lookup data',
                }));
            }
        };

        if (isOpen) {
            loadLookups();
        }
    }, [isOpen, sharePointService]);

    // Auto-generate PR Number using ID sequencing (replicates main PR creation)
    React.useEffect(() => {
        if (isOpen && !state.formData.Title) {
            const generatePRCode = async (): Promise<void> => {
                try {
                    const allItems = await sharePointService.getListData('PRC_Purchase_Requisition_Register', undefined, 500, 0, undefined);
                    let nextNumber = 1;
                    if (allItems && allItems.length > 0) {
                        allItems.forEach((item: any) => {
                            const title = item.Title || item.PR_Number || '';
                            const matches = title.match(/\d+$/);
                            if (matches) {
                                const num = parseInt(matches[0], 10);
                                if (num >= nextNumber) {
                                    nextNumber = num + 1;
                                }
                            }
                        });
                    }
                    const newCode = `PR-${nextNumber.toString().padStart(4, '0')}`;
                    setState((prev) => ({
                        ...prev,
                        formData: { ...prev.formData, Title: newCode },
                    }));
                } catch (error) {
                    console.error('Error generating PR code:', error);
                    const timestamp = Date.now().toString().slice(-6);
                    setState((prev) => ({
                        ...prev,
                        formData: { ...prev.formData, Title: `PR-${timestamp}` },
                    }));
                }
            };
            generatePRCode();
        }
    }, [isOpen, sharePointService]);

    const handleInputChange = (field: keyof IFormData, value: any): void => {
        setState((prev) => ({
            ...prev,
            formData: {
                ...prev.formData,
                [field]: value,
            },
        }));
    };

    const getLookupId = (val: any): number | undefined => {
        if (typeof val === 'number' || typeof val === 'string') return Number(val) || undefined;
        if (val && typeof val === 'object') return val.ID || val.Id || val.id;
        return undefined;
    };

    // Auto-populate from selected Material Request (enhanced for Convert-to-PR flow + duplicate safe)
    const handleMaterialRequestChange = (materialRequestId: number | undefined): void => {
        if (!materialRequestId) {
            handleInputChange('Material_RequestId', undefined);
            return;
        }

        const selectedMR = state.materialRequests.find((mr: any) => (mr.Id || mr.ID) === materialRequestId);
        if (!selectedMR) {
            handleInputChange('Material_RequestId', materialRequestId);
            return;
        }

        handleInputChange('Material_RequestId', materialRequestId);
        handleInputChange('Quantity', selectedMR.Quantity ? String(selectedMR.Quantity) : state.formData.Quantity);

        const projId = getLookupId(selectedMR.Project_CodeId) || getLookupId(selectedMR.Project_Code) || state.formData.Project_CodeId;
        handleInputChange('Project_CodeId', projId);

        // Replicate project dropdown logic: auto-populate Project Manager from selected project
        if (projId) {
            const selectedProject = state.projects.find((p: any) => (p.Id || p.ID) === projId);
            if (selectedProject && selectedProject.Project_ManagerId) {
                handleInputChange('Project_ManagerId', selectedProject.Project_ManagerId);
            }
        }

        if (selectedMR.MaterialId || selectedMR.Description || selectedMR.Material) {
            handleInputChange('Description', selectedMR.Material?.Title || selectedMR.Description || selectedMR.Title || `Material Request ${selectedMR.Title || selectedMR.Id || selectedMR.ID}`);
        }

        const estCost = selectedMR.Estimated_Cost !== undefined
            ? selectedMR.Estimated_Cost
            : (selectedMR.Quantity ? Number(selectedMR.Quantity) * 100 : undefined);
        handleInputChange('Estimated_Cost', estCost);

        // Pre-populate Requested By from MR if present (Requested_By lookup)
        const reqBy = getLookupId(selectedMR.Requested_ById) || getLookupId(selectedMR.Requested_By);
        if (reqBy) handleInputChange('Requested_ById', reqBy);
    };

    const classNames = mergeStyleSets({
        formContainer: { padding: '8px 0' },
        fieldGroup: { marginBottom: '12px' },
        sectionHeader: { fontWeight: 600, margin: '16px 0 8px' },
        buttonGroup: { display: 'flex', gap: '8px', marginTop: '24px' },
    });

    // Convert lookup items to dropdown options
    const projectOptions: IDropdownOption[] = state.projects.map((item) => ({
        key: item.Id || item.ID,
        text: item.Project_Code || `Project ${item.Id || item.ID}`,
    }));

    const materialRequestOptions: IDropdownOption[] = state.materialRequests.map((item: any) => ({
        key: item.Id || item.ID,
        text: item.Title || item.MR_Number || item.Request_No || `MR ${item.Id || item.ID}`,
    }));

    const isDisabled = state.isSubmitting || state.isLoadingLookups;

    const validateForm = (): boolean => {
        if (!state.formData.Title || !state.formData.Title.trim()) {
            setState((prev) => ({ ...prev, error: 'PR Number is required' }));
            return false;
        }
        if (!state.formData.Project_CodeId) {
            setState((prev) => ({ ...prev, error: 'Project Code is required' }));
            return false;
        }
        if (!state.formData.Quantity || !state.formData.Quantity.trim()) {
            setState((prev) => ({ ...prev, error: 'Quantity is required' }));
            return false;
        }
        return true;
    };

    const handleSubmit = React.useCallback(async (): Promise<void> => {
        if (!validateForm()) {
            return;
        }

        setState((prev) => ({ ...prev, isSubmitting: true, error: undefined }));

        try {
            // Duplicate prevention check for MR->PR conversion
            if (state.formData.Material_RequestId) {
                const existingPRs = await sharePointService.getListData(
                    'PRC_Purchase_Requisition_Register',
                    `Material_RequestId eq ${state.formData.Material_RequestId}`,
                    1
                );
                if (existingPRs && existingPRs.length > 0) {
                    setState((prev) => ({
                        ...prev,
                        isSubmitting: false,
                        error: 'This Material Request has already been converted to a Purchase Requisition',
                    }));
                    return;
                }
            }

            const prNumber = state.formData.Title || `PR-${Date.now().toString().slice(-6)}`;

            const itemData: any = {
                Title: prNumber,
                PR_Number: prNumber,
                Project_CodeId: state.formData.Project_CodeId,
                Material_RequestId: state.formData.Material_RequestId,
                Description: state.formData.Description || '',
                Quantity: parseInt(state.formData.Quantity, 10) || 0,
                Estimated_Cost: state.formData.Estimated_Cost,
                Status: 'Draft',
            };

            // If converting from MR, mark source MR as Completed (replicates Convert button)
            if (state.formData.Material_RequestId) {
                try {
                    await sharePointService.updateListItem('PRC_Material_Request_Register', state.formData.Material_RequestId, {
                        Approval_Status: 'Approved',
                        Status: 'Completed',
                    });
                } catch (markErr) {
                    console.warn('Could not mark source MR as completed:', markErr);
                }
            }

            const createdItem = await sharePointService.createListItem('PRC_Purchase_Requisition_Register', itemData);

            // Initialize the standard Approval Workflow (updates Current_Approver, Approval_Status etc.)
            let currentApprover: string | undefined;
            if (createdItem && createdItem.ID) {
                await sharePointService.initializePurchaseRequisitionApproval(createdItem.ID);

                // Fetch the updated item to populate Current_Approver for notification
                const updatedItems = await sharePointService.getListData(
                    'PRC_Purchase_Requisition_Register',
                    `ID eq ${createdItem.ID}`,
                    1
                );
                if (updatedItems.length > 0) {
                    const approverField = updatedItems[0].Current_Approver;
                    currentApprover = typeof approverField === 'object'
                        ? (approverField.Email || approverField.Title || approverField.LoginName)
                        : approverField;
                }
            }

            // Send notification (reuse existing service; approval workflow is the key requirement)
            if (webPartContext && currentApprover) {
                try {
                    const graphClient: MSGraphClientV3 = await webPartContext.msGraphClientFactory.getClient('3');
                    const notificationService = new NotificationService(graphClient);

                    const projectName = state.projects.find((p: any) => (p.Id || p.ID) === state.formData.Project_CodeId)?.Project_Code || 'Unknown Project';
                    const deepLink = `${window.location.origin}${window.location.pathname}?pr=${createdItem?.ID || ''}`;

                    console.log('[PurchaseRequisitionForm] Sending approval notification to:', currentApprover);

                    // Use existing MR notifier (text will indicate Material but delivers the alert; dedicated PR notifier can be added later)
                    await notificationService.sendMaterialRequestApprovalNotification({
                        requestTitle: prNumber,
                        project: projectName,
                        amount: state.formData.Estimated_Cost,
                        approverEmail: currentApprover,
                        deepLink,
                    });

                    console.log('[PurchaseRequisitionForm] Notification sent successfully');
                } catch (notifyError) {
                    console.error('[PurchaseRequisitionForm] Notification failed (workflow still initialized):', notifyError);
                }
            }

            setState((prev) => ({
                ...prev,
                isSubmitting: false,
                successMessage: 'Purchase Requisition created successfully!',
                formData: {
                    Title: '',
                    Project_CodeId: undefined,
                    Material_RequestId: undefined,
                    Requested_ById: prev.currentUserId,
                    Project_ManagerId: undefined,
                    Description: '',
                    Quantity: '',
                    Estimated_Cost: undefined,
                },
            }));

            setTimeout(() => {
                onDismiss();
                if (onSubmitSuccess) {
                    onSubmitSuccess();
                }
            }, 1500);
        } catch (error) {
            const errorMessage = error instanceof Error ? error.message : 'Failed to create Purchase Requisition';
            setState((prev) => ({
                ...prev,
                isSubmitting: false,
                error: errorMessage,
            }));
        }
    }, [state.formData, sharePointService, onDismiss, onSubmitSuccess, webPartContext, state.projects]);

    const handleCancel = (): void => {
        setState({
            formData: {
                Title: '',
                Project_CodeId: undefined,
                Material_RequestId: undefined,
                Requested_ById: state.currentUserId,
                Project_ManagerId: undefined,
                Description: '',
                Quantity: '',
                Estimated_Cost: undefined,
            },
            isSubmitting: false,
            error: undefined,
            successMessage: undefined,
            projects: state.projects,
            materialRequests: state.materialRequests,
            isLoadingLookups: false,
            currentUserId: state.currentUserId,
        });
        onDismiss();
    };

    return (
        <Panel
            isOpen={isOpen}
            onDismiss={handleCancel}
            type={PanelType.medium}
            headerText="New Purchase Requisition"
            closeButtonAriaLabel="Close"
            isLightDismiss={!state.isSubmitting}
        >
            <div style={{ padding: '0 20px', height: '100%', overflowY: 'auto' }}>
                <div className={classNames.formContainer}>
                    {state.error && (
                        <div style={{ marginBottom: '16px' }}>
                            <MessageBar messageBarType={MessageBarType.error} isMultiline>
                                {state.error}
                            </MessageBar>
                        </div>
                    )}

                    {state.successMessage && (
                        <div style={{ marginBottom: '16px' }}>
                            <MessageBar messageBarType={MessageBarType.success} isMultiline>
                                {state.successMessage}
                            </MessageBar>
                        </div>
                    )}

                    {state.isLoadingLookups && (
                        <div style={{ textAlign: 'center', padding: '20px' }}>
                            <Spinner size={SpinnerSize.medium} label="Loading data..." />
                        </div>
                    )}

                    {!state.isLoadingLookups && (
                        <Stack tokens={{ childrenGap: 16 }}>
                            {/* Section 1: Request Details */}
                            <div className={classNames.sectionHeader}>Request Details</div>

                            <div className={classNames.fieldGroup}>
                                <TextField
                                    label="PR Number"
                                    value={state.formData.Title}
                                    onChange={(_, value) => handleInputChange('Title', value)}
                                    disabled={isDisabled}
                                    placeholder="Auto-generated"
                                />
                            </div>

                            <div className={classNames.fieldGroup}>
                                <Dropdown
                                    label="Project Code"
                                    options={projectOptions}
                                    selectedKey={state.formData.Project_CodeId}
                                    onChange={(e, option) => {
                                        handleInputChange('Project_CodeId', option?.key as number);
                                        // Auto-populate Project Manager from selected project
                                         const selectedProject = state.projects.find((p: any) => (p.Id || p.ID) === (option?.key as any));
                                        if (selectedProject) {
                                            handleInputChange('Project_ManagerId', selectedProject.Project_ManagerId);
                                        }
                                    }}
                                    disabled={isDisabled}
                                    required
                                    placeholder="Select Project"
                                />
                            </div>

                            <div className={classNames.fieldGroup}>
                                <Dropdown
                                    label="Material Request Reference"
                                    options={materialRequestOptions}
                                    selectedKey={state.formData.Material_RequestId}
                                    onChange={(e, option) => handleMaterialRequestChange(option?.key as number)}
                                    disabled={isDisabled}
                                    placeholder="Select Material Request (Optional)"
                                />
                            </div>

                            <div className={classNames.fieldGroup}>
                                <TextField
                                    label="Description"
                                    value={state.formData.Description}
                                    onChange={(_, value) => handleInputChange('Description', value)}
                                    disabled={isDisabled}
                                    multiline
                                    rows={3}
                                />
                            </div>

                            <div className={classNames.fieldGroup}>
                                <TextField
                                    label="Quantity"
                                    type="number"
                                    value={state.formData.Quantity}
                                    onChange={(_, value) => handleInputChange('Quantity', value)}
                                    disabled={isDisabled}
                                    required
                                />
                            </div>

                            {/* Section 2: Financial Details */}
                            <div className={classNames.sectionHeader}>Financial Details</div>

                            <div className={classNames.fieldGroup}>
                                <TextField
                                    label="Estimated Cost"
                                    type="number"
                                    value={state.formData.Estimated_Cost?.toString() || ''}
                                    onChange={(_, value) => handleInputChange('Estimated_Cost', value ? Number(value) : undefined)}
                                    disabled={isDisabled}
                                    prefix="$"
                                />
                            </div>

                            {/* Section 3: Personnel */}
                            <div className={classNames.sectionHeader}>Personnel</div>

                            <div className={classNames.fieldGroup}>
                                <PeoplePicker
                                    key="pr-requested-by-picker"
                                    titleText="Requested By"
                                    selectedUsers={state.formData.Requested_ById ? [{ id: typeof state.formData.Requested_ById === 'string' ? parseInt(state.formData.Requested_ById, 10) : state.formData.Requested_ById, Title: '', Email: '', LoginName: '' }] : []}
                                    personSelectionLimit={1}
                                     required={false}
                                    disabled={isDisabled}
                                    spHttpClient={spHttpClient}
                                    pageContext={pageContext}
                                    webPartContext={webPartContext}
                                    onChange={(items) => {
                                        if (items && items.length > 0 && items[0].id) {
                                            handleInputChange('Requested_ById', items[0].id);
                                        } else {
                                            handleInputChange('Requested_ById', undefined);
                                        }
                                    }}
                                />
                            </div>

                            <div className={classNames.fieldGroup}>
                                <PeoplePicker
                                    key="pr-project-manager-picker"
                                    titleText="Project Manager"
                                    selectedUsers={state.formData.Project_ManagerId ? [{ id: typeof state.formData.Project_ManagerId === 'string' ? parseInt(state.formData.Project_ManagerId, 10) : state.formData.Project_ManagerId, Title: '', Email: '', LoginName: '' }] : []}
                                    personSelectionLimit={1}
                                    disabled={isDisabled}
                                    spHttpClient={spHttpClient}
                                    pageContext={pageContext}
                                    webPartContext={webPartContext}
                                    onChange={(items) => {
                                        if (items && items.length > 0 && items[0].id) {
                                            handleInputChange('Project_ManagerId', items[0].id);
                                        } else {
                                            handleInputChange('Project_ManagerId', undefined);
                                        }
                                    }}
                                />
                            </div>

                            {/* Buttons */}
                            <div className={classNames.buttonGroup}>
                                <PrimaryButton
                                    text="Submit"
                                    onClick={handleSubmit}
                                    disabled={isDisabled}
                                    iconProps={{ iconName: 'CheckMark' }}
                                />
                                <DefaultButton
                                    text="Cancel"
                                    onClick={handleCancel}
                                    disabled={isDisabled}
                                    iconProps={{ iconName: 'Cancel' }}
                                />
                            </div>

                            {state.isSubmitting && (
                                <div style={{ textAlign: 'center', marginTop: '16px' }}>
                                    <Spinner size={SpinnerSize.medium} label="Submitting..." />
                                </div>
                            )}
                        </Stack>
                    )}
                </div>
            </div>
        </Panel>
    );
};

export default PurchaseRequisitionForm;
