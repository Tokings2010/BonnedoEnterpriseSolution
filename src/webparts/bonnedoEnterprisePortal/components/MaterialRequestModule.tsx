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

export interface IMaterialRequestFormProps {
    isOpen: boolean;
    onDismiss: () => void;
    onSubmitSuccess?: () => void;
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    webPartContext?: WebPartContext;
}

interface IFormData {
    Project_CodeId: number | undefined;
    MaterialId: number | undefined;
    Quantity: string;
    UOM: string;
    Description: string;
    Required_Date: string;
    Estimated_Cost: number | undefined;
    Requested_ById: number | undefined;
    Project_ManagerId: number | undefined;
}

interface IFormState {
    formData: IFormData;
    isSubmitting: boolean;
    error: string | undefined;
    successMessage: string | undefined;
    projects: IListItem[];
    materials: IListItem[];
    isLoadingLookups: boolean;
    currentUserId: number | undefined;
}

const UOM_OPTIONS: IDropdownOption[] = [
    { key: 'PCS', text: 'Pieces' },
    { key: 'KG', text: 'Kilograms' },
    { key: 'L', text: 'Liters' },
    { key: 'M', text: 'Meters' },
    { key: 'BOX', text: 'Box' },
    { key: 'PACK', text: 'Pack' },
    { key: 'ROLL', text: 'Roll' },
    { key: 'SHEET', text: 'Sheet' },
];

const MaterialRequestForm: React.FC<IMaterialRequestFormProps> = ({
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
            Project_CodeId: undefined,
            MaterialId: undefined,
            Quantity: '',
            UOM: 'PCS',
            Description: '',
            Required_Date: '',
            Estimated_Cost: undefined,
            Requested_ById: undefined,
            Project_ManagerId: undefined,
        },
        isSubmitting: false,
        error: undefined,
        successMessage: undefined,
        projects: [],
        materials: [],
        isLoadingLookups: true,
        currentUserId: undefined,
    });

    const sharePointService = React.useMemo(
        () => new SharePointService(spHttpClient, pageContext),
        [spHttpClient, pageContext]
    );

    // Load lookup data on mount
    React.useEffect(() => {
        const loadLookups = async (): Promise<void> => {
            try {
                const [projects, materials, currentUser] = await Promise.all([
                    sharePointService.getProjects(),
                    sharePointService.getMaterials(),
                    sharePointService.getCurrentUser(),
                ]);
                setState((prev) => ({
                    ...prev,
                    projects,
                    materials,
                    currentUserId: currentUser?.ID,
                    formData: {
                        ...prev.formData,
                        // Auto-populate Requested By with current user
                        Requested_ById: currentUser?.ID,
                    },
                    isLoadingLookups: false,
                }));
            } catch (error) {
                console.error('Error loading lookups:', error);
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

    // Auto-generate MR Number
    React.useEffect(() => {
        if (isOpen && !state.formData.Description) {
            const generateCode = async (): Promise<void> => {
                try {
                    // Get all items and find the highest number
                    const allItems = await sharePointService.getListData('PRC_Material_Request_Register', undefined, 500, 0, undefined);
                    let nextNumber = 1;
                    if (allItems && allItems.length > 0) {
                        allItems.forEach((item) => {
                            if (item.Title) {
                                const matches = item.Title.match(/\d+$/);
                                if (matches) {
                                    const num = parseInt(matches[0], 10);
                                    if (num >= nextNumber) {
                                        nextNumber = num + 1;
                                    }
                                }
                            }
                        });
                    }
                    const newCode = `MR-${nextNumber.toString().padStart(4, '0')}`;
                    setState((prev) => ({
                        ...prev,
                        formData: { ...prev.formData, Description: newCode },
                    }));
                } catch (error) {
                    console.error('Error generating MR code:', error);
                    // Fallback to timestamp-based code
                    const timestamp = Date.now().toString().slice(-6);
                    setState((prev) => ({
                        ...prev,
                        formData: { ...prev.formData, Description: `MR-${timestamp}` },
                    }));
                }
            };
            generateCode();
        }
    }, [isOpen, sharePointService]);

    const classNames = mergeStyleSets({
        formContainer: {
            padding: '20px',
        },
        fieldGroup: {
            marginBottom: '16px',
        },
        sectionHeader: {
            fontSize: '16px',
            fontWeight: 600,
            color: theme.palette.neutralPrimary,
            marginBottom: '16px',
            marginTop: '24px',
            borderBottom: `1px solid ${theme.palette.neutralLight}`,
            paddingBottom: '8px',
        },
        buttonGroup: {
            display: 'flex',
            gap: '12px',
            marginTop: '24px',
            flexWrap: 'wrap',
        },
    });

    const handleInputChange = (fieldName: keyof IFormData, value: string | number | undefined): void => {
        setState((prev) => ({
            ...prev,
            formData: {
                ...prev.formData,
                [fieldName]: value,
            },
        }));
    };

    const validateForm = (): boolean => {
        if (!state.formData.Project_CodeId) {
            setState((prev) => ({ ...prev, error: 'Project Code is required' }));
            return false;
        }
        if (!state.formData.MaterialId) {
            setState((prev) => ({ ...prev, error: 'Material is required' }));
            return false;
        }
        if (!state.formData.Quantity.trim()) {
            setState((prev) => ({ ...prev, error: 'Quantity is required' }));
            return false;
        }
        if (!state.formData.Requested_ById) {
            setState((prev) => ({ ...prev, error: 'Requested By is required' }));
            return false;
        }
        return true;
    };

    const handleSubmit = React.useCallback(async () => {
        if (!validateForm()) {
            return;
        }

        setState((prev) => ({ ...prev, isSubmitting: true, error: undefined }));

        try {
            const itemData = {
                Title: state.formData.Description || `MR-${Date.now()}`,
                Request_No: state.formData.Description || `MR-${Date.now()}`,
                Project_CodeId: state.formData.Project_CodeId,
                MaterialId: state.formData.MaterialId,
                Quantity: parseInt(state.formData.Quantity, 10),
                UOM: state.formData.UOM,
                Request_Date: state.formData.Required_Date ? new Date(state.formData.Required_Date).toISOString() : null,
                Requested_ById: state.formData.Requested_ById,
                Status: 'Draft',
            };

            const createdItem = await sharePointService.createListItem('PRC_Material_Request_Register', itemData);
            
            // Initialize approval workflow
            let currentApprover: string | undefined;
            if (createdItem && createdItem.ID) {
              await sharePointService.initializeMaterialRequestApproval(createdItem.ID);
              
              // Fetch the updated item to get Current_Approver
              const updatedItems = await sharePointService.getListData(
                'PRC_Material_Request_Register',
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

            // Send MSGraph notifications (Email + Teams)
            if (webPartContext && currentApprover) {
              try {
                const graphClient: MSGraphClientV3 = await webPartContext.msGraphClientFactory.getClient('3');
                const notificationService = new NotificationService(graphClient);

                const projectName = state.projects.find(p => p.Id === state.formData.Project_CodeId)?.Project_Code || 'Unknown Project';
                const deepLink = `${window.location.origin}${window.location.pathname}?mr=${createdItem?.ID || ''}`;

                console.log('[MaterialRequest] Sending approval notification to:', currentApprover);

                await notificationService.sendMaterialRequestApprovalNotification({
                  requestTitle: state.formData.Description || `MR-${Date.now()}`,
                  project: projectName,
                  amount: state.formData.Estimated_Cost,
                  approverEmail: currentApprover,
                  deepLink,
                });

                console.log('[MaterialRequest] Notification sent successfully');
              } catch (notifyError) {
                console.error('[MaterialRequest] Notification failed:', notifyError);
              }
            }

            setState((prev) => ({
                ...prev,
                isSubmitting: false,
                successMessage: 'Material Request created successfully!',
                formData: {
                    Project_CodeId: undefined,
                    MaterialId: undefined,
                    Quantity: '',
                    UOM: 'PCS',
                    Description: '',
                    Required_Date: '',
                    Estimated_Cost: undefined,
                    Requested_ById: prev.currentUserId,
                    Project_ManagerId: undefined,
                },
            }));

            setTimeout(() => {
                onDismiss();
                if (onSubmitSuccess) {
                    onSubmitSuccess();
                }
            }, 1500);
        } catch (error) {
            const errorMessage = error instanceof Error ? error.message : 'Failed to create Material Request';
            setState((prev) => ({
                ...prev,
                isSubmitting: false,
                error: errorMessage,
            }));
        }
    }, [state.formData, sharePointService, onDismiss, onSubmitSuccess]);

    const handleCancel = (): void => {
        setState({
            formData: {
                Project_CodeId: undefined,
                MaterialId: undefined,
                Quantity: '',
                UOM: 'PCS',
                Description: '',
                Required_Date: '',
                Estimated_Cost: undefined,
                Requested_ById: state.currentUserId,
                Project_ManagerId: undefined,
            },
            isSubmitting: false,
            error: undefined,
            successMessage: undefined,
            projects: state.projects,
            materials: state.materials,
            isLoadingLookups: false,
            currentUserId: state.currentUserId,
        });
        onDismiss();
    };

    // Convert lookup items to dropdown options
    const projectOptions: IDropdownOption[] = state.projects.map((item) => ({
        key: item.Id,
        text: item.Project_Code || `Project ${item.Id}`,
    }));

    const materialOptions: IDropdownOption[] = state.materials.map((item) => ({
        key: item.Id,
        text: item.Material_Name || item.Title || `Material ${item.Id}`,
    }));

    const isDisabled = state.isSubmitting || state.isLoadingLookups;

    return (
        <Panel
            isOpen={isOpen}
            onDismiss={handleCancel}
            type={PanelType.medium}
            headerText="New Material Request"
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
                                    label="MR Number"
                                    value={state.formData.Description}
                                    onChange={(_, value) => handleInputChange('Description', value)}
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
                                        const projectId = option?.key as number;
                                        handleInputChange('Project_CodeId', projectId);
                                        // Auto-populate Project Manager from selected project
                                        if (projectId) {
                                            const selectedProject = state.projects.find(p => p.Id === projectId);
                                            if (selectedProject && selectedProject.Project_ManagerId) {
                                                handleInputChange('Project_ManagerId', selectedProject.Project_ManagerId);
                                            } else {
                                                handleInputChange('Project_ManagerId', undefined);
                                            }
                                        } else {
                                            handleInputChange('Project_ManagerId', undefined);
                                        }
                                    }}
                                    disabled={isDisabled}
                                    required
                                    placeholder="Select Project"
                                />
                            </div>

                            <div className={classNames.fieldGroup}>
                                <Dropdown
                                    label="Material"
                                    options={materialOptions}
                                    selectedKey={state.formData.MaterialId}
                                    onChange={(e, option) => handleInputChange('MaterialId', option?.key as number)}
                                    disabled={isDisabled}
                                    required
                                    placeholder="Select Material"
                                />
                            </div>

                            <div className={classNames.fieldGroup}>
                                <Stack horizontal tokens={{ childrenGap: 12 }}>
                                    <TextField
                                        label="Quantity"
                                        type="number"
                                        value={state.formData.Quantity}
                                        onChange={(_, value) => handleInputChange('Quantity', value)}
                                        disabled={isDisabled}
                                        required
                                        styles={{ root: { flex: 1 } }}
                                    />
                                    <Dropdown
                                        label="UOM"
                                        options={UOM_OPTIONS}
                                        selectedKey={state.formData.UOM}
                                        onChange={(e, option) => handleInputChange('UOM', option?.key as string)}
                                        disabled={isDisabled}
                                        styles={{ root: { flex: 1 } }}
                                    />
                                </Stack>
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

                            <div className={classNames.fieldGroup}>
                                <TextField
                                    label="Required Date"
                                    type="date"
                                    value={state.formData.Required_Date}
                                    onChange={(_, value) => handleInputChange('Required_Date', value)}
                                    disabled={isDisabled}
                                />
                            </div>

                            {/* Section 3: Personnel */}
                            <div className={classNames.sectionHeader}>Personnel</div>

                            <div className={classNames.fieldGroup}>
                                <PeoplePicker
                                    key="requested-by-picker"
                                    titleText="Requested By"
                                    selectedUsers={state.formData.Requested_ById ? [{ id: typeof state.formData.Requested_ById === 'string' ? parseInt(state.formData.Requested_ById, 10) : state.formData.Requested_ById, Title: '', Email: '', LoginName: '' }] : []}
                                    personSelectionLimit={1}
                                    required={true}
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
                                    key="project-manager-picker"
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

export default MaterialRequestForm;
