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
    mergeStyleSets,
    Dropdown,
    IDropdownOption,
    getTheme,
} from '@fluentui/react';
import PeoplePicker from './PeoplePicker';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointService } from '../services/SharePointService';

export interface IApprovalMatrixFormProps {
    isOpen: boolean;
    onDismiss: () => void;
    onSubmitSuccess?: () => void;
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    webPartContext?: WebPartContext;
    editMode?: boolean;
    editRecord?: {
        ID: number;
        Module: string;
        Min_Amount: number;
        Max_Amount: number;
        Stage?: number;
        Approver_Role: string;
        Approver_UserId: number;
        Approver_User?: { Title: string; ID: number };
    };
    moduleOptions?: IDropdownOption[];
    approverRoleOptions?: IDropdownOption[];
}

interface IFormData {
    Module: string;
    Min_Amount: number | undefined;
    Max_Amount: number | undefined;
    Stage: number | undefined;
    Approver_Role: string;
    Approver_UserId: number | string | undefined;
}

interface IFormState {
    formData: IFormData;
    isSubmitting: boolean;
    error: string | undefined;
    successMessage: string | undefined;
}


const ApprovalMatrixForm: React.FC<IApprovalMatrixFormProps> = ({
    isOpen,
    onDismiss,
    onSubmitSuccess,
    spHttpClient,
    pageContext,
    webPartContext,
    editMode,
    editRecord,
    moduleOptions = [],
    approverRoleOptions = [],
}) => {
    const theme = getTheme();
    const [state, setState] = React.useState<IFormState>({
        formData: {
            Module: '',
            Min_Amount: undefined,
            Max_Amount: undefined,
            Stage: undefined,
            Approver_Role: '',
            Approver_UserId: undefined,
        },
        isSubmitting: false,
        error: undefined,
        successMessage: undefined,
    });

    const [internalModuleOptions, setInternalModuleOptions] = React.useState<IDropdownOption[]>([]);
    const [internalApproverRoleOptions, setInternalApproverRoleOptions] = React.useState<IDropdownOption[]>([]);
    const [isLoadingChoices, setIsLoadingChoices] = React.useState(false);

    const sharePointService = React.useMemo(
        () => new SharePointService(spHttpClient, pageContext),
        [spHttpClient, pageContext]
    );

    React.useEffect(() => {
        if (!isOpen) return;
        const fetchChoices = async () => {
            setIsLoadingChoices(true);
            try {
                const webUrl = pageContext.web.absoluteUrl;
                const url = `${webUrl}/_api/web/lists/getByTitle('SYS_Approval_Matrix')/fields?$select=Title,Choices&$filter=TypeAsString eq 'Choice'`;
                const resp = await spHttpClient.get(url, SPHttpClient.configurations.v1);
                if (resp.ok) {
                    const data = await resp.json();
                    (data.value || []).forEach((f: any) => {
                        const opts = (f.Choices || []).map((c: string) => ({ key: c, text: c }));
                        if (f.Title === 'Module') setInternalModuleOptions(opts);
                        if (f.Title === 'Approver_Role') setInternalApproverRoleOptions(opts);
                    });
                }
            } catch (e) {
                // silent fallback to empty
            } finally {
                setIsLoadingChoices(false);
            }
        };
        fetchChoices();
    }, [isOpen, spHttpClient, pageContext]);

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

    // Initialize form data when opening in edit mode
    React.useEffect(() => {
        if (isOpen && editMode && editRecord) {
            setState({
                formData: {
                    Module: editRecord.Module || '',
                    Min_Amount: editRecord.Min_Amount,
                    Max_Amount: editRecord.Max_Amount,
                    Stage: editRecord.Stage,
                    Approver_Role: editRecord.Approver_Role || '',
                    Approver_UserId: editRecord.Approver_UserId,
                },
                isSubmitting: false,
                error: undefined,
                successMessage: undefined,
            });
        } else if (isOpen && !editMode) {
            // Reset form for new entry
            setState({
                formData: {
                    Module: '',
                    Min_Amount: undefined,
                    Max_Amount: undefined,
                    Stage: undefined,
                    Approver_Role: '',
                    Approver_UserId: undefined,
                },
                isSubmitting: false,
                error: undefined,
                successMessage: undefined,
            });
        }
    }, [isOpen, editMode, editRecord]);

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
        if (!state.formData.Module) {
            setState((prev) => ({ ...prev, error: 'Module is required' }));
            return false;
        }
        if (!state.formData.Min_Amount && state.formData.Min_Amount !== 0) {
            setState((prev) => ({ ...prev, error: 'Minimum Amount is required' }));
            return false;
        }
        if (!state.formData.Max_Amount && state.formData.Max_Amount !== 0) {
            setState((prev) => ({ ...prev, error: 'Maximum Amount is required' }));
            return false;
        }
        if (state.formData.Min_Amount && state.formData.Max_Amount && state.formData.Min_Amount > state.formData.Max_Amount) {
            setState((prev) => ({ ...prev, error: 'Minimum Amount cannot be greater than Maximum Amount' }));
            return false;
        }
        if (!state.formData.Approver_Role && !state.formData.Approver_UserId) {
            setState((prev) => ({ ...prev, error: 'Either Approver Role or Approver User is required' }));
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
                Module: state.formData.Module,
                Min_Amount: state.formData.Min_Amount || 0,
                Max_Amount: state.formData.Max_Amount || 0,
                Stage: state.formData.Stage || null,
                Approver_Role: state.formData.Approver_Role,
                Approver_UserId: state.formData.Approver_UserId || null,
            };

            if (editMode && editRecord) {
                // Update existing item
                const webUrl = pageContext.web.absoluteUrl;
                const updateUrl = `${webUrl}/_api/web/lists/getByTitle('SYS_Approval_Matrix')/items(${editRecord.ID})`;

                const response = await spHttpClient.post(
                    updateUrl,
                    SPHttpClient.configurations.v1,
                    {
                        headers: {
                            'Accept': 'application/json',
                            'Content-Type': 'application/json',
                            'X-HTTP-Method': 'MERGE',
                            'If-Match': '*'
                        },
                        body: JSON.stringify(itemData)
                    }
                );

                if (response.ok) {
                    setState((prev) => ({
                        ...prev,
                        isSubmitting: false,
                        successMessage: 'Approval Rule updated successfully!',
                    }));
                } else {
                    throw new Error('Failed to update approval rule');
                }
            } else {
                // Create new item
                await sharePointService.createListItem('SYS_Approval_Matrix', itemData);
                setState((prev) => ({
                    ...prev,
                    isSubmitting: false,
                    successMessage: 'Approval Rule created successfully!',
                    formData: {
                        Module: '',
                        Min_Amount: undefined,
                        Max_Amount: undefined,
                        Stage: undefined,
                        Approver_Role: '',
                        Approver_UserId: undefined,
                    },
                }));
            }

            setTimeout(() => {
                onDismiss();
                if (onSubmitSuccess) {
                    onSubmitSuccess();
                }
            }, 1500);
        } catch (error) {
            const errorMessage = error instanceof Error ? error.message : (editMode ? 'Failed to update Approval Rule' : 'Failed to create Approval Rule');
            setState((prev) => ({
                ...prev,
                isSubmitting: false,
                error: errorMessage,
            }));
        }
    }, [state.formData, sharePointService, onDismiss, onSubmitSuccess, editMode, editRecord, spHttpClient, pageContext]);

    const handleCancel = (): void => {
        setState({
            formData: {
                Module: '',
                Min_Amount: undefined,
                Max_Amount: undefined,
                Stage: undefined,
                Approver_Role: '',
                Approver_UserId: undefined,
            },
            isSubmitting: false,
            error: undefined,
            successMessage: undefined,
        });
        onDismiss();
    };

    const isDisabled = state.isSubmitting;

    return (
        <Panel
            isOpen={isOpen}
            onDismiss={handleCancel}
            type={PanelType.medium}
            headerText={editMode ? 'Edit Approval Rule' : 'New Approval Rule'}
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

                    <Stack tokens={{ childrenGap: 16 }}>
                        {/* Section 1: Approval Details */}
                        <div className={classNames.sectionHeader}>Approval Details</div>

                        <div className={classNames.fieldGroup}>
                            <Dropdown
                                label="Module"
                                options={internalModuleOptions}
                                selectedKey={state.formData.Module || undefined}
                                onChange={(e, option) => handleInputChange('Module', option?.key as string)}
                                disabled={isDisabled || isLoadingChoices}
                                required
                                placeholder="Select Module"
                            />
                        </div>

                        <div className={classNames.fieldGroup}>
                            <Stack horizontal tokens={{ childrenGap: 12 }}>
                                <TextField
                                    label="Minimum Amount"
                                    type="number"
                                    value={state.formData.Min_Amount?.toString() || ''}
                                    onChange={(e, value) => handleInputChange('Min_Amount', value ? Number(value) : undefined)}
                                    disabled={isDisabled}
                                    required
                                    prefix="₦"
                                    styles={{ root: { flex: 1 } }}
                                />
                                <TextField
                                    label="Maximum Amount"
                                    type="number"
                                    value={state.formData.Max_Amount?.toString() || ''}
                                    onChange={(e, value) => handleInputChange('Max_Amount', value ? Number(value) : undefined)}
                                    disabled={isDisabled}
                                    required
                                    prefix="₦"
                                    styles={{ root: { flex: 1 } }}
                                />
                            </Stack>
                        </div>

                        <div className={classNames.fieldGroup}>
                            <TextField
                                label="Stage"
                                type="number"
                                value={state.formData.Stage?.toString() || ''}
                                onChange={(e, value) => handleInputChange('Stage', value ? Number(value) : undefined)}
                                disabled={isDisabled}
                                placeholder="Stage number"
                            />
                        </div>

                        {/* Section 2: Approver */}
                        <div className={classNames.sectionHeader}>Approver</div>

                        <div className={classNames.fieldGroup}>
                            <Dropdown
                                label="Approver Role"
                                options={internalApproverRoleOptions}
                                selectedKey={state.formData.Approver_Role || undefined}
                                onChange={(e, option) => handleInputChange('Approver_Role', option?.key as string)}
                                disabled={isDisabled || isLoadingChoices}
                                placeholder="Select Role"
                            />
                        </div>

                        <div className={classNames.fieldGroup}>
                            <PeoplePicker
                                key="approver-user-picker"
                                titleText="Or Select Approver User"
                                selectedUsers={state.formData.Approver_UserId ? [{ id: state.formData.Approver_UserId, Title: '', Email: '', LoginName: '' }] : []}
                                personSelectionLimit={1}
                                disabled={isDisabled}
                                spHttpClient={spHttpClient}
                                pageContext={pageContext}
                                webPartContext={webPartContext}
                                onChange={(items) => {
                                    if (items && items.length > 0 && items[0].id) {
                                        handleInputChange('Approver_UserId', items[0].id);
                                    } else {
                                        handleInputChange('Approver_UserId', undefined);
                                    }
                                }}
                            />
                        </div>

                        {/* Buttons */}
                        <div className={classNames.buttonGroup}>
                            <PrimaryButton
                                text={editMode ? 'Update' : 'Submit'}
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
                    </Stack>
                </div>
            </div>
        </Panel>
    );
};

export default ApprovalMatrixForm;
