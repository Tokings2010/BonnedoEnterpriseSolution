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

export interface IFinancePaymentFormProps {
    isOpen: boolean;
    onDismiss: () => void;
    onSubmitSuccess?: () => void;
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    webPartContext?: WebPartContext;
}

interface IFormData {
    Project_CodeId: number | undefined;
    VendorId: number | undefined;
    PR_NumberId: number | undefined;
    PO_NumberId: number | undefined;
    GRN_ReferenceId: number | undefined;
    Amount: number | undefined;
    Requested_ById: number | undefined;
    Current_ApproverId: number | undefined;
    Notes: string;
}

interface IFormState {
    formData: IFormData;
    isSubmitting: boolean;
    error: string | undefined;
    successMessage: string | undefined;
    projects: IListItem[];
    vendors: IListItem[];
    purchaseRequisitions: IListItem[];
    purchaseOrders: IListItem[];
    isLoadingLookups: boolean;
    currentUserId: number | undefined;
}

const FinancePaymentForm: React.FC<IFinancePaymentFormProps> = ({
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
            VendorId: undefined,
            PR_NumberId: undefined,
            PO_NumberId: undefined,
            GRN_ReferenceId: undefined,
            Amount: undefined,
            Requested_ById: undefined,
            Current_ApproverId: undefined,
            Notes: '',
        },
        isSubmitting: false,
        error: undefined,
        successMessage: undefined,
        projects: [],
        vendors: [],
        purchaseRequisitions: [],
        purchaseOrders: [],
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
                const [projects, vendors, prs, pos, currentUser] = await Promise.all([
                    sharePointService.getProjects(),
                    sharePointService.getVendors(),
                    sharePointService.getPurchaseRequisitions(),
                    sharePointService.getPurchaseOrders(),
                    sharePointService.getCurrentUser(),
                ]);
                setState((prev) => ({
                    ...prev,
                    projects,
                    vendors,
                    purchaseRequisitions: prs,
                    purchaseOrders: pos,
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
        if (!state.formData.VendorId) {
            setState((prev) => ({ ...prev, error: 'Vendor is required' }));
            return false;
        }
        if (!state.formData.PO_NumberId) {
            setState((prev) => ({ ...prev, error: 'Purchase Order is required' }));
            return false;
        }
        if (!state.formData.Amount) {
            setState((prev) => ({ ...prev, error: 'Amount is required' }));
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

        // STEP 2: Validation - Cannot proceed unless GRN exists for the linked PO
        if (state.formData.PO_NumberId) {
            const hasGRN = await sharePointService.validateGRNExists(state.formData.PO_NumberId);
            if (!hasGRN) {
                setState((prev) => ({
                    ...prev,
                    error: 'Cannot create Payment Request: A Goods Received Note (GRN) must exist for the selected Purchase Order.',
                }));
                return;
            }
        }

        setState((prev) => ({ ...prev, isSubmitting: true, error: undefined }));

        try {
            const amount = state.formData.Amount || 0;

            const itemData: any = {
                Project_CodeId: state.formData.Project_CodeId,
                VendorId: state.formData.VendorId,
                PR_NumberId: state.formData.PR_NumberId || null,
                PO_NumberId: state.formData.PO_NumberId,
                GRN_ReferenceId: state.formData.GRN_ReferenceId || null,
                Amount: amount,
                Requested_ById: state.formData.Requested_ById,
                Notes: state.formData.Notes,
                Payment_Status: 'Pending',
            };

            const createdItem = await sharePointService.createListItem('FIN_Payment_Request_Register', itemData);

            // STEP 3: Initialize multi-stage approval (Manager → Finance Lead → Director)
            if (createdItem && createdItem.ID) {
                await sharePointService.initializePaymentRequestApproval(createdItem.ID, amount);
            }

            setState((prev) => ({
                ...prev,
                isSubmitting: false,
                successMessage: 'Payment Request created successfully!',
                formData: {
                    Project_CodeId: undefined,
                    VendorId: undefined,
                    PR_NumberId: undefined,
                    PO_NumberId: undefined,
                    GRN_ReferenceId: undefined,
                    Amount: undefined,
                    Requested_ById: prev.currentUserId,
                    Current_ApproverId: undefined,
                    Notes: '',
                },
            }));

            setTimeout(() => {
                onDismiss();
                if (onSubmitSuccess) {
                    onSubmitSuccess();
                }
            }, 1500);
        } catch (error) {
            const errorMessage = error instanceof Error ? error.message : 'Failed to create Payment Request';
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
                        VendorId: undefined,
                        PR_NumberId: undefined,
                        PO_NumberId: undefined,
                        GRN_ReferenceId: undefined,
                        Amount: undefined,
                        Requested_ById: state.currentUserId,
                        Current_ApproverId: undefined,
                        Notes: '',
                    },
            isSubmitting: false,
            error: undefined,
            successMessage: undefined,
            projects: state.projects,
            vendors: state.vendors,
            purchaseRequisitions: state.purchaseRequisitions,
            purchaseOrders: state.purchaseOrders,
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

    const vendorOptions: IDropdownOption[] = state.vendors.map((item) => ({
        key: item.Id,
        text: item.Vendor_Name || item.Title || `Vendor ${item.Id}`,
    }));

    const prOptions: IDropdownOption[] = state.purchaseRequisitions.map((item) => ({
        key: item.Id,
        text: item.PR_Number || item.Title || `PR ${item.Id}`,
    }));

    const poOptions: IDropdownOption[] = state.purchaseOrders.map((item) => ({
        key: item.Id,
        text: item.PO_Number || item.Title || `PO ${item.Id}`,
    }));

    const isDisabled = state.isSubmitting || state.isLoadingLookups;

    return (
        <Panel
            isOpen={isOpen}
            onDismiss={handleCancel}
            type={PanelType.medium}
            headerText="New Payment Request"
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
                            {/* Section 1: Payment Details */}
                            <div className={classNames.sectionHeader}>Payment Details</div>

                            <div className={classNames.fieldGroup}>
                                <Dropdown
                                    label="Project Code"
                                    options={projectOptions}
                                    selectedKey={state.formData.Project_CodeId}
                                    onChange={(e, option) => {
                                        handleInputChange('Project_CodeId', option?.key as number);
                                    }}
                                    disabled={isDisabled}
                                    required
                                    placeholder="Select Project"
                                />
                            </div>

                            <div className={classNames.fieldGroup}>
                                <Dropdown
                                    label="Vendor"
                                    options={vendorOptions}
                                    selectedKey={state.formData.VendorId}
                                    onChange={(e, option) => handleInputChange('VendorId', option?.key as number)}
                                    disabled={isDisabled}
                                    required
                                    placeholder="Select Vendor"
                                />
                            </div>

                            <div className={classNames.fieldGroup}>
                                <Dropdown
                                    label="Purchase Requisition (Optional)"
                                    options={prOptions}
                                    selectedKey={state.formData.PR_NumberId}
                                    onChange={(e, option) => handleInputChange('PR_NumberId', option?.key as number)}
                                    disabled={isDisabled}
                                    placeholder="Select PR"
                                />
                            </div>

                            <div className={classNames.fieldGroup}>
                                <Dropdown
                                    label="Purchase Order"
                                    options={poOptions}
                                    selectedKey={state.formData.PO_NumberId}
                                    onChange={(e, option) => handleInputChange('PO_NumberId', option?.key as number)}
                                    disabled={isDisabled}
                                    required
                                    placeholder="Select PO"
                                />
                            </div>

                            <div className={classNames.fieldGroup}>
                                <TextField
                                    label="Amount"
                                    type="number"
                                    value={state.formData.Amount?.toString() || ''}
                                    onChange={(e, value) => handleInputChange('Amount', value ? Number(value) : undefined)}
                                    disabled={isDisabled}
                                    required
                                    prefix="₦"
                                />
                            </div>

                            <div className={classNames.fieldGroup}>
                                <TextField
                                    label="Notes"
                                    value={state.formData.Notes}
                                    onChange={(e, value) => handleInputChange('Notes', value || '')}
                                    disabled={isDisabled}
                                    multiline
                                    rows={3}
                                />
                            </div>

                            {/* Section 2: Personnel */}
                            <div className={classNames.sectionHeader}>Personnel</div>

                            <div className={classNames.fieldGroup}>
                                <PeoplePicker
                                    key="requested-by-picker"
                                    titleText="Requested By"
                                    selectedUsers={state.formData.Requested_ById ? [{ id: state.formData.Requested_ById, Title: '', Email: '', LoginName: '' }] : []}
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
                                    key="current-approver-picker"
                                    titleText="Current Approver (Optional)"
                                    selectedUsers={state.formData.Current_ApproverId ? [{ id: state.formData.Current_ApproverId, Title: '', Email: '', LoginName: '' }] : []}
                                    personSelectionLimit={1}
                                    disabled={isDisabled}
                                    spHttpClient={spHttpClient}
                                    pageContext={pageContext}
                                    webPartContext={webPartContext}
                                    onChange={(items) => {
                                        if (items && items.length > 0 && items[0].id) {
                                            handleInputChange('Current_ApproverId', items[0].id);
                                        } else {
                                            handleInputChange('Current_ApproverId', undefined);
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
                        </Stack>
                    )}
                </div>
            </div>
        </Panel>
    );
};

export default FinancePaymentForm;
